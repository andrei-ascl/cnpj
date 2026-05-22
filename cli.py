#!/usr/bin/env python
"""CLI para execução do CNPJ consultation system com argumentos customizáveis."""

import argparse
import asyncio
import json
import logging
import re
import random
import time
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Literal, NamedTuple, Optional

import pandas as pd
import aiohttp

# =======================
# LOGGING
# =======================
def setup_logging(log_dir: Path) -> logging.Logger:
    """Configura logging com console (INFO) e arquivo (DEBUG)."""
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("cnpj")
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        return logger

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    fh = logging.FileHandler(log_dir / "cnpj.log", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger

# =======================
# CONFIGURAÇÕES
# =======================
@dataclass
class Config:
    """Configuração centralizada do sistema."""
    arquivo_entrada : Path  = field(default_factory=lambda: Path("data/input/pesquisa_cnpj.xlsx"))
    arquivo_saida   : Path  = field(default_factory=lambda: Path("data/output/pesquisa_cnpj_resultado.xlsx"))
    arquivo_cache   : Path  = field(default_factory=lambda: Path("data/logs/cache_cnpj.json"))
    arquivo_falhas  : Path  = field(default_factory=lambda: Path("data/logs/falhas_cnpj.csv"))
    aba_excel       : int   = 0
    base_sleep      : float = 1.0
    tentativas      : int   = 3
    timeout_http    : int   = 15
    sleep_fixo      : float = 0.6
    max_concurrent  : int   = 5
    formato_socios  : Literal["lista_nomes", "principal", "quantidade", "detalhado"] = "lista_nomes"

    def __post_init__(self) -> None:
        for pasta in [self.arquivo_entrada.parent, self.arquivo_saida.parent, self.arquivo_cache.parent]:
            pasta.mkdir(parents=True, exist_ok=True)

# =======================
# ESTRUTURA DO CACHE
# =======================
class DadosCNPJ(NamedTuple):
    """Estrutura imutável dos dados retornados pela BrasilAPI."""
    data_abertura : Optional[str]
    ano_abertura  : Optional[int]
    idade_anos    : Optional[int]
    situacao_desc : Optional[str]
    situacao_cod  : Optional[int]
    data_situacao : Optional[str]
    socios        : Optional[object]
    bairro        : Optional[str]
    setor         : Optional[str]
    porte         : Optional[str]
    ativa         : Optional[bool]

CACHE_VAZIO    = DadosCNPJ(*[None] * len(DadosCNPJ._fields))
CACHE_N_CAMPOS = len(DadosCNPJ._fields)


def cache_tem_resultado(valor) -> bool:
    """Determina se uma entrada de cache contém dados úteis."""
    return isinstance(valor, DadosCNPJ) and any(campo is not None for campo in valor)

MAPA_PORTE = {
    "MICRO EMPRESA"            : "ME",
    "EMPRESA DE PEQUENO PORTE" : "EPP",
    "DEMAIS"                   : "DEMAIS",
}

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module=r"openpyxl\.styles\.stylesheet",
)

# =======================
# UTILIDADES (extraídas da notebook)
# =======================
def cnpj_tem_digitos_validos(cnpj: str) -> bool:
    """Valida os dígitos verificadores de um CNPJ com 14 dígitos."""
    if not re.fullmatch(r"\d{14}", cnpj):
        return False
    if cnpj == cnpj[0] * 14:
        return False

    def calcula_digito(base: str, pesos: list[int]) -> str:
        soma = sum(int(d) * p for d, p in zip(base, pesos))
        resto = soma % 11
        return "0" if resto < 2 else str(11 - resto)

    digito_1 = calcula_digito(cnpj[:12], [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    digito_2 = calcula_digito(cnpj[:13], [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    return cnpj[-2:] == digito_1 + digito_2


def limpa_cnpj(valor) -> Optional[str]:
    """Remove nao-digitos, recompoe zeros a esquerda e valida o CNPJ."""
    if pd.isna(valor):
        return None

    texto = str(int(valor)) if isinstance(valor, float) and valor.is_integer() else str(valor)
    digitos = re.sub(r"\D", "", texto)
    if not digitos or len(digitos) > 14:
        return None

    cnpj = digitos.zfill(14)
    return cnpj if cnpj_tem_digitos_validos(cnpj) else None

def formata_data(data_str: str) -> Optional[str]:
    """Converte string de data para DD/MM/YYYY."""
    try:
        dt = pd.to_datetime(data_str, errors="coerce", utc=False)
        return None if pd.isna(dt) else dt.strftime("%d/%m/%Y")
    except Exception:
        return None

def processar_socios(qsa_list, formato: str = "lista_nomes") -> Optional[object]:
    """Processa a lista de sócios (QSA)."""
    if not isinstance(qsa_list, list):
        return None
    validos = [s for s in qsa_list if isinstance(s, dict) and s.get("nome_socio")]
    if not validos:
        return None

    if formato == "quantidade":
        return len(validos)
    if formato == "principal":
        return validos[0].get("nome_socio", "").strip() or None
    if formato == "lista_nomes":
        nomes = [s["nome_socio"].strip() for s in validos if s.get("nome_socio")]
        return "; ".join(nomes) or None
    if formato == "detalhado":
        partes = []
        for s in validos:
            nome = s.get("nome_socio", "").strip()
            qual = s.get("qualificacao_socio", "").strip()
            if nome:
                partes.append(f"{nome} ({qual})" if qual else nome)
        return "; ".join(partes) or None
    return None

async def backoff_sleep(base: float, tentativa: int) -> None:
    """Pausa assíncrona com backoff exponencial + jitter."""
    await asyncio.sleep(base * (2 ** tentativa) + random.uniform(0, 0.4))

# =======================
# CONSULTA API
# =======================
async def consulta_cnpj(cnpj: str, sess: aiohttp.ClientSession, cfg: Config, logger: logging.Logger) -> DadosCNPJ:
    """Consulta BrasilAPI para obter dados de um CNPJ (assincronamente)."""
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    for i in range(cfg.tentativas):
        try:
            async with sess.get(url, timeout=aiohttp.ClientTimeout(total=cfg.timeout_http)) as r:
                if r.status == 200:
                    data = await r.json() or {}

                    abertura_raw = data.get("data_inicio_atividade")
                    if abertura_raw:
                        abertura_fmt = formata_data(str(abertura_raw))
                        try:
                            ano   = int(str(abertura_raw)[:4])
                            idade = datetime.now().year - ano
                        except (ValueError, TypeError):
                            ano, idade = None, None
                    else:
                        abertura_fmt = ano = idade = None

                    situacao_desc = data.get("descricao_situacao_cadastral")
                    situacao_cod  = data.get("situacao_cadastral")
                    situacao_data = data.get("data_situacao_cadastral")

                    return DadosCNPJ(
                        data_abertura = abertura_fmt,
                        ano_abertura  = ano,
                        idade_anos    = idade,
                        situacao_desc = situacao_desc,
                        situacao_cod  = situacao_cod,
                        data_situacao = formata_data(str(situacao_data)) if situacao_data else None,
                        socios        = processar_socios(data.get("qsa", []), cfg.formato_socios),
                        bairro        = data.get("bairro") or None,
                        setor         = data.get("cnae_fiscal_descricao") or None,
                        porte         = data.get("porte") or None,
                        ativa         = str(situacao_desc or "").upper() == "ATIVA",
                    )

                if r.status in (429, 500, 502, 503, 504):
                    await backoff_sleep(cfg.base_sleep, i)
                    continue

        except (aiohttp.ClientError, asyncio.TimeoutError):
            await backoff_sleep(cfg.base_sleep, i)

    return CACHE_VAZIO

# =======================
# CACHE
# =======================
def carregar_cache(caminho: Path, logger: logging.Logger) -> dict:
    """Carrega cache JSON."""
    if not caminho.exists():
        return {}
    try:
        with caminho.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        if not isinstance(raw, dict):
            raise ValueError("cache JSON nao contem um objeto/dicionario")

        out      : dict = {}
        migrados : int  = 0
        for k, v in raw.items():
            if not isinstance(v, list) or len(v) < CACHE_N_CAMPOS:
                out[k] = None
                migrados += 1
                continue

            dados = DadosCNPJ(*v[:CACHE_N_CAMPOS])
            out[k] = dados

        if migrados:
            logger.warning(f"Cache: {migrados} entrada(s) desatualizada(s) serao reprocessadas.")
        return out
    except Exception as e:
        logger.warning(f"Falha ao carregar cache ({e}). Iniciando cache vazio.")
        return {}

def salvar_cache(caminho: Path, cache: dict, logger: logging.Logger, max_tentativas: int = 3) -> None:
    """Salva cache como JSON com retry."""
    dados = {k: list(v) for k, v in cache.items() if isinstance(v, DadosCNPJ)}

    for tentativa in range(max_tentativas):
        try:
            with caminho.open("w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False)
            return
        except (IOError, OSError) as e:
            if tentativa < max_tentativas - 1:
                tempo_espera = 0.5 * (2 ** tentativa)
                logger.warning(f"Falha ao salvar cache (tentativa {tentativa + 1}/{max_tentativas}), aguardando {tempo_espera:.1f}s...")
                time.sleep(tempo_espera)
            else:
                logger.error(f"Falha ao salvar cache apos {max_tentativas} tentativas: {e}")
                raise ValueError(f"Cache nao pode ser salvo: {e}") from e
        except Exception as e:
            logger.error(f"Erro inesperado ao salvar cache: {e}")
            raise ValueError(f"Cache nao pode ser salvo: {e}") from e

def salvar_excel_com_fallback(df: pd.DataFrame, caminho: Path) -> None:
    """Salva DataFrame em Excel com fallback."""
    try:
        df.to_excel(caminho, index=False, engine="xlsxwriter")
    except Exception:
        df.to_excel(caminho, index=False)

def verificar_colunas_obrigatorias(df: pd.DataFrame, colunas: list) -> None:
    """Verifica colunas obrigatórias."""
    faltantes = [c for c in colunas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas obrigatorias nao encontradas: {', '.join(faltantes)}")

def validar_estrutura_entrada(df: pd.DataFrame, logger: logging.Logger) -> None:
    """Loga estatísticas do arquivo de entrada."""
    logger.info("VALIDACAO DO ARQUIVO DE ENTRADA")
    logger.info(f"  Total de linhas: {len(df)}")
    logger.info(f"  CNPJs preenchidos: {df['cnpj'].notna().sum()}")
    logger.info(f"  CNPJs vazios: {df['cnpj'].isna().sum()}")
    logger.info(f"  Colunas encontradas: {len(df.columns)}")

def limpar_e_validar_cnpjs(df: pd.DataFrame, logger: logging.Logger) -> list:
    """Limpa e valida CNPJs."""
    df["cnpj_limpo"] = df["cnpj"].apply(limpa_cnpj)
    preenchidos = df["cnpj"].dropna().astype(str).str.strip()
    validos     = df["cnpj_limpo"].dropna().astype(str)
    unicos      = validos.unique().tolist()

    logger.info("LIMPEZA E VALIDACAO DE CNPJs")
    logger.info(f"  CNPJs originais: {len(df)}")
    logger.info(f"  CNPJs preenchidos: {df['cnpj'].notna().sum()}")
    logger.info(f"  CNPJs validos apos limpeza: {len(validos)}")
    logger.info(f"  CNPJs descartados (invalidos): {len(preenchidos) - len(validos)}")
    logger.info(f"  CNPJs duplicados removidos: {len(validos) - len(unicos)}")
    logger.info(f"  CNPJs validos unicos: {len(unicos)}")
    return unicos

def verificar_cache_existente(cache: dict, cnpjs_validos: list, logger: logging.Logger) -> list:
    """Separa CNPJs em cache vs a consultar."""
    em_cache    = [c for c in cnpjs_validos if cache_tem_resultado(cache.get(c))]
    a_consultar = [c for c in cnpjs_validos if not cache_tem_resultado(cache.get(c))]
    pct = len(em_cache) / len(cnpjs_validos) * 100 if cnpjs_validos else 0.0
    logger.info("STATUS DO CACHE")
    logger.info(f"  CNPJs em cache valido: {len(em_cache)} ({pct:.1f}%)")
    logger.info(f"  CNPJs a consultar: {len(a_consultar)}")
    return a_consultar

def aplicar_resultados_ao_dataframe(df: pd.DataFrame, cache: dict) -> pd.DataFrame:
    """Aplica dados de cache ao DataFrame."""
    def resolve(cnpj) -> Optional[DadosCNPJ]:
        if pd.isna(cnpj):
            return None
        r = cache.get(cnpj)
        return r if isinstance(r, DadosCNPJ) else None

    dados = df["cnpj_limpo"].apply(resolve)

    df["DataAbertura"] = dados.apply(lambda d: d.data_abertura if d else None)
    df["AnoFundacao"]  = dados.apply(lambda d: d.ano_abertura  if d else None)
    df["IdadeEmpresa"] = dados.apply(lambda d: d.idade_anos    if d else None)
    df["SituacaoDesc"] = dados.apply(lambda d: d.situacao_desc if d else None)
    df["SituacaoCod"]  = dados.apply(lambda d: d.situacao_cod  if d else None)
    df["DataSituacao"] = dados.apply(lambda d: d.data_situacao if d else None)
    df["Socios"]       = dados.apply(lambda d: d.socios        if d else None)
    df["Bairro"]       = dados.apply(lambda d: d.bairro        if d else None)
    df["Setor"]        = dados.apply(lambda d: d.setor         if d else None)
    df["Porte"]        = dados.apply(lambda d: d.porte         if d else None)
    df["Ativa"]        = dados.apply(lambda d: d.ativa         if d else None)

    df["Porte"] = df["Porte"].apply(
        lambda v: MAPA_PORTE.get(str(v).upper().strip(), v) if pd.notna(v) else v
    )
    for col in ["AnoFundacao", "IdadeEmpresa", "SituacaoCod"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    df["cnpj"] = df["cnpj"].astype(str)
    df["cnpj_limpo"] = df["cnpj_limpo"].astype(str)

    return df

def gerar_relatorio_final(df: pd.DataFrame, cnpjs_validos: list, falhas: list, logger: logging.Logger) -> None:
    """Loga relatório final."""
    total      = len(df)
    qtd_ativas = df["Ativa"].eq(True).sum()

    logger.info("RELATORIO FINAL DE PROCESSAMENTO")
    logger.info(f"  Total de registros: {total}")
    logger.info(f"  CNPJs unicos: {len(cnpjs_validos)}")
    logger.info(f"  Com data de abertura: {df['DataAbertura'].notna().sum()}")
    logger.info(f"  Com situacao cadastral: {df['SituacaoDesc'].notna().sum()}")
    if total:
        logger.info(f"  Empresas ATIVAS: {qtd_ativas} ({qtd_ativas / total * 100:.1f}%)")
    logger.info(f"  Com socios: {df['Socios'].notna().sum()}")
    logger.info(f"  Com bairro: {df['Bairro'].notna().sum()}")
    logger.info(f"  Com setor (CNAE): {df['Setor'].notna().sum()}")
    logger.info(f"  Com porte: {df['Porte'].notna().sum()}")

    logger.info("DISTRIBUICAO - Situacao Cadastral")
    for sit, cnt in df["SituacaoDesc"].value_counts().items():
        logger.info(f"  {sit}: {cnt} ({cnt / total * 100:.1f}%)")

    logger.info("DISTRIBUICAO - Porte")
    for p, cnt in df["Porte"].value_counts().items():
        logger.info(f"  {p}: {cnt} ({cnt / total * 100:.1f}%)")

    n_falhas = len(falhas)
    if falhas:
        pct = n_falhas / len(cnpjs_validos) * 100 if cnpjs_validos else 0.0
        logger.warning(f"CNPJs sem retorno: {n_falhas} ({pct:.1f}%)")
    else:
        logger.info("Sem falhas detectadas!")

def gerar_excel_resumido(df: pd.DataFrame, caminho: Path, logger: logging.Logger) -> None:
    """Gera arquivo Excel resumido."""
    COLUNAS = [
        "dataAtendimento", "razaoSocial", "nomeFantasia", "cnpj",
        "Bairro", "Porte", "AnoFundacao", "IdadeEmpresa", "SituacaoDesc", "Socios",
    ]
    existentes = [c for c in COLUNAS if c in df.columns]
    faltantes  = [c for c in COLUNAS if c not in df.columns]

    if faltantes:
        logger.warning(f"Colunas ausentes (serao ignoradas): {', '.join(faltantes)}")

    df_resumido = df[existentes].copy()
    df_resumido["cnpj"] = df_resumido["cnpj"].astype(str)
    n = len(df_resumido)

    logger.info("Arquivo resumido - preenchimento por coluna")
    for col in existentes:
        preench = df_resumido[col].notna().sum()
        pct = preench / n * 100 if n else 0.0
        logger.info(f"  {col:<20}: {preench}/{n} ({pct:.1f}%)")

    try:
        df_resumido.to_excel(caminho, index=False, engine="openpyxl")
        logger.info(f"Arquivo resumido salvo: {caminho}")
    except Exception as e:
        fallback = caminho.with_suffix(".csv")
        df_resumido.to_csv(fallback, index=False, encoding="utf-8-sig")
        logger.warning(f"Falha ao salvar Excel: {e}")
        logger.info(f"Salvo como CSV: {fallback}")

async def worker_consulta(cnpj: str, sess: aiohttp.ClientSession, cfg: Config, semaphore: asyncio.Semaphore, cache: dict, results: list, lock: asyncio.Lock, logger: logging.Logger) -> None:
    """Worker que consulta um CNPJ assincronamente."""
    async with semaphore:
        try:
            dados = await consulta_cnpj(cnpj, sess, cfg, logger)
            async with lock:
                cache[cnpj] = dados
                results.append(cnpj)
        except Exception as e:
            logger.debug(f"Erro ao consultar {cnpj}: {e}")
            async with lock:
                cache[cnpj] = CACHE_VAZIO
                results.append(cnpj)

        await asyncio.sleep(cfg.sleep_fixo)

async def main(cfg: Config, logger: logging.Logger) -> None:
    """Pipeline principal de consulta de CNPJs."""
    try:
        logger.info(f"Lendo arquivo: {cfg.arquivo_entrada}")
        df = pd.read_excel(cfg.arquivo_entrada, sheet_name=cfg.aba_excel)

        verificar_colunas_obrigatorias(df, ["cnpj"])
        validar_estrutura_entrada(df, logger)

        cnpjs_validos = limpar_e_validar_cnpjs(df, logger)
        if not cnpjs_validos:
            raise ValueError("Nenhum CNPJ valido encontrado apos limpeza e validacao!")

        cache           = carregar_cache(cfg.arquivo_cache, logger)
        cnpjs_pendentes = verificar_cache_existente(cache, cnpjs_validos, logger)

        if cnpjs_pendentes:
            total = len(cnpjs_pendentes)
            logger.info(f"Consultando {total} CNPJs na BrasilAPI (ate {cfg.max_concurrent} paralelos)...")
            logger.info(f"Formato de socios: {cfg.formato_socios}")

            connector = aiohttp.TCPConnector(limit_per_host=cfg.max_concurrent)
            timeout = aiohttp.ClientTimeout(total=cfg.timeout_http * cfg.tentativas * 3)

            async with aiohttp.ClientSession(connector=connector, timeout=timeout) as sess:
                sess.headers.update({
                    "User-Agent": "Mozilla/5.0 (compatible; CNPJBatchBot/1.0; +https://brasilapi.com.br)"
                })

                semaphore = asyncio.Semaphore(cfg.max_concurrent)
                lock = asyncio.Lock()
                results = []

                tasks = [
                    worker_consulta(cnpj, sess, cfg, semaphore, cache, results, lock, logger)
                    for cnpj in cnpjs_pendentes
                ]

                for i, task in enumerate(asyncio.as_completed(tasks), start=1):
                    await task
                    if i % 25 == 0 or i == total:
                        salvar_cache(cfg.arquivo_cache, cache, logger)
                        logger.info(f"Progresso: {i}/{total} CNPJs processados, cache salvo")

                logger.info(f"Concluido: {len(results)}/{total} CNPJs consultados")
        else:
            logger.info("Todos os CNPJs ja estao em cache valido!")

        logger.info("Aplicando resultados ao DataFrame...")
        df = aplicar_resultados_ao_dataframe(df, cache)

        if "dataAtendimento" in df.columns:
            logger.info("Ordenando por data de atendimento...")
            df["_dt_sort"] = pd.to_datetime(df["dataAtendimento"], errors="coerce")
            df = df.sort_values("_dt_sort", ascending=False, na_position="last")
            df = df.drop("_dt_sort", axis=1).reset_index(drop=True)

        falhas = [c for c in cnpjs_validos if not cache_tem_resultado(cache.get(c))]

        if falhas:
            pd.DataFrame({"cnpj_limpo": falhas}).to_csv(
                cfg.arquivo_falhas, index=False, encoding="utf-8-sig"
            )
            logger.warning(f"{len(falhas)} CNPJs sem retorno. Log: {cfg.arquivo_falhas}")

        salvar_cache(cfg.arquivo_cache, cache, logger)
        salvar_excel_com_fallback(df, cfg.arquivo_saida)
        logger.info(f"Arquivo completo salvo: {cfg.arquivo_saida}")

        gerar_relatorio_final(df, cnpjs_validos, falhas, logger)

        logger.info("GERANDO ARQUIVO RESUMIDO")
        arquivo_resumido = cfg.arquivo_saida.with_name(
            cfg.arquivo_saida.stem + "_RESUMIDO" + cfg.arquivo_saida.suffix
        )
        gerar_excel_resumido(df, arquivo_resumido, logger)

        logger.info("PROCESSAMENTO CONCLUIDO COM SUCESSO!")
        logger.info(f"Arquivos gerados:")
        logger.info(f"  1. Completo: {cfg.arquivo_saida}")
        logger.info(f"  2. Resumido: {arquivo_resumido}")
        if falhas:
            logger.info(f"  3. Falhas: {cfg.arquivo_falhas}")

    except FileNotFoundError as e:
        logger.error(f"Arquivo nao encontrado: {e}")
    except ValueError as e:
        logger.error(f"Erro de validacao: {e}")
    except Exception as e:
        logger.exception(f"Erro inesperado: {type(e).__name__}: {e}")

# =======================
# CLI
# =======================
def main_cli():
    """Ponto de entrada da CLI."""
    parser = argparse.ArgumentParser(
        description="CNPJ Consultation System - Consulta dados de empresas via CNPJ",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python cli.py                                          # Usa config padrão
  python cli.py --input dados.xlsx --concurrent 10      # 10 paralelos
  python cli.py --format detalhado --ttl 15             # Sócios detalhados, cache 15 dias
  python cli.py --input entrada.xlsx --output saida.xlsx # Custom I/O paths
        """,
    )

    parser.add_argument(
        "--input", "-i",
        type=Path,
        default=Path("data/input/pesquisa_cnpj.xlsx"),
        help="Arquivo Excel de entrada com CNPJs (default: data/input/pesquisa_cnpj.xlsx)"
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=Path("data/output/pesquisa_cnpj_resultado.xlsx"),
        help="Arquivo Excel de saída (default: data/output/pesquisa_cnpj_resultado.xlsx)"
    )
    parser.add_argument(
        "--concurrent", "-c",
        type=int,
        default=5,
        help="Máximo de requisições paralelas (default: 5, max: 10)"
    )
    parser.add_argument(
        "--format", "-f",
        choices=["lista_nomes", "principal", "quantidade", "detalhado"],
        default="lista_nomes",
        help="Formato de retorno dos sócios (default: lista_nomes)"
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=15,
        help="Timeout de conexão HTTP em segundos (default: 15)"
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.6,
        help="Sleep entre requisições em segundos (default: 0.6)"
    )

    args = parser.parse_args()

    # Validar argumentos
    if not args.input.exists():
        print(f"ERRO: Arquivo de entrada não encontrado: {args.input}")
        return 1

    if args.concurrent < 1 or args.concurrent > 10:
        print(f"ERRO: --concurrent deve estar entre 1 e 10")
        return 1

    # Criar config
    cfg = Config(
        arquivo_entrada=args.input,
        arquivo_saida=args.output,
        max_concurrent=args.concurrent,
        formato_socios=args.format,
        timeout_http=args.timeout,
        sleep_fixo=args.sleep,
    )

    # Setup logging
    log_dir = cfg.arquivo_cache.parent
    logger = setup_logging(log_dir)

    logger.info(f"Config: input={cfg.arquivo_entrada}, concurrent={cfg.max_concurrent}, format={cfg.formato_socios}")

    # Executar
    try:
        asyncio.run(main(cfg, logger))
        return 0
    except Exception as e:
        logger.exception(f"Erro fatal: {e}")
        return 1

if __name__ == "__main__":
    exit(main_cli())
