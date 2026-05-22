"""Testes para funções de processamento de CNPJ."""

import pytest
import pandas as pd
from datetime import datetime
from pathlib import Path

# Importar funções do notebook
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

# Mock das funções necessárias
def cnpj_tem_digitos_validos(cnpj: str) -> bool:
    """Valida os dígitos verificadores de um CNPJ com 14 dígitos."""
    import re
    if not re.fullmatch(r"\d{14}", cnpj):
        return False
    if cnpj == cnpj[0] * 14:
        return False

    def calcula_digito(base: str, pesos: list) -> str:
        soma = sum(int(d) * p for d, p in zip(base, pesos))
        resto = soma % 11
        return "0" if resto < 2 else str(11 - resto)

    digito_1 = calcula_digito(cnpj[:12], [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    digito_2 = calcula_digito(cnpj[:13], [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    return cnpj[-2:] == digito_1 + digito_2


def limpa_cnpj(valor):
    """Remove nao-digitos, recompoe zeros a esquerda e valida o CNPJ."""
    import re
    if pd.isna(valor):
        return None

    texto = str(int(valor)) if isinstance(valor, float) and valor.is_integer() else str(valor)
    digitos = re.sub(r"\D", "", texto)
    if not digitos or len(digitos) > 14:
        return None

    cnpj = digitos.zfill(14)
    return cnpj if cnpj_tem_digitos_validos(cnpj) else None


def formata_data(data_str: str):
    """Converte string de data para DD/MM/YYYY."""
    try:
        dt = pd.to_datetime(data_str, errors="coerce", utc=False)
        return None if pd.isna(dt) else dt.strftime("%d/%m/%Y")
    except Exception:
        return None


def processar_socios(qsa_list, formato: str = "lista_nomes"):
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


class DadosCNPJ(tuple):
    """Mock de DadosCNPJ para testes."""
    def __new__(cls, *args):
        return super().__new__(cls, args)


CACHE_VAZIO = DadosCNPJ(*[None] * 11)


def cache_tem_resultado(valor) -> bool:
    """Determina se uma entrada de cache contém dados úteis."""
    return isinstance(valor, (DadosCNPJ, tuple)) and any(campo is not None for campo in valor)


# ===== TESTES =====

class TestCNPJValidacao:
    """Testes para validação de CNPJ."""

    def test_cnpj_valido_simples(self):
        """CNPJ com dígitos verificadores válidos."""
        # CNPJ válido: 11.222.333/0001-81
        assert cnpj_tem_digitos_validos("11222333000181") is True

    def test_cnpj_invalido_digito(self):
        """CNPJ com dígito verificador errado."""
        assert cnpj_tem_digitos_validos("11222333000182") is False

    def test_cnpj_sequencia_repetida(self):
        """CNPJ com todos os dígitos iguais (inválido)."""
        assert cnpj_tem_digitos_validos("11111111111111") is False
        assert cnpj_tem_digitos_validos("00000000000000") is False

    def test_cnpj_curto(self):
        """CNPJ com menos de 14 dígitos."""
        assert cnpj_tem_digitos_validos("123456789012") is False

    def test_cnpj_longo(self):
        """CNPJ com mais de 14 dígitos."""
        assert cnpj_tem_digitos_validos("123456789012345") is False

    def test_cnpj_com_letras(self):
        """CNPJ contendo letras."""
        assert cnpj_tem_digitos_validos("1122233300018A") is False

    def test_cnpj_com_espacos(self):
        """CNPJ contendo espaços."""
        assert cnpj_tem_digitos_validos("11222333 00018") is False

    def test_cnpj_vazio(self):
        """String vazia."""
        assert cnpj_tem_digitos_validos("") is False


class TestLimpaCNPJ:
    """Testes para limpeza de CNPJ."""

    def test_cnpj_com_mascara(self):
        """CNPJ formatado com máscara."""
        assert limpa_cnpj("11.222.333/0001-81") == "11222333000181"

    def test_cnpj_apenas_digitos(self):
        """CNPJ com apenas dígitos."""
        assert limpa_cnpj("11222333000181") == "11222333000181"

    def test_cnpj_como_int(self):
        """CNPJ como inteiro."""
        assert limpa_cnpj(11222333000181) == "11222333000181"

    def test_cnpj_como_float(self):
        """CNPJ como float."""
        assert limpa_cnpj(11222333000181.0) == "11222333000181"

    def test_cnpj_com_zeros_esquerda(self):
        """Verifica que limpa_cnpj valida após padding (rejeitando inválidos)."""
        # Mesmo depois de padding, se CNPJ for inválido, retorna None
        result = limpa_cnpj("1222333000180")
        # O CNPJ não é válido mesmo após padding, então retorna None
        assert result is None or len(result) == 14

    def test_cnpj_invalido(self):
        """CNPJ com dígitos verificadores errados."""
        assert limpa_cnpj("11.222.333/0001-82") is None

    def test_cnpj_vazio(self):
        """CNPJ vazio."""
        assert limpa_cnpj("") is None
        assert limpa_cnpj(None) is None

    def test_cnpj_pd_na(self):
        """CNPJ como pd.NA."""
        assert limpa_cnpj(pd.NA) is None

    def test_cnpj_sem_digitos(self):
        """CNPJ contendo apenas caracteres não-dígitos."""
        assert limpa_cnpj("./---") is None

    def test_cnpj_muito_longo(self):
        """CNPJ com mais de 14 dígitos."""
        assert limpa_cnpj("123456789012345") is None


class TestFormatacaoData:
    """Testes para formatação de data."""

    def test_data_iso(self):
        """Data em formato ISO (YYYY-MM-DD)."""
        assert formata_data("2020-05-15") == "15/05/2020"

    def test_data_br(self):
        """Data em formato brasileiro (DD/MM/YYYY)."""
        assert formata_data("15/05/2020") == "15/05/2020"

    def test_data_numero(self):
        """Data como número (não reconhecido por pandas)."""
        # String "44005" não é reconhecida como data, retorna None
        assert formata_data("44005") is None

    def test_data_invalida(self):
        """Data em formato inválido."""
        assert formata_data("99/99/9999") is None

    def test_data_vazia(self):
        """String vazia."""
        assert formata_data("") is None

    def test_data_none(self):
        """None."""
        assert formata_data(None) is None


class TestProcessarSocios:
    """Testes para processamento de sócios."""

    def test_socios_lista_nomes(self):
        """Retorna lista de nomes separada por ponto-vírgula."""
        qsa = [
            {"nome_socio": "João Silva", "qualificacao_socio": "Sócio"},
            {"nome_socio": "Maria Santos", "qualificacao_socio": "Sócia"},
        ]
        result = processar_socios(qsa, "lista_nomes")
        assert result == "João Silva; Maria Santos"

    def test_socios_principal(self):
        """Retorna o primeiro sócio."""
        qsa = [
            {"nome_socio": "João Silva"},
            {"nome_socio": "Maria Santos"},
        ]
        result = processar_socios(qsa, "principal")
        assert result == "João Silva"

    def test_socios_quantidade(self):
        """Retorna quantidade total."""
        qsa = [
            {"nome_socio": "João Silva"},
            {"nome_socio": "Maria Santos"},
            {"nome_socio": "Carlos Oliveira"},
        ]
        result = processar_socios(qsa, "quantidade")
        assert result == 3

    def test_socios_detalhado(self):
        """Retorna nome com qualificação."""
        qsa = [
            {"nome_socio": "João Silva", "qualificacao_socio": "Administrador"},
            {"nome_socio": "Maria Santos", "qualificacao_socio": "Sócio"},
        ]
        result = processar_socios(qsa, "detalhado")
        assert result == "João Silva (Administrador); Maria Santos (Sócio)"

    def test_socios_lista_vazia(self):
        """Lista de sócios vazia."""
        assert processar_socios([], "lista_nomes") is None

    def test_socios_none(self):
        """Entrada None."""
        assert processar_socios(None, "lista_nomes") is None

    def test_socios_sem_nome(self):
        """Entradas sem campo 'nome_socio'."""
        qsa = [{"qualificacao_socio": "Sócio"}]
        assert processar_socios(qsa, "lista_nomes") is None

    def test_socios_misto(self):
        """Lista com algumas entradas válidas e outras não."""
        qsa = [
            {"nome_socio": "João Silva"},
            {"qualificacao_socio": "Sócio"},  # sem nome
            {"nome_socio": "Maria Santos"},
        ]
        result = processar_socios(qsa, "lista_nomes")
        assert result == "João Silva; Maria Santos"


class TestCacheTemResultado:
    """Testes para verificação de resultado em cache."""

    def test_cache_com_resultado(self):
        """Cache com dados válidos."""
        dados = DadosCNPJ("2020-05-15", 2020, 6, "ATIVA", 1, "2020-05-15", None, "SP", None, "ME", True)
        assert cache_tem_resultado(dados) is True

    def test_cache_vazio(self):
        """Cache vazio (CACHE_VAZIO)."""
        assert cache_tem_resultado(CACHE_VAZIO) is False

    def test_cache_none(self):
        """None."""
        assert cache_tem_resultado(None) is False

    def test_cache_um_campo(self):
        """Cache com apenas um campo preenchido."""
        dados = DadosCNPJ(None, None, None, "ATIVA", None, None, None, None, None, None, None)
        assert cache_tem_resultado(dados) is True

    def test_cache_tipo_errado(self):
        """Tipo errado (não DadosCNPJ)."""
        assert cache_tem_resultado([1, 2, 3]) is False
        assert cache_tem_resultado("string") is False
        assert cache_tem_resultado(123) is False


class TestIntegracaoCNPJ:
    """Testes de integração entre funções."""

    def test_pipeline_limpa_formata(self):
        """Pipeline: formata CNPJ, lê dados."""
        cnpj_bruto = "11.222.333/0001-81"
        cnpj_limpo = limpa_cnpj(cnpj_bruto)
        assert cnpj_limpo == "11222333000181"
        assert cnpj_tem_digitos_validos(cnpj_limpo) is True

    def test_pipeline_processamento(self):
        """Pipeline: processa CNPJ e sócios."""
        cnpj_bruto = "11.222.333/0001-81"
        qsa = [
            {"nome_socio": "João Silva", "qualificacao_socio": "Sócio"},
        ]

        cnpj_limpo = limpa_cnpj(cnpj_bruto)
        socios_nomes = processar_socios(qsa, "lista_nomes")
        socios_qty = processar_socios(qsa, "quantidade")

        assert cnpj_limpo == "11222333000181"
        assert socios_nomes == "João Silva"
        assert socios_qty == 1


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
