# CNPJ Consultation System

Sistema automatizado para consultar dados de empresas brasileiras via CNPJ usando a BrasilAPI.

## Características

- ✅ **Consultas assíncronas**: até 5 requisições paralelas (5x mais rápido)
- ✅ **Cache persistente**: evita re-consultas desnecessárias
- ✅ **Logging estruturado**: arquivo de log com timestamps e níveis
- ✅ **Retry automático**: backoff exponencial em caso de falha
- ✅ **Validação CNPJ**: verificação de dígitos verificadores
- ✅ **Excel com formato correto**: CNPJ armazenado como texto
- ✅ **Memoização de falhas**: não re-consulta CNPJs que falharam

## Requisitos

```bash
pip install pandas aiohttp openpyxl xlsxwriter
```

## Uso

### 1. Preparar arquivo de entrada

Crie um arquivo Excel com uma coluna `cnpj`:

```
data/input/pesquisa_cnpj.xlsx
├─ cnpj (obrigatório)
├─ dataAtendimento (opcional)
├─ razaoSocial (opcional)
└─ nomeFantasia (opcional)
```

### 2. Executar o script

**Jupyter Notebook:**
```python
# Abrir: consulta-cnpj-lista-porte_revisado.ipynb
# Executar todas as células
```

**Python direto:**
```bash
python -c "
import asyncio
from pathlib import Path
import sys

# Adicionar caminho do projeto
sys.path.insert(0, str(Path.cwd()))

# Importar e executar
exec(open('consulta-cnpj-lista-porte_revisado.ipynb').read().replace('null', 'None'))
asyncio.run(main())
"
```

Ou extrair a célula do notebook e criar um script `.py`.

### 3. Saídas geradas

```
data/output/
├─ pesquisa_cnpj_resultado.xlsx          # Arquivo completo com todos os campos
├─ pesquisa_cnpj_resultado_RESUMIDO.xlsx # Resumido (colunas selecionadas)

data/logs/
├─ cnpj.log              # Log com timestamps e níveis
├─ cache_cnpj.json       # Cache de consultas
└─ falhas_cnpj.csv       # CNPJs que não retornaram dados
```

## Configuração

Edite `Config` no notebook para customizar:

```python
@dataclass
class Config:
    arquivo_entrada   = Path("data/input/pesquisa_cnpj.xlsx")
    arquivo_saida     = Path("data/output/pesquisa_cnpj_resultado.xlsx")
    aba_excel         = 0              # Aba do Excel (0-indexed)
    
    # API
    base_sleep        = 1.0            # Sleep base para backoff (segundos)
    tentativas        = 3              # Tentativas por CNPJ
    timeout_http      = 15             # Timeout de conexão (segundos)
    sleep_fixo        = 0.6            # Sleep entre requisições (segundos)
    max_concurrent    = 5              # Máximo de requisições paralelas
    
    # Processamento
    formato_socios    = "lista_nomes"  # "lista_nomes" | "principal" | "quantidade" | "detalhado"
```

### Parâmetros de Performance

- **`max_concurrent`**: aumentar para ~10 se a API permitir (respeitar rate limits)
- **`sleep_fixo`**: diminuir para ~0.3 se a API permitir, aumentar se tomar rate limit (429)
- **`tentativas`**: aumentar para ~5 se houver muitas falhas transitórias

## Dados Retornados

Para cada CNPJ, o sistema consulta:

```
data_abertura        (str)   - Data de abertura formatada (DD/MM/YYYY)
ano_abertura         (int)   - Ano de fundação
idade_anos           (int)   - Idade da empresa em anos
situacao_desc        (str)   - Status da empresa (ATIVA, INAPTA, BAIXADA, etc)
situacao_cod         (int)   - Código do status
data_situacao        (str)   - Data da situação (DD/MM/YYYY)
socios               (str)   - Sócios (formato configurável)
bairro               (str)   - Bairro da sede
setor                (str)   - Setor/CNAE
porte                (str)   - Porte (ME, EPP, DEMAIS)
ativa                (bool)  - Booleano: é ativa?
```

## Logging

O sistema gera logs em dois níveis:

**Console** (INFO e acima):
```
2026-05-21 21:45:47 [INFO] Lendo arquivo: data/input/pesquisa_cnpj.xlsx
2026-05-21 21:45:48 [INFO] Consultando 50 CNPJs na BrasilAPI (ate 5 paralelos)...
2026-05-21 21:45:53 [WARNING] 1 CNPJ sem retorno. Log: data/logs/falhas_cnpj.csv
```

**Arquivo** (`data/logs/cnpj.log`, DEBUG e acima):
```
2026-05-21 21:45:47 [DEBUG] Worker iniciado para CNPJ 12345678901234
2026-05-21 21:45:47 [DEBUG] Tentativa 1/3 para CNPJ 12345678901234: status 200
```

Níveis:
- `DEBUG`: detalhes granulares (tentativas de API, retentativas)
- `INFO`: progresso geral (arquivos lidos, processamento concluído)
- `WARNING`: situações recuperáveis (cache expirado, save retry)
- `ERROR`: falhas críticas (arquivo não encontrado, cache corrompido)

## Tratamento de Erros

### Cache Corrompido
Se `cache_cnpj.json` estiver malformado, o sistema:
1. Log `WARNING: Falha ao carregar cache`
2. Inicia com cache vazio
3. Re-consulta todos os CNPJs

### Falha de API
Para cada CNPJ que falha:
1. Tenta 3 vezes com backoff exponencial (1s → 2s → 4s + jitter)
2. Se todas falharem, armazena `CACHE_VAZIO` (memoiza failure)
3. Próxima execução não tenta novamente (evita loop infinito)

### Excel não consegue salvar
Se openpyxl falhar, tenta fallback com xlsxwriter. Se ambos falharem, salva como CSV.

## Performance

### Cenários

**100 CNPJs novos, 5 paralelos, 0.6s sleep:**
- Sequencial (requests): ~65 segundos
- Assíncrono (aiohttp): ~13 segundos
- **Ganho: 5x mais rápido**

**Fila com cache válido:**
- Tempo: ~2 segundos (só processamento de dados)

### Otimizações

1. **Cache persistente**: não re-consulta CNPJs já processados
2. **Requisições paralelas**: até 5 por padrão (configurável)
3. **Async/await**: não bloqueia em I/O
4. **Semáforo**: limita concorrência (respeita rate limits)

## Exemplos de Uso

### Consultar novos CNPJs

```python
# Adicionar novos CNPJs ao arquivo de entrada
# Executar o script
# Saída: resultado.xlsx com dados novos + cache anterior
```

### Forçar re-consulta

```python
# Opção 1: deletar data/logs/cache_cnpj.json
# Opção 2: editar cache_cnpj.json e remover CNPJs específicos
# Opção 3: implementar TTL no cache (futura melhoria)
```

### Filtrar por status

```python
# No notebook, após executar:
df_ativas = df[df['Ativa'] == True]
df_ativas.to_excel('resultado_apenas_ativas.xlsx', index=False)
```

## Estrutura do Projeto

```
.
├── consulta-cnpj-lista-porte_revisado.ipynb  # Notebook principal
├── README.md                                  # Este arquivo
├── data/
│   ├── input/
│   │   └── pesquisa_cnpj.xlsx                # Arquivo de entrada
│   ├── output/
│   │   ├── pesquisa_cnpj_resultado.xlsx
│   │   └── pesquisa_cnpj_resultado_RESUMIDO.xlsx
│   └── logs/
│       ├── cnpj.log
│       ├── cache_cnpj.json
│       └── falhas_cnpj.csv
└── .git/                                     # Git repository
```

## Troubleshooting

| Problema | Solução |
|----------|---------|
| `ModuleNotFoundError: aiohttp` | `pip install aiohttp` |
| `No module named 'pandas'` | `pip install pandas` |
| "CNPJ inválido" | Verificar dígitos verificadores (algoritmo modulo-11) |
| Status 429 (rate limit) | Aumentar `sleep_fixo` ou diminuir `max_concurrent` |
| Log muito grande | Rotacionar `data/logs/cnpj.log` manualmente ou implementar `RotatingFileHandler` |
| Excel abre com aviso | Ignorar - openpyxl gera aviso de estilo padrão (inofensivo) |

## Melhorias Futuras

- [ ] **Cache com TTL**: expirar dados após X dias
- [ ] **Testes unitários**: cobertura de `limpa_cnpj()`, `processar_socios()`
- [ ] **Validação de schema**: verificar estrutura de resposta da API
- [ ] **Dashboard**: gerar relatório HTML com métricas
- [ ] **CLI**: interface de linha de comando para customizar parâmetros
- [ ] **Notificações**: alertas por email/Slack em caso de falha

## Autores

Desenvolvido com [Claude Code](https://claude.com/claude-code)

## Licença

MIT
