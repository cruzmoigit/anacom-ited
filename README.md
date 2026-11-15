# Pesquisa de Termos de Responsabilidade ITED (ANACOM)

Aplicação web (Python + HTML/JS) que permite pesquisar termos de responsabilidade ITED (Instalações de Telecomunicações em Edifícios) no site da ANACOM, extrair os PDFs dos termos encontrados e filtrar/exportar aqueles com 3 ou mais fogos.

## Funcionalidades

- **Pesquisa** por Distrito e Concelho (com opção "Todos os concelhos" do distrito) e período de emissão (1 a 4 semanas anteriores à data atual).
- Fixo: Tipo de termo = **Execução** | Tipo de edifício = **Residencial**.
- Paginação automática (todas as páginas de resultados).
- **Extração de PDFs**: descarrega cada "Termo de responsabilidade" e extrai:
  - Identificação do termo e data de emissão
  - N.º de fogos
  - Morada do edifício (Rua, Localidade, Código Postal)
  - Coordenadas geográficas
  - Identificação do dono de obra
  - Data do certificado (fallback: data de emissão)
- **Filtro**: mantém apenas termos com **≥ 3 fogos**.
- **Tabela** visualizável + **exportação para Excel** (`.xlsx`).

## Requisitos

- Python 3.10+
- Dependências: `requests`, `pypdf`, `openpyxl` (instaladas via `pip`).

## Instalação

```bash
pip install requests pypdf openpyxl
```

## Execução

```bash
cd anacom-ited
python server.py
```

Abrir no browser: `http://127.0.0.1:8765`

## Estrutura

```
anacom-ited/
├── server.py          # backend (http.server + requests + pypdf + openpyxl)
├── distritos.json     # mapeamento Distrito -> Concelhos (extraído do site ANACOM)
└── static/
    └── index.html     # frontend
```

## Endpoints

- `GET /` — frontend
- `GET /api/distritos` — lista de distritos
- `GET /api/concelhos?distrito=...` — concelhos do distrito (inclui opção `-` = todos)
- `POST /api/pesquisar` — `{distrito, concelho, semanas}` → resultados paginados
- `POST /api/extrair` — `{items}` → extração de PDFs (≥3 fogos)
- `GET /api/exportar.xlsx` — download do Excel

## Notas

- O site ANACOM requer sessão (cookies) e `Referer`; o backend trata disso automaticamente.
- A codificação do formulário é ISO-8859-15.
- Os PDFs são cacheados em `%TEMP%/anacom_ited_cache`.
- A paginação do ANACOM usa o parâmetro `pagina=N` sobre a URL de resultados.

## Licença

Uso pessoal/educacional.
