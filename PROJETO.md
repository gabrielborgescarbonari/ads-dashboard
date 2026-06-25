# Dashboard de Mídia — Google + Meta

Projeto unificado (junho/2026): juntou o pipeline de API (antes em `ads-dashboard`),
o repositório de publicação (antes em `dash-deploy`), o gerador de HTML (antes solto
no `Downloads`) e a documentação da integração. A parte Streamlit (`app.py`) foi
descartada — não era mais usada.

## Estrutura da pasta
```
src/                    clientes de API (google_ads_client, meta_ads, config, ssl_patch)
scripts/                pipeline + scripts auxiliares de relatório
  dash_sync.py          puxa Google+Meta da API e monta o cache data/dash_data.xlsx
  dash_fetch.py         coleta bruta das APIs
  dash_derive.py        deriva Público/Formato do nome da campanha
data/                   caches (dash_data.xlsx, raw_*.csv) e a planilha histórica GROWTH(2).xlsx
deploy/                 repositório git que publica index.html no GitHub Pages (.../dash)
gerar_dashboard_midia.py  lê o cache e gera o index.html
atualizar_dashboard_api.bat  orquestra tudo (usa %~dp0, caminhos relativos ao .bat)
.env, service_account.json   credenciais (fora do git)
```

## Automação
- Tarefa Windows `Dashboard-Midia-Sync` roda `atualizar_dashboard_api.bat`: sync da
  API → cache → gera HTML em `deploy/index.html` → commit e push para o GitHub Pages.
- Tarefa Windows `MetaAds-Sheets-Sync` roda `scripts/sync_meta_to_sheets.py`
  (atualmente desativada).
- `dash_sync.py` usa a API até o cutover (2026-06-18) com Meta vindo da planilha
  histórica `data/GROWTH_Estudo plataformas de mídia (2).xlsx`; depois, da API.

## Arquitetura
```
Meta Ads API → Python script → Google Sheets → Looker Studio
Google Ads                   → Conector nativo → Looker Studio
```

---

## Meta Ads

### Conta
- Conta: GNDI (Hapvida / Notredame Intermédica)
- Account ID: 407816857488424
- Token: System User (longa duração)

### Google Sheets
- Planilha: Meta Ads - Dados
- ID: 1FH8KDKYFj4JCe-q9ECOUnftIAh0XHhwqRWsRNJCXAfg
- Link: https://docs.google.com/spreadsheets/d/1FH8KDKYFj4JCe-q9ECOUnftIAh0XHhwqRWsRNJCXAfg

### Agendamento
- Tarefa Windows: MetaAds-Sheets-Sync
- Frequência: a cada 3 horas
- Período de dados: últimos 90 dias + hoje (lotes de 30 dias)
- O PC precisa estar ligado para rodar

### Colunas da planilha
| Coluna | O que representa |
|--------|-----------------|
| Investimento (R$) | Gasto total no período |
| Impressões | Vezes que o anúncio foi exibido |
| Cliques | Cliques no anúncio |
| CTR (%) | % de pessoas que viram e clicaram |
| CPC (R$) | Custo médio por clique |
| CPM (R$) | Custo a cada 1.000 impressões |
| Leads Formulário Meta | Preenchimento de formulário nativo dentro do Facebook/Instagram |
| Leads Onsite (Meta) | Versão consolidada do Meta — agrupa leadforms e interações de mensagem |
| Conversas WhatsApp | Conversas iniciadas no WhatsApp em até 7 dias após ver o anúncio |
| Pixel Custom Total | Qualquer evento customizado do pixel no site (tudo junto) |
| Custom: Tráfego Qualif. PF | Conversão customizada para Hapvida PF |
| Custom: Tráfego Qualif. PME | Conversão customizada para Hapvida PME |
| Custom: LP Smart | Conversão customizada para landing pages LP Smart |
| Custom: Notrelife | Conversão customizada para campanhas Notrelife |

### Pixels e eventos por tipo de campanha
- Conecta → purchase, form_conversion — Hapvida PF LP
- Conecta - PME → escolha_cep — Hapvida PME LP
- Pixel Hapvida NDI → form_conversion, form_conversion_lp_smart — maioria das campanhas
- Sem pixel — campanhas de WhatsApp e Leadform nativo

---

## Google Ads

### Conta
- MCC: MCC Notredame Intermédica (374-817-3931)
- Developer Token: Ua13bIZXUKI1TCgP9PaF0Q
- Nível de acesso API: Básico

### Observação importante
A API do Google Ads está bloqueada pela rede corporativa da Hapvida.
O Google Ads só é acessível via Looker Studio (conector nativo) ou fora da rede corporativa.

### Como ver eventos personalizados no Looker Studio
1. Abrir o relatório no Looker Studio
2. Adicionar dimensão: Nome da ação de conversão
3. Isso quebra os dados por cada evento de conversão cadastrado

---

## Scripts disponíveis

| Script | O que faz |
|--------|-----------|
| sync_meta_to_sheets.py | Sincroniza 90 dias de Meta Ads com o Google Sheets |
| listar_conversoes.py | Lista todos os eventos de conversão da conta Meta |
| listar_pixels_eventos.py | Lista pixels e eventos configurados por conjunto de anúncios |
| listar_exclusoes_publico.py | Lista exclusões de público em todos os conjuntos |
| listar_negativacao_beneficiarios.py | Lista campanhas ativas com negativação específica |

---

## Configuração técnica

### Requisitos
- Python 3.13+
- Bibliotecas: streamlit, pandas, requests, python-dotenv, google-ads, google-auth-oauthlib, gspread

### Arquivos de credenciais (não estão no repositório)
- `.env` — tokens Meta e Google Ads
- `service_account.json` — credenciais Google Sheets

### Repositório GitHub
https://github.com/gabrielborgescarbonari/ads-dashboard

---

## Negativações de público (Meta)
- 120 de 265 conjuntos têm exclusões de público configuradas
- Público mais usado como negativo: beneficiarios_ativos_integracao_api_mkt
- Campanhas ativas com essa negativação:
  - e_meta_hapvida_pf_lp-conecta-conversao
  - e_meta_hapvida_pf_lp-conecta-conversao_advantage
  - e_meta_mar-aberto_pf_leadform-native
  - e_meta_mar-aberto_pme_leadform-native
  - e_meta_mar-aberto_whatsapp
