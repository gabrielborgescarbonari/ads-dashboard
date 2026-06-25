"""
gerar_dashboard_midia.py — v3
Novas funcionalidades:
- % do total nos tooltips
- Filtro de campanha por texto (contém)
- Um filtro multiselect para CADA segmento (pipe) do nome da campanha:
  Frente | Tipo de Funil | Temperatura | Rede | Local | Objetivo | Data | Descrição
  Nomes sem taxonomia (< 6 segmentos) caem em "Sem <campo>".
- Múltipla seleção em todos os filtros categóricos
"""
import os, json
import pandas as pd

# Fonte e saida podem ser sobrescritas por variavel de ambiente (sync via API usa o cache).
XL_PATH  = os.environ.get('DASH_XL_PATH',  os.path.join(r'C:\Users\Usuario\Downloads', 'GROWTH_Estudo plataformas de mídia.xlsx'))
HTML_OUT = os.environ.get('DASH_HTML_OUT', os.path.join(r'C:\Users\Usuario\Downloads', 'dashboard_midia.html'))

# De-para de campanhas renomeadas (nome legado antigo -> nome novo com taxonomia).
# Migracao ocorreu so na conta Hapvida (Search), corte em 28/29-05. As variacoes
# curtas (horario-comercial, jornadamini-v2, sprj, h_*) NAO entram: sao separadas.
DEPARA = {
    'e_google_sch_hapvida_pf_generica-conversao':
        'PF | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | GENERICA',
    'e_google_sch_hapvida_pf_institucional-foco-conversao':
        'PF | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | INSTITUCIONAL',
    'e_google_sch_hapvida_pme_generica-conversao':
        'PME | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | GENERICA',
    'e_google_sch_hapvida_pme_institucional-conversao':
        'PME | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | INSTITUCIONAL',
    # Migracao jun/2026: legados que pararam em maio -> nomes novos de junho.
    'e_google_dgen_hapvida_pf_geral-conversao':
        'PF | FULLFUNNEL | T4-FRIO | GDN | BR | LED | MAI26 | CONECTA',
    'e_google_dgen_hapvida_pme_geral-conversao':
        'PME | FULLFUNNEL | T4-FRIO | GDN | BR | LED | MAI26 | LP-SMART',
    'e_google_pmax_hapvida_pme_geral-conversao':
        'PME | FULLFUNNEL | T4-FRIO | PMAX | BR | LED | MAI26 | LP-SMART',
    'e_google_pmax_notredame_pme_produtosppo-conversao':
        'PME | FULLFUNNEL | T4-FRIO | PMAX | BR | CPA | MAI26 | PRODUTOS-PPO',
    'e_google_sch_notredame_pme_produtosppo-conversao':
        'PME | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | PRODUTOS-PPO',
    'e_google_sch_notredame_pme_produtosppo-guarulhos-conversao':
        'PME | FULLFUNNEL | TS-SCH | SCH | BR | LED | MAI26 | PRODUTOS-PPO-GRARULHOS',
    # Meta: legado -> novo (caso A) e padronizacao barra->hifen (caso B).
    'e_meta_hapvida_pf_lp/conecta/conversao':
        'PF | FULLFUNNEL | T4-FRIO | META | BR | LED | MAI26 | LP-CONECTA',
    'e_meta_mar/aberto_leadform/native':
        'PF | FULLFUNNEL | T4-FRIO | META | BR | LED | MAI26 | LEADFORM',
    'PF | FULLFUNNEL | T4/FRIO | META | BR | LED | MAI26 | LP/CONECTA':
        'PF | FULLFUNNEL | T4-FRIO | META | BR | LED | MAI26 | LP-CONECTA',
    'PF | FULLFUNNEL | T4/FRIO | META | BR | LED | MAI26 | LEADFORM':
        'PF | FULLFUNNEL | T4-FRIO | META | BR | LED | MAI26 | LEADFORM',
    'PF | AWARENESS | T4/FRIO | META | BR | LED | MAI26 | ZICO':
        'PF | AWARENESS | T4-FRIO | META | BR | LED | MAI26 | ZICO',
    'h_meta_hapvida_pme_lp/conecta/conversao_advantage':
        'h_meta_hapvida_pme_lp-conecta-conversao_advantage',
    # PME SPRJ e de PPO, mas o nome na planilha nao tem PPO. So no dash.
    'PME | FULLFUNNEL | TS-SCH | SCH | BR | CPA | JUN26 | INSTITUCIONAL | SPRJ':
        'PME | FULLFUNNEL | TS-SCH | SCH | BR | CPA | JUN26 | INSTITUCIONAL | SPRJ-PPO',
}

print("Lendo Excel...")
_all  = pd.read_excel(XL_PATH, sheet_name=None, header=0)
g_raw = _all['Google']

# O Meta pode vir numa unica aba 'Meta' (legado) ou separado por conta
# (ex.: '407816857488424 -> GNDI', 'NNN GNDI Vendas'). Junta toda aba que
# nao seja Google/Leads/Procv e que tenha dados.
_NAO_META    = {'Google', 'Leads interno', 'Procv'}
_meta_sheets = [d for n, d in _all.items()
                if n not in _NAO_META and d.shape[1] >= 17 and len(d) > 0]
if not _meta_sheets:
    raise SystemExit('Nenhuma aba de Meta com dados encontrada no Excel.')
m_raw = pd.concat(_meta_sheets, ignore_index=True)


def clean_sheet(df):
    df = df.copy()
    n = df.shape[1]
    if n >= 18:  # cache: Conjunto apos Campanha (+ Visualizacoes/Inscricoes no cache novo)
        names = ['pl','ct','dt','ca','cj','pb','fm','iv','al','im','cl','ctr','cpc','cpm',
                 'lp','cpl_p','li','cpl_i','vv','insc']
        take = min(n, len(names))
        df = df.iloc[:, :take]
        df.columns = names[:take]
    else:  # planilha/cache legado de 17 colunas: sem Conjunto
        df = df.iloc[:, :17]
        df.columns = ['pl','ct','dt','ca','pb','fm','iv','al','im','cl','ctr','cpc','cpm','lp','cpl_p','li','cpl_i']
        df['cj'] = '(sem conjunto)'
    for extra in ['vv','insc']:  # tolera cache antigo (sem metricas de video)
        if extra not in df.columns:
            df[extra] = 0
    df = df[df['pl'].notna() & (df['pl'].astype(str).str.strip() != 'Plataforma')]
    df['dt'] = pd.to_datetime(df['dt'], errors='coerce')
    for col in ['iv','al','im','cl','lp','li','vv','insc']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['cj'] = df['cj'].fillna('(sem conjunto)')
    for col in ['pl','ct','ca','cj','pb','fm']:
        df[col] = df[col].astype(str).str.strip()
    df['ca'] = df['ca'].map(lambda c: DEPARA.get(c, c))  # unifica campanhas renomeadas
    return df.dropna(subset=['dt'])


# Campos da taxonomia do nome da campanha (separador ' | ').
# Cada posição do nome vira um filtro multiselect.
# (key_js, rótulo exibido, índice no nome, rótulo quando vazio)
TAX_FIELDS = [
    ('t_funil',  'Tipo de Funil', 1, 'Sem funil'),
    ('t_temp',   'Temperatura',   2, 'Sem temperatura'),
    ('t_obj',    'Objetivo',      5, 'Sem objetivo'),
    ('t_desc',   'Descrição',     7, 'Sem descrição'),
]
MIN_TAX_PARTS = 6  # nº mínimo de segmentos para o nome ser tratado como taxonomia

# Nova nomenclatura (rollout jun/2026): nome = slug minúsculo com underscore, mesmo
# conteúdo da taxonomia pipe. Normaliza-se para a forma canônica (CAPSLOCK com ' | ')
# para os valores se mesclarem com os nomes antigos nos filtros (senão META≠meta etc.).
NEW_FRENTES = {'pf', 'pme', 'all', 'ppo'}
NEW_FUNIS   = {'perpetuo', 'distribuicao', 'awareness', 'fullfunnel'}


def tax_parts(ca):
    """Devolve os segmentos da taxonomia (CAPSLOCK), tanto do formato pipe quanto do
    slug novo. 7 campos fixos + descrição (que absorve os segmentos extras)."""
    ca = str(ca)
    if ' | ' in ca:
        return [p.strip() for p in ca.split(' | ')]
    toks = ca.lower().split('_')
    if len(toks) >= 7 and toks[0] in NEW_FRENTES and toks[1] in NEW_FUNIS:
        parts = [t.upper() for t in toks[:7]]
        if len(toks) > 7:
            parts.append(' | '.join(t.upper() for t in toks[7:]))  # descrição
        return parts
    return [ca]  # fora da taxonomia → cai em "Sem <campo>"


def tax_value(ca, idx):
    parts = tax_parts(ca)
    if len(parts) < MIN_TAX_PARTS:
        return None
    if idx == 7:  # Descrição absorve qualquer segmento extra (campanhas com 9+ pipes)
        tail = [p for p in parts[7:] if p]
        return ' | '.join(tail) if tail else None
    if idx < len(parts) and parts[idx]:
        return parts[idx]
    return None


g  = clean_sheet(g_raw)
m  = clean_sheet(m_raw)
df = pd.concat([g, m], ignore_index=True)
print(f"  Google: {len(g)} | Meta: {len(m)} | Total: {len(df)}")

records = []
for _, row in df.iterrows():
    rec = {
        'pl': row['pl'], 'ct': row['ct'], 'dt': row['dt'].strftime('%Y-%m-%d'),
        'ca': row['ca'], 'cj': row['cj'],
        'pb': ('PPO' if 'PPO' in str(row['ca']).upper() else row['pb']),  # PPO = PME premium, separado
        'fm': row['fm'],
        'iv': round(float(row['iv']), 2),
        'al': int(row['al']), 'im': int(row['im']), 'cl': int(row['cl']),
        'lp': round(float(row['lp']), 2),
        'vv': int(row['vv']), 'insc': round(float(row['insc']), 2),
    }
    for key, _lbl, idx, empty in TAX_FIELDS:
        rec[key] = tax_value(row['ca'], idx) or empty
    records.append(rec)

data_json = json.dumps(records, ensure_ascii=False, separators=(',',':'))
min_date  = df['dt'].min().strftime('%Y-%m-%d')
max_date  = df['dt'].max().strftime('%Y-%m-%d')
print(f"  Periodo: {min_date} a {max_date} | Invest: R$ {df['iv'].sum():,.2f}")

contas    = sorted(df['ct'].unique())
formatos  = sorted(df['fm'].unique())


def ms_opts(values):
    return '\n'.join(
        f'              <label class="ms-item"><input type="checkbox" value="{v}"><span>{v}</span></label>'
        for v in values
    )


def field_values(key, empty):
    vals = sorted(set(r[key] for r in records))
    # empurra o rótulo de "vazio" para o fim da lista
    return [v for v in vals if v != empty] + ([empty] if empty in vals else [])


conta_opts   = ms_opts(contas)
formato_opts = ms_opts(formatos)


def tax_block(key, label, empty):
    opts = ms_opts(field_values(key, empty))
    return f'''  <div class="filter-group">
    <label>{label}</label>
    <div class="ms-wrap" id="ms-{key}">
      <button class="ms-btn" id="ms-{key}-btn"><span class="ms-label">Todos</span><span class="ms-arrow">▾</span></button>
      <div class="ms-dropdown" id="ms-{key}-dd">
        <div class="ms-clear">Limpar seleção</div>
{opts}
      </div>
    </div>
  </div>'''


tax_filters_html = '\n'.join(tax_block(k, l, e) for k, l, _i, e in TAX_FIELDS)
tax_state_keys   = ', '.join(f'{k}:[]' for k, *_ in TAX_FIELDS)
tax_init_js      = '\n  '.join(f"initMultiselect('ms-{k}', '{k}', 'Todos');" for k, *_ in TAX_FIELDS)
tax_filter_js    = '\n    '.join(
    f"if (state.{k}.length && !state.{k}.includes(r.{k})) return false;" for k, *_ in TAX_FIELDS
)

HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard de Mídia Paga — Hapvida</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:#f1f5f9; --card:#ffffff; --header-bg:#0f172a; --header-text:#f8fafc;
    --text:#1e293b; --text-light:#64748b; --border:#e2e8f0;
    --shadow:0 1px 3px rgba(0,0,0,.08),0 2px 8px rgba(0,0,0,.06);
    --shadow-md:0 4px 16px rgba(0,0,0,.1); --radius:12px;
    --c-blue:#3b82f6; --c-purple:#8b5cf6; --c-teal:#14b8a6; --c-orange:#f97316;
    --c-green:#22c55e; --c-red:#ef4444; --c-indigo:#6366f1; --c-pink:#ec4899;
    --font:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  }}
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh}}

  header{{background:var(--header-bg);color:var(--header-text);padding:28px 32px 24px;
    display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap}}
  header h1{{font-size:1.6rem;font-weight:700;letter-spacing:-.5px}}
  header p{{font-size:.95rem;color:#94a3b8;margin-top:4px}}
  #periodo-exibido{{font-size:.85rem;color:#cbd5e1;background:rgba(255,255,255,.07);
    border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:8px 14px;white-space:nowrap}}

  .filters-bar{{position:sticky;top:0;z-index:100;background:var(--header-bg);
    border-bottom:1px solid rgba(255,255,255,.08);padding:12px 32px;
    display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}}
  .filter-group{{display:flex;flex-direction:column;gap:4px}}
  .filter-group label{{font-size:.7rem;font-weight:600;text-transform:uppercase;
    letter-spacing:.5px;color:#94a3b8}}
  .filters-bar input[type=date]{{background:#1e293b;color:#f1f5f9;border:1px solid #334155;
    border-radius:7px;padding:6px 10px;font-size:.82rem;font-family:var(--font);
    cursor:pointer;min-width:130px}}
  .filters-bar input[type=date]:focus{{outline:2px solid var(--c-blue);border-color:transparent}}
  .gran-btns{{display:flex;gap:4px}}
  .gran-btn{{background:#1e293b;color:#94a3b8;border:1px solid #334155;border-radius:7px;
    padding:6px 14px;font-size:.82rem;font-family:var(--font);cursor:pointer;transition:all .15s}}
  .gran-btn.active{{background:var(--c-blue);color:#fff;border-color:var(--c-blue)}}
  .gran-btn:hover:not(.active){{background:#334155;color:#f1f5f9}}
  #date-custom{{display:none;flex-wrap:wrap;gap:8px;align-items:flex-end}}
  #date-custom.visible{{display:flex}}
  .filter-divider{{width:1px;height:36px;background:#334155;margin:0 4px;align-self:flex-end}}

  /* MULTISELECT */
  .ms-wrap{{position:relative;display:inline-block}}
  .ms-btn{{display:flex;align-items:center;justify-content:space-between;gap:8px;
    background:#1e293b;color:#f1f5f9;border:1px solid #334155;border-radius:7px;
    padding:6px 10px;font-size:.82rem;font-family:var(--font);cursor:pointer;
    min-width:120px;max-width:200px}}
  .ms-btn:focus{{outline:2px solid var(--c-blue);border-color:transparent}}
  .ms-label{{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left}}
  .ms-arrow{{flex-shrink:0;font-size:.6rem;color:#94a3b8;transition:transform .2s}}
  .ms-btn.open .ms-arrow{{transform:rotate(180deg)}}
  .ms-dropdown{{display:none;position:absolute;top:calc(100% + 4px);left:0;min-width:190px;
    background:#1e293b;border:1px solid #334155;border-radius:8px;z-index:300;
    padding:4px 0;max-height:260px;overflow-y:auto;
    box-shadow:0 8px 32px rgba(0,0,0,.5)}}
  .ms-dropdown.open{{display:block}}
  .ms-clear{{padding:7px 14px;font-size:.75rem;color:#94a3b8;
    border-bottom:1px solid #334155;cursor:pointer;transition:color .15s,background .15s}}
  .ms-clear:hover{{color:#f1f5f9;background:#334155}}
  .ms-item{{display:flex;align-items:center;gap:8px;padding:7px 14px;cursor:pointer;
    font-size:.82rem;color:#f1f5f9;user-select:none}}
  .ms-item:hover{{background:#334155}}
  .ms-item input[type=checkbox]{{width:14px;height:14px;accent-color:var(--c-blue);
    cursor:pointer;flex-shrink:0}}

  /* CAMPANHA SEARCH */
  .camp-search{{background:#1e293b;color:#f1f5f9;border:1px solid #334155;
    border-radius:7px;padding:6px 10px;font-size:.82rem;font-family:var(--font);
    min-width:150px}}
  .camp-search::placeholder{{color:#64748b}}
  .camp-search:focus{{outline:2px solid var(--c-blue);border-color:transparent}}

  main{{padding:28px 32px;max-width:none}}
  .filter-break{{flex-basis:100%;height:0;margin:0}}

  .kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px}}
  .kpi-card{{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);
    padding:20px 22px 18px;border-top:4px solid transparent;transition:box-shadow .2s}}
  .kpi-card:hover{{box-shadow:var(--shadow-md)}}
  .kpi-card.blue{{border-top-color:var(--c-blue)}} .kpi-card.purple{{border-top-color:var(--c-purple)}}
  .kpi-card.teal{{border-top-color:var(--c-teal)}}  .kpi-card.orange{{border-top-color:var(--c-orange)}}
  .kpi-card.green{{border-top-color:var(--c-green)}} .kpi-card.red{{border-top-color:var(--c-red)}}
  .kpi-card.indigo{{border-top-color:var(--c-indigo)}} .kpi-card.pink{{border-top-color:var(--c-pink)}}
  .kpi-label{{font-size:.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.5px;
    color:var(--text-light);margin-bottom:8px}}
  .kpi-value{{font-size:1.55rem;font-weight:700;color:var(--text);line-height:1}}
  .kpi-sub{{font-size:.8rem;color:var(--text-light);margin-top:6px}}

  .charts-row{{display:grid;gap:20px;margin-bottom:20px}}
  .charts-row.col3{{grid-template-columns:1fr 1fr 1fr}}
  .charts-row.col1{{grid-template-columns:1fr}}
  .charts-row.col6040{{grid-template-columns:3fr 2fr}}
  .charts-row.col2eq{{grid-template-columns:1fr 1fr}}
  .chart-card{{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:22px}}
  .chart-card h3{{font-size:.82rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;
    color:var(--text-light);margin-bottom:4px}}
  .chart-note{{font-size:.72rem;font-style:italic;font-weight:400;text-transform:none;
    letter-spacing:0;color:#94a3b8;margin-bottom:16px}}
  .chart-wrap{{position:relative}} .chart-wrap canvas{{max-width:100%}}

  .metric-btns{{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px}}
  .metric-btn{{background:var(--bg);color:var(--text-light);border:1px solid var(--border);
    border-radius:7px;padding:5px 12px;font-size:.75rem;font-family:var(--font);cursor:pointer;transition:all .15s}}
  .metric-btn.active{{background:var(--c-blue);color:#fff;border-color:var(--c-blue)}}
  .metric-btn:hover:not(.active){{background:#e2e8f0;color:var(--text)}}

  .funnel-wrap{{padding:8px 0}}
  .funnel-item{{margin-bottom:14px}}
  .funnel-label{{display:flex;justify-content:space-between;font-size:.78rem;margin-bottom:5px;color:var(--text)}}
  .funnel-label span{{color:var(--text-light)}}
  .funnel-bar-bg{{background:var(--bg);border-radius:6px;height:20px;overflow:hidden}}
  .funnel-bar{{height:100%;border-radius:6px;transition:width .4s ease}}

  .table-section{{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);
    padding:22px;margin-top:20px}}
  .table-header{{display:flex;justify-content:space-between;align-items:center;
    margin-bottom:16px;flex-wrap:wrap;gap:12px}}
  .table-header h3{{font-size:.82rem;font-weight:700;text-transform:uppercase;
    letter-spacing:.5px;color:var(--text-light)}}
  .table-info{{font-size:.78rem;color:var(--text-light)}}
  /* barra de checkboxes de colunas */
  .col-toggle{{display:flex;flex-wrap:wrap;gap:6px 14px;margin-bottom:12px;
    padding:10px 12px;background:var(--bg);border-radius:8px}}
  .col-toggle label{{display:flex;align-items:center;gap:5px;font-size:.72rem;
    color:var(--text-light);cursor:pointer;user-select:none}}
  .col-toggle input{{accent-color:var(--c-blue);cursor:pointer}}
  /* barra de rolagem no topo, sincronizada */
  .top-scroll{{overflow-x:auto;overflow-y:hidden}}
  .top-scroll>div{{height:1px}}
  .table-wrap{{overflow-x:auto}}
  table{{min-width:100%;border-collapse:collapse;font-size:.8rem;table-layout:fixed}}
  th{{background:var(--bg);padding:10px 12px;text-align:left;font-size:.7rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.5px;color:var(--text-light);cursor:pointer;
    user-select:none;white-space:nowrap;border-bottom:2px solid var(--border);
    position:relative;overflow:hidden;text-overflow:ellipsis}}
  th:hover{{background:#e2e8f0;color:var(--text)}}
  th .sort-icon{{margin-left:4px;opacity:.4}} th.sorted .sort-icon{{opacity:1;color:var(--c-blue)}}
  /* alça de redimensionamento da coluna */
  th .col-resize{{position:absolute;top:0;right:0;width:6px;height:100%;cursor:col-resize;
    user-select:none}}
  td{{padding:9px 12px;border-bottom:1px solid var(--border);color:var(--text);
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
  tr:last-child td{{border-bottom:none}} tr:hover td{{background:#f8fafc}}
  .badge{{display:inline-block;padding:2px 8px;border-radius:99px;font-size:.7rem;
    font-weight:700;text-transform:uppercase;letter-spacing:.3px}}
  .badge-google{{background:#dbeafe;color:#1d4ed8}} .badge-meta{{background:#ede9fe;color:#6d28d9}}
  .num{{text-align:right;font-variant-numeric:tabular-nums}}
  td[title]{{cursor:default}}

  ::-webkit-scrollbar{{width:6px;height:6px}}
  ::-webkit-scrollbar-track{{background:var(--bg)}}
  ::-webkit-scrollbar-thumb{{background:#cbd5e1;border-radius:3px}}

  @media(max-width:1100px){{
    .kpi-grid{{grid-template-columns:repeat(2,1fr)}}
    .charts-row.col3,.charts-row.col6040,.charts-row.col2eq{{grid-template-columns:1fr}}
  }}
  @media(max-width:640px){{
    main{{padding:16px}} .filters-bar{{padding:10px 16px}} header{{padding:20px 16px 18px}}
    .kpi-grid{{grid-template-columns:1fr 1fr}}
  }}
</style>
</head>
<body>

<header>
  <div>
    <h1>Dashboard de Mídia Paga — Hapvida</h1>
    <p>Google Ads &amp; Meta Ads</p>
  </div>
  <div id="periodo-exibido">Carregando...</div>
</header>

<div class="filters-bar">
  <div class="filter-group">
    <label>Granularidade</label>
    <div class="gran-btns">
      <button class="gran-btn active" data-gran="dia">Dia</button>
      <button class="gran-btn" data-gran="semana">Semana</button>
      <button class="gran-btn" data-gran="mes">Mês</button>
    </div>
  </div>
  <div class="filter-divider"></div>
  <div class="filter-group">
    <label>Período</label>
    <select id="fil-periodo" style="background:#1e293b;color:#f1f5f9;border:1px solid #334155;border-radius:7px;padding:6px 10px;font-size:.82rem;font-family:var(--font);cursor:pointer;min-width:130px">
      <option value="all">Todos</option>
      <option value="01">Janeiro</option>
      <option value="02">Fevereiro</option>
      <option value="03">Março</option>
      <option value="04">Abril</option>
      <option value="05">Maio</option>
      <option value="06">Junho</option>
      <option value="07">Julho</option>
      <option value="08">Agosto</option>
      <option value="09">Setembro</option>
      <option value="10">Outubro</option>
      <option value="11">Novembro</option>
      <option value="12">Dezembro</option>
      <option value="custom">Personalizado</option>
    </select>
  </div>
  <div id="date-custom">
    <div class="filter-group"><label>De</label><input type="date" id="fil-date-start" value="{min_date}"></div>
    <div class="filter-group"><label>Até</label><input type="date" id="fil-date-end" value="{max_date}"></div>
  </div>
  <div class="filter-divider"></div>

  <div class="filter-group">
    <label>Plataforma</label>
    <div class="ms-wrap" id="ms-plat">
      <button class="ms-btn" id="ms-plat-btn"><span class="ms-label">Todas</span><span class="ms-arrow">▾</span></button>
      <div class="ms-dropdown" id="ms-plat-dd">
        <div class="ms-clear">Limpar seleção</div>
        <label class="ms-item"><input type="checkbox" value="Google"><span>Google</span></label>
        <label class="ms-item"><input type="checkbox" value="Meta"><span>Meta</span></label>
      </div>
    </div>
  </div>

  <div class="filter-group">
    <label>Conta</label>
    <div class="ms-wrap" id="ms-conta">
      <button class="ms-btn" id="ms-conta-btn"><span class="ms-label">Todas</span><span class="ms-arrow">▾</span></button>
      <div class="ms-dropdown" id="ms-conta-dd">
        <div class="ms-clear">Limpar seleção</div>
{conta_opts}
      </div>
    </div>
  </div>

  <div class="filter-group">
    <label>Público</label>
    <div class="ms-wrap" id="ms-pub">
      <button class="ms-btn" id="ms-pub-btn"><span class="ms-label">Todos</span><span class="ms-arrow">▾</span></button>
      <div class="ms-dropdown" id="ms-pub-dd">
        <div class="ms-clear">Limpar seleção</div>
        <label class="ms-item"><input type="checkbox" value="PF"><span>PF</span></label>
        <label class="ms-item"><input type="checkbox" value="PME"><span>PME</span></label>
        <label class="ms-item"><input type="checkbox" value="PPO"><span>PPO</span></label>
        <label class="ms-item"><input type="checkbox" value="Awareness"><span>Awareness</span></label>
        <label class="ms-item"><input type="checkbox" value="Odonto"><span>Odonto</span></label>
      </div>
    </div>
  </div>

  <div class="filter-group">
    <label>Formato</label>
    <div class="ms-wrap" id="ms-fmt">
      <button class="ms-btn" id="ms-fmt-btn"><span class="ms-label">Todos</span><span class="ms-arrow">▾</span></button>
      <div class="ms-dropdown" id="ms-fmt-dd">
        <div class="ms-clear">Limpar seleção</div>
{formato_opts}
      </div>
    </div>
  </div>

  <div class="filter-break"></div>
{tax_filters_html}

  <div class="filter-group">
    <label>Campanha</label>
    <input type="text" id="fil-camp-search" class="camp-search" placeholder="Contém...">
  </div>
  <div class="filter-group">
    <label>Conjunto de anúncios</label>
    <input type="text" id="fil-cj-search" class="camp-search" placeholder="Contém...">
  </div>
</div>

<main>
  <div class="kpi-grid">
    <div class="kpi-card blue"><div class="kpi-label">Investimento Total</div><div class="kpi-value" id="kpi-inv">—</div><div class="kpi-sub" id="kpi-inv-sub">—</div></div>
    <div class="kpi-card purple"><div class="kpi-label">Impressões</div><div class="kpi-value" id="kpi-imp">—</div><div class="kpi-sub" id="kpi-imp-sub">—</div></div>
    <div class="kpi-card teal"><div class="kpi-label">Cliques</div><div class="kpi-value" id="kpi-clk">—</div><div class="kpi-sub" id="kpi-clk-sub">—</div></div>
    <div class="kpi-card orange"><div class="kpi-label">CPM Médio</div><div class="kpi-value" id="kpi-cpm">—</div><div class="kpi-sub" id="kpi-cpm-sub">—</div></div>
    <div class="kpi-card pink"><div class="kpi-label">CPC Médio</div><div class="kpi-value" id="kpi-cpc">—</div><div class="kpi-sub" id="kpi-cpc-sub">—</div></div>
    <div class="kpi-card green"><div class="kpi-label">Conversões Plataforma</div><div class="kpi-value" id="kpi-lp">—</div><div class="kpi-sub" id="kpi-lp-sub">—</div></div>
    <div class="kpi-card red"><div class="kpi-label">Visualizações</div><div class="kpi-value" id="kpi-vv">—</div><div class="kpi-sub" id="kpi-vv-sub">—</div></div>
    <div class="kpi-card teal"><div class="kpi-label">Inscrições</div><div class="kpi-value" id="kpi-insc">—</div><div class="kpi-sub" id="kpi-insc-sub">—</div></div>
    <div class="kpi-card indigo"><div class="kpi-label">CTR Médio</div><div class="kpi-value" id="kpi-ctr">—</div><div class="kpi-sub" id="kpi-ctr-sub">—</div></div>
  </div>

  <div class="charts-row col2eq">
    <div class="chart-card"><h3>Investimento por Plataforma</h3><div class="chart-note">Como o investimento se divide entre Google e Meta?</div><div class="chart-wrap" style="height:240px"><canvas id="chart-donut-inv"></canvas></div></div>
    <div class="chart-card"><h3>Conversões Plataforma por Público</h3><div class="chart-note">Para qual público vão as conversões?</div><div class="chart-wrap" style="height:240px"><canvas id="chart-donut-pub"></canvas></div></div>
  </div>

  <div class="charts-row col1"><div class="chart-card"><h3>Evolução de Investimento por Plataforma</h3><div class="chart-note">Como o investimento de cada plataforma evolui no tempo?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-inv-line"></canvas></div></div></div>
  <div class="charts-row col1"><div class="chart-card"><h3>Evolução de Conversões e CPL</h3><div class="chart-note">As conversões crescem enquanto o CPL cai?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-leads-line"></canvas></div></div></div>

  <div class="charts-row col1"><div class="chart-card">
    <h3>Evolução por Campanha</h3>
    <div class="chart-note">Compare a trajetória de cada campanha numa métrica ao longo do tempo (respeita os filtros e a granularidade; passe o mouse sobre a linha para ver a campanha).</div>
    <div class="metric-btns" id="camp-metric-btns"></div>
    <div class="chart-wrap" style="height:360px"><canvas id="chart-camp-line"></canvas></div>
  </div></div>

  <div class="charts-row col1"><div class="chart-card">
    <h3>Comparativo Mês a Mês (por dia do mês)</h3>
    <div class="chart-note">Cada linha é um mês, alinhados pelo dia do mês (dia 1, 2, 3…), para comparar na mesma quantidade de dias. Ignora o filtro de período (sempre de janeiro ao mês atual); respeita os demais filtros. Acumulado = total do mês até o dia (CPL acum. = investimento acum. ÷ leads acum.).</div>
    <div class="metric-btns" id="mm-metric-btns"></div>
    <div class="metric-btns" id="mm-mode-btns"></div>
    <div class="chart-wrap" style="height:380px"><canvas id="chart-mm-line"></canvas></div>
  </div></div>

  <div class="charts-row col1"><div class="chart-card">
    <h3>Evolução Comparativa (duas métricas)</h3>
    <div class="chart-note">Compare duas métricas no tempo, cada uma no seu eixo (esquerdo e direito). Respeita os filtros e a granularidade.</div>
    <div class="metric-btns">
      <span style="font-size:11px;color:#64748b;align-self:center">Eixo esquerdo:</span>
      <select id="dual-a" style="font-size:12px;padding:4px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text)"></select>
      <span style="font-size:11px;color:#64748b;align-self:center;margin-left:8px">Eixo direito:</span>
      <select id="dual-b" style="font-size:12px;padding:4px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text)"></select>
    </div>
    <div class="chart-wrap" style="height:300px"><canvas id="chart-dual-line"></canvas></div>
  </div></div>

  <div class="charts-row col2eq">
    <div class="chart-card"><h3>Share Google vs Meta no Tempo</h3><div class="chart-note">O mix de investimento entre as plataformas está mudando?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-share"></canvas></div></div>
    <div class="chart-card"><h3>Comparativo por Público</h3><div class="chart-note">Qual público rende mais conversão por real investido?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-pub-comp"></canvas></div></div>
  </div>

  <div class="charts-row col2eq">
    <div class="chart-card"><h3>Análise por Dia da Semana</h3><div class="chart-note">Em quais dias da semana o resultado aparece?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-dow"></canvas></div></div>
    <div class="chart-card"><h3>Conversões por R$ 1 mil — Formato</h3><div class="chart-note">Qual formato entrega mais conversões por real investido?</div><div class="chart-wrap" style="height:280px"><canvas id="chart-fmt"></canvas></div></div>
  </div>

  <div class="charts-row col1"><div class="chart-card"><h3>Dispersão de Campanhas</h3><div class="chart-note">Quais campanhas gastam muito com CPL ruim? (alto investimento + alto CPL = canto superior direito)</div><div class="chart-wrap" style="height:360px"><canvas id="chart-bubble"></canvas></div></div></div>

  <div class="charts-row col2eq">
    <div class="chart-card"><h3>Top 10 Campanhas por Investimento</h3><div class="chart-note">Quais campanhas concentram o investimento?</div><div class="chart-wrap" style="height:340px"><canvas id="chart-top-camp"></canvas></div></div>
    <div class="chart-card"><h3>Ranking de Eficiência</h3><div class="chart-note">Quais campanhas escalar (topo) e quais cortar (base)? Conversões por R$ 1 mil.</div><div class="chart-wrap" style="height:340px"><canvas id="chart-rank"></canvas></div></div>
  </div>

  <div class="table-section">
    <div class="table-header"><h3>Campanhas</h3><div class="table-info" id="table-info">—</div></div>
    <div class="chart-note">Detalhe por campanha. Arraste a borda das colunas pra ver o nome completo; use os checkboxes pra escolher quais colunas aparecem.</div>
    <div class="col-toggle" id="col-toggle"></div>
    <div class="top-scroll" id="top-scroll"><div id="top-scroll-inner"></div></div>
    <div class="table-wrap" id="table-wrap">
      <table id="camp-table">
        <thead><tr>
          <th data-col="pl">Plataforma <span class="sort-icon">↕</span></th>
          <th data-col="ct">Conta <span class="sort-icon">↕</span></th>
          <th data-col="ca">Campanha <span class="sort-icon">↕</span></th>
          <th data-col="cj">Conjunto <span class="sort-icon">↕</span></th>
          <th data-col="pb">Público <span class="sort-icon">↕</span></th>
          <th data-col="fm">Formato <span class="sort-icon">↕</span></th>
          <th data-col="ob">Objetivo <span class="sort-icon">↕</span></th>
          <th data-col="iv" class="num sorted">Investimento <span class="sort-icon">↓</span></th>
          <th data-col="im" class="num">Impressões <span class="sort-icon">↕</span></th>
          <th data-col="cl" class="num">Cliques <span class="sort-icon">↕</span></th>
          <th data-col="ctr" class="num">CTR <span class="sort-icon">↕</span></th>
          <th data-col="cpc" class="num">CPC <span class="sort-icon">↕</span></th>
          <th data-col="lp" class="num">Conversões Plat <span class="sort-icon">↕</span></th>
          <th data-col="cpl_p" class="num">CPL Plat <span class="sort-icon">↕</span></th>
          <th data-col="vv" class="num">Visualizações <span class="sort-icon">↕</span></th>
          <th data-col="insc" class="num">Inscrições <span class="sort-icon">↕</span></th>
        </tr></thead>
        <tbody id="camp-tbody"></tbody>
      </table>
    </div>
  </div>
</main>

<script>
const DATA = {data_json};
const MIN_DATE = '{min_date}';
const MAX_DATE = '{max_date}';

// Default ao abrir: do 1o dia do mes ATE ONTEM (o dia de hoje fica fora por ser parcial,
// mas continua nos dados; o usuario pode incluir estendendo a data final ou escolhendo
// o mes/Todos). Granularidade Dia.
const _now = new Date();
function _ds(d){{ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }}
const _yest = new Date(_now); _yest.setDate(_now.getDate()-1);
let DEFAULT_START = _ds(new Date(_now.getFullYear(), _now.getMonth(), 1));
let DEFAULT_END = _ds(_yest);
if (DEFAULT_END < DEFAULT_START) {{ DEFAULT_START = DEFAULT_END = _ds(_now); }}  // dia 1 do mes

const state = {{
  gran: 'dia', periodo: 'custom',
  dateStart: DEFAULT_START, dateEnd: DEFAULT_END,
  plataforma: [], conta: [], publico: [], formato: [], {tax_state_keys},
  campSearch: '', cjSearch: '',
  campMetric: 'iv',
  mmMetric: 'iv', mmMode: 'cum',
  dualA: 'ctr', dualB: 'cpl',
  tableSortCol: 'iv', tableSortAsc: false
}};

const charts = {{}};
let PERIOD_TOTALS = {{iv:0,im:0,cl:0,lp:0}};
const COL_GOOGLE='#ca8a04', COL_META='#1e40af';
let HIDDEN_COLS = new Set(['cj']);  // conjunto comeca oculto: tabela em nivel de campanha

function fmtBRL(v)  {{ return 'R$ '+Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function fmtNum(v)  {{ return Number(v).toLocaleString('pt-BR'); }}
function fmtPct(v)  {{ return Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:2,maximumFractionDigits:2}})+'%'; }}

function getTimeKey(dt, gran) {{
  const d = new Date(dt+'T00:00:00');
  if (gran==='dia') return dt;
  if (gran==='semana') {{
    const day=d.getDay(), diff=(day===0)?-6:1-day, mon=new Date(d);
    mon.setDate(d.getDate()+diff); return mon.toISOString().slice(0,10);
  }}
  return dt.slice(0,7);
}}
function getTimeLabel(key, gran) {{
  if (gran==='dia')    {{ const [y,m,d]=key.split('-'); return d+'/'+m+'/'+y.slice(2); }}
  if (gran==='semana') {{ const [y,m,d]=key.split('-'); return 'Sem '+d+'/'+m; }}
  const months=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const [y,m]=key.split('-'); return months[parseInt(m)-1]+'/'+y.slice(2);
}}

function filterDataPeriodOnly() {{
  return DATA.filter(r => {{
    if (state.periodo==='custom') {{
      if (r.dt<state.dateStart||r.dt>state.dateEnd) return false;
    }} else if (state.periodo!=='all') {{
      if (r.dt.slice(5,7)!==state.periodo) return false;
    }}
    return true;
  }});
}}

function computePeriodTotals() {{
  const d = filterDataPeriodOnly();
  PERIOD_TOTALS = {{
    iv: d.reduce((s,r)=>s+r.iv,0),
    im: d.reduce((s,r)=>s+r.im,0),
    cl: d.reduce((s,r)=>s+r.cl,0),
    lp: d.reduce((s,r)=>s+r.lp,0),
  }};
}}

function afterLabelPct(key) {{
  return ctx => {{
    const total = PERIOD_TOTALS[key];
    if (!total || !ctx.raw) return undefined;
    return '  '+(ctx.raw/total*100).toFixed(1)+'% do total';
  }};
}}

function filterData() {{
  return DATA.filter(r => {{
    if (state.periodo==='custom') {{
      if (r.dt<state.dateStart||r.dt>state.dateEnd) return false;
    }} else if (state.periodo!=='all') {{
      if (r.dt.slice(5,7)!==state.periodo) return false;
    }}
    if (state.plataforma.length && !state.plataforma.includes(r.pl)) return false;
    if (state.conta.length     && !state.conta.includes(r.ct))        return false;
    if (state.publico.length   && !state.publico.includes(r.pb))      return false;
    if (state.formato.length   && !state.formato.includes(r.fm))      return false;
    {tax_filter_js}
    if (state.campSearch && !r.ca.toLowerCase().includes(state.campSearch.toLowerCase())) return false;
    if (state.cjSearch && !(r.cj||'').toLowerCase().includes(state.cjSearch.toLowerCase())) return false;
    return true;
  }});
}}

// Como filterData, mas SEM o filtro de período (o comparativo mês a mês precisa
// sempre de todos os meses). Mantém os filtros categóricos.
function filterDataCatOnly() {{
  return DATA.filter(r => {{
    if (state.plataforma.length && !state.plataforma.includes(r.pl)) return false;
    if (state.conta.length     && !state.conta.includes(r.ct))        return false;
    if (state.publico.length   && !state.publico.includes(r.pb))      return false;
    if (state.formato.length   && !state.formato.includes(r.fm))      return false;
    {tax_filter_js}
    if (state.campSearch && !r.ca.toLowerCase().includes(state.campSearch.toLowerCase())) return false;
    if (state.cjSearch && !(r.cj||'').toLowerCase().includes(state.cjSearch.toLowerCase())) return false;
    return true;
  }});
}}

function groupByField(data, field) {{
  const map={{}};
  for (const r of data) {{
    const k=r[field];
    if (!map[k]) map[k]={{iv:0,im:0,cl:0,lp:0}};
    map[k].iv+=r.iv; map[k].im+=r.im; map[k].cl+=r.cl; map[k].lp+=r.lp;
  }}
  return map;
}}

function groupByTime(data, gran) {{
  const map={{}};
  for (const r of data) {{
    const k=getTimeKey(r.dt,gran);
    if (!map[k]) map[k]={{iv:0,im:0,cl:0,lp:0,Google:0,Meta:0,PF:0,PME:0,PPO:0,Awareness:0,Odonto:0}};
    map[k].iv+=r.iv; map[k].im+=r.im; map[k].cl+=r.cl; map[k].lp+=r.lp;
    if (r.pl==='Google') map[k].Google+=r.iv;
    if (r.pl==='Meta')   map[k].Meta+=r.iv;
    if (r.pb==='PF')        map[k].PF+=r.lp;
    if (r.pb==='PME')       map[k].PME+=r.lp;
    if (r.pb==='PPO')       map[k].PPO+=r.lp;
    if (r.pb==='Awareness') map[k].Awareness+=r.lp;
    if (r.pb==='Odonto')    map[k].Odonto+=r.lp;
  }}
  const keys=Object.keys(map).sort();
  return {{keys,map}};
}}

function aggregateTable(data) {{
  const byCj = !HIDDEN_COLS.has('cj');  // conjunto ligado -> quebra a tabela por conjunto
  const map={{}};
  for (const r of data) {{
    const cj = byCj ? (r.cj||'(sem conjunto)') : '';
    const k=r.pl+'||'+r.ct+'||'+r.ca+'||'+cj+'||'+r.pb+'||'+r.fm;
    if (!map[k]) map[k]={{pl:r.pl,ct:r.ct,ca:r.ca,cj:byCj?cj:'—',pb:r.pb,fm:r.fm,ob:r.t_obj,iv:0,im:0,cl:0,lp:0,vv:0,insc:0}};
    map[k].iv+=r.iv; map[k].im+=r.im; map[k].cl+=r.cl; map[k].lp+=r.lp; map[k].vv+=(r.vv||0); map[k].insc+=(r.insc||0);
  }}
  return Object.values(map).map(r => {{
    r.ctr=r.im>0?r.cl/r.im*100:0; r.cpc=r.cl>0?r.iv/r.cl:0;
    r.cpl_p=r.lp>0?r.iv/r.lp:0;
    return r;
  }});
}}

const CHART_DEFAULTS = {{
  responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{ labels:{{ font:{{family:'system-ui',size:11}},boxWidth:12 }} }} }},
}};
function destroyChart(id) {{ if (charts[id]) {{ charts[id].destroy(); delete charts[id]; }} }}

function getPeriodoLabel(data) {{
  if (!data.length) return 'Sem dados';
  const dates=data.map(r=>r.dt).sort();
  const fmt=dt=>{{ const [y,m,d]=dt.split('-'); return d+'/'+m+'/'+y; }};
  return fmt(dates[0])+' → '+fmt(dates[dates.length-1]);
}}

function updateKPIs(data) {{
  const inv=data.reduce((s,r)=>s+r.iv,0), imp=data.reduce((s,r)=>s+r.im,0),
        clk=data.reduce((s,r)=>s+r.cl,0), lp=data.reduce((s,r)=>s+r.lp,0);
  const cpm=imp>0?inv/imp*1000:0, ctr=imp>0?clk/imp*100:0, cpl_p=lp>0?inv/lp:0, cpc=clk>0?inv/clk:0;
  document.getElementById('kpi-inv').textContent=fmtBRL(inv);
  document.getElementById('kpi-inv-sub').textContent='Total investido no período';
  document.getElementById('kpi-imp').textContent=fmtNum(Math.round(imp));
  document.getElementById('kpi-imp-sub').textContent='Total de impressões';
  document.getElementById('kpi-clk').textContent=fmtNum(Math.round(clk));
  document.getElementById('kpi-clk-sub').textContent='CTR: '+fmtPct(ctr);
  document.getElementById('kpi-cpm').textContent=fmtBRL(cpm);
  document.getElementById('kpi-cpm-sub').textContent='Custo por mil impressões';
  document.getElementById('kpi-cpc').textContent=fmtBRL(cpc);
  document.getElementById('kpi-cpc-sub').textContent='Custo por clique';
  document.getElementById('kpi-lp').textContent=fmtNum(Math.round(lp));
  document.getElementById('kpi-lp-sub').textContent='CPL: '+fmtBRL(cpl_p);
  const vv=data.reduce((s,r)=>s+(r.vv||0),0), insc=data.reduce((s,r)=>s+(r.insc||0),0);
  const ivVV=data.reduce((s,r)=>s+(r.vv>0?r.iv:0),0), ivIN=data.reduce((s,r)=>s+(r.insc>0?r.iv:0),0);
  const cpv=vv>0?ivVV/vv:0, cpin=insc>0?ivIN/insc:0;
  document.getElementById('kpi-vv').textContent=fmtNum(Math.round(vv));
  document.getElementById('kpi-vv-sub').textContent='CPV: '+fmtBRL(cpv);
  document.getElementById('kpi-insc').textContent=fmtNum(Math.round(insc));
  document.getElementById('kpi-insc-sub').textContent='Custo/insc.: '+fmtBRL(cpin);
  document.getElementById('kpi-ctr').textContent=fmtPct(ctr);
  document.getElementById('kpi-ctr-sub').textContent='Cliques / Impressões';
  document.getElementById('periodo-exibido').textContent=getPeriodoLabel(data);
}}

function updateDonutInv(data) {{
  destroyChart('donut-inv');
  const byPl=groupByField(data,'pl'), labels=Object.keys(byPl);
  const vals=labels.map(k=>byPl[k].iv), colors=labels.map(k=>k==='Google'?COL_GOOGLE:COL_META);
  charts['donut-inv']=new Chart(document.getElementById('chart-donut-inv'),{{
    type:'doughnut',
    data:{{labels,datasets:[{{data:vals,backgroundColor:colors,borderWidth:2,borderColor:'#fff',hoverOffset:6}}]}},
    options:{{...CHART_DEFAULTS,cutout:'62%',plugins:{{
      legend:CHART_DEFAULTS.plugins.legend,
      tooltip:{{callbacks:{{
        label:ctx=>' '+ctx.label+': '+fmtBRL(ctx.raw),
        afterLabel:afterLabelPct('iv')
      }}}}
    }}}}
  }});
}}

function updateDonutPub(data) {{
  destroyChart('donut-pub');
  const byPb=groupByField(data,'pb'), labels=Object.keys(byPb);
  const vals=labels.map(k=>byPb[k].lp);
  const palette={{'PF':'#22c55e','PME':'#3b82f6','PPO':'#8b5cf6','Awareness':'#f97316','Odonto':'#06b6d4'}};
  const colors=labels.map(k=>palette[k]||'#94a3b8');
  charts['donut-pub']=new Chart(document.getElementById('chart-donut-pub'),{{
    type:'doughnut',
    data:{{labels,datasets:[{{data:vals,backgroundColor:colors,borderWidth:2,borderColor:'#fff',hoverOffset:6}}]}},
    options:{{...CHART_DEFAULTS,cutout:'62%',plugins:{{
      legend:CHART_DEFAULTS.plugins.legend,
      tooltip:{{callbacks:{{
        label:ctx=>' '+ctx.label+': '+fmtNum(Math.round(ctx.raw)),
        afterLabel:afterLabelPct('lp')
      }}}}
    }}}}
  }});
}}

function updateInvLine(data) {{
  destroyChart('inv-line');
  const {{keys,map}}=groupByTime(data,state.gran), labels=keys.map(k=>getTimeLabel(k,state.gran));
  charts['inv-line']=new Chart(document.getElementById('chart-inv-line'),{{
    type:'line',
    data:{{labels,datasets:[
      {{label:'Google',data:keys.map(k=>map[k].Google),borderColor:COL_GOOGLE,backgroundColor:'rgba(202,138,4,.10)',tension:.35,pointRadius:3,fill:false}},
      {{label:'Meta',  data:keys.map(k=>map[k].Meta),  borderColor:COL_META,backgroundColor:'rgba(30,64,175,.10)',tension:.35,pointRadius:3,fill:false}},
      {{label:'Total', data:keys.map(k=>map[k].iv),    borderColor:'#94a3b8',borderDash:[6,3],tension:.35,pointRadius:3,fill:false}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{ticks:{{callback:v=>'R$ '+fmtNum(Math.round(v/1000))+'k',font:{{size:10}}}},grid:{{color:'#f1f5f9'}}}},
        x:{{ticks:{{font:{{size:10}},maxRotation:45}},grid:{{display:false}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>' '+ctx.dataset.label+': '+fmtBRL(ctx.raw),
        afterLabel:afterLabelPct('iv')
      }}}}}}
    }}
  }});
}}

function updateLeadsLine(data) {{
  destroyChart('leads-line');
  const {{keys,map}}=groupByTime(data,state.gran), labels=keys.map(k=>getTimeLabel(k,state.gran));
  const cplData=keys.map(k=>map[k].lp>0?map[k].iv/map[k].lp:null);
  charts['leads-line']=new Chart(document.getElementById('chart-leads-line'),{{
    type:'bar',
    data:{{labels,datasets:[
      {{label:'Conversões PF', data:keys.map(k=>map[k].PF), backgroundColor:'rgba(34,197,94,.7)', stack:'leads',order:2}},
      {{label:'Conversões PME',data:keys.map(k=>map[k].PME),backgroundColor:'rgba(59,130,246,.7)',stack:'leads',order:2}},
      {{label:'Conversões PPO',data:keys.map(k=>map[k].PPO),backgroundColor:'rgba(139,92,246,.7)',stack:'leads',order:2}},
      {{label:'Conversões Odonto',data:keys.map(k=>map[k].Odonto),backgroundColor:'rgba(6,182,212,.7)',stack:'leads',order:2}},
      {{label:'CPL',data:cplData,type:'line',borderColor:'#f97316',backgroundColor:'transparent',
        yAxisID:'y2',tension:.35,pointRadius:4,order:1,borderWidth:2}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{stacked:true,ticks:{{font:{{size:10}}}},grid:{{color:'#f1f5f9'}}}},
        y2:{{position:'right',ticks:{{callback:v=>fmtBRL(v),font:{{size:10}}}},grid:{{display:false}}}},
        x:{{ticks:{{font:{{size:10}},maxRotation:45}},grid:{{display:false}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>ctx.dataset.label==='CPL'?' CPL: '+fmtBRL(ctx.raw):' '+ctx.dataset.label+': '+fmtNum(Math.round(ctx.raw)),
        afterLabel:ctx=>ctx.dataset.label==='CPL'?undefined:afterLabelPct('lp')(ctx),
        footer:items=>{{ if(!items.length) return ''; const k=keys[items[0].dataIndex]; return 'Total de conversões: '+fmtNum(Math.round(map[k].lp)); }}
      }}}}}}
    }}
  }});
}}

const CAMP_METRICS=[
  {{k:'iv', label:'Investimento', calc:a=>a.iv, fmt:fmtBRL, tick:v=>'R$ '+fmtNum(Math.round(v/1000))+'k'}},
  {{k:'lp', label:'Conversões',   calc:a=>a.lp, fmt:v=>fmtNum(Math.round(v)), tick:v=>fmtNum(Math.round(v))}},
  {{k:'cpl',label:'CPL',          calc:a=>a.lp>0?a.iv/a.lp:null, fmt:fmtBRL, tick:v=>fmtBRL(v)}},
  {{k:'im', label:'Impressões',   calc:a=>a.im, fmt:v=>fmtNum(Math.round(v)), tick:v=>fmtNum(Math.round(v/1000))+'k'}},
  {{k:'cl', label:'Cliques',      calc:a=>a.cl, fmt:v=>fmtNum(Math.round(v)), tick:v=>fmtNum(Math.round(v))}},
  {{k:'ctr',label:'CTR',          calc:a=>a.im>0?a.cl/a.im*100:null, fmt:fmtPct, tick:v=>fmtPct(v)}},
  {{k:'cpc',label:'CPC',          calc:a=>a.cl>0?a.iv/a.cl:null, fmt:fmtBRL, tick:v=>fmtBRL(v)}},
  {{k:'cpm',label:'CPM',          calc:a=>a.im>0?a.iv/a.im*1000:null, fmt:fmtBRL, tick:v=>fmtBRL(v)}},
];

function groupByCampTime(data, gran) {{
  const camps={{}}; const keySet=new Set();
  for (const r of data) {{
    const tk=getTimeKey(r.dt,gran); keySet.add(tk);
    if (!camps[r.ca]) camps[r.ca]={{}};
    if (!camps[r.ca][tk]) camps[r.ca][tk]={{iv:0,im:0,cl:0,lp:0}};
    const a=camps[r.ca][tk]; a.iv+=r.iv; a.im+=r.im; a.cl+=r.cl; a.lp+=r.lp;
  }}
  return {{keys:[...keySet].sort(), camps}};
}}

function campColor(i, n) {{ return 'hsl('+Math.round(360*i/Math.max(1,n))+',62%,50%)'; }}

function updateCampLine(data) {{
  destroyChart('camp-line');
  const metric=CAMP_METRICS.find(m=>m.k===state.campMetric)||CAMP_METRICS[0];
  const {{keys,camps}}=groupByCampTime(data,state.gran);
  const labels=keys.map(k=>getTimeLabel(k,state.gran));
  // ordena por investimento total desc (cores estáveis e legenda do maior pro menor)
  const names=Object.keys(camps).sort((a,b)=>{{
    const sa=keys.reduce((s,k)=>s+((camps[a][k]||{{}}).iv||0),0);
    const sb=keys.reduce((s,k)=>s+((camps[b][k]||{{}}).iv||0),0);
    return sb-sa;
  }});
  const n=names.length;
  const datasets=names.map((ca,i)=>{{
    const color=campColor(i,n);
    return {{
      label:ca.length>46?ca.slice(0,46)+'…':ca, fullName:ca,
      data:keys.map(k=>camps[ca][k]?metric.calc(camps[ca][k]):null),
      borderColor:color, backgroundColor:color, tension:.3, pointRadius:2, borderWidth:1.5, spanGaps:true
    }};
  }});
  charts['camp-line']=new Chart(document.getElementById('chart-camp-line'),{{
    type:'line',
    data:{{labels,datasets}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{ticks:{{callback:metric.tick,font:{{size:10}}}},grid:{{color:'#f1f5f9'}}}},
        x:{{ticks:{{font:{{size:10}},maxRotation:45}},grid:{{display:false}}}}
      }},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{
          title:items=>items.length?labels[items[0].dataIndex]:'',
          label:ctx=>' '+(ctx.dataset.fullName||ctx.dataset.label)+': '+(ctx.raw==null?'—':metric.fmt(ctx.raw))
        }}}}
      }}
    }}
  }});
}}

function initCampMetricBtns() {{
  const box=document.getElementById('camp-metric-btns');
  CAMP_METRICS.forEach(m=>{{
    const b=document.createElement('button');
    b.className='metric-btn'+(m.k===state.campMetric?' active':'');
    b.textContent=m.label; b.dataset.k=m.k;
    b.addEventListener('click',()=>{{
      state.campMetric=m.k;
      box.querySelectorAll('.metric-btn').forEach(x=>x.classList.remove('active'));
      b.classList.add('active');
      updateCampLine(filterData());
    }});
    box.appendChild(b);
  }});
}}

// ---- Comparativo mês a mês (cada linha = um mês; eixo X = dia do mês) ----
const MM_METRICS=[
  {{k:'iv',  label:'Investimento', ratio:false, pick:a=>a.iv, fmt:fmtBRL,                   tick:v=>'R$ '+fmtNum(Math.round(v/1000))+'k'}},
  {{k:'lp',  label:'Leads',        ratio:false, pick:a=>a.lp, fmt:v=>fmtNum(Math.round(v)), tick:v=>fmtNum(Math.round(v))}},
  {{k:'cpl', label:'CPL',          ratio:true,                fmt:fmtBRL,                   tick:v=>fmtBRL(v)}},
];
const MM_MONTHS=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function mmColor(monthNum) {{ return 'hsl('+Math.round((monthNum-1)/12*360)+',64%,48%)'; }}

function updateMMLine(data) {{
  destroyChart('mm-line');
  const metric=MM_METRICS.find(m=>m.k===state.mmMetric)||MM_METRICS[0];
  const cum=state.mmMode==='cum';
  const months={{}};
  for (const r of data) {{
    const ym=r.dt.slice(0,7), dom=parseInt(r.dt.slice(8,10),10);
    if (!months[ym]) months[ym]={{}};
    const cell=months[ym][dom]||(months[ym][dom]={{iv:0,lp:0}});
    cell.iv+=r.iv; cell.lp+=r.lp;
  }}
  const yms=Object.keys(months).sort();
  const labels=[]; for (let d=1; d<=31; d++) labels.push(String(d));
  const datasets=yms.map(ym=>{{
    const mnum=parseInt(ym.slice(5,7),10);
    const lastDay=Math.max(...Object.keys(months[ym]).map(Number));
    let cumIv=0, cumLp=0; const arr=[];
    for (let d=1; d<=31; d++) {{
      const cell=months[ym][d];
      if (cell) {{ cumIv+=cell.iv; cumLp+=cell.lp; }}
      let v=null;
      if (cum) {{
        if (d<=lastDay) v = metric.ratio ? (cumLp>0?cumIv/cumLp:null) : metric.pick({{iv:cumIv,lp:cumLp}});
      }} else if (cell) {{
        v = metric.ratio ? (cell.lp>0?cell.iv/cell.lp:null) : metric.pick(cell);
      }}
      arr.push(v);
    }}
    const color=mmColor(mnum);
    return {{label:MM_MONTHS[mnum-1]+'/'+ym.slice(2,4), data:arr, borderColor:color, backgroundColor:color,
             tension:.25, pointRadius:1.5, borderWidth:1.8, spanGaps:true}};
  }});
  charts['mm-line']=new Chart(document.getElementById('chart-mm-line'),{{
    type:'line',
    data:{{labels,datasets}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{ticks:{{callback:metric.tick,font:{{size:10}}}},grid:{{color:'#f1f5f9'}},beginAtZero:!metric.ratio}},
        x:{{title:{{display:true,text:'dia do mês',font:{{size:10}}}},ticks:{{font:{{size:9}}}},grid:{{display:false}}}}
      }},
      plugins:{{
        legend:{{display:true,position:'top',labels:{{font:{{size:10}},boxWidth:12}}}},
        tooltip:{{callbacks:{{
          title:items=>items.length?'Dia '+labels[items[0].dataIndex]:'',
          label:ctx=>' '+ctx.dataset.label+': '+(ctx.raw==null?'—':metric.fmt(ctx.raw))
        }}}}
      }}
    }}
  }});
}}

function initMMControls() {{
  const mkBtn=(box,label,active,onClick)=>{{
    const b=document.createElement('button');
    b.className='metric-btn'+(active?' active':'');
    b.textContent=label;
    b.addEventListener('click',()=>{{
      box.querySelectorAll('.metric-btn').forEach(x=>x.classList.remove('active'));
      b.classList.add('active'); onClick();
      updateMMLine(filterDataCatOnly());
    }});
    box.appendChild(b);
  }};
  const mkLabel=(box,txt)=>{{
    const s=document.createElement('span'); s.textContent=txt;
    s.style.cssText='font-size:11px;color:#64748b;align-self:center;margin-right:2px';
    box.appendChild(s);
  }};
  const mbox=document.getElementById('mm-metric-btns');
  mkLabel(mbox,'Métrica:');
  MM_METRICS.forEach(m=>mkBtn(mbox,m.label,m.k===state.mmMetric,()=>state.mmMetric=m.k));
  const obox=document.getElementById('mm-mode-btns');
  mkLabel(obox,'Valores:');
  [['cum','Acumulado'],['day','Diário']].forEach(([k,lbl])=>
    mkBtn(obox,lbl,k===state.mmMode,()=>state.mmMode=k));
}}

function updateTopCamp(data) {{
  destroyChart('top-camp');
  const byCamp=groupByField(data,'ca');
  const sorted=Object.entries(byCamp).sort((a,b)=>b[1].iv-a[1].iv).slice(0,10);
  const labels=sorted.map(([k])=>k.length>40?k.slice(0,40)+'…':k), vals=sorted.map(([,v])=>v.iv);
  charts['top-camp']=new Chart(document.getElementById('chart-top-camp'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Investimento',data:vals,backgroundColor:'rgba(59,130,246,.75)',borderRadius:4}}]}},
    options:{{...CHART_DEFAULTS,indexAxis:'y',
      scales:{{
        x:{{ticks:{{callback:v=>'R$ '+fmtNum(Math.round(v/1000))+'k',font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}},
        y:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}}
      }},
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{
        label:ctx=>' '+fmtBRL(ctx.raw),
        afterLabel:afterLabelPct('iv')
      }}}}}}
    }}
  }});
}}

function updateFmt(data) {{
  destroyChart('fmt');
  const byFmt=groupByField(data,'fm');
  const arr=Object.keys(byFmt).filter(k=>byFmt[k].iv>0)
    .map(k=>({{k,v:byFmt[k].lp/byFmt[k].iv*1000}})).sort((a,b)=>b.v-a.v);
  const labels=arr.map(x=>x.k), vals=arr.map(x=>x.v);
  charts['fmt']=new Chart(document.getElementById('chart-fmt'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Conversões por R$ 1 mil',data:vals,backgroundColor:'rgba(34,197,94,.78)',borderRadius:4}}]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{ticks:{{font:{{size:9}}}},grid:{{color:'#f1f5f9'}},title:{{display:true,text:'conv. por R$ 1 mil',font:{{size:9}}}}}},
        x:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}}
      }},
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{
        label:ctx=>' '+Number(ctx.raw).toLocaleString('pt-BR',{{maximumFractionDigits:2}})+' conv. por R$ 1 mil'
      }}}}}}
    }}
  }});
}}

// Evolução comparativa: o usuário escolhe duas métricas (eixos esq./dir.), reaproveitando
// as definições de CAMP_METRICS (calc opera sobre o agregado {{iv,im,cl,lp}} do groupByTime).
function updateDualLine(data) {{
  destroyChart('dual-line');
  const mA=CAMP_METRICS.find(m=>m.k===state.dualA)||CAMP_METRICS[0];
  const mB=CAMP_METRICS.find(m=>m.k===state.dualB)||CAMP_METRICS[2];
  const {{keys,map}}=groupByTime(data,state.gran), labels=keys.map(k=>getTimeLabel(k,state.gran));
  const dA=keys.map(k=>mA.calc(map[k])), dB=keys.map(k=>mB.calc(map[k]));
  charts['dual-line']=new Chart(document.getElementById('chart-dual-line'),{{
    type:'line',
    data:{{labels,datasets:[
      {{label:mA.label+' (esq.)',data:dA,borderColor:'#6366f1',backgroundColor:'transparent',yAxisID:'y',tension:.35,pointRadius:3,borderWidth:2,spanGaps:true}},
      {{label:mB.label+' (dir.)',data:dB,borderColor:'#f97316',backgroundColor:'transparent',yAxisID:'y2',tension:.35,pointRadius:3,borderWidth:2,spanGaps:true}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{position:'left',ticks:{{callback:mA.tick,font:{{size:10}}}},grid:{{color:'#f1f5f9'}}}},
        y2:{{position:'right',ticks:{{callback:mB.tick,font:{{size:10}}}},grid:{{display:false}}}},
        x:{{ticks:{{font:{{size:10}},maxRotation:45}},grid:{{display:false}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>{{const m=ctx.datasetIndex===0?mA:mB; return ' '+m.label+': '+(ctx.raw==null?'—':m.fmt(ctx.raw));}}
      }}}}}}
    }}
  }});
}}

function initDualControls() {{
  const fill=(sel,cur)=>CAMP_METRICS.forEach(m=>{{
    const o=document.createElement('option'); o.value=m.k; o.textContent=m.label;
    if (m.k===cur) o.selected=true; sel.appendChild(o);
  }});
  const a=document.getElementById('dual-a'), b=document.getElementById('dual-b');
  fill(a,state.dualA); fill(b,state.dualB);
  a.addEventListener('change',()=>{{ state.dualA=a.value; updateDualLine(filterData()); }});
  b.addEventListener('change',()=>{{ state.dualB=b.value; updateDualLine(filterData()); }});
}}

function updateShare(data) {{
  destroyChart('share');
  const {{keys,map}}=groupByTime(data,state.gran), labels=keys.map(k=>getTimeLabel(k,state.gran));
  const gp=keys.map(k=>{{const t=map[k].Google+map[k].Meta;return t>0?map[k].Google/t*100:0;}});
  const mp=keys.map(k=>{{const t=map[k].Google+map[k].Meta;return t>0?map[k].Meta/t*100:0;}});
  charts['share']=new Chart(document.getElementById('chart-share'),{{
    type:'bar',
    data:{{labels,datasets:[
      {{label:'Google',data:gp,backgroundColor:COL_GOOGLE,stack:'s',borderRadius:2,maxBarThickness:46}},
      {{label:'Meta',data:mp,backgroundColor:COL_META,stack:'s',borderRadius:2,maxBarThickness:46}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        x:{{stacked:true,ticks:{{font:{{size:10}},maxRotation:45}},grid:{{display:false}}}},
        y:{{stacked:true,min:0,max:100,ticks:{{callback:v=>v+'%',font:{{size:10}}}},grid:{{color:'#f1f5f9'}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>' '+ctx.dataset.label+': '+fmtPct(ctx.raw)
      }}}}}}
    }}
  }});
}}

function updatePubComp(data) {{
  destroyChart('pub-comp');
  const byPb=groupByField(data,'pb'), labels=Object.keys(byPb);
  const invV=labels.map(k=>byPb[k].iv), lpV=labels.map(k=>byPb[k].lp);
  const cplV=labels.map(k=>byPb[k].lp>0?byPb[k].iv/byPb[k].lp:0);
  charts['pub-comp']=new Chart(document.getElementById('chart-pub-comp'),{{
    type:'bar',
    data:{{labels,datasets:[
      {{label:'Investimento',data:invV,backgroundColor:'rgba(99,102,241,.75)',yAxisID:'y',borderRadius:3}},
      {{label:'Conversões',data:lpV,backgroundColor:'rgba(34,197,94,.75)',yAxisID:'y2',borderRadius:3}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{position:'left',ticks:{{callback:v=>'R$ '+fmtNum(Math.round(v/1000))+'k',font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}},
        y2:{{position:'right',ticks:{{font:{{size:9}}}},grid:{{display:false}}}},
        x:{{ticks:{{font:{{size:10}}}},grid:{{display:false}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>ctx.dataset.label==='Investimento'?' Invest: '+fmtBRL(ctx.raw):' Conversões: '+fmtNum(Math.round(ctx.raw)),
        afterBody:items=>{{const i=items[0].dataIndex;return 'CPL: '+fmtBRL(cplV[i]);}}
      }}}}}}
    }}
  }});
}}

function updateDow(data) {{
  destroyChart('dow');
  const names=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  const agg=names.map(()=>({{iv:0,lp:0,dates:new Set()}}));
  for (const r of data) {{
    const wd=new Date(r.dt+'T00:00:00').getDay();
    agg[wd].iv+=r.iv; agg[wd].lp+=r.lp; agg[wd].dates.add(r.dt);
  }}
  const order=[1,2,3,4,5,6,0];
  const labels=order.map(i=>names[i]);
  const invAvg=order.map(i=>agg[i].dates.size>0?agg[i].iv/agg[i].dates.size:0);
  const lpAvg=order.map(i=>agg[i].dates.size>0?agg[i].lp/agg[i].dates.size:0);
  charts['dow']=new Chart(document.getElementById('chart-dow'),{{
    type:'bar',
    data:{{labels,datasets:[
      {{label:'Investimento médio/dia',data:invAvg,backgroundColor:'rgba(99,102,241,.75)',yAxisID:'y',borderRadius:3}},
      {{label:'Conversões médias/dia',data:lpAvg,backgroundColor:'rgba(34,197,94,.75)',yAxisID:'y2',borderRadius:3}}
    ]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        y:{{position:'left',ticks:{{callback:v=>'R$ '+fmtNum(Math.round(v/1000))+'k',font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}},
        y2:{{position:'right',ticks:{{font:{{size:9}}}},grid:{{display:false}}}},
        x:{{ticks:{{font:{{size:10}}}},grid:{{display:false}}}}
      }},
      plugins:{{legend:CHART_DEFAULTS.plugins.legend,tooltip:{{callbacks:{{
        label:ctx=>ctx.dataset.label.indexOf('Invest')===0?' Invest médio: '+fmtBRL(ctx.raw):' Conv. médias: '+fmtNum(Math.round(ctx.raw))
      }}}}}}
    }}
  }});
}}

function updateRank(data) {{
  destroyChart('rank');
  const byCamp=groupByField(data,'ca');
  const arr=Object.keys(byCamp).filter(k=>byCamp[k].iv>0&&byCamp[k].lp>0)
    .map(k=>({{k,v:byCamp[k].lp/byCamp[k].iv*1000}})).sort((a,b)=>b.v-a.v);
  const sel=arr.length<=10?arr:arr.slice(0,5).concat(arr.slice(-5));
  const labels=sel.map(x=>x.k.length>34?x.k.slice(0,34)+'…':x.k), vals=sel.map(x=>x.v);
  const colors=sel.map((x,i)=>(arr.length>10&&i>=5)?'rgba(239,68,68,.78)':'rgba(34,197,94,.78)');
  charts['rank']=new Chart(document.getElementById('chart-rank'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Conversões por R$ 1 mil',data:vals,backgroundColor:colors,borderRadius:4}}]}},
    options:{{...CHART_DEFAULTS,indexAxis:'y',
      scales:{{
        x:{{ticks:{{font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}},
        y:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}}
      }},
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{
        label:ctx=>' '+Number(ctx.raw).toLocaleString('pt-BR',{{maximumFractionDigits:2}})+' conv. por R$ 1 mil'
      }}}}}}
    }}
  }});
}}

function quantile(sorted, p) {{
  if (!sorted.length) return 0;
  const i=(sorted.length-1)*p, lo=Math.floor(i), hi=Math.ceil(i);
  return sorted[lo]+(sorted[hi]-sorted[lo])*(i-lo);
}}

function updateBubble(data) {{
  destroyChart('bubble');
  const byCamp=groupByField(data,'ca');
  const pts=Object.keys(byCamp).filter(k=>byCamp[k].iv>0&&byCamp[k].lp>0).map(k=>({{
    ca:k, x:byCamp[k].iv, cpl:byCamp[k].iv/byCamp[k].lp, conv:byCamp[k].lp
  }}));
  // teto do eixo Y pela cerca de outliers (Q3 + 1.5*IQR); evita que 1 campanha estique tudo
  const cpls=pts.map(p=>p.cpl).sort((a,b)=>a-b);
  const q1=quantile(cpls,.25), q3=quantile(cpls,.75);
  let cap=q3+1.5*(q3-q1);
  if (!(cap>0) || cap>=Math.max(...cpls,0)) cap=0;   // sem outliers -> escala normal
  const yMax = cap>0 ? cap*1.08 : undefined;
  const maxConv=Math.max(1,...pts.map(p=>p.conv));
  const bubbles=pts.map(p=>{{
    const over=cap>0 && p.cpl>cap;
    return {{x:p.x, y:over?cap:p.cpl, r:4+Math.sqrt(p.conv/maxConv)*22, ca:p.ca, conv:p.conv, cpl:p.cpl, over}};
  }});
  charts['bubble']=new Chart(document.getElementById('chart-bubble'),{{
    type:'bubble',
    data:{{datasets:[{{label:'Campanhas',data:bubbles,
      backgroundColor:bubbles.map(b=>b.over?'rgba(239,68,68,.55)':'rgba(59,130,246,.45)'),
      borderColor:bubbles.map(b=>b.over?'#ef4444':'#3b82f6'),borderWidth:1}}]}},
    options:{{...CHART_DEFAULTS,
      scales:{{
        x:{{title:{{display:true,text:'Investimento',font:{{size:10}}}},ticks:{{callback:v=>'R$ '+fmtNum(Math.round(v/1000))+'k',font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}},
        y:{{title:{{display:true,text:'CPL (R$ por conversão)',font:{{size:10}}}},max:yMax,ticks:{{callback:v=>fmtBRL(v),font:{{size:9}}}},grid:{{color:'#f1f5f9'}}}}
      }},
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{
        label:ctx=>{{const d=ctx.raw;return [d.ca,'Invest: '+fmtBRL(d.x),'CPL: '+fmtBRL(d.cpl)+(d.over?' (fora de escala)':''),'Conversões: '+fmtNum(Math.round(d.conv))];}}
      }}}}}}
    }}
  }});
}}

function updateTable(data) {{
  let rows=aggregateTable(data);
  const col=state.tableSortCol;
  rows.sort((a,b)=>{{
    const va=a[col],vb=b[col];
    if (typeof va==='string') return state.tableSortAsc?va.localeCompare(vb):vb.localeCompare(va);
    return state.tableSortAsc?va-vb:vb-va;
  }});
  const tbody=document.getElementById('camp-tbody');
  tbody.innerHTML='';
  rows.forEach(r=>{{
    const tr=document.createElement('tr');
    const badge=r.pl==='Google'?'<span class="badge badge-google">Google</span>':'<span class="badge badge-meta">Meta</span>';
    tr.innerHTML=`
      <td data-col="pl">${{badge}}</td>
      <td data-col="ct" title="${{r.ct}}">${{r.ct}}</td>
      <td data-col="ca" title="${{r.ca}}">${{r.ca}}</td>
      <td data-col="cj" title="${{r.cj}}">${{r.cj}}</td>
      <td data-col="pb">${{r.pb}}</td>
      <td data-col="fm" title="${{r.fm}}">${{r.fm}}</td>
      <td data-col="ob" title="${{r.ob}}">${{r.ob}}</td>
      <td data-col="iv" class="num">${{fmtBRL(r.iv)}}</td>
      <td data-col="im" class="num">${{fmtNum(Math.round(r.im))}}</td>
      <td data-col="cl" class="num">${{fmtNum(Math.round(r.cl))}}</td>
      <td data-col="ctr" class="num">${{fmtPct(r.ctr)}}</td>
      <td data-col="cpc" class="num">${{fmtBRL(r.cpc)}}</td>
      <td data-col="lp" class="num">${{fmtNum(Math.round(r.lp))}}</td>
      <td data-col="cpl_p" class="num">${{r.cpl_p>0?fmtBRL(r.cpl_p):'—'}}</td>
      <td data-col="vv" class="num">${{r.vv>0?fmtNum(Math.round(r.vv)):'—'}}</td>
      <td data-col="insc" class="num">${{r.insc>0?fmtNum(Math.round(r.insc)):'—'}}</td>
    `;
    tbody.appendChild(tr);
  }});
  document.getElementById('table-info').textContent=rows.length+(HIDDEN_COLS.has('cj')?' campanhas':' conjuntos');
  document.querySelectorAll('#camp-table th').forEach(th=>{{
    th.classList.remove('sorted'); const si=th.querySelector('.sort-icon'); if(si) si.textContent='↕';
  }});
  const activeTh=document.querySelector(`#camp-table th[data-col="${{state.tableSortCol}}"]`);
  if (activeTh) {{ activeTh.classList.add('sorted'); const si=activeTh.querySelector('.sort-icon'); if(si) si.textContent=state.tableSortAsc?'↑':'↓'; }}
  applyHiddenCols();
  syncTopScrollWidth();
}}

function renderAll() {{
  computePeriodTotals();
  const data=filterData();
  updateKPIs(data); updateDonutInv(data); updateDonutPub(data);
  updateInvLine(data); updateLeadsLine(data); updateCampLine(data);
  updateMMLine(filterDataCatOnly());
  updateDualLine(data); updateShare(data);
  updatePubComp(data); updateDow(data); updateFmt(data);
  updateBubble(data); updateTopCamp(data); updateRank(data);
  updateTable(data);
}}

function initMultiselect(wrapId, stateKey, placeholder) {{
  const btn=document.getElementById(wrapId+'-btn');
  const dd=document.getElementById(wrapId+'-dd');
  const cbs=dd.querySelectorAll('input[type=checkbox]');
  const clr=dd.querySelector('.ms-clear');

  btn.addEventListener('click', e=>{{
    e.stopPropagation();
    const opening=!dd.classList.contains('open');
    document.querySelectorAll('.ms-dropdown.open').forEach(d=>d.classList.remove('open'));
    document.querySelectorAll('.ms-btn.open').forEach(b=>b.classList.remove('open'));
    if (opening) {{ dd.classList.add('open'); btn.classList.add('open'); }}
  }});

  function refresh() {{
    const sel=[...cbs].filter(c=>c.checked).map(c=>c.value);
    state[stateKey]=sel;
    const lbl=btn.querySelector('.ms-label');
    lbl.textContent=sel.length===0?placeholder:sel.length===1?sel[0]:sel.length+' sel.';
    renderAll();
  }}
  cbs.forEach(cb=>cb.addEventListener('change',refresh));
  if (clr) clr.addEventListener('click',e=>{{ e.stopPropagation(); cbs.forEach(c=>c.checked=false); refresh(); }});
}}

const COL_LABELS={{pl:'Plataforma',ct:'Conta',ca:'Campanha',cj:'Conjunto',pb:'Público',fm:'Formato',ob:'Objetivo',iv:'Investimento',im:'Impressões',cl:'Cliques',ctr:'CTR',cpc:'CPC',lp:'Conversões Plat',cpl_p:'CPL Plat',vv:'Visualizações',insc:'Inscrições'}};
const COL_DEFAULT_W={{pl:90,ct:120,ca:280,cj:220,pb:90,fm:130,ob:100,iv:130,im:110,cl:90,ctr:80,cpc:90,lp:120,cpl_p:100,vv:120,insc:100}};

function applyHiddenCols() {{
  document.querySelectorAll('#camp-table th, #camp-table td').forEach(c=>{{
    const col=c.getAttribute('data-col');
    c.style.display=(col && HIDDEN_COLS.has(col))?'none':'';
  }});
}}

function syncTopScrollWidth() {{
  const tbl=document.getElementById('camp-table');
  const inner=document.getElementById('top-scroll-inner');
  if (tbl && inner) inner.style.width=tbl.offsetWidth+'px';
}}

function initTableUI() {{
  document.querySelectorAll('#camp-table th').forEach(th=>{{
    const col=th.getAttribute('data-col');
    if (COL_DEFAULT_W[col]) th.style.width=COL_DEFAULT_W[col]+'px';
  }});
  const box=document.getElementById('col-toggle');
  Object.keys(COL_LABELS).forEach(col=>{{
    const lab=document.createElement('label');
    lab.innerHTML=`<input type="checkbox" data-col="${{col}}" ${{HIDDEN_COLS.has(col)?'':'checked'}}> ${{COL_LABELS[col]}}`;
    box.appendChild(lab);
  }});
  box.querySelectorAll('input').forEach(cb=>{{
    cb.addEventListener('change',()=>{{
      const col=cb.getAttribute('data-col');
      if (cb.checked) HIDDEN_COLS.delete(col); else HIDDEN_COLS.add(col);
      applyHiddenCols(); syncTopScrollWidth();
      if (col==='cj') updateTable(filterData());  // conjunto muda o agrupamento da tabela
    }});
  }});
  const top=document.getElementById('top-scroll'), wrap=document.getElementById('table-wrap');
  let lock=false;
  top.addEventListener('scroll',()=>{{ if(lock)return; lock=true; wrap.scrollLeft=top.scrollLeft; lock=false; }});
  wrap.addEventListener('scroll',()=>{{ if(lock)return; lock=true; top.scrollLeft=wrap.scrollLeft; lock=false; }});
  document.querySelectorAll('#camp-table th').forEach(th=>{{
    const handle=document.createElement('span');
    handle.className='col-resize';
    th.appendChild(handle);
    let startX, startW;
    handle.addEventListener('mousedown',e=>{{
      e.preventDefault(); e.stopPropagation();
      startX=e.pageX; startW=th.offsetWidth; th.dataset.resizing='1';
      const move=ev=>{{ const w=Math.max(50,startW+(ev.pageX-startX)); th.style.width=w+'px'; syncTopScrollWidth(); }};
      const up=()=>{{ document.removeEventListener('mousemove',move); document.removeEventListener('mouseup',up);
        setTimeout(()=>{{ delete th.dataset.resizing; }},0); }};
      document.addEventListener('mousemove',move); document.addEventListener('mouseup',up);
    }});
    handle.addEventListener('click',e=>e.stopPropagation());
  }});
  window.addEventListener('resize',syncTopScrollWidth);
  syncTopScrollWidth();
}}

function initEvents() {{
  initTableUI();
  initCampMetricBtns();
  initMMControls();
  initDualControls();
  document.getElementById('fil-periodo').value = state.periodo;
  document.getElementById('fil-date-start').value = state.dateStart;
  document.getElementById('fil-date-end').value = state.dateEnd;
  document.getElementById('date-custom').classList.toggle('visible', state.periodo==='custom');
  document.querySelectorAll('.gran-btn').forEach(btn=>{{
    btn.addEventListener('click',()=>{{
      document.querySelectorAll('.gran-btn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active'); state.gran=btn.dataset.gran; renderAll();
    }});
  }});

  document.getElementById('fil-periodo').addEventListener('change',e=>{{
    state.periodo=e.target.value;
    document.getElementById('date-custom').classList.toggle('visible',state.periodo==='custom');
    renderAll();
  }});
  document.getElementById('fil-date-start').addEventListener('change',e=>{{ state.dateStart=e.target.value; if(state.periodo==='custom') renderAll(); }});
  document.getElementById('fil-date-end').addEventListener('change',e=>{{ state.dateEnd=e.target.value; if(state.periodo==='custom') renderAll(); }});

  initMultiselect('ms-plat',  'plataforma', 'Todas');
  initMultiselect('ms-conta', 'conta',      'Todas');
  initMultiselect('ms-pub',   'publico',    'Todos');
  initMultiselect('ms-fmt',   'formato',    'Todos');
  {tax_init_js}

  let campTimer;
  document.getElementById('fil-camp-search').addEventListener('input',e=>{{
    clearTimeout(campTimer);
    campTimer=setTimeout(()=>{{ state.campSearch=e.target.value.trim(); renderAll(); }},250);
  }});
  let cjTimer;
  document.getElementById('fil-cj-search').addEventListener('input',e=>{{
    clearTimeout(cjTimer);
    cjTimer=setTimeout(()=>{{ state.cjSearch=e.target.value.trim(); renderAll(); }},250);
  }});

  document.querySelectorAll('#camp-table th[data-col]').forEach(th=>{{
    th.addEventListener('click',()=>{{
      if (th.dataset.resizing) return;
      const col=th.dataset.col;
      if (state.tableSortCol===col) state.tableSortAsc=!state.tableSortAsc;
      else {{ state.tableSortCol=col; state.tableSortAsc=false; }}
      updateTable(filterData());
    }});
  }});

  document.addEventListener('click',()=>{{
    document.querySelectorAll('.ms-dropdown.open').forEach(d=>d.classList.remove('open'));
    document.querySelectorAll('.ms-btn.open').forEach(b=>b.classList.remove('open'));
  }});
}}

document.addEventListener('DOMContentLoaded',()=>{{ initEvents(); renderAll(); }});
</script>
</body>
</html>"""

with open(HTML_OUT, 'w', encoding='utf-8') as f:
    f.write(HTML)

size_mb = os.path.getsize(HTML_OUT) / 1024 / 1024
print(f"\nHTML gerado: {HTML_OUT}")
print(f"Tamanho: {size_mb:.2f} MB")
print("Concluido!")
