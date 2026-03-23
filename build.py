import pandas as pd
import json
from pathlib import Path

def n(v):
    try: f=float(v); return 0 if pd.isna(f) else round(f,4)
    except: return 0

df_p = pd.read_excel('pipeline.xlsx', sheet_name='Pipeline', header=0)
df_p.columns = ['cliente','parceiro','status','valor','entrada','liquido','tcTotal','tcLiquida','tc','financiado','data_desembolso','prioridade']

pipeline = []
for _, r in df_p.iterrows():
    dd_raw = r['data_desembolso']
    dd = None
    dd_label = None
    if pd.notna(dd_raw):
        s = str(dd_raw).strip()
        if s in ['1S2026','2S2026','1S2025','2S2025']:
            dd_label = s
        else:
            try: dd = pd.to_datetime(dd_raw).strftime('%Y-%m-%d')
            except: pass
    pipeline.append({
        'cliente': str(r['cliente']).strip() if pd.notna(r['cliente']) else '—',
        'parceiro': str(r['parceiro']).strip() if pd.notna(r['parceiro']) else '—',
        'status': str(r['status']).strip().lower() if pd.notna(r['status']) else '',
        'valor': n(r['valor']), 'entrada': n(r['entrada']), 'liquido': n(r['liquido']),
        'tcTotal': n(r['tcTotal']), 'tcLiq': n(r['tcLiquida']), 'tc': n(r['tc']),
        'financiado': n(r['financiado']),
        'dataDesembolso': dd,
        'semestreLabel': dd_label,
        'prioridade': str(r['prioridade']).strip() if pd.notna(r['prioridade']) else '',
    })

df_c = pd.read_excel('pipeline.xlsx', sheet_name='Carteira', header=0)
df_c.columns = ['cliente','parceiro','localidade','data_emissao','valor','vp','taxa','prazo','garantias']

today = pd.Timestamp.now()
carteira = []
for _, r in df_c.iterrows():
    de = pd.to_datetime(r['data_emissao'])
    meses_ativos = int((today - de).days / 30)
    meses_rest = max(0, int(r['prazo']) - meses_ativos)
    vp = round(float(r['vp']),2) if pd.notna(r['vp']) else 0
    valor = round(float(r['valor']),2)
    carteira.append({
        'cliente': str(r['cliente']).strip(),
        'parceiro': str(r['parceiro']).strip(),
        'localidade': str(r['localidade']).strip(),
        'dataEmissao': de.strftime('%Y-%m-%d'),
        'valor': valor, 'vp': vp,
        'taxa': round(float(r['taxa']),4),
        'prazo': int(r['prazo']),
        'mesesAtivos': meses_ativos,
        'mesesRestantes': meses_rest,
    })

P = json.dumps(pipeline, ensure_ascii=False)
C = json.dumps(carteira, ensure_ascii=False)

template = Path('template.html').read_text(encoding='utf-8')
html = template.replace('__PIPELINE_DATA__', P).replace('__CARTEIRA_DATA__', C)
Path('index.html').write_text(html, encoding='utf-8')
print(f"✅ index.html gerado — {len(pipeline)} ops pipeline, {len(carteira)} ops carteira")
