#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lee el PDF del mail diario de Looker Studio ("Google Ads Brum") y escribe
google_raw.json con el mismo formato que usa generate.py:

  {"by_campaign": [{"name", "spend", "conversions", "lost_budget_pct"}], "as_of": "YYYY-MM-DD",
   "source": "looker_pdf"}

Uso:  python3 -I parse_looker_pdf.py <ruta_pdf> <as_of YYYY-MM-DD> <salida_json>
Sale con código != 0 (y no escribe nada) si no puede leer al menos una campaña,
así la tarea diaria puede caer al respaldo (planilla puente).
"""
import json
import re
import subprocess
import sys

NUM = r'[\d,]+(?:\.\d+)?'
ROW = re.compile(
    r'^\s*\d+\.\s+(?P<name>.+?)\s+\$\s*(?P<cost>' + NUM + r')\s+(?P<impr>' + NUM + r')\s+(?P<clicks>' + NUM + r')'
    r'\s+(?P<ctr>' + NUM + r')%\s+(?P<conv>' + NUM + r'|-+)\s+(?:\$\s*(?P<cpa>' + NUM + r')|-+)'
    r'(?:\s+(?P<lost>' + NUM + r')%|\s+-+)?\s*$')
PAGER = re.compile(r'(\d+)-(\d+)\s*/\s*(\d+)')


def f(s):
    if s is None or set(s) <= {'-'}:
        return None
    return float(s.replace(',', ''))


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    pdf, as_of, out = sys.argv[1:]
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', as_of):
        print('as_of inválido:', as_of)
        return 2
    text = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
    rows = []
    for line in text.splitlines():
        m = ROW.match(line)
        if not m:
            continue
        rows.append({
            'name': m.group('name').strip(),
            'spend': f(m.group('cost')) or 0.0,
            'conversions': f(m.group('conv')) or 0.0,
            'lost_budget_pct': f(m.group('lost')),
        })
    if not rows:
        print('No se encontraron campañas en el PDF')
        return 1
    pager = PAGER.search(text)
    if pager and int(pager.group(2)) < int(pager.group(3)):
        print('AVISO: el PDF muestra solo {}-{} de {} campañas'.format(*pager.groups()))
    with open(out, 'w', encoding='utf-8') as fh:
        json.dump({'by_campaign': rows, 'as_of': as_of, 'source': 'looker_pdf'}, fh, ensure_ascii=False, indent=1)
    print('OK', len(rows), 'campañas, gasto total', round(sum(r['spend'] for r in rows)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
