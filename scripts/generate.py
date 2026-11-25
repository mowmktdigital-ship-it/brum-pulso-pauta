#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador estatico del "Pulso de Pauta BRUM" para Netlify.
"""
import json
import os
import sys
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))


def load(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return json.load(f)


def parse_num(s):
    if s is None:
        return None
    t = str(s).strip()
    if t == '' or 'DIV' in t:
        return None
    t = t.replace('$', '').replace(' ', '').replace('%', '')
    if t == '' or t == '-':
        return None
    t = t.replace('.', '').replace(',', '.')
    try:
        return float(t)
    except ValueError:
        return None


def build_meta_rows(raw):
    rows = []
    for e in raw:
        if e.get('effective_status') != 'ACTIVE':
            continue
        spend = float(e.get('amount_spent', {}).get('value') or 0)
        if not spend:
            continue
        indicator = (e.get('results') or {}).get('indicator') or ''
        if 'leadgen' in indicator or 'fb_pixel_lead' in indicator:
            resultado_tipo = 'leads'
        elif 'landing_page_view' in indicator:
            resultado_tipo = 'vistas_landing'
        elif 'messaging_conversation_started' in indicator:
            resultado_tipo = 'conversaciones'
        elif indicator == 'profile_visit_view':
            resultado_tipo = 'visitas_perfil'
        else:
            resultado_tipo = 'otro'
        resultado = None
        results = e.get('results') or {}
        if results.get('values'):
            resultado = float(results['values'][0]['value'])
        elif results.get('value') not in (None, 'Not available'):
            resultado = float(results['value'])
        cpl = None
        cpr = (e.get('cost_per_result') or {}).get('value')
        if cpr is not None and not (isinstance(cpr, str) and cpr.startswith('Not available')):
            cpl = float(cpr)
        rows.append({
            'campana': e['name'], 'estado': 'Activa', 'gasto': spend, 'objetivo': e.get('objective') or '',
            'resultado_tipo': resultado_tipo, 'resultado': resultado, 'cpl': cpl,
        })
    return rows


def build_google_rows(raw):
    rows = []
    for c in raw.get('by_campaign', []):
        conv = float(c.get('conversions') or 0)
        spend = float(c.get('spend') or 0)
        rows.append({
            'campana': c['name'], 'gasto': spend, 'conversiones': conv,
            'cpa': (spend / conv) if conv > 0 else None,
        })
    return rows


def build_excel_rows(values):
    data_rows = values[4:]
    rows = []
    for r in data_rows:
        def g(i):
            return r[i] if i < len(r) else None
        marca = g(0)
        if not marca:
            continue
        rows.append({
            'marca': marca,
            'objetivo_mes': parse_num(g(1)),
            'leads_reales': parse_num(g(2)),
            'proyectado_fin_mes': parse_num(g(6)),
            'ventas': parse_num(g(4)),
            'inv_total': parse_num(g(7)),
            'cpl_total': parse_num(g(8)),
            'google_pauta_leads': parse_num(g(20)),
            'google_pauta_inv': parse_num(g(21)),
            'meta_op_leads': parse_num(g(27)),
            'meta_op_inv': parse_num(g(28)),
            'meta_form_leads': parse_num(g(34)),
            'meta_form_inv': parse_num(g(35)),
            'wsp_leads': parse_num(g(44)),
            'wsp_inv': parse_num(g(45)),
        })
    return rows


# Definicion de canales para la tabla de "leads x canal": (nombre, idx_leads, idx_inversion,
# idx_ventas, idx_tasa_conv) sobre la fila TOTAL del Excel de Lu (misma hoja "INV x CONCE
# Proyeccion" que usa build_excel_rows). idx_inversion=None para canales organicos, que no
# tienen gasto de pauta asociado.
CHANNEL_DEFS = [
    ('Google Ads (pauta)', 20, 21, 24, 25),
    ('Meta Oportunidades', 27, 28, 31, 32),
    ('Meta Formulario', 34, 35, 38, 39),
    ('Filtrado de Leads / WhatsApp', 44, 45, 48, 49),
    ('Orgánico Google', 11, None, 12, 13),
    ('Orgánico Redes', 16, None, 17, 18),
    ('Otros orígenes', 40, None, 42, 43),
]


def build_channel_rows(excel_values):
    total = next((r for r in excel_values[4:] if r and r[0] == 'TOTAL'), None)
    if not total:
        return [], None

    def g(i):
        return total[i] if i is not None and i < len(total) else None

    rows = []
    for nombre, i_leads, i_inv, i_ventas, i_tasa in CHANNEL_DEFS:
        leads = parse_num(g(i_leads))
        inv = parse_num(g(i_inv)) if i_inv is not None else None
        ventas = parse_num(g(i_ventas))
        tasa = parse_num(g(i_tasa))
        cpl = (inv / leads) if (inv is not None and leads) else None
        rows.append({
            'canal': nombre, 'leads': leads, 'inversion': inv, 'cpl': cpl,
            'ventas': ventas, 'tasa_conv': tasa,
        })
    overall = {
        'canal': 'TOTAL (incluye orgánicos)',
        'leads': parse_num(g(2)), 'inversion': parse_num(g(7)), 'cpl': parse_num(g(8)),
        'ventas': parse_num(g(4)), 'tasa_conv': None,
    }
    if overall['ventas'] is not None and overall['leads']:
        overall['tasa_conv'] = overall['ventas'] / overall['leads'] * 100
    return rows, overall


def calc_account_summary(excel_rows, meta_rows, google_rows, now):
    total = next((r for r in excel_rows if r['marca'] == 'TOTAL'), None)
    gasto_meta = sum(r['gasto'] for r in meta_rows)
    gasto_google = sum(r['gasto'] for r in google_rows)
    gasto_total = gasto_meta + gasto_google
    inv_excel = (total or {}).get('inv_total') or 0
    gap = gasto_total - inv_excel
    days_in_month = (datetime.date(now.year + (now.month == 12), (now.month % 12) + 1, 1) - datetime.timedelta(days=1)).day
    objetivo_mes = (total or {}).get('objetivo_mes')
    leads_totales = (total or {}).get('leads_reales')
    proyeccion_fin_mes = (total or {}).get('proyectado_fin_mes')
    return {
        'periodo_fin': now.strftime('%Y-%m-%d'),
        'dias_transcurridos': now.day,
        'dias_totales': days_in_month,
        'leads_totales': leads_totales,
        'ventas_totales': (total or {}).get('ventas'),
        'inversion_excel': inv_excel,
        'cpl_promedio': (total or {}).get('cpl_total'),
        'gasto_meta': gasto_meta,
        'gasto_google': gasto_google,
        'gasto_real_total': gasto_total,
        'gap_total': gap,
        'gap_pct': (gap / gasto_total * 100) if gasto_total else 0,
        'objetivo_mes': objetivo_mes,
        'avance_pct': (leads_totales / objetivo_mes * 100) if objetivo_mes and leads_totales is not None else None,
        'proyeccion_fin_mes': proyeccion_fin_mes,
        'proyeccion_pct': (proyeccion_fin_mes / objetivo_mes * 100) if objetivo_mes and proyeccion_fin_mes is not None else None,
    }


# CORREGIDO 20/11 segun aclaracion directa de Agos: Campana 1 (Prioritarios) y Campana 2
# (Generica BYD + no prioritarios) son un trabajo aparte que se le hace directamente a la
# terminal BYD, no a los concesionarios Brum. No corresponden a este tablero en absoluto:
# ni a los reportes de Meta, ni a la inversion total, ni al cruce con el Excel de Lu (nunca
# van a aparecer ahi). Se excluyen por completo de meta_rows en main() antes de cualquier
# calculo.
# 25/11 (Agos): se suma Campaña 3 — Escala Seal 5, que también es del negocio BYD terminal.
EXCLUDED_CAMPAIGN_NAMES = ['Campaña 1 — Prioritarios', 'Campaña 2 — Genérica BYD + no prioritarios',
                           'Campaña 3 — Escala Seal 5']
FORMULARIO_NAMES = ['Meta Formulario Por Marca']
TRAFICO_NAMES = ['Tráfico al Sitio', 'Orgánicas - Tráfico IG Perfil | Test', 'Remarketing']
FILTRADO_NAMES = ['Campañas Filtrado de Leads']


def _es_oportunidades_real(nombre):
    # Campanas "Leads <concesionario>" = la categoria real de Meta Oportunidades del Excel.
    return nombre.strip().lower().startswith('leads ')


def calc_excel_gap(excel_rows, meta_rows, google_rows):
    total = next((r for r in excel_rows if r['marca'] == 'TOTAL'), None)
    if not total:
        return []
    b_op = b_form = b_traf = b_filt = b_otros = 0.0
    for r in meta_rows:
        g = r['gasto']
        nombre = r['campana']
        if _es_oportunidades_real(nombre):
            b_op += g
        elif nombre in FORMULARIO_NAMES:
            b_form += g
        elif nombre in TRAFICO_NAMES:
            b_traf += g
        elif nombre in FILTRADO_NAMES:
            b_filt += g
        else:
            b_otros += g
    gasto_google = sum(r['gasto'] for r in google_rows)

    comps = [
        {'componente': 'Meta Oportunidades (Leads + concesionario) no cargado', 'monto': b_op - (total.get('meta_op_inv') or 0),
         'explicacion': 'Campanas "Leads + nombre del concesionario" (la categoria real de Oportunidades). Hoy estan todas pausadas en Meta, asi que el gasto de plataforma en esta categoria es minimo o nulo - coherente con lo que carga el Excel. Si se reactivan, esta fila va a mostrar gasto real de nuevo.'},
        {'componente': 'Campañas de tráfico/engagement sin bucket en el Excel', 'monto': b_traf,
         'explicacion': 'Trafico al Sitio + Organicas IG + Remarketing no estan pensadas para generar leads trackeables asi que el Excel no les asigna bucket. Esperable - no es un error.'},
        {'componente': 'Meta Formulario — desfasaje de carga', 'monto': b_form - (total.get('meta_form_inv') or 0),
         'explicacion': 'El Excel va levemente atras del gasto real de Meta Formulario Por Marca (lag normal de uno o dos dias de carga).'},
        {'componente': 'Google Ads — el Excel usa un CPL fijo', 'monto': gasto_google - (total.get('google_pauta_inv') or 0),
         'explicacion': 'El Excel calcula la inversion de Google como leads reales multiplicados por un CPL fijo en vez de usar el gasto real de la plataforma. Es un artefacto de esa formula - no una medicion real de desfasaje.'},
        {'componente': 'Filtrado de Leads/WhatsApp — desfasaje menor', 'monto': b_filt - (total.get('wsp_inv') or 0),
         'explicacion': 'Mismo lag de carga que Meta Formulario - montos chicos.'},
    ]
    if abs(b_otros) > 1:
        comps.append({'componente': 'Campañas Meta sin clasificar', 'monto': b_otros,
                       'explicacion': 'Campana nueva, renombrada o pausada que este tablero todavia no clasifica en ningun bucket. Avisar para mapearla.'})
    gap_total = sum(c['monto'] for c in comps)
    for c in comps:
        c['pct_gap'] = (c['monto'] / gap_total * 100) if gap_total else 0
    return comps


def calc_dealer_cpl(excel_rows):
    total = next((r for r in excel_rows if r['marca'] == 'TOTAL'), None)
    if not total:
        return []
    inv_excel_total = total.get('inv_total') or 0
    avg_cpl = total.get('cpl_total')
    rows = []
    for r in excel_rows:
        if r['marca'] == 'TOTAL':
            continue
        inv = r.get('inv_total') or 0
        if inv <= 0:
            continue
        leads = r.get('leads_reales') or 0
        cpl = (inv / leads) if leads > 0 else None
        ratio = (cpl / avg_cpl) if (cpl is not None and avg_cpl) else None
        if ratio is None:
            estado = 'Sin datos'
        elif ratio > 2.0:
            estado = 'Revisar prioritario'
        elif ratio > 1.3:
            estado = 'Revisar'
        elif ratio > 1.0:
            estado = 'A vigilar'
        else:
            estado = 'OK'
        rows.append({
            'marca': r['marca'], 'inversion': inv,
            'pct_total': (inv / inv_excel_total * 100) if inv_excel_total else None,
            'cpl': cpl, 'vs_promedio': ratio, 'ventas': r.get('ventas'), 'estado': estado,
        })
    rows.sort(key=lambda r: -r['inversion'])
    return rows[:10]


def calc_alerts(meta_rows, google_rows):
    out = []
    for r in meta_rows:
        gasto = r['gasto']
        resultado = r['resultado']
        if gasto > 0 and (resultado is None or resultado == 0):
            out.append({'plataforma': 'Meta Ads', 'campana': r['campana'], 'gasto': gasto})
    for r in google_rows:
        if r['gasto'] > 0 and (r['conversiones'] or 0) == 0:
            out.append({'plataforma': 'Google Ads', 'campana': r['campana'], 'gasto': r['gasto']})
    out.sort(key=lambda r: -r['gasto'])
    return out


def money(v):
    if v is None:
        return '—'
    return '$' + format(round(v), ',').replace(',', '.')


def pct(v, dec=1):
    if v is None:
        return '—'
    return (f'{v:,.{dec}f}%').replace(',', '_').replace('.', ',').replace('_', '.')


def num(v):
    if v is None:
        return '—'
    return format(round(v), ',').replace(',', '.')


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


RESULTADO_LABELS = {
    'leads': 'Leads', 'vistas_landing': 'Vistas de landing',
    'conversaciones': 'Conversaciones iniciadas', 'visitas_perfil': 'Visitas de perfil',
    'otro': 'Otro resultado',
}


def badge(estado):
    e = estado or ''
    if 'prioritario' in e.lower():
        cls = 'bad'
    elif e.lower().startswith('revisar') or 'vigilar' in e.lower():
        cls = 'warn'
    else:
        cls = 'ok'
    return f'<span class="badge {cls}">{esc(estado)}</span>'


# ---------------------------------------------------------------------------
# SEÑALES DE CUENTA (agregado 25/11). Tres indicadores con semáforo arriba de
# todo + comparación contra el valor de hace una semana (history.json).
# Archivos opcionales: resumen_values.json (pestaña "Resumen" del Excel de Lu),
# presupuesto.json ({"YYYY-MM": monto}) e history.json. Si faltan, la señal
# correspondiente muestra "Sin dato" en vez de abortar.
# ---------------------------------------------------------------------------
MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
         'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

# Umbrales de semáforo (fijos)
PROY_VERDE, PROY_AMARILLO = 100.0, 80.0          # % del objetivo proyectado
CPL_AMARILLO_MAX = 15.0                          # % arriba del mes anterior
GASTO_VERDE = (90.0, 105.0)                      # % del presupuesto proyectado
GASTO_AMARILLO = (80.0, 115.0)
ATRASADO_PCT = 60.0                              # concesionario "atrasado"


def load_optional(name, default):
    p = os.path.join(HERE, name)
    if not os.path.exists(p):
        return default
    try:
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    except (ValueError, OSError):
        return default


def resumen_total_paid(values, month_idx):
    """Devuelve {'inv','leads','cpl'} de la fila 'Total Paid' del bloque del mes
    (month_idx 0-11) en la pestaña Resumen, o None si no está."""
    if not values:
        return None
    target = MESES[month_idx]
    in_block = False
    for r in values:
        label = (r[1] if len(r) > 1 else '').strip().upper()
        if label in MESES:
            in_block = (label == target)
            continue
        if in_block and label == 'TOTAL PAID':
            cpl = parse_num(r[4] if len(r) > 4 else None)
            if cpl is None:
                return None
            return {'inv': parse_num(r[2] if len(r) > 2 else None),
                    'leads': parse_num(r[3] if len(r) > 3 else None), 'cpl': cpl}
    return None


def history_week_ago(history, today):
    """Entrada de hace ~7 días (6 a 8), del mismo mes. Prefiere exactamente 7."""
    best = None
    for h in history:
        try:
            d = datetime.date.fromisoformat(h['fecha'])
        except (KeyError, ValueError):
            continue
        diff = (today - d).days
        if 6 <= diff <= 8 and d.month == today.month and d.year == today.year:
            if best is None or abs(diff - 7) < abs(best[0] - 7):
                best = (diff, h)
    return best[1] if best else None


def money_m(v):
    if v is None:
        return '—'
    if abs(v) >= 1_000_000:
        return '$' + f'{v / 1_000_000:.1f}'.replace('.', ',') + 'M'
    return money(v)


def calc_signals(excel_rows, summary, google_raw, resumen_values, presupuesto, now, ritmo_dias=(None, None)):
    total = next((r for r in excel_rows if r['marca'] == 'TOTAL'), None) or {}
    sig = {}

    # 1 · Proyección de leads vs objetivo
    proy_pct = summary['proyeccion_pct']
    con_obj = [r for r in excel_rows if r['marca'] != 'TOTAL' and (r.get('objetivo_mes') or 0) > 0]
    atrasados = [r for r in con_obj if ((r.get('proyectado_fin_mes') or 0) / r['objetivo_mes'] * 100) < ATRASADO_PCT]
    if proy_pct is None:
        estado = 'nodata'
    elif proy_pct >= PROY_VERDE:
        estado = 'ok'
    elif proy_pct >= PROY_AMARILLO:
        estado = 'warn'
    else:
        estado = 'bad'
    sig['proy'] = {
        'valor': proy_pct, 'estado': estado,
        'contexto': '{} de {} concesionarios con objetivo proyectan menos del {}% de su meta'.format(
            len(atrasados), len(con_obj), int(ATRASADO_PCT)) if con_obj else '',
    }
    sig['proy']['detalle'] = '{} leads proyectados a fin de mes, objetivo {}'.format(
        num(summary['proyeccion_fin_mes']), num(summary['objetivo_mes'])) if proy_pct is not None else 'Sin datos del Excel de Lu'

    # 2 · Costo por lead de la pauta vs mes anterior (pestaña Resumen de Lu)
    cur = resumen_total_paid(resumen_values, now.month - 1)
    prev_idx = (now.month - 2) % 12
    prev = resumen_total_paid(resumen_values, prev_idx)
    cpl = cur['cpl'] if cur else None
    var = ((cpl / prev['cpl'] - 1) * 100) if (cpl is not None and prev and prev['cpl']) else None
    if cpl is None:
        estado = 'nodata'
    elif var is None:
        estado = 'nometa'
    elif var <= 0:
        estado = 'ok'
    elif var <= CPL_AMARILLO_MAX:
        estado = 'warn'
    else:
        estado = 'bad'
    if cpl is None:
        det = 'No se encontró el bloque del mes en la pestaña Resumen del Excel de Lu'
    elif prev:
        det ='{} vs. {} ({})'.format(('+' if var > 0 else '') + pct(var, 0), MESES[prev_idx].lower(), money(prev['cpl']))
    else:
        det = 'Sin dato del mes anterior para comparar'
    sig['cpl'] = {
        'valor': cpl, 'estado': estado, 'detalle': det,
        'contexto': '{} leads pagos · inversión {}'.format(num(cur['leads']), money_m(cur['inv'])) if cur else '',
    }

    # 3 · Ritmo de gasto vs presupuesto del mes
    dias = summary['dias_totales']
    gasto_proy = None
    if not google_raw.get('unavailable'):
        meta_proj = summary['gasto_meta'] / now.day * dias
        g_day = now.day
        try:
            asof = datetime.date.fromisoformat(google_raw.get('as_of', ''))
            if asof.year == now.year and asof.month == now.month:
                g_day = asof.day
        except ValueError:
            pass
        gasto_proy = meta_proj + summary['gasto_google'] / g_day * dias
    pres = (presupuesto or {}).get(now.strftime('%Y-%m'))
    gasto_pct = (gasto_proy / pres * 100) if (gasto_proy is not None and pres) else None
    if gasto_proy is None:
        estado, det = 'nodata', 'Google Ads no disponible hoy: no se puede proyectar el gasto total'
    elif gasto_pct is None:
        estado, det = 'nodata', 'Falta el presupuesto del mes para comparar'
    else:
        lo, hi = GASTO_VERDE
        alo, ahi = GASTO_AMARILLO
        estado = 'ok' if lo <= gasto_pct <= hi else ('warn' if alo <= gasto_pct <= ahi else 'bad')
        det = '{} del presupuesto ({})'.format(pct(gasto_pct, 0), money_m(pres))
    sig['gasto'] = {
        'valor': gasto_proy, 'pct': gasto_pct, 'estado': estado, 'detalle': det,
        'contexto': 'Gasto real Meta + Google al día {} de {}, proyectado a fin de mes'.format(now.day, dias),
    }

    # 4 · Presupuesto extra estimado para cumplir el objetivo (sobre la proyección)
    objetivo = summary['objetivo_mes']
    proyectado = summary['proyeccion_fin_mes']
    if objetivo is None or proyectado is None or cpl is None:
        sig['extra'] = {'valor': None, 'estado': 'nodata', 'detalle': 'Falta la proyección o el costo por lead',
                        'contexto': ''}
    else:
        faltan = objetivo - proyectado
        if faltan <= 0:
            sig['extra'] = {'valor': 0.0, 'estado': 'ok', 'texto': 'No hace falta',
                            'detalle': 'La proyección ya supera el objetivo por {} leads'.format(num(-faltan)),
                            'contexto': 'Estimado sobre la proyección del Excel de Lu.'}
        else:
            sig['extra'] = {
                'valor': faltan * cpl, 'estado': sig['proy']['estado'],
                'detalle': 'Faltan {} leads (proyectado {} vs. objetivo {}) × {} por lead'.format(
                    num(faltan), num(proyectado), num(objetivo), money(cpl)),
                'contexto': 'Estimado: supone que los leads extra cuestan lo mismo que los actuales. Al subir la inversión el costo por lead suele subir, así que es un piso.',
            }

    # 5 · Ritmo diario de leads: actual vs. necesario para llegar al objetivo
    leads = summary['leads_totales']
    d_trans, d_rest = ritmo_dias
    if leads is None or objetivo is None or not d_trans:
        sig['ritmo'] = {'valor': None, 'estado': 'nodata', 'detalle': 'Sin datos de leads del Excel de Lu', 'contexto': ''}
    else:
        actual = leads / d_trans
        if d_rest and d_rest > 0:
            necesario = max(objetivo - leads, 0) / d_rest
            det = 'Harían falta {} por día en los {} días que quedan'.format(num(necesario), d_rest)
        else:
            det = 'Último día del mes'
        sig['ritmo'] = {'valor': actual, 'estado': sig['proy']['estado'], 'detalle': det,
                        'contexto': '{} leads reales en {} días (Excel de Lu)'.format(num(leads), d_trans)}
    return sig


ESTADO_LABELS = {'ok': 'Cumple', 'warn': 'Atención', 'bad': 'No cumple',
                 'nodata': 'Sin dato', 'nometa': 'Sin comparación'}


def render_signals(sig, week_ago):
    def delta(key, fmt, higher_is_good):
        if not week_ago or week_ago.get(key) is None or sig_val(key) is None:
            return '<span class="sig-delta">Hace una semana: —</span>'
        now_v, old_v = sig_val(key), week_ago[key]
        if abs(now_v - old_v) < 1e-9:
            arrow, cls = '=', ''
        else:
            up = now_v > old_v
            arrow = '↑' if up else '↓'
            cls = '' if higher_is_good is None else ('good' if up == higher_is_good else 'bad')
        return '<span class="sig-delta {cls}">{a} hace una semana: {v}</span>'.format(cls=cls, a=arrow, v=fmt(old_v))

    def sig_val(key):
        return {'proy_pct': sig['proy']['valor'], 'cpl_pauta': sig['cpl']['valor'],
                'gasto_proy': sig['gasto']['valor'], 'extra': sig['extra']['valor'],
                'ritmo': sig['ritmo']['valor']}[key]

    cards = [
        ('¿Llegamos al objetivo de leads?', pct(sig['proy']['valor'], 0) if sig['proy']['valor'] is not None else '—',
         'proyectado', sig['proy'], delta('proy_pct', lambda v: pct(v, 0), True)),
        ('¿El lead pago está caro o barato?', money(sig['cpl']['valor']) if sig['cpl']['valor'] is not None else '—',
         'costo por lead de la pauta', sig['cpl'], delta('cpl_pauta', money, False)),
        ('¿Vamos bien de ritmo de gasto?', money_m(sig['gasto']['valor']),
         'gasto proyectado a fin de mes', sig['gasto'], delta('gasto_proy', money_m, None)),
        ('Presupuesto extra para cumplir el objetivo · estimado',
         sig['extra'].get('texto') or money_m(sig['extra']['valor']),
         'más en el mes' if not sig['extra'].get('texto') and sig['extra']['valor'] else '',
         sig['extra'], delta('extra', money_m, False)),
        ('¿A qué ritmo entran los leads?', num(sig['ritmo']['valor']) if sig['ritmo']['valor'] is not None else '—',
         'leads por día', sig['ritmo'], delta('ritmo', num, True)),
    ]
    out = ''
    for titulo, valor, unidad, s, d in cards:
        out += (
            '<div class="sig-card {estado}"><div class="sig-head"><span class="sig-title">{titulo}</span>'
            '<span class="badge {estado}">{label}</span></div>'
            '<div class="sig-value">{valor} <span class="sig-unit">{unidad}</span></div>'
            '<div class="sig-detail">{detalle}</div>{delta}'
            '<div class="sig-context">{contexto}</div></div>'
        ).format(estado=s['estado'], titulo=esc(titulo), label=ESTADO_LABELS[s['estado']], valor=esc(valor),
                 unidad=esc(unidad), detalle=esc(s['detalle']), delta=d, contexto=esc(s.get('contexto', '')))
    return out


def update_history(history, sig, today):
    entry = {'fecha': today.isoformat(), 'proy_pct': sig['proy']['valor'],
             'cpl_pauta': sig['cpl']['valor'], 'gasto_proy': sig['gasto']['valor'],
             'extra': sig['extra']['valor'], 'ritmo': sig['ritmo']['valor']}
    history = [h for h in history if h.get('fecha') != entry['fecha']] + [entry]
    history.sort(key=lambda h: h.get('fecha', ''))
    return history[-120:]


def main():
    meta_raw = load('meta_raw.json')
    google_raw = load('google_raw.json')
    excel_values = load('excel_values.json')

    meta_rows = build_meta_rows(meta_raw)
    meta_rows = [r for r in meta_rows if r['campana'] not in EXCLUDED_CAMPAIGN_NAMES]
    google_rows = build_google_rows(google_raw)
    excel_rows = build_excel_rows(excel_values)

    now = datetime.datetime.now()
    summary = calc_account_summary(excel_rows, meta_rows, google_rows, now)
    gap_rows = calc_excel_gap(excel_rows, meta_rows, google_rows)
    dealer_rows = calc_dealer_cpl(excel_rows)
    alert_rows = calc_alerts(meta_rows, google_rows)
    channel_rows, channel_total = build_channel_rows(excel_values)

    google_unavailable = bool(google_raw.get('unavailable'))
    google_stale = (not google_unavailable) and google_raw.get('as_of') != now.strftime('%Y-%m-%d')

    resumen_values = load_optional('resumen_values.json', [])
    presupuesto = load_optional('presupuesto.json', {})
    history = load_optional('history.json', [])
    try:
        ritmo_dias = (int(parse_num(excel_values[1][7])), int(parse_num(excel_values[1][8])))
    except (IndexError, TypeError, ValueError):
        ritmo_dias = (now.day, summary['dias_totales'] - now.day)
    signals = calc_signals(excel_rows, summary, google_raw, resumen_values, presupuesto, now, ritmo_dias)
    signals_html = render_signals(signals, history_week_ago(history, now.date()))
    history = update_history(history, signals, now.date())
    with open(os.path.join(HERE, 'history.json'), 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=1)

    meta_tr = ''
    grupos = [
        ('Campañas de leads', [r for r in meta_rows if r['objetivo'] == 'OUTCOME_LEADS']),
        ('Otras campañas (tráfico, interacción, mensajes)', [r for r in meta_rows if r['objetivo'] != 'OUTCOME_LEADS']),
    ]
    for titulo_grupo, filas in grupos:
        if not filas:
            continue
        meta_tr += ('<tr class="group-row"><td colspan="2">{t}</td><td class="num">{g}</td>'
                    '<td colspan="3"></td></tr>').format(t=esc(titulo_grupo), g=money(sum(r['gasto'] for r in filas)))
        for r in sorted(filas, key=lambda r: -r['gasto']):
            meta_tr += (
                '<tr><td>{campana}</td><td>{estado}</td><td class="num">{gasto}</td>'
                '<td>{rtipo}</td><td class="num">{resultado}</td><td class="num">{cpl}</td></tr>'
            ).format(
                campana=esc(r['campana']), estado=esc(r['estado']), gasto=money(r['gasto']),
                rtipo=esc(RESULTADO_LABELS.get(r['resultado_tipo'], 'Otro resultado')),
                resultado=num(r['resultado']) if r['resultado'] is not None else '—',
                cpl=money(r['cpl']) if r['cpl'] is not None else '—',
            )
    if not meta_rows:
        meta_tr = '<tr><td colspan="6">—</td></tr>'

    google_tr = ''
    for r in sorted(google_rows, key=lambda r: -r['gasto']):
        google_tr += (
            '<tr><td>{campana}</td><td class="num">{gasto}</td><td class="num">{conv}</td><td class="num">{cpa}</td></tr>'
        ).format(
            campana=esc(r['campana']), gasto=money(r['gasto']),
            conv=f"{r['conversiones']:.2f}".replace('.', ','),
            cpa=money(r['cpa']) if r['cpa'] is not None else '—',
        )
    if not google_rows:
        google_tr = '<tr><td colspan="4">—</td></tr>'

    gap_tr = ''
    max_gap_abs = max([abs(c['monto']) for c in gap_rows] + [1])
    gap_chart = ''
    for c in gap_rows:
        gap_tr += (
            '<tr><td class="wrap">{comp}</td><td class="num">{monto}</td><td class="num">{pctg}</td>'
            '<td class="wrap">{expl}</td></tr>'
        ).format(comp=esc(c['componente']), monto=money(c['monto']), pctg=pct(c['pct_gap']),
                  expl=esc(c['explicacion']))
        color = 'var(--color-bad)' if c['monto'] >= 0 else 'var(--color-ok)'
        gap_chart += (
            '<div class="hbar-row"><span class="hbar-label">{label}</span>'
            '<div class="bar-track"><div class="bar-fill" style="width:{w:.1f}%;background:{color}"></div></div>'
            '<span class="hbar-value">{val}</span></div>'
        ).format(label=esc(c['componente']), w=(abs(c['monto']) / max_gap_abs * 100), color=color, val=money(c['monto']))
    if not gap_rows:
        gap_tr = '<tr><td colspan="4">—</td></tr>'

    dealer_tr = ''
    max_cpl = max([d['cpl'] for d in dealer_rows if d['cpl']] + [1])
    avg_cpl = summary['cpl_promedio'] or 0
    dealer_chart = ''
    for d in dealer_rows:
        dealer_tr += (
            '<tr><td>{marca}</td><td class="num">{inv}</td><td class="num">{pcttot}</td>'
            '<td class="num">{cpl}</td><td class="num">{ratio}</td><td class="num">{ventas}</td><td>{estado}</td></tr>'
        ).format(
            marca=esc(d['marca']), inv=money(d['inversion']), pcttot=pct(d['pct_total']),
            cpl=money(d['cpl']) if d['cpl'] is not None else '—',
            ratio=(f"{d['vs_promedio']:.2f}x".replace('.', ',')) if d['vs_promedio'] is not None else '—',
            ventas=num(d['ventas']) if d['ventas'] is not None else '—', estado=badge(d['estado']),
        )
        cls_color = {'OK': 'var(--color-ok)', 'A vigilar': 'var(--color-warn)',
                     'Revisar': 'var(--color-warn)', 'Revisar prioritario': 'var(--color-bad)'}.get(d['estado'], 'var(--color-fg-muted)')
        cpl_val = d['cpl'] or 0
        dealer_chart += (
            '<div class="hbar-row"><span class="hbar-label">{label}</span>'
            '<div class="bar-track"><div class="bar-fill" style="width:{w:.1f}%;background:{color}"></div></div>'
            '<span class="hbar-value">{val}</span></div>'
        ).format(label=esc(d['marca']), w=(cpl_val / max_cpl * 100), color=cls_color, val=money(d['cpl']) if d['cpl'] else '—')
    if not dealer_rows:
        dealer_tr = '<tr><td colspan="7">—</td></tr>'

    channel_tr = ''
    for c in channel_rows:
        channel_tr += (
            '<tr><td>{canal}</td><td class="num">{leads}</td><td class="num">{inv}</td>'
            '<td class="num">{cpl}</td><td class="num">{ventas}</td><td class="num">{tasa}</td></tr>'
        ).format(
            canal=esc(c['canal']), leads=num(c['leads']) if c['leads'] is not None else '—',
            inv=money(c['inversion']) if c['inversion'] is not None else '—',
            cpl=money(c['cpl']) if c['cpl'] is not None else '—',
            ventas=num(c['ventas']) if c['ventas'] is not None else '—',
            tasa=pct(c['tasa_conv'], 2) if c['tasa_conv'] is not None else '—',
        )
    if channel_total:
        channel_tr += (
            '<tr style="font-weight:600"><td>{canal}</td><td class="num">{leads}</td><td class="num">{inv}</td>'
            '<td class="num">{cpl}</td><td class="num">{ventas}</td><td class="num">{tasa}</td></tr>'
        ).format(
            canal=esc(channel_total['canal']),
            leads=num(channel_total['leads']) if channel_total['leads'] is not None else '—',
            inv=money(channel_total['inversion']) if channel_total['inversion'] is not None else '—',
            cpl=money(channel_total['cpl']) if channel_total['cpl'] is not None else '—',
            ventas=num(channel_total['ventas']) if channel_total['ventas'] is not None else '—',
            tasa=pct(channel_total['tasa_conv'], 2) if channel_total['tasa_conv'] is not None else '—',
        )
    if not channel_tr:
        channel_tr = '<tr><td colspan="6">—</td></tr>'

    alert_tr = ''
    if alert_rows:
        for a in alert_rows:
            alert_tr += '<tr><td>{plat}</td><td>{camp}</td><td class="num">{gasto}</td></tr>'.format(
                plat=esc(a['plataforma']), camp=esc(a['campana']), gasto=money(a['gasto']))
    else:
        alert_tr = '<tr><td colspan="3" class="alert-empty">Ninguna campaña activa gastó sin generar resultados. 👍</td></tr>'

    leads = summary['leads_totales'] or 0
    objetivo = summary['objetivo_mes'] or 0
    avance_pct = summary['avance_pct']
    proyeccion = summary['proyeccion_fin_mes']
    proyeccion_pct = summary['proyeccion_pct']
    bar_fill_pct = min(100, avance_pct) if avance_pct is not None else 0
    bar_proj_pct = min(100, proyeccion_pct) if proyeccion_pct is not None else 0

    periodo_label = f"01/{now.month:02d} al {now.day:02d}/{now.month:02d}/{now.year}"
    gap_pct_label = f"{summary['gap_pct']:.1f}".replace('.', ',') + '% del gasto real'
    headline = 'El gasto real de plataforma es {} mayor que lo que el Excel de Lu tiene cargado.'.format(money(summary['gap_total'])) if summary['gap_total'] >= 0 else 'El Excel de Lu tiene cargado {} más que el gasto real de plataforma.'.format(money(abs(summary['gap_total'])))
    cruce_resumen = (
        'Gasto real de plataforma: {real} · Excel de Lu: {excel} · Diferencia: {gap} ({pctg})'
    ).format(real=money(summary['gasto_real_total']), excel=money(summary['inversion_excel']),
              gap=money(summary['gap_total']), pctg=gap_pct_label)

    google_warning_html = ''
    if google_unavailable:
        google_warning_html = (
            '<p class="note warn-note">⚠️ Google Ads no disponible en este snapshot ({reason}). '
            'Excluido de los totales de gasto y del cruce con el Excel de hoy — no es gasto real $0, '
            'es que no se pudo consultar. Se vuelve a intentar mañana automáticamente.</p>'
        ).format(reason=esc(google_raw.get('reason', 'error de conexión')))
    elif google_stale:
        google_warning_html = (
            '<p class="note warn-note">⚠️ Google Ads: mostrando el último dato disponible del '
            '{as_of}. En cuanto se pueda volver a consultar, se actualiza solo.</p>'
        ).format(as_of=esc(google_raw.get('as_of', '?')))

    updated_at = now.strftime('%d/%m/%Y %H:%M')

    html = TEMPLATE.format(
        updated_at=updated_at,
        periodo_label=periodo_label,
        dia=now.day,
        dias_totales=summary['dias_totales'],
        headline=esc(headline),
        gasto_meta=money(summary['gasto_meta']),
        gasto_google=money(summary['gasto_google']),
        gasto_total=money(summary['gasto_real_total']),
        gap_total=money(summary['gap_total']),
        gap_pct_sub=gap_pct_label,
        google_warning_html=google_warning_html,
        avance_leads=num(leads), avance_objetivo=num(objetivo),
        avance_pct=pct(avance_pct) if avance_pct is not None else '—',
        avance_proyeccion=num(proyeccion) if proyeccion is not None else '—',
        avance_proyeccion_pct=pct(proyeccion_pct) if proyeccion_pct is not None else '—',
        bar_fill_pct=bar_fill_pct, bar_proj_pct=bar_proj_pct,
        alert_rows=alert_tr,
        meta_rows=meta_tr,
        google_rows=google_tr,
        cruce_resumen=esc(cruce_resumen),
        gap_chart=gap_chart,
        gap_rows=gap_tr,
        dealer_chart=dealer_chart,
        avg_cpl_label=money(avg_cpl),
        dealer_rows=dealer_tr,
        channel_rows=channel_tr,
        signals_html=signals_html,
    )

    out_path = os.path.join(HERE, 'output.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print('Wrote', out_path, len(html), 'bytes')


TEMPLATE = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pulso de Pauta BRUM</title>
<style>
  :root {{
    --color-fg: #1a1a18; --color-fg-muted: #6b6b63; --color-bg: #ffffff; --color-panel: #f7f7f3;
    --color-border-line: #e6e4dd; --color-ok: #2f7a3c; --color-warn: #b5750f; --color-bad: #b93c3c;
    --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
    --font-serif: Georgia, "Times New Roman", serif;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:not([data-theme="light"]) {{
      --color-fg: #ece9e2; --color-fg-muted: #9c9a92; --color-bg: #15140f; --color-panel: #1d1c16;
      --color-border-line: #332f23; --color-ok: #6fc37c; --color-warn: #e2a640; --color-bad: #e2726a;
    }}
  }}
  :root[data-theme="dark"] {{
    --color-fg: #ece9e2; --color-fg-muted: #9c9a92; --color-bg: #15140f; --color-panel: #1d1c16;
    --color-border-line: #332f23; --color-ok: #6fc37c; --color-warn: #e2a640; --color-bad: #e2726a;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; background: var(--color-bg); color: var(--color-fg); font-family: var(--font-sans); }}
  .wrap {{ max-width: 980px; margin: 0 auto; padding: 32px 16px 64px; }}
  h1, h2 {{ font-family: var(--font-serif); font-weight: 400; }}
  .eyebrow {{ font-size: 0.75em; letter-spacing: 1.5px; text-transform: uppercase; color: var(--color-fg-muted); margin: 0 0 6px; }}
  h1 {{ font-size: 1.9em; margin: 0 0 10px; }}
  .lead {{ font-size: 1.15em; margin: 0 0 8px; }}
  .note {{ color: var(--color-fg-muted); font-size: 0.85em; margin: 4px 0; max-width: 60em; }}
  .note.warn-note {{ color: var(--color-warn); }}
  header {{ margin-bottom: 32px; }}
  section {{ margin-bottom: 36px; }}
  h2 {{ font-size: 1.3em; margin: 0 0 8px; }}
  .section-note {{ color: var(--color-fg-muted); font-size: 0.9em; margin: 0 0 14px; max-width: 60em; }}

  .kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(13em, 1fr)); gap: 14px; }}
  .kpi-card {{ background: var(--color-panel); border: 1px solid var(--color-border-line); border-radius: 10px; padding: 16px; display: flex; flex-direction: column; gap: 4px; }}
  .kpi-card .kpi-label {{ font-size: 0.8em; color: var(--color-fg-muted); }}
  .kpi-card .kpi-value {{ font-family: var(--font-serif); font-size: 1.9em; font-variant-numeric: tabular-nums; line-height: 1.1; }}
  .kpi-card .kpi-sub {{ font-size: 0.8em; color: var(--color-fg-muted); }}
  .kpi-card.bad .kpi-value {{ color: var(--color-bad); }}

  .dash-table-wrap {{ overflow-x: auto; }}
  table.pulso-table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; font-variant-numeric: tabular-nums; }}
  table.pulso-table th, table.pulso-table td {{ text-align: left; padding: 8px 10px; border-bottom: 1px solid var(--color-border-line); white-space: nowrap; }}
  table.pulso-table th {{ color: var(--color-fg-muted); font-weight: 600; font-size: 0.75em; text-transform: uppercase; letter-spacing: 0.5px; }}
  table.pulso-table td.num, table.pulso-table th.num {{ text-align: right; }}
  table.pulso-table td.wrap, table.pulso-table th.wrap {{ white-space: normal; max-width: 28em; }}

  .badge {{ display: inline-block; padding: 0.15em 0.6em; border-radius: 999px; font-size: 0.78em; font-weight: 600; }}
  .badge.ok {{ background: color-mix(in srgb, var(--color-ok) 18%, transparent); color: var(--color-ok); }}
  .badge.warn {{ background: color-mix(in srgb, var(--color-warn) 18%, transparent); color: var(--color-warn); }}
  .badge.bad {{ background: color-mix(in srgb, var(--color-bad) 18%, transparent); color: var(--color-bad); }}

  .hbar-row {{ display: grid; grid-template-columns: 16em 1fr 7em; gap: 10px; align-items: center; margin-bottom: 8px; font-size: 0.85em; }}
  .hbar-label {{ color: var(--color-fg-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .hbar-value {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .bar-track {{ position: relative; height: 10px; border-radius: 999px; background: var(--color-border-line); overflow: visible; }}
  .bar-fill {{ position: absolute; left: 0; top: 0; height: 100%; border-radius: 999px; }}
  .chart-box {{ background: var(--color-panel); border: 1px solid var(--color-border-line); border-radius: 10px; padding: 16px; margin-bottom: 14px; }}
  .chart-box .avg-label {{ font-size: 0.75em; color: var(--color-fg-muted); margin-top: 6px; }}

  details.methodology {{ background: var(--color-panel); border: 1px solid var(--color-border-line); border-radius: 10px; padding: 16px; }}
  details.methodology summary {{ cursor: pointer; font-weight: 600; }}
  details.methodology ul {{ margin: 10px 0 0; padding-left: 1.2em; }}
  details.methodology li {{ margin-bottom: 6px; color: var(--color-fg-muted); font-size: 0.9em; }}

  .legend {{ display: flex; flex-wrap: wrap; gap: 8px 20px; margin-top: 10px; font-size: 0.8em; color: var(--color-fg-muted); align-items: baseline; }}
  .legend .legend-item {{ display: inline-flex; align-items: baseline; gap: 6px; }}

  .goal-card {{ background: var(--color-panel); border: 1px solid var(--color-border-line); border-radius: 10px; padding: 16px; display: flex; flex-direction: column; gap: 10px; }}
  .goal-row {{ display: flex; justify-content: space-between; align-items: baseline; font-size: 0.9em; gap: 14px; flex-wrap: wrap; }}
  .goal-value {{ font-variant-numeric: tabular-nums; font-weight: 600; }}
  .goal-bar-track {{ position: relative; height: 10px; border-radius: 999px; background: var(--color-border-line); overflow: visible; }}
  .goal-bar-fill {{ position: absolute; left: 0; top: 0; height: 100%; border-radius: 999px; background: var(--color-ok); }}
  .goal-bar-proj {{ position: absolute; top: -3px; width: 2px; height: 16px; background: var(--color-fg-muted); }}
  .goal-sub {{ font-size: 0.8em; color: var(--color-fg-muted); }}
  .alert-empty {{ color: var(--color-fg-muted); font-size: 0.9em; font-style: italic; }}

  .signals {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(17em, 1fr)); gap: 14px; }}
  .sig-card {{ background: var(--color-panel); border: 1px solid var(--color-border-line); border-left-width: 4px; border-radius: 10px; padding: 16px; display: flex; flex-direction: column; gap: 6px; }}
  .sig-card.ok {{ border-left-color: var(--color-ok); }}
  .sig-card.warn {{ border-left-color: var(--color-warn); }}
  .sig-card.bad {{ border-left-color: var(--color-bad); }}
  .sig-card.nodata, .sig-card.nometa {{ border-left-color: var(--color-border-line); border-left-style: dashed; }}
  .sig-head {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }}
  .sig-title {{ font-size: 0.85em; font-weight: 600; }}
  .sig-value {{ font-family: var(--font-serif); font-size: 2.2em; line-height: 1.1; font-variant-numeric: tabular-nums; }}
  .sig-unit {{ font-family: var(--font-sans); font-size: 0.4em; color: var(--color-fg-muted); }}
  .sig-detail {{ font-size: 0.85em; }}
  .sig-delta {{ font-size: 0.8em; color: var(--color-fg-muted); font-variant-numeric: tabular-nums; }}
  .sig-delta.good {{ color: var(--color-ok); }}
  .sig-delta.bad {{ color: var(--color-bad); }}
  .sig-context {{ font-size: 0.78em; color: var(--color-fg-muted); border-top: 1px solid var(--color-border-line); padding-top: 6px; margin-top: 2px; }}
  table.pulso-table tr.group-row td {{ font-weight: 600; font-size: 0.8em; text-transform: uppercase; letter-spacing: 0.5px; color: var(--color-fg-muted); background: var(--color-panel); }}
  .badge.nodata, .badge.nometa {{ background: transparent; color: var(--color-fg-muted); border: 1px dashed var(--color-fg-muted); }}

  footer {{ margin-top: 40px; color: var(--color-fg-muted); font-size: 0.78em; }}
  @media (max-width: 640px) {{
    .hbar-row {{ grid-template-columns: 8em 1fr 5.5em; }}
    table.pulso-table th, table.pulso-table td {{ padding: 6px 7px; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <p class="eyebrow">BRUM — Pauta · Pulso (snapshot diario)</p>
    <h1>Pulso de Pauta</h1>
    <p class="lead">{headline}</p>
    <p class="note">
      Mes en curso: <b>{periodo_label}</b> (día {dia} de {dias_totales}).
      Esta página se regenera automáticamente una vez por día con datos de Meta Ads, Google Ads y el Excel de Lu.
      El diagnóstico narrativo, las recomendaciones y el seguimiento de pedidos siguen exclusivamente en el PDF semanal.
    </p>
    <p class="note">Datos actualizados: <b>{updated_at}</b></p>
    {google_warning_html}
  </header>

  <section id="senales">
    <h2>Cómo viene la cuenta</h2>
    <div class="signals">{signals_html}</div>
  </section>

  <section class="kpis" id="resumen">
    <div class="kpi-card">
      <span class="kpi-label">Gasto Meta Ads</span>
      <span class="kpi-value">{gasto_meta}</span>
      <span class="kpi-sub">Campañas activas, mes en curso</span>
    </div>
    <div class="kpi-card">
      <span class="kpi-label">Gasto Google Ads</span>
      <span class="kpi-value">{gasto_google}</span>
      <span class="kpi-sub">Campañas Search, mes en curso</span>
    </div>
    <div class="kpi-card">
      <span class="kpi-label">Gasto real total</span>
      <span class="kpi-value">{gasto_total}</span>
      <span class="kpi-sub">Meta + Google, plataforma</span>
    </div>
    <div class="kpi-card bad">
      <span class="kpi-label">Gap vs. Excel de Lu</span>
      <span class="kpi-value">{gap_total}</span>
      <span class="kpi-sub">{gap_pct_sub}</span>
    </div>
  </section>

  <section id="objetivo">
    <h2>Avance hacia el objetivo del mes</h2>
    <div class="goal-card">
      <div class="goal-row">
        <span>Leads reales: <span class="goal-value">{avance_leads}</span> de <span class="goal-value">{avance_objetivo}</span> ({avance_pct})</span>
        <span class="goal-sub">Proyección a fin de mes: <span class="goal-value">{avance_proyeccion}</span> ({avance_proyeccion_pct} del objetivo)</span>
      </div>
      <div class="goal-bar-track">
        <div class="goal-bar-fill" style="width:{bar_fill_pct:.1f}%"></div>
        <div class="goal-bar-proj" style="left:{bar_proj_pct:.1f}%"></div>
      </div>
    </div>
  </section>

  <section id="alertas">
    <h2>Campañas con gasto y cero resultados</h2>
    <p class="section-note">Campañas activas que gastaron en el periodo pero no sumaron ningún resultado ni conversión. Es un hecho, no un diagnóstico de causa.</p>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th>Plataforma</th><th>Campaña</th><th class="num">Gasto</th></tr></thead>
        <tbody>{alert_rows}</tbody>
      </table>
    </div>
  </section>

  <section id="meta">
    <h2>Meta Ads por campaña</h2>
    <p class="section-note">Gasto y costo por resultado de cada campaña activa, mes en curso, agrupadas según el objetivo de la campaña en Meta: primero las de leads, después el resto. No incluye las Campañas 1, 2 y 3 (BYD terminal).</p>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th>Campaña</th><th>Estado</th><th class="num">Gasto</th><th>Resultado</th><th class="num">Cantidad</th><th class="num">Costo / resultado</th></tr></thead>
        <tbody>{meta_rows}</tbody>
      </table>
    </div>
  </section>

  <section id="google">
    <h2>Google Ads por campaña</h2>
    <p class="section-note">Campañas Search activas de la cuenta, mes en curso.</p>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th>Campaña</th><th class="num">Gasto</th><th class="num">Conversiones</th><th class="num">CPA</th></tr></thead>
        <tbody>{google_rows}</tbody>
      </table>
    </div>
  </section>

  <section id="cruce">
    <h2>Cruce con el Excel de Lu</h2>
    <p class="section-note">{cruce_resumen}</p>
    <div class="chart-box">{gap_chart}</div>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th class="wrap">Componente del gap</th><th class="num">Monto</th><th class="num">% del gap</th><th class="wrap">Explicación</th></tr></thead>
        <tbody>{gap_rows}</tbody>
      </table>
    </div>
  </section>

  <section id="canales">
    <h2>Leads por canal: CPL promedio y conversión a venta</h2>
    <p class="section-note">Leads reales, inversión, costo por lead y tasa de conversión a venta de cada canal, según el Excel de Lu (fila TOTAL), mes en curso. Los canales orgánicos no tienen gasto de pauta asociado, por eso no muestran CPL. No incluye las Campañas 1, 2 y 3 (BYD terminal), que son un negocio aparte y no están en este Excel.</p>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th>Canal</th><th class="num">Leads</th><th class="num">Inversión</th><th class="num">CPL</th><th class="num">Ventas</th><th class="num">Tasa de conv. a venta</th></tr></thead>
        <tbody>{channel_rows}</tbody>
      </table>
    </div>
  </section>

  <section id="dealers">
    <h2>CPL por concesionario vs. promedio de cuenta</h2>
    <p class="section-note">Top 10 marcas por inversión en el Excel, mes en curso. El promedio de cuenta surge de dividir la inversión total pauta por los leads reales totales del mes (incluye leads orgánicos, que diluyen el promedio).</p>
    <div class="chart-box">
      {dealer_chart}
      <div class="avg-label">Línea de promedio de cuenta: {avg_cpl_label}</div>
    </div>
    <div class="dash-table-wrap">
      <table class="pulso-table">
        <thead><tr><th>Marca</th><th class="num">Inversión</th><th class="num">% del total</th><th class="num">CPL</th><th class="num">vs. promedio</th><th class="num">Ventas</th><th>Estado</th></tr></thead>
        <tbody>{dealer_rows}</tbody>
      </table>
    </div>
    <div class="legend">
      <span class="legend-item"><span class="badge ok">OK</span> CPL igual o por debajo del promedio de cuenta (1x o menos)</span>
      <span class="legend-item"><span class="badge warn">A vigilar</span> Entre 1x y 1,3x el promedio</span>
      <span class="legend-item"><span class="badge warn">Revisar</span> Entre 1,3x y 2x el promedio</span>
      <span class="legend-item"><span class="badge bad">Revisar prioritario</span> Más de 2x el promedio</span>
    </div>
  </section>

  <details class="methodology">
    <summary>Cómo se construye este pulso</summary>
    <ul>
      <li>Señales de cuenta: (1) proyección de leads a fin de mes ÷ objetivo, del Excel de Lu — verde desde 100%, amarillo entre 80% y 100%, rojo debajo de 80%; el contexto cuenta los concesionarios con objetivo que proyectan menos del 60% de su meta. (2) Costo por lead de la pauta del mes (fila "Total Paid" de la pestaña Resumen de Lu) contra el mes anterior — verde si es igual o más barato, amarillo hasta 15% más caro, rojo si sube más. (3) Gasto real Meta + Google proyectado a fin de mes contra el presupuesto del mes — verde entre 90% y 105%, amarillo entre 80% y 115%, rojo fuera de eso; sin presupuesto cargado muestra "Sin dato". (4) Presupuesto extra estimado: leads que faltan entre la proyección a fin de mes y el objetivo × costo por lead actual de la pauta; se recalcula cada día sobre la proyección, sin esperar al cierre. Es un piso: supone que los leads extra cuestan lo mismo que los actuales. (5) Ritmo diario: leads reales ÷ días transcurridos, contra los leads por día que harían falta en los días restantes (días según el Excel de Lu); toma el mismo color que la señal 1. La comparación "hace una semana" usa el valor guardado 7 días antes dentro del mismo mes.</li>
      <li>Meta Ads y Google Ads: gasto y resultados reales desde el día 1 del mes en curso hasta hoy. Esta página se regenera una vez por día (no en vivo, a diferencia de la versión que vive en Claude).</li>
      <li>Excel de Lu: fila TOTAL y detalle por concesionario de la hoja "INV x CONCE Proyección". Si Lu arma una pestaña nueva para el próximo mes con otro nombre, hay que avisarle a Claude para actualizar el rango.</li>
      <li>Campaña 1 (Prioritarios), Campaña 2 (Genérica BYD + no prioritarios) y Campaña 3 (Escala Seal 5, sumada el 25/11) son un trabajo aparte que se le hace directamente a la terminal BYD, no a los concesionarios Brum. Se excluyen por completo de este tablero — de la tabla de Meta Ads, de la inversión total y de cualquier cruce — porque nunca van a aparecer en el Excel de Lu (corregido 20/11 por aclaración directa de Agos).</li>
      <li>El gap se reparte en varios componentes con la misma metodología del informe semanal. Meta Oportunidades = campañas "Leads + nombre del concesionario" (corregido el 18/11 por aclaración directa de Lu). Las demás campañas de Meta se mapean por su nombre exacto a un bucket; si Meta cambia o agrega campañas, puede aparecer una fila "Campañas Meta sin clasificar": es la señal de que hay que revisar el mapeo.</li>
      <li>El CPL por concesionario usa inversión reportada en el Excel ÷ leads reales reportados en el Excel, igual que en el informe semanal. El semáforo es puramente por CPL vs. promedio: más de 2x es "Revisar prioritario", entre 1,3x y 2x es "Revisar", entre 1x y 1,3x es "A vigilar", 1x o menos es "OK".</li>
      <li>Leads por canal: toma directamente las columnas que el Excel de Lu ya calcula por canal (Google Pauta, Meta Oportunidades, Meta Formulario, Filtrado de Leads/WhatsApp, y los dos orgánicos). El CPL de cada canal es inversión ÷ leads reales de ese canal; la tasa de conversión a venta es ventas ÷ leads reales de ese canal, tal cual las carga Lu. La fila TOTAL incluye los leads orgánicos, que diluyen el CPL y la tasa de conversión promedio de cuenta.</li>
      <li>El avance hacia el objetivo del mes y la proyección a fin de mes usan las columnas que el propio Excel de Lu ya calcula (objetivo, reales y proyectados a fin de mes).</li>
      <li>La alerta de campañas con gasto y cero resultados es un hecho objetivo (gasto &gt; 0 y resultado = 0), sin interpretar la causa.</li>
      <li>Esta vista es solo números: diagnóstico, recomendaciones y seguimiento de pedidos siguen en el PDF semanal, no acá.</li>
    </ul>
  </details>

  <footer>Pulso de Pauta BRUM · snapshot generado automáticamente por MOW Digital · no requiere inicio de sesión</footer>
</div>
</body>
</html>
"""

if __name__ == '__main__':
    main()
