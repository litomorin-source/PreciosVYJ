import base64
import hashlib
import io
import json
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title='Historial de precios', page_icon='📈', layout='wide')
ROOT = Path(__file__).parent
INDEX_PATH = 'data/index.json'


def secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


TOKEN = secret('GITHUB_TOKEN')
REPO = secret('GITHUB_REPO')
BRANCH = secret('GITHUB_BRANCH', 'main')
APP_PIN = str(secret('APP_PIN', '')).strip()
USE_GITHUB = bool(TOKEN and REPO)


def gh_headers():
    return {
        'Authorization': f'Bearer {TOKEN}',
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
    }


def gh_url(path):
    return f'https://api.github.com/repos/{REPO}/contents/{path}'


def load_bytes(path):
    if USE_GITHUB:
        r = requests.get(gh_url(path), headers=gh_headers(), params={'ref': BRANCH}, timeout=30)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return base64.b64decode(r.json()['content'])
    p = ROOT / path
    return p.read_bytes() if p.exists() else None


def save_bytes(path, content, message):
    if USE_GITHUB:
        current = requests.get(gh_url(path), headers=gh_headers(), params={'ref': BRANCH}, timeout=30)
        payload = {
            'message': message,
            'content': base64.b64encode(content).decode('utf-8'),
            'branch': BRANCH,
        }
        if current.status_code == 200:
            payload['sha'] = current.json()['sha']
        elif current.status_code != 404:
            current.raise_for_status()
        r = requests.put(gh_url(path), headers=gh_headers(), json=payload, timeout=45)
        r.raise_for_status()
    else:
        p = ROOT / path
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(content)


def load_index():
    raw = load_bytes(INDEX_PATH)
    return json.loads(raw.decode('utf-8')) if raw else {'version': 1, 'listas': []}


def save_index(index):
    save_bytes(INDEX_PATH, json.dumps(index, ensure_ascii=False, indent=2).encode(), 'Actualizar historial')


def require_admin():
    if not APP_PIN or st.session_state.get('admin_ok'):
        return True
    pin = st.text_input('PIN de administrador', type='password')
    if st.button('Ingresar', type='primary'):
        if pin == APP_PIN:
            st.session_state.admin_ok = True
            st.rerun()
        st.error('PIN incorrecto')
    return False


def read_excel(file_bytes, filename):
    engine = 'xlrd' if filename.lower().endswith('.xls') else 'openpyxl'
    return pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine)


def norm_text(value):
    return (str(value).strip().lower().replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u'))


def find_header(raw):
    aliases = {
        'codigo': {'codigo', 'cod', 'cod.'},
        'producto': {'producto', 'descripcion', 'articulo'},
        'precio': {'precio final', 'precio', 'precio venta', 'importe'},
    }
    for row_idx in range(min(20, len(raw))):
        vals = [norm_text(v) for v in raw.iloc[row_idx].tolist()]
        found = {}
        for key, names in aliases.items():
            for i, val in enumerate(vals):
                if val in names:
                    found[key] = i
                    break
        if len(found) == 3:
            return row_idx, found
    raise ValueError('No encontré las columnas Código, Producto y Precio final en las primeras 20 filas.')


def parse_price(value):
    if pd.isna(value) or value == '':
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace('$','').replace(' ','')
    if ',' in text and '.' in text:
        text = text.replace('.','').replace(',','.') if text.rfind(',') > text.rfind('.') else text.replace(',','')
    elif ',' in text:
        text = text.replace('.','').replace(',','.')
    return float(text)


def normalize_file(file_bytes, filename):
    raw = read_excel(file_bytes, filename)
    header_row, cols = find_header(raw)
    records = []
    for _, row in raw.iloc[header_row+1:].iterrows():
        code, product, price = row.iloc[cols['codigo']], row.iloc[cols['producto']], row.iloc[cols['precio']]
        if pd.isna(code) and pd.isna(product):
            continue
        code = str(code).strip()
        if code.endswith('.0'):
            code = code[:-2]
        try:
            price = parse_price(price)
        except Exception:
            continue
        if not code or code.lower() == 'nan' or price is None:
            continue
        records.append({'codigo': code, 'producto': str(product).strip(), 'precio': price})
    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError('No se encontraron productos válidos.')
    dup = df[df.duplicated('codigo', keep=False)].copy()
    df = df.drop_duplicates('codigo', keep='last').reset_index(drop=True)
    preview = ' '.join(str(v) for v in raw.head(10).fillna('').astype(str).values.flatten())
    return df, dup, preview


def detect_date(filename, preview):
    for source in (preview, filename):
        for m in re.finditer(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b', source):
            d, mo, y = map(int, m.groups())
            y = y + 2000 if y < 100 else y
            try:
                return date(y, mo, d)
            except ValueError:
                pass
    m = re.search(r'(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)', filename)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(2000+y, mo, d)
        except ValueError:
            pass
    return date.today()


def compare_lists(old, new):
    a = old.rename(columns={'producto':'producto_anterior','precio':'precio_anterior'})
    b = new.rename(columns={'producto':'producto_actual','precio':'precio_actual'})
    m = a.merge(b, on='codigo', how='outer', indicator=True)
    def kind(r):
        if r['_merge'] == 'right_only': return 'Agregado'
        if r['_merge'] == 'left_only': return 'Retirado'
        if r['precio_anterior'] == 0 and r['precio_actual'] > 0: return 'De $0 a precio'
        if r['precio_actual'] > r['precio_anterior']: return 'Aumento'
        if r['precio_actual'] < r['precio_anterior']: return 'Baja'
        if str(r['producto_anterior']) != str(r['producto_actual']): return 'Cambio de nombre'
        return 'Sin cambio'
    m['tipo'] = m.apply(kind, axis=1)
    m['diferencia'] = m['precio_actual'] - m['precio_anterior']
    m['variacion_pct'] = m.apply(lambda r: r['diferencia']/r['precio_anterior'] if pd.notna(r['precio_anterior']) and r['precio_anterior'] != 0 else None, axis=1)
    m['producto'] = m['producto_actual'].fillna(m['producto_anterior'])
    return m


def report_excel(comp, old_date, new_date):
    out = io.BytesIO()
    inc = comp[comp.tipo == 'Aumento']
    summary = pd.DataFrame([
        ['Lista anterior', old_date], ['Lista nueva', new_date],
        ['Productos lista anterior', comp.precio_anterior.notna().sum()],
        ['Productos lista nueva', comp.precio_actual.notna().sum()],
        ['Sin cambio', (comp.tipo=='Sin cambio').sum()], ['Aumentos', len(inc)],
        ['De $0 a precio', (comp.tipo=='De $0 a precio').sum()],
        ['Bajas', (comp.tipo=='Baja').sum()], ['Agregados', (comp.tipo=='Agregado').sum()],
        ['Retirados', (comp.tipo=='Retirado').sum()],
        ['Aumento promedio', inc.variacion_pct.mean() if len(inc) else 0],
    ], columns=['Dato','Resultado'])
    cols = ['codigo','producto','precio_anterior','precio_actual','diferencia','variacion_pct','tipo']
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        summary.to_excel(w, sheet_name='Resumen', index=False)
        comp[comp.tipo!='Sin cambio'][cols].to_excel(w, sheet_name='Todos los cambios', index=False)
        for kind, sheet in [('Agregado','Agregados'),('Retirado','Retirados'),('Cambio de nombre','Cambios de nombre')]:
            comp[comp.tipo==kind][cols].to_excel(w, sheet_name=sheet, index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            for c in ws[1]: c.font = c.font.copy(bold=True)
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col)+2, 65)
    return out.getvalue()


@st.cache_data(ttl=30)
def load_normalized(path):
    raw = load_bytes(path)
    if not raw: raise FileNotFoundError(path)
    return pd.read_csv(io.BytesIO(raw), dtype={'codigo': str})


def sorted_lists(index):
    return sorted(index.get('listas', []), key=lambda x: x['fecha'])


def by_date(index, iso):
    return next((x for x in index.get('listas',[]) if x['fecha']==iso), None)


def save_new(index, file_bytes, filename, selected_date, df, replace=False):
    iso = selected_date.isoformat()
    if by_date(index, iso) and not replace:
        raise ValueError('Ya existe una lista para esa fecha.')
    ext = Path(filename).suffix.lower()
    original_path = f'data/listas/lista_{iso}{ext}'
    normalized_path = f'data/normalizados/lista_{iso}.csv'
    save_bytes(original_path, file_bytes, f'Guardar lista {iso}')
    save_bytes(normalized_path, df.to_csv(index=False).encode(), f'Guardar normalizado {iso}')
    entry = {
        'fecha': iso, 'archivo_original': original_path, 'archivo_normalizado': normalized_path,
        'nombre_original': filename, 'productos': int(len(df)),
        'sha256': hashlib.sha256(file_bytes).hexdigest(), 'cargado_el': datetime.now().isoformat(timespec='seconds')
    }
    index['listas'] = [x for x in index.get('listas',[]) if x['fecha'] != iso] + [entry]
    index['listas'] = sorted_lists(index)
    save_index(index)
    st.cache_data.clear()



def render_comparison(comp, old_date, new_date, accumulated=False):
    old_label = datetime.fromisoformat(old_date).strftime('%d/%m/%Y')
    new_label = datetime.fromisoformat(new_date).strftime('%d/%m/%Y')
    title = f"{old_label} vs. {new_label}"
    if accumulated:
        title += " · Variación acumulada"
    st.subheader(title)

    inc = comp[comp.tipo == 'Aumento']
    metrics = [
        ('Sin cambios', (comp.tipo == 'Sin cambio').sum()),
        ('Aumentos', len(inc)),
        ('Bajas', (comp.tipo == 'Baja').sum()),
        ('Agregados', (comp.tipo == 'Agregado').sum()),
        ('Retirados', (comp.tipo == 'Retirado').sum()),
    ]
    for col, (label, val) in zip(st.columns(5), metrics):
        col.metric(label, int(val))

    zero_count = int((comp.tipo == 'De $0 a precio').sum())
    rename_count = int((comp.tipo == 'Cambio de nombre').sum())
    avg = inc.variacion_pct.mean() if len(inc) else 0
    med = inc.variacion_pct.median() if len(inc) else 0
    st.write(
        f"**Aumento promedio:** {avg:.2%} · "
        f"**Mediana:** {med:.2%} · "
        f"**De $0 a precio:** {zero_count} · "
        f"**Cambios de nombre:** {rename_count}"
    )

    a, b = st.columns(2)
    for target, heading, frame in [
        (a, 'Mayores aumentos', inc.nlargest(10, 'variacion_pct')),
        (b, 'Menores aumentos', inc.nsmallest(10, 'variacion_pct')),
    ]:
        with target:
            st.markdown(f'#### {heading}')
            if len(frame):
                st.dataframe(
                    frame[['codigo','producto','precio_anterior','precio_actual','variacion_pct']],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info('No hubo aumentos en esta comparación.')

    for heading, kind in [
        ('Productos que bajaron', 'Baja'),
        ('Productos agregados', 'Agregado'),
        ('Productos retirados', 'Retirado'),
        ('Pasaron de $0 a precio', 'De $0 a precio'),
        ('Cambios de nombre', 'Cambio de nombre'),
    ]:
        sub = comp[comp.tipo == kind]
        with st.expander(f'{heading} ({len(sub)})'):
            if len(sub):
                st.dataframe(sub, use_container_width=True, hide_index=True)
            else:
                st.caption('No hubo casos en esta comparación.')

    st.download_button(
        'Descargar Excel comparativo',
        report_excel(comp, old_date, new_date),
        f'comparativo_{old_date}_vs_{new_date}.xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f'download_{old_date}_{new_date}_{"acc" if accumulated else "last"}',
    )

st.title('📈 Historial de Listas de Precios')
st.caption('Carga, guarda y compara listas usando el código del producto.')
if USE_GITHUB:
    st.success(f'Guardado permanente conectado a GitHub: {REPO}')
else:
    st.warning('Modo local. En Streamlit Cloud configurá GITHUB_TOKEN y GITHUB_REPO en Secrets.')

index = load_index()
lists = sorted_lists(index)
tabs = st.tabs(['📤 Nueva lista','📊 Última comparación','🕘 Historial','🔎 Producto'])

with tabs[0]:
    if require_admin():
        uploaded = st.file_uploader('Subí la lista nueva', type=['xls','xlsx'])
        if uploaded:
            try:
                data = uploaded.getvalue()
                df, dup, preview = normalize_file(data, uploaded.name)
                detected = detect_date(uploaded.name, preview)
                c1, c2 = st.columns(2)
                c1.write(f'**Archivo:** {uploaded.name}')
                c1.write(f'**Productos detectados:** {len(df):,}'.replace(',','.'))
                selected = c2.date_input('¿Este archivo corresponde a esta fecha?', value=detected, format='DD/MM/YYYY')
                if not dup.empty:
                    st.warning(f'Encontré {dup.codigo.nunique()} códigos duplicados. Se conservará la última aparición.')
                    with st.expander('Ver duplicados'): st.dataframe(dup, use_container_width=True)
                existing = by_date(index, selected.isoformat())
                replace = False
                if existing:
                    st.error(f'Ya existe una lista del {selected.strftime("%d/%m/%Y")}.')
                    replace = st.checkbox('Reemplazar la lista existente')
                if lists and selected.isoformat() < lists[-1]['fecha'] and not existing:
                    st.warning('La fecha es anterior a la última lista. Se guardará respetando el orden cronológico.')
                if st.button('Confirmar, guardar y comparar', type='primary', disabled=bool(existing and not replace)):
                    with st.spinner('Guardando...'):
                        save_new(index, data, uploaded.name, selected, df, replace)
                    st.success('Lista guardada correctamente.')
                    st.rerun()
            except Exception as e:
                st.error(f'No pude procesar el archivo: {e}')

with tabs[1]:
    index = load_index(); lists = sorted_lists(index)
    if len(lists) < 2:
        st.info('Se necesitan al menos dos listas.')
    else:
        old_meta, new_meta = lists[-2], lists[-1]
        comp = compare_lists(
            load_normalized(old_meta['archivo_normalizado']),
            load_normalized(new_meta['archivo_normalizado'])
        )
        render_comparison(comp, old_meta['fecha'], new_meta['fecha'], accumulated=False)

with tabs[2]:
    index = load_index(); lists = sorted_lists(index)
    if not lists:
        st.info('Todavía no hay listas.')
    else:
        hist = pd.DataFrame(lists)
        hist['fecha'] = pd.to_datetime(hist['fecha']).dt.strftime('%d/%m/%Y')
        hist['cargado_el'] = pd.to_datetime(hist['cargado_el']).dt.strftime('%d/%m/%Y %H:%M')
        st.dataframe(
            hist[['fecha','productos','nombre_original','cargado_el']].rename(
                columns={
                    'fecha':'Fecha',
                    'productos':'Productos',
                    'nombre_original':'Archivo',
                    'cargado_el':'Cargado'
                }
            ),
            use_container_width=True,
            hide_index=True
        )
        st.markdown('#### Comparar dos fechas')
        dates = [x['fecha'] for x in lists]
        if len(dates) >= 2:
            c1, c2 = st.columns(2)
            old_date = c1.selectbox(
                'Lista anterior',
                dates[:-1],
                format_func=lambda x: datetime.fromisoformat(x).strftime('%d/%m/%Y')
            )
            valid = [x for x in dates if x > old_date]
            new_date = c2.selectbox(
                'Lista nueva',
                valid,
                format_func=lambda x: datetime.fromisoformat(x).strftime('%d/%m/%Y')
            )
            if st.button('Comparar fechas seleccionadas', type='primary'):
                st.session_state['history_comparison'] = (old_date, new_date)

            selected_pair = st.session_state.get('history_comparison')
            if selected_pair:
                selected_old, selected_new = selected_pair
                old_item = by_date(index, selected_old)
                new_item = by_date(index, selected_new)
                if old_item and new_item:
                    comp = compare_lists(
                        load_normalized(old_item['archivo_normalizado']),
                        load_normalized(new_item['archivo_normalizado'])
                    )
                    old_pos = dates.index(selected_old)
                    new_pos = dates.index(selected_new)
                    accumulated = (new_pos - old_pos) > 1
                    render_comparison(
                        comp, selected_old, selected_new, accumulated=accumulated
                    )

with tabs[3]:
    index = load_index(); lists = sorted_lists(index)
    if not lists: st.info('Todavía no hay productos.')
    else:
        latest = load_normalized(lists[-1]['archivo_normalizado'])
        q = st.text_input('Buscá por código o nombre')
        if q:
            matches = latest[latest.codigo.str.contains(q,case=False,na=False,regex=False) | latest.producto.str.contains(q,case=False,na=False,regex=False)].head(50)
            st.dataframe(matches, use_container_width=True, hide_index=True)
            if len(matches):
                code = st.selectbox('Elegí un producto', matches.codigo.tolist(), format_func=lambda x: f"{x} — {matches.loc[matches.codigo==x,'producto'].iloc[0]}")
                evo = []
                for item in lists:
                    df = load_normalized(item['archivo_normalizado'])
                    row = df[df.codigo==code]
                    if len(row): evo.append({'fecha':item['fecha'],'precio':float(row.iloc[0].precio),'producto':row.iloc[0].producto})
                evo = pd.DataFrame(evo)
                if len(evo):
                    evo['fecha'] = pd.to_datetime(evo.fecha); evo['variacion'] = evo.precio.pct_change()
                    st.line_chart(evo.set_index('fecha').precio)
                    st.dataframe(evo, use_container_width=True, hide_index=True)
