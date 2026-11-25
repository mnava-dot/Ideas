"""Asigna gestor a la base AXA (pestañas OFICINAS y TIENDAS) y genera data/axa.json.

Uso:
  python3 tools/build_axa.py Seguimiento_AXA.xlsx data/municipios.topo.json data/inmuebles.json data/axa.json [salida.xlsx]

Criterio de asignación, en este orden:
  1. REGLAS_AXA: asignaciones fijas por estado o municipio (clave INEGI) que da el usuario.
  2. municipio: el gestor que ya atiende ese municipio en la cartera BBVA (el de más inmuebles).
  3. estado: si en ese estado hay un solo gestor en BBVA.
  4. cercania: estado con varios gestores -> el gestor del municipio BBVA más cercano dentro del estado.
  5. PENDIENTE: estado sin gestor.

Si se da salida.xlsx, copia el Excel original agregando la columna "GESTOR 2027" al final de cada
tabla (edita el XML directamente para conservar listas desplegables, formatos y tablas).
"""
import openpyxl, json, unicodedata, collections, sys, math, re, zipfile, uuid, warnings
warnings.filterwarnings('ignore')

def N(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().upper()
    for a in '.,-"()\n': s = s.replace(a, ' ')
    return ' '.join(s.split())

STATES = {'BAJA CALIFORNIA NORTE': 'BAJA CALIFORNIA', 'EDO MEXICO': 'MEXICO', 'ESTADO DE MEXICO': 'MEXICO',
          'PUEBLA SONATA MAGESTIC': 'PUEBLA', 'COAHUILA': 'COAHUILA DE ZARAGOZA',
          'MICHOACAN': 'MICHOACAN DE OCAMPO', 'VERACRUZ': 'VERACRUZ DE IGNACIO DE LA LLAVE'}
# ciudad del Excel -> municipio oficial (se revisó contra la dirección)
CITY = {'CIUDAD JUAREZ': 'JUAREZ', 'PACHUCA': 'PACHUCA DE SOTO', 'LOS MOCHIS': 'AHOME', 'CIUDAD OBREGON': 'CAJEME',
        'TLANEPANTLA': 'TLALNEPANTLA DE BAZ', 'TLANEPANTLA SANTA MONICA': 'TLALNEPANTLA DE BAZ',
        'TOLUCA': 'LERMA',  # la dirección es Parque Industrial Lerma
        'CHOLULA': 'SAN ANDRES CHOLULA', 'PUEBLA HUEXOCHITLA': 'PUEBLA', 'ECATEPEC': 'ECATEPEC DE MORELOS',
        'NAUCALPAN': 'NAUCALPAN DE JUAREZ', 'QUERETARO JURIQUILLA': 'QUERETARO',
        'TLAQUEPAQUE PLAZA RIO NILO': 'SAN PEDRO TLAQUEPAQUE', 'TLAJOMULCO': 'TLAJOMULCO DE ZUNIGA',
        'VILLAHERMOSA': 'CENTRO', 'ESCOBEDO': 'GENERAL ESCOBEDO'}
# Cuautitlán: según la dirección puede ser Cuautitlán o Cuautitlán Izcalli
def city_fix(st, city, addr):
    c = N(city)
    if c == 'CUAUTITLAN' and 'IZCALLI' in N(addr): return 'CUAUTITLAN IZCALLI'
    return CITY.get(c, c)

# Asignaciones fijas del usuario. Clave: 2 dígitos = estado, 5 dígitos = municipio.
REGLAS_AXA = {
    '07': 'NOE',       # Chiapas
    '27': 'NOE',       # Tabasco
    '04': 'NOE',       # Campeche
    '31050': 'MERCEDES',  # Mérida, Yucatán
    '17': 'EDGAR',     # Morelos
    '10': 'CATA',      # Durango
}

def main(xlsx, topo_path, bbva_path, out_path, xlsx_out=None):
    topo = json.load(open(topo_path, encoding='utf-8'))
    P = {g['properties']['id']: g['properties'] for g in topo['objects']['mun']['geometries']}
    byst = collections.defaultdict(dict); stname = {}
    for p in P.values():
        byst[N(p['s'])][N(p['n'])] = p['id']; stname[p['e']] = p['s']
    bb = json.load(open(bbva_path, encoding='utf-8')); C = {c: i for i, c in enumerate(bb['cols'])}
    mun_g = collections.defaultdict(collections.Counter); edo_g = collections.defaultdict(collections.Counter)
    for r in bb['rows']:
        cve, g = r[C['cve']], r[C['gestor']]
        if cve and g: mun_g[cve][g] += 1; edo_g[cve[:2]][g] += 1
    def km(a, b):
        pa, pb = P[a], P[b]; t = math.pi / 180
        h = math.sin((pb['cy'] - pa['cy']) * t / 2) ** 2 + math.cos(pa['cy'] * t) * math.cos(pb['cy'] * t) * math.sin((pb['cx'] - pa['cx']) * t / 2) ** 2
        return 2 * 6371 * math.asin(math.sqrt(h))
    def assign(cve):
        if cve in REGLAS_AXA: return REGLAS_AXA[cve], 'regla'
        if cve[:2] in REGLAS_AXA: return REGLAS_AXA[cve[:2]], 'regla'
        if cve in mun_g: return mun_g[cve].most_common(1)[0][0], 'municipio'
        gs = edo_g.get(cve[:2])
        if not gs: return 'PENDIENTE', 'sin gestor en el estado'
        if len(gs) == 1: return next(iter(gs)), 'estado'
        near = min((m for m in mun_g if m[:2] == cve[:2]), key=lambda m: km(cve, m))
        return mun_g[near].most_common(1)[0][0], f'cercania ({P[near]["n"]}, {round(km(cve, near))} km)'

    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    rows, per_sheet, report = [], {}, []
    for sheet, tipo in [('OFICINAS', 'OFICINA'), ('TIENDAS', 'TIENDA')]:
        per_sheet[sheet] = {}
        for i, r in enumerate(wb[sheet].iter_rows(values_only=True), start=1):
            if i < 3 or r[0] is None: continue
            st = N(r[2]); st = STATES.get(st, st)
            mu = city_fix(st, r[3], r[4])
            cve = byst.get(st, {}).get(mu, '')
            g, how = assign(cve) if cve else ('PENDIENTE', 'municipio no reconocido')
            per_sheet[sheet][i] = g
            report.append((sheet, i, r[0], str(r[2]).strip(), str(r[3]).strip().replace('\n', ' '), P[cve]['n'] if cve else '?', g, how))
            modelo = str(r[7 if sheet == 'OFICINAS' else 6] or '').upper()
            rows.append([f'AXA-{tipo[:2]}-{r[0]}', f'AXA {tipo.title()} {r[0]} · {str(r[3]).strip().splitlines()[0]}', tipo, 'AXA ' + sheet,
                         str(r[5] or '').strip().upper(), str(r[0]), g, f'Territorial {r[1]} · modelo {modelo.lower()}' if r[1] else '',
                         cve, str(r[4] or '').strip()[:220], modelo, str(r[1] or ''), str(r[2]).strip(), str(r[3]).strip(), '', 'AXA', how])
    cols = bb['cols'] + ['cliente', 'criterio']
    json.dump({'cols': cols, 'rows': rows, 'fuente': 'Seguimiento AXA · OFICINAS y TIENDAS'}, open(out_path, 'w', encoding='utf-8'),
              ensure_ascii=False, separators=(',', ':'))
    for x in report: print('\t'.join(map(str, x)))
    if xlsx_out: write_xlsx(xlsx, xlsx_out, per_sheet)

def col_idx(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(): n = n * 26 + ord(ch) - 64
    return n
def col_name(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def xesc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def write_xlsx(src, dst, per_sheet):
    zin = zipfile.ZipFile(src)
    wbx = zin.read('xl/workbook.xml').decode('utf-8')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    sheet_file = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
        sheet_file[name] = 'xl/' + tgt.group(1).lstrip('/').replace('xl/', '')
    edits = {}
    for sheet, vals in per_sheet.items():
        sf = sheet_file[sheet]
        srels = zin.read(sf.replace('worksheets/', 'worksheets/_rels/') + '.rels').decode('utf-8')
        tfile = 'xl/' + re.search(r'Target="\.\./(tables/table\d+\.xml)"', srels).group(1)
        tx = zin.read(tfile).decode('utf-8')
        ref = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tx)
        last = col_idx(ref.group(3)); newc = col_name(last + 1); hdr_row = int(ref.group(2)); end_row = int(ref.group(4))
        newref = f'{ref.group(1)}{hdr_row}:{newc}{end_row}'
        tx = tx.replace(f' ref="{ref.group(0)[6:-1]}"', f' ref="{newref}"')
        cnt = int(re.search(r'<tableColumns count="(\d+)"', tx).group(1))
        maxid = max(int(x) for x in re.findall(r'<tableColumn id="(\d+)"', tx))
        tx = tx.replace(f'<tableColumns count="{cnt}"', f'<tableColumns count="{cnt + 1}"')
        tx = tx.replace('</tableColumns>', f'<tableColumn id="{maxid + 1}" xr3:uid="{{{str(uuid.uuid4()).upper()}}}" name="GESTOR 2027"/></tableColumns>')
        edits[tfile] = tx
        sx = zin.read(sf).decode('utf-8')
        def fix_row(m):
            row = m.group(0); rn = int(re.search(r' r="(\d+)"', row).group(1))
            if rn < hdr_row or rn > end_row: return row
            text = 'GESTOR 2027' if rn == hdr_row else vals.get(rn, '')
            prev = re.search(r'<c r="%s%d"[^>]*?(?: s="(\d+)")?[^>]*?(?:/>|>)' % (ref.group(3), rn), row)
            sm = re.search(r' s="(\d+)"', prev.group(0)) if prev else None
            sattr = f' s="{sm.group(1)}"' if sm else ''
            cell = f'<c r="{newc}{rn}"{sattr} t="inlineStr"><is><t>{xesc(text)}</t></is></c>' if text else f'<c r="{newc}{rn}"{sattr}/>'
            # quitar celda existente en la nueva columna (vacía) e insertar en orden
            row = re.sub(r'<c r="%s%d"(?:[^>]*/>|[^>]*>.*?</c>)' % (newc, rn), '', row)
            cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', row))
            after = [c for c in cells if col_idx(c.group(1)) > last + 1]
            pos = after[0].start() if after else row.rfind('</row>')
            if row.endswith('/>') and '</row>' not in row:
                row = row[:-2] + '>' + cell + '</row>'
            else:
                row = row[:pos] + cell + row[pos:]
            row = re.sub(r' spans="\d+:\d+"', '', row, count=1)
            return row
        sx = re.sub(r'<row [^>]*?(?:/>|>.*?</row>)', fix_row, sx, flags=re.S)
        cols = [(int(a), int(b)) for a, b in re.findall(r'<col min="(\d+)" max="(\d+)"', sx)]
        if '</cols>' in sx and not any(a <= last + 1 <= b for a, b in cols):
            sx = sx.replace('</cols>', f'<col min="{last + 1}" max="{last + 1}" width="18" customWidth="1"/></cols>', 1)
        dm = re.search(r'<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"/>', sx)
        if dm and col_idx(dm.group(2)) < last + 1:
            sx = sx.replace(dm.group(0), f'<dimension ref="{dm.group(1)}:{newc}{dm.group(3)}"/>')
        edits[sf] = sx
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = edits[item.filename].encode('utf-8') if item.filename in edits else zin.read(item.filename)
            zout.writestr(item, data)

if __name__ == '__main__':
    main(*sys.argv[1:])
