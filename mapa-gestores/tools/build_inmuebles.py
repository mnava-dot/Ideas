import openpyxl,json,unicodedata,collections,sys,datetime
def N(s):
    s=unicodedata.normalize('NFD',str(s or '')).encode('ascii','ignore').decode().upper()
    for a in '.,-"': s=s.replace(a,' ')
    return ' '.join(s.split())
topo=json.load(open(sys.argv[2]))
geos=topo['objects']['mun']['geometries']
byst=collections.defaultdict(dict)
for g in geos:
    p=g['properties']; byst[N(p['s'])][N(p['n'])]=p['id']
stmap={'COAHUILA':'COAHUILA DE ZARAGOZA','MICHOACAN':'MICHOACAN DE OCAMPO','VERACRUZ':'VERACRUZ DE IGNACIO DE LA LLAVE','EDO MEXICO':'MEXICO','ESTADO DE MEXICO':'MEXICO','CDMX':'CIUDAD DE MEXICO'}
# alias: (excel_state, excel_mun) -> (official_state or None, official_mun)
A={
('AGUASCALIENTES','AGUASCALIENTES CENTRO'):(None,'AGUASCALIENTES'),
('BAJA CALIFORNIA','CABO SAN LUCAS'):('BAJA CALIFORNIA SUR','LOS CABOS'),
('BAJA CALIFORNIA','ENSENADA JUAREZ'):(None,'ENSENADA'),
('BAJA CALIFORNIA','LA PAZ'):(None,'ENSENADA'),
('BAJA CALIFORNIA','LOS CABOS'):('BAJA CALIFORNIA SUR','LOS CABOS'),
('BAJA CALIFORNIA','LOTERO'):('BAJA CALIFORNIA SUR','LORETO'),
('BAJA CALIFORNIA','ROSARITO'):(None,'PLAYAS DE ROSARITO'),
('BAJA CALIFORNIA SUR','CABO SAN LUCAS'):(None,'LOS CABOS'),
('BAJA CALIFORNIA SUR','SAN JOSE DEL CABO'):(None,'LOS CABOS'),
('BAJA CALIFORNIA SUR','SANTA ROSALIA'):(None,'MULEGE'),
('CHIHUAHUA','ASCENCION'):(None,'ASCENSION'),
('CHIHUAHUA','CD CUAUHTEMOC'):(None,'CUAUHTEMOC'),
('CHIHUAHUA','CD DELICIAS'):(None,'DELICIAS'),
('CHIHUAHUA','CD JIMENEZ'):(None,'JIMENEZ'),
('CHIHUAHUA','CD JUAREZ'):(None,'JUAREZ'),
('CHIHUAHUA','CD MEOQUI'):(None,'MEOQUI'),
('CHIHUAHUA','MISIONES I ETAPA'):(None,'CHIHUAHUA'),
('CHIHUAHUA','PARRAL'):(None,'HIDALGO DEL PARRAL'),
('COAHUILA','NUEVA ROSITA'):(None,'SAN JUAN DE SABINAS'),
('COAHUILA','PARRAS DE LA FUENTE'):(None,'PARRAS'),
('COAHUILA','SAN BAUTISTA'):(None,'MONCLOVA'),
('JALISCO','ACATLAN'):(None,'ACATLAN DE JUAREZ'),
('JALISCO','AHUALULCO'):(None,'AHUALULCO DE MERCADO'),
('JALISCO','CABO CORRIENTE'):(None,'CABO CORRIENTES'),
('JALISCO','ESCOBEDO'):(None,'SAN JUANITO DE ESCOBEDO'),
('JALISCO','OCOTLAN DE MORELOS'):(None,'OCOTLAN'),
('JALISCO','OJUELOS'):(None,'OJUELOS DE JALISCO'),
('JALISCO','TEPATITLAN'):(None,'TEPATITLAN DE MORELOS'),
('JALISCO','TLAQUEPAQUE'):(None,'SAN PEDRO TLAQUEPAQUE'),
('JALISCO','VALLARTA'):(None,'PUERTO VALLARTA'),
('JALISCO','ZACOALCO'):(None,'ZACOALCO DE TORRES'),
('JALISCO','ZEMPOALA'):('HIDALGO','ZEMPOALA'),
('MICHOACAN','ARIO DE ROSALES'):(None,'ARIO'),
('MICHOACAN','CD HIDALGO'):(None,'HIDALGO'),
('MICHOACAN','LA CIUDAD DE MORELIA'):(None,'MORELIA'),
('MICHOACAN','NUEVA ITALIA'):(None,'MUGICA'),
('MICHOACAN','PERIBAN DE RAMOS'):(None,'PERIBAN'),
('MICHOACAN','PUREPERO DE ECHANIZ'):(None,'PUREPERO'),
('MICHOACAN','PURUPEREO'):(None,'PUREPERO'),
('MICHOACAN','SAN JOSE DE GRACIA'):(None,'MARCOS CASTELLANOS'),
('MICHOACAN','TLAPUJAHUA'):(None,'TLALPUJAHUA'),
('MICHOACAN','TUXPAN MORELOS'):(None,'TUXPAN'),
('MICHOACAN','VILLA MORELOS'):(None,'MORELOS'),
('MICHOACAN','ZINAPECUARO DE FIGUEROA'):(None,'ZINAPECUARO'),
('NUEVO LEON','CADEREYTA'):(None,'CADEREYTA JIMENEZ'),
('NUEVO LEON','ESCOBEDO'):(None,'GENERAL ESCOBEDO'),
('NUEVO LEON','ZUAZUA'):(None,'GENERAL ZUAZUA'),
('OAXACA','BAHIAS DE HUATULCO'):(None,'SANTA MARIA HUATULCO'),
('OAXACA','HUATULCO'):(None,'SANTA MARIA HUATULCO'),
('OAXACA','CD IXTEPEC'):(None,'CIUDAD IXTEPEC'),
('OAXACA','IXTEPEC'):(None,'CIUDAD IXTEPEC'),
('OAXACA','ESPINAL'):(None,'EL ESPINAL'),
('OAXACA','JUCHITAN'):(None,'JUCHITAN DE ZARAGOZA'),
('OAXACA','MELCHOR OCAMPO'):(None,'VILLA DE TUTUTEPEC'),
('OAXACA','MIAHUATLAN'):(None,'MIAHUATLAN DE PORFIRIO DIAZ'),
('OAXACA','OAXACA'):(None,'OAXACA DE JUAREZ'),
('OAXACA','PINOTEPA NACIONAL'):(None,'SANTIAGO PINOTEPA NACIONAL'),
('OAXACA','POCHUTLA'):(None,'SAN PEDRO POCHUTLA'),
('OAXACA','PUTLA'):(None,'PUTLA VILLA DE GUERRERO'),
('OAXACA','TEHUANTEPEC'):(None,'SANTO DOMINGO TEHUANTEPEC'),
('QUERETARO','AMEALCO CENTRO'):(None,'AMEALCO DE BONFIL'),
('QUERETARO','BENITO JUAREZ'):(None,'QUERETARO'),
('QUERETARO','EL MARQUEZ'):(None,'EL MARQUES'),
('QUERETARO','MARQUEZ'):(None,'EL MARQUES'),
('QUERETARO','EL PUEBLITO'):(None,'CORREGIDORA'),
('QUERETARO','EL ROCIO'):(None,'QUERETARO'),
('QUERETARO','EPIGMENIO GONZALEZ'):(None,'QUERETARO'),
('QUERETARO','FELIX OSORES SOTOMAYO'):(None,'QUERETARO'),
('QUERETARO','SAN JUAN RIO'):(None,'SAN JUAN DEL RIO'),
('QUERETARO','SANTA ROSA JAUREGUI'):(None,'QUERETARO'),
('QUERETARO','SANTIAGO DE QUERETARO'):(None,'QUERETARO'),
('QUERETARO','TEPEJI DEL RIO'):('HIDALGO','TEPEJI DEL RIO DE OCAMPO'),
('SINALOA','LOS MOCHIS'):(None,'AHOME'),
('SONORA','CD OBREGON'):(None,'CAJEME'),
('SONORA','MAGDALENA DE KINO'):(None,'MAGDALENA'),
('SONORA','BENITO JUAREZ'):(None,'BENITO JUAREZ'),
('TAMAULIPAS','CD MADERO'):(None,'CIUDAD MADERO'),
('TAMAULIPAS','CD MANTE'):(None,'EL MANTE'),
('TAMAULIPAS','CD MIGUEL ALEMAN'):(None,'MIGUEL ALEMAN'),
('TAMAULIPAS','CD VICTORIA'):(None,'VICTORIA'),
('TAMAULIPAS','ESTACION MANUEL'):(None,'GONZALEZ'),
('TAMAULIPAS','HEORICA'):(None,'MATAMOROS'),
('VERACRUZ','ALTO LUCERO GUTIERREZ BARRIOS'):(None,'ALTO LUCERO DE GUTIERREZ BARRIOS'),
('VERACRUZ','BOCA DEL MONTE'):(None,'COMAPA'),
('VERACRUZ','CARLOS CARRILLO'):(None,'CARLOS A CARRILLO'),
('VERACRUZ','CHOCOMAN'):(None,'CHOCAMAN'),
('VERACRUZ','DE ISLA'):(None,'ISLA'),
('VERACRUZ','FORTIN DE LAS FLORES'):(None,'FORTIN'),
('VERACRUZ','IXTAZOQUITLAN'):(None,'IXTACZOQUITLAN'),
('VERACRUZ','LA RUEDA'):(None,'POZA RICA DE HIDALGO'),
('VERACRUZ','MIGUEL LERDO'):(None,'COSAUTLAN DE CARVAJAL'),
('VERACRUZ','OZULUAMA'):(None,'OZULUAMA DE MASCARENAS'),
('VERACRUZ','PAPANTLA DE OLARTE'):(None,'PAPANTLA'),
('VERACRUZ','POZA RICA'):(None,'POZA RICA DE HIDALGO'),
('VERACRUZ','SUCHILAPAM'):(None,'JESUS CARRANZA'),
('VERACRUZ','TARIMOYA'):(None,'VERACRUZ'),
('VERACRUZ','TEJERIA'):(None,'VERACRUZ'),
('VERACRUZ','TEPEZINTLA'):(None,'TEPETZINTLA'),
}
# Reglas de asignación por estado (clave INEGI de entidad): mandan sobre la columna GESTOR 2027 del Excel
REGLAS={'03':'JC',      # Baja California Sur (Los Cabos, Loreto, La Paz...) la atiende JC
        '13':'GERARDO'} # Todo Hidalgo lo atiende GERARDO
wb=openpyxl.load_workbook(sys.argv[1],read_only=True,data_only=True)
rows=list(wb['OPERACIONES'].iter_rows(values_only=True))[1:]
out=[];miss=collections.Counter()
def fmt(v):
    if isinstance(v,datetime.datetime): return v.strftime('%d/%m/%Y')
    return '' if v is None else str(v)
for r in rows:
    if not r[1]: continue
    st,mu=N(r[8]),N(r[9])
    mu2=mu
    if 'TOCUMBO' in mu: mu2='TOCUMBO'
    st2=stmap.get(st,st)
    if (st,mu) in A:
        s,m=A[(st,mu)]; st2=s or st2; mu2=m
    code=byst.get(st2,{}).get(mu2)
    if not code:
        # VERACRUZ BENITO JUAREZ address is Veracruz city when CP 917xx; keep official BENITO JUAREZ otherwise
        miss[(st,mu)]+=1
    g_xls=fmt(r[50]).strip()
    g=REGLAS.get((code or '')[:2],g_xls)
    out.append([r[1],fmt(r[5]),fmt(r[6]),fmt(r[0]),fmt(r[12]),fmt(r[4]),g,fmt(r[51]),code or '',fmt(r[10])[:220],fmt(r[11]),fmt(r[7]),fmt(r[8]),fmt(r[9]),g_xls if g!=g_xls else ''])
print('unmatched',miss, file=sys.stderr)
cols=['id','nombre','tipo','proyecto','estatus','cr','gestor','nota','cve','direccion','subclas','division','estado_xls','mun_xls','gestor_xls']
json.dump({'cols':cols,'rows':out,'fuente':'ASIGNACION_GESTORES_2027.xlsx · hoja OPERACIONES','generado':datetime.date.today().isoformat()},open(sys.argv[3],'w'),ensure_ascii=False,separators=(',',':'))
print(len(out),'rows',sum(1 for o in out if o[8]),'matched',len(set(o[8] for o in out if o[8])),'municipios',file=sys.stderr)
