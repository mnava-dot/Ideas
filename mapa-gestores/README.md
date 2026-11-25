# Atlas de Gestores 2027

Mapa interactivo de los 2,475 municipios de México (INEGI 2023) con la asignación de inmuebles por gestor.

## Qué hace
- **Gestores**: al presionar un gestor se pintan sus municipios con su color (intensidad = número de inmuebles) y el mapa se acerca a su zona.
- **Vista general**: cada municipio con el color de su gestor; rayas cuando el municipio lo comparten varios gestores.
- **Municipio**: clic en el mapa o buscador → inmuebles, gestor(es), vecinos más cercanos y municipios con inmuebles más próximos.
- **Distancias**: origen y destino → km en línea recta, km de carretera aproximados (recta × 1.3) y tiempo aproximado (75 km/h). Tabla de distancias a todo el estado, a todos los municipios con inmuebles o a la cartera del gestor; se copia como CSV.
- **Planificador**: arma el plan de visitas del gestor por jornadas, con fecha y hora:
  - Parámetros: base de salida, primer día, hora de salida, horas por jornada, minutos por inmueble, velocidad promedio, solo días hábiles.
  - Viajes de un día (algoritmo de ahorros Clarke-Wright) para municipios cercanos; giras con pernocta por zona para los lejanos o con muchos inmuebles; jornadas locales en la base; tramos de más de una jornada se marcan como traslado largo (considerar vuelo).
  - Orden de cada salida optimizado (vecino más cercano + 2-opt), itinerario hora por hora con los inmuebles a visitar, enlace a Google Maps por día y exportación CSV.
  - Se pueden quitar o agregar municipios del plan desde la lista o dando clic en el mapa.
- Filtros por proyecto y estatus del inmueble.

## Correr localmente
```
cd mapa-gestores
python3 -m http.server 8000   # abrir http://localhost:8000
```

## Actualizar los datos
`tools/build_inmuebles.py` lee la hoja `OPERACIONES` del Excel de asignación y genera `data/inmuebles.json`,
relacionando cada municipio con su clave INEGI (incluye una tabla de equivalencias para nombres como "CD VICTORIA" → Victoria).
También aplica reglas de asignación por estado (`REGLAS`): Baja California Sur → JC, Hidalgo → GERARDO, Chihuahua → BRENDA (se quita "BRENDA / CRIS").

```
python3 tools/build_inmuebles.py ASIGNACION_GESTORES_2027.xlsx data/municipios.topo.json data/inmuebles.json
```

Límites municipales: CONABIO/INEGI 2023 vía github.com/PhantomInsights/mexico-geojson (MIT), simplificados con mapshaper.
