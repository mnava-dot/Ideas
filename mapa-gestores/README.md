# Atlas de Gestores 2027

Mapa interactivo de los 2,475 municipios de México (INEGI 2023) con la asignación de inmuebles por gestor.

## Qué hace
- **Gestores**: al presionar un gestor se pintan sus municipios con su color (intensidad = número de inmuebles) y el mapa se acerca a su zona.
- **Vista general**: cada municipio con el color de su gestor; rayas cuando el municipio lo comparten varios gestores.
- **Municipio**: clic en el mapa o buscador → inmuebles, gestor(es), vecinos más cercanos y municipios con inmuebles más próximos.
- **Distancias**: origen y destino → km en línea recta, km de carretera aproximados (recta × 1.3) y tiempo aproximado (75 km/h). Tabla de distancias a todo el estado, a todos los municipios con inmuebles o a la cartera del gestor; se copia como CSV.
- **Ruta gestor**: ruta de visita sugerida (vecino más cercano + 2-opt) desde una base elegible y matriz de distancias de su cartera.
- Filtros por proyecto y estatus del inmueble.

## Correr localmente
```
cd mapa-gestores
python3 -m http.server 8000   # abrir http://localhost:8000
```

## Actualizar los datos
`tools/build_inmuebles.py` lee la hoja `OPERACIONES` del Excel de asignación y genera `data/inmuebles.json`,
relacionando cada municipio con su clave INEGI (incluye una tabla de equivalencias para nombres como "CD VICTORIA" → Victoria).

```
python3 tools/build_inmuebles.py ASIGNACION_GESTORES_2027.xlsx data/municipios.topo.json data/inmuebles.json
```

Límites municipales: CONABIO/INEGI 2023 vía github.com/PhantomInsights/mexico-geojson (MIT), simplificados con mapshaper.
