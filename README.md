# STCLE Dashboard

Dashboard estático para GitHub Pages + GitHub Actions.

## Uso

1. Sube los XLSX mensuales a `data/raw/`.
2. Ejecuta localmente `npm install` y `npm run build:data`, o deja que GitHub Actions lo haga.
3. Publica con GitHub Pages.

## Archivos generados

- `data/kpis.json`: KPIs por persona/cargo/período.
- `data/metadata.json`: filtros y resumen general.
- `data/alerts.json`: alertas automáticas.
- `data/analytics.json`: análisis sindicato, rankings, evolución y heatmap.
- `data/detail/YYYY-MM.json`: detalle mensual compacto opcional.

No se genera `dashboard.json` completo para evitar archivos sobre 100 MB.
