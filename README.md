# Índice de Precios Productor (IPP) e Índice de Precios al Por Mayor (IPM) — Bolivia

Dashboard interactivo y datos abiertos de los índices de precios productor y mayorista de Bolivia.

**Live:** [centro-de-estudios-populi.github.io/populi-precios-productor](https://centro-de-estudios-populi.github.io/populi-precios-productor/)

## Datos

Fuente: [Instituto Nacional de Estadística (INE)](https://www.ine.gob.bo/) — Base 2016=100

- **IPP**: Índice de Precios Productor — mide la variación de precios a la salida de producción
- **IPM**: Índice de Precios al Por Mayor — mide la variación de precios en el canal mayorista

Los datos se actualizan automáticamente mediante GitHub Actions.

## Estructura

```
data/           → JSON procesados (IPP, IPM, metadata)
scripts/        → Scraper Python
embed/          → Widgets embebibles para populi.org.bo
.github/        → GitHub Actions workflows
index.html      → Dashboard principal
```

## Centro de Estudios POPULI

[populi.org.bo](https://populi.org.bo)
