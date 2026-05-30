# Contratos Menores — Spain Tracker

Este proyecto nació por curiosidad en Vélez-Málaga, al descubrir la cantidad de contratos menores que se adjudican cada semana. A partir de ahí surgió la idea de automatizarlo y hacerlo extensible para que cualquiera pueda usarlo con su propio municipio.

Scraper automático que recoge todos los contratos menores adjudicados por un ayuntamiento desde la plataforma pública de contratación del Estado y los publica en una hoja de cálculo actualizada semanalmente.

## 📊 Ver los datos — Vélez-Málaga

> Este Sheet contiene los datos del Ayuntamiento de Vélez-Málaga. Si usas el proyecto con otro municipio, este enlace no te corresponde.

[**Abrir Google Sheet →**](https://docs.google.com/spreadsheets/d/1oRhwJBzAx8C5-LALaJ7hK4IEMN5kK4CUr_VoTM3ZH64)

- Una pestaña por año (2018, 2019, ... 2026)
- Ranking de adjudicatarios ordenado por importe total
- Detalle de cada contrato por empresa
- Actualización automática cada domingo

## ¿Qué son los contratos menores?

Según la [Ley 9/2017 de Contratos del Sector Público](https://www.boe.es/buscar/act.php?id=BOE-A-2017-12902), son contratos de valor estimado (sin IVA) inferior a **40.000 € en obras** o **15.000 € en servicios y suministros** (art. 118). Pueden adjudicarse directamente a cualquier empresario sin licitación pública (art. 131.3).

La ley prohíbe expresamente fraccionar contratos para eludir estos límites (art. 99.2), y la Instrucción 1/2019 de la OIReScon exige solicitar al menos tres presupuestos como medida antifraude.

## ¿Cómo funciona?

1. GitHub Actions ejecuta el scraper cada domingo a la 1:00 AM
2. El script navega con Firefox headless por [contrataciondelestado.es](https://contrataciondelestado.es)
3. Extrae todos los contratos del año en curso
4. Borra la hoja y reescribe los datos desde cero (sin duplicados)

## Datos

| Campo | Descripción |
|-------|-------------|
| # | Posición en el ranking |
| Adjudicatario | Empresa o persona que recibe el contrato |
| Nº Contratos | Número de contratos adjudicados |
| Proyectos | Detalle de cada contrato con su importe |
| Total (€) | Suma total adjudicada |

## Usar con otro ayuntamiento

Si quieres replicar esto con tu municipio, sigue los pasos del [**SETUP.md →**](SETUP.md): crear una cuenta de servicio en Google Cloud, configurar el Google Sheet y añadir los secretos en GitHub Actions.

## Fuente de datos

[Plataforma de Contratación del Sector Público](https://contrataciondelestado.es) — [Perfil del contratante del Ayuntamiento de Vélez-Málaga](https://contrataciondelestado.es/wps/poc?uri=deeplink:perfilContratante&idBp=VW8fwBSzF%2FEQK2TEfXGy%2BA%3D%3D).

---

*Proyecto de transparencia ciudadana. Los datos son públicos y provienen de fuentes oficiales.*
