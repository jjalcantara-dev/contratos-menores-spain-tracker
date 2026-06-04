# Contratos Menores — Spain Tracker

Este proyecto nació por curiosidad en Vélez-Málaga, al descubrir la cantidad de contratos menores que se adjudican cada semana. A partir de ahí surgió la idea de automatizarlo y hacerlo extensible para que cualquiera pueda usarlo con su propio municipio.

Scraper automático que recoge todos los contratos menores adjudicados por un ayuntamiento desde la plataforma pública de contratación del Estado y los publica en una hoja de cálculo actualizada semanalmente.

## 📊 Ver los datos — Vélez-Málaga

> Este Sheet contiene los datos del Ayuntamiento de Vélez-Málaga. Si usas el proyecto con otro municipio, este enlace no te corresponde.

[**Abrir Google Sheet →**](https://docs.google.com/spreadsheets/d/1oRhwJBzAx8C5-LALaJ7hK4IEMN5kK4CUr_VoTM3ZH64)

- **Contratos {año}** — ranking de adjudicatarios por año (2018, 2019, ... 2026)
- **Estadísticas** — gráficas de gasto, contratos, adjudicatarios y top 10 histórico
- **Registro Total** — unión de todos los años en una sola tabla
- Actualización automática cada domingo

## ¿Qué son los contratos menores?

Según la [Ley 9/2017 de Contratos del Sector Público](https://www.boe.es/buscar/act.php?id=BOE-A-2017-12902), son contratos de valor estimado (sin IVA) inferior a **40.000 € en obras** o **15.000 € en servicios y suministros** (art. 118). Pueden adjudicarse directamente a cualquier empresario sin licitación pública (art. 131.3).

La ley prohíbe expresamente fraccionar contratos para eludir estos límites (art. 99.2), y la Instrucción 1/2019 de la OIReScon exige solicitar al menos tres presupuestos como medida antifraude.

## ¿Cómo funciona?

**Modo automático (GitHub Actions):**

1. GitHub Actions ejecuta el scraper cada domingo a la 1:00 AM
2. El script navega con Firefox headless por [contrataciondelestado.es](https://contrataciondelestado.es)
3. Extrae todos los contratos del año en curso y actualiza su pestaña
4. Regenera **Registro Total** y **Estadísticas** con el histórico completo

Los nombres de empresa se normalizan automáticamente: `AXAPLAY, S.L.` y `AXAPLAY S.L.` se agrupan como el mismo adjudicatario.

**Modo local (sin configurar nada):**

Necesitas Python, Firefox y ejecutar:

```bash
pip install selenium webdriver-manager openpyxl
python scraper_xlsx.py
```

Genera `output/contratos.xlsx` con pestañas por año, Estadísticas y Registro Total. Para cargar varios años de una vez: `python scraper_xlsx.py 2012 2025`.

## Datos

| Campo | Descripción |
|-------|-------------|
| # | Posición en el ranking |
| Adjudicatario | Empresa o persona que recibe el contrato |
| Nº Contratos | Número de contratos adjudicados |
| Proyectos | Detalle de cada contrato con su importe |
| Total (€) | Suma total adjudicada |

## Usar con otro ayuntamiento

Para probarlo rápido en local con tu municipio, solo necesitas la URL de su perfil de contratante y ejecutar `scraper_xlsx.py`. Para la versión automatizada con Google Sheets y GitHub Actions, sigue los pasos del [**SETUP.md →**](SETUP.md).

## Fuente de datos

[Plataforma de Contratación del Sector Público](https://contrataciondelestado.es) — [Perfil del contratante del Ayuntamiento de Vélez-Málaga](https://contrataciondelestado.es/wps/poc?uri=deeplink:perfilContratante&idBp=VW8fwBSzF%2FEQK2TEfXGy%2BA%3D%3D).

## Autor

Creado por **Jesús Jiménez** — [Connect](https://jjalcantara.dev)

- Web: [jjalcantara.dev](https://jjalcantara.dev)
- LinkedIn: [linkedin.com/in/jjalcantara](https://linkedin.com/in/jjalcantara)
- GitHub: [github.com/jjalcantara-dev](https://github.com/jjalcantara-dev)
- Email: [jesusjimalc98@gmail.com](mailto:jesusjimalc98@gmail.com)

Si haces un fork, por favor mantén la atribución al autor original.

---

*Proyecto de transparencia ciudadana. Los datos son públicos y provienen de fuentes oficiales.*
