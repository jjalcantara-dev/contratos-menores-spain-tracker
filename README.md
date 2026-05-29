# Contratos Menores — Vélez-Málaga

Scraper automático que recoge todos los contratos menores adjudicados por el **Ayuntamiento de Vélez-Málaga** desde la plataforma pública de contratación del Estado y los publica en una hoja de cálculo actualizada semanalmente.

## 📊 Ver los datos

[**Abrir Google Sheet →**](https://docs.google.com/spreadsheets/d/1oRhwJBzAx8C5-LALaJ7hK4IEMN5kK4CUr_VoTM3ZH64)

- Una pestaña por año (2018, 2019, ... 2026)
- Ranking de adjudicatarios ordenado por importe total
- Detalle de cada contrato por empresa
- Actualización automática cada domingo

## ¿Qué son los contratos menores?

Los contratos menores son adjudicaciones directas sin licitación pública de hasta 15.000 € en servicios y 40.000 € en obras. Son de obligatoria publicación en la plataforma de contratación del Estado.

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

## Fuente de datos

[Plataforma de Contratación del Sector Público](https://contrataciondelestado.es) — Perfil del contratante del Ayuntamiento de Vélez-Málaga.

---

*Proyecto de transparencia ciudadana. Los datos son públicos y provienen de fuentes oficiales.*
