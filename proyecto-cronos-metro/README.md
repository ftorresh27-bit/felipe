# Proyecto Cronos — Metro de Santiago

Informe ejecutivo para la evaluación práctica **"Proyecto Cronos – Metro de Santiago"** (asignatura Mando de Control & Balanced Scorecard, prof. Germán Droguett).

El trabajo analiza cómo el **Balanced Scorecard (BSC)** de Metro S.A. debió **pivotar de una estrategia de crecimiento a una de supervivencia y recuperación** ante tres "Cisnes Negros" consecutivos:

- **Escenario A (Oct 2019):** Crisis de infraestructura — estallido social y destrucción de estaciones.
- **Escenario B (2020–2021):** Colapso de demanda y liquidez — COVID-19.
- **Escenario C (2023–2024):** Crisis de seguridad y evasión masiva.

Para **cada escenario** se desarrollan las 5 secciones exigidas por la pauta:
1. Auditoría Histórica (datos reales de la crisis)
2. El Pivot Estratégico (mapa de emergencia, 4 perspectivas, causa-efecto bottom-up)
3. Definición de Métricas (KPI con tipo Lag/Lead)
4. Ficha de Control Crítica (indicador Lead, fórmula, frecuencia, meta, responsable)
5. Iniciativa Ágil (proyecto concreto de corto plazo)

Además incluye el material de preparación para la dinámica presencial **"Shark Tank"** (pitch de 8 min + defensa por rol: Ministerio de Hacienda, Sindicato y Asociación de Pasajeros).

## Estructura del repositorio

```
proyecto-cronos-metro/
├── informe/
│   ├── Proyecto_Cronos_Metro_Santiago.md     # Informe (fuente Markdown)
│   └── Proyecto_Cronos_Metro_Santiago.docx   # Informe exportado a Word
├── tools/
│   └── build_docx.py                         # Conversor Markdown -> .docx (Python puro)
└── README.md
```

## Regenerar el documento Word

No requiere dependencias externas (solo la librería estándar de Python 3):

```bash
python3 tools/build_docx.py informe/Proyecto_Cronos_Metro_Santiago.md informe/Proyecto_Cronos_Metro_Santiago.docx
```

## Importante sobre las cifras financieras

Las cifras de Estados Financieros (EERR) marcadas con `[VERIFICAR EN EERR]` deben **confirmarse y completarse** con los valores exactos de las Memorias Anuales / EERR auditados de Metro S.A., disponibles en:

- Metro S.A. — Informe Financiero Anual: https://metro.cl/gobierno-corporativo/informe-financiero-anual
- Comisión para el Mercado Financiero (CMF): https://www.cmfchile.cl

El informe documenta exactamente **qué línea del EERR usar y cómo construir cada indicador**, de modo que solo reste pegar el valor final. Esto responde a la advertencia de la pauta sobre la veracidad de los números.

> Las fuentes externas fueron parafraseadas y resumidas; no se reproducen textualmente. El listado completo de referencias está en la sección 9 del informe.
