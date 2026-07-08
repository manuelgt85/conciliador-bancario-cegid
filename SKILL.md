---
name: conciliador-bancario-cegid
description: >
  Concilia extractos bancarios contra el libro mayor y plan de cuentas de Cegid Diez,
  asignando la cuenta de contrapartida a cada movimiento. USA ESTA SKILL cuando Manuel
  diga "concilia el extracto", "conciliar banco", "contrapartidas del extracto",
  "procesa el extracto de [banco]", "conciliación bancaria de [cliente]" o similar.
  Soporta cualquier formato bancario español (BBVA, CaixaBank, Santander, Sabadell…),
  detecta el layout automáticamente, cruza con el libro mayor y PGC de Cegid, aplica
  la lógica contable PGC 2007 y genera un Excel con la columna CUENTA_CONTRAPARTIDA
  lista para importar en Cegid, más una hoja REVISAR con los movimientos que requieren
  criterio manual.
---

# Conciliador Bancario Cegid — EMETE Asesoría

## Cuándo se activa

Frases de activación:
- "Concilia el extracto de [banco] de [cliente]"
- "Tengo el extracto de [banco], el mayor y el plan de cuentas"
- "Contrapartidas del extracto"
- "Conciliación bancaria de [cliente] / [período]"
- "Procesa este extracto para Cegid"

## Ficheros que necesita

Siempre tres. Pídelos si no los ha adjuntado:

| Fichero | Descripción | Formatos válidos |
|---|---|---|
| **Extracto bancario** | Movimientos del período | .xlsx, .xls, .csv (aunque sea xlsx renombrado) |
| **Libro mayor** | Exportación de Cegid — hoja `Extracto` | .xlsx |
| **Plan de cuentas** | Exportación de Cegid — hoja `Datos` | .xlsx |

Opcional:
- **Filtro de fechas**: si solo quiere un rango (ej. "solo enero-marzo"), aplicarlo al cargar el extracto (`FECHA_DESDE` / `FECHA_HASTA` en `conciliar.py`).

---

## Flujo completo

### Paso 1 — Análisis de ficheros
Lee los tres ficheros. Para cada uno informa de:
- Total de movimientos del extracto (y rango de fechas)
- Total de terceros en el libro mayor (cuentas 400/410/430)
- Total de cuentas en el PGC

Si el extracto no tiene las columnas estándar, la auto-detección de formato de `conciliar.py` lo resuelve.

### Paso 2 — Clasificación automática
Ejecuta el script (ver sección técnica). La cascada de clasificación es:
```
1. Identidad propia (cuenta → 572 traspaso)
2. Observaciones/concepto inequívoco (nómina → 465, préstamo → 551, SS → 476…)
3. Cruce con Libro Mayor (400/410/430 del tercero)
4. Excepciones directas (comisiones banco → 626, impuestos → 475, ONG → 678…)
5. Acreedor genérico 41000000 con nota del establecimiento
```

### Paso 3 — Reporte previo (NO generar Excel todavía)
Presenta a Manuel un reporte estructurado con:

**Bloque A — Clasificados sin dudas**: tabla con cuenta propuesta y N movimientos.

**Bloque B — Requieren criterio**: los movimientos donde la contrapartida depende de contexto que el agente no puede inferir. Para cada uno:
- Fecha, concepto, importe
- Propuesta del agente con razonamiento
- Preguntar si confirma o corrige

Ir **pregunta por pregunta** si hay varios dudosos, no todos en una batería. Tras cada respuesta, confirmar y pasar al siguiente.

**Bloque C — Proveedores sin LM**: listado de pagos a establecimientos no identificados que van a `41000000` genérico, con el nombre extraído de las observaciones.

### Paso 4 — Aplicar correcciones y generar Excel
Una vez confirmados todos los criterios:
1. Aplicar las correcciones manuales al resultado de la clasificación automática
2. Generar el Excel con:
   - Hoja original del extracto + columnas añadidas al final: `CUENTA_CONTRAPARTIDA`, `DESCRIPCION_CUENTA`, `CONFIANZA`, `JUSTIFICACION`, `METODO`
   - Hoja `RESUMEN`: estadísticas de confianza y top cuentas
   - Hoja `REVISAR`: movimientos BAJA/REVISAR filtrados
3. Presentar el fichero para descarga

---

## Reglas contables fundamentales (PGC 2007)

**Regla de oro**: ningún pago bancario va directamente a grupo 6, salvo:
- Comisiones/intereses del propio banco → `62600000`
- Donativos ONG → `67800000` (liberalidades, no son pagos a proveedor)

Todo pago a proveedor/acreedor no identificado en el LM → `41000000 ACREEDORES` con nota.

**Dirección de los movimientos**:
- Cargo (negativo) + beneficiario en LM → pago que cancela deuda con acreedor (400/410)
- Abono (positivo) + ordenante en LM → cobro que cancela deuda de cliente (430)
- Transferencia a socio sin referencia → `55100000` REVISAR
- Ingreso en efectivo → `55100000` REVISAR (aportación socio, no caja)
- Retirada en cajero → `57000000` CAJA (572→570)

---

## Sección técnica — Script

Toda la lógica vive en **`conciliar.py`** (junto a este SKILL.md). Es el equivalente
del `trimestre.py` del organizador: contiene el parser de importes, la auto-detección
de formato, el matching con el libro mayor y las reglas de clasificación.

**Antes de conciliar, adapta el bloque CONFIGURACIÓN de `conciliar.py`**:
- `EXTRACTO_PATH`, `LM_PATH`, `PGC_PATH`, `OUTPUT_PATH` → rutas de los ficheros subidos.
- `FECHA_DESDE` / `FECHA_HASTA` → si el usuario pidió un rango.
- `API_KEY` → opcional; solo si se quiere refinar los `ACREEDOR_GENERICO` con Claude.

Verificar la instalación (sin ficheros):
```
python conciliar.py        # debe imprimir: self-check OK
```
Conciliar de verdad (con los tres ficheros y las rutas ya ajustadas):
```
python conciliar.py        # genera OUTPUT_PATH e imprime el reparto ALTA/MEDIA/BAJA/REVISAR
```

## Modo lote (varias empresas)

Estructura de entrada: `carpeta_madre/<empresa>/` con, en cada subcarpeta, su(s)
extracto(s) (varios bancos = varios ficheros), su `MAYOR` y su `PLAN`. Autodetección por
nombre (`*mayor*`, `*plan*`/`*cuentas*`, resto → extractos; formatos `.xlsx/.xls/.csv/.pdf`).

Opcional por empresa: un `criterios.json` (alias marca→cuenta, `cuenta_ventas_tpv`,
`cuenta_prl`, `match_por_importe`…) que evita repreguntar criterios ya tomados.

Ejecutar: fijar `CARPETA_MADRE` en la CONFIGURACIÓN de `conciliar.py` y `python conciliar.py`.

Salida por empresa: `<empresa>_conciliado.xlsx` con hojas
`Historico + RESUMEN + REVISAR + CEGID + ACCIONES_CEGID`. Global: `INFORME_YYYYMMDD.md`.
Tras el lote, Claude presenta **una sola ronda** de dudas de alto impacto (`dudas_alto_impacto`).

### Cascada actualizada
`ALIAS (criterios) → observaciones (R_OBS) → socios → LM por nombre → LM por importe →
excepciones → acreedor genérico`. El match por importe solo asigna si el candidato es único.

### Soporte PDF
Extractos en PDF (BBVA) se leen con `pdfplumber`; tras cargar se ejecuta un check de
continuidad de saldo que avisa si falta/sobra algún movimiento.

---

## Adaptaciones por cliente

Antes de ejecutar, revisar y ajustar en `conciliar.py`:

| Parámetro | Dónde | Descripción |
|---|---|---|
| `SOCIOS` | constante `SOCIOS` | Apellidos o nombres de socios/administradores del cliente |
| Empresa propia (traspaso) | Nivel 1 de `clasificar_fila` | Añadir condición con nombre de la empresa para detectar traspasos entre cuentas propias |
| `41000001` (PRL) | `R_OBS` | Si el cliente tiene proveedor PRL distinto de Prevento, cambiar la cuenta |

---

## Criterios contables que Claude debe conocer (briefing al inicio de cada conciliación)

Al activar la skill, Claude informa a Manuel de estos criterios y pregunta si hay alguno que el cliente trate diferente:

1. **Seguros domiciliados** (SegurCaixa, Mapfre, AXA, Allianz…): ¿gasto directo `62500000` o acreedor `41000000`?
2. **Retiradas de cajero**: siempre `57000000 CAJA` (572→570), salvo indicación contraria.
3. **Ingresos en efectivo**: `55100000` por defecto (aportación socio), no `57000000`.
4. **Transferencias recibidas sin nombre**: preguntar siempre — pueden ser clientes (430), socios (551) o préstamos (521).
5. **Pagos a socio/administrador sin concepto**: `55100000 REVISAR` — no asumir que es nómina.
6. **Compras con tarjeta de importe pequeño y frecuente** (gasolineras, restaurantes, parking): preguntar si van a `55100000` (socio) o al acreedor del establecimiento.

---

## Outputs esperados

| Fichero | Contenido |
|---|---|
| `conciliado.xlsx` | Extracto original + columnas CUENTA_CONTRAPARTIDA, DESCRIPCION_CUENTA, CONFIANZA, JUSTIFICACION, METODO |
| Hoja RESUMEN | Estadísticas: ALTA/MEDIA/BAJA/REVISAR + top cuentas |
| Hoja REVISAR | Movimientos que requieren validación manual |

---

## Notas de mantenimiento

- Añadir nuevos patrones a `R_OBS` o `R_EXCEPCIONES` cuando aparezcan conceptos recurrentes no cubiertos. Al tocar reglas, corre `python conciliar.py` (self-check) para no romper la cascada.
- Los terceros del LM se indexan automáticamente en cada ejecución — no hay fichero de configuración por cliente.
- Si el banco usa un CSV real (no xlsx renombrado), ajustar `cargar_extracto` para leer con `pd.read_csv` detectando separador y encoding.
- El threshold de matching del LM es 0.6. Si hay demasiados falsos positivos, subir a 0.65.
