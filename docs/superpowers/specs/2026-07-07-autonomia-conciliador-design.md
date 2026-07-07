# Diseño — Autonomía del Conciliador Bancario Cegid

**Fecha:** 2026-07-07
**Skill:** `conciliador-bancario-cegid` (`conciliar.py` + `SKILL.md`)
**Origen:** brainstorming sobre la retrospectiva real *Elegance Group / BBVA 2025* (522 movimientos)
más la petición de procesar **una carpeta con varias empresas y varios bancos** y devolver
**informes** y **acciones para Cegid** (crear cuentas antes de importar).

## Objetivo

Que la skill haga sola lo que hoy exige criterio manual (~15% de los movimientos caen a
REVISAR/BAJA) y escale a lote multi-empresa, produciendo por empresa un Excel conciliado
listo para importar en Cegid Diez + un bloque de acciones previas, y un informe global.

**Principio rector:** todo es **aditivo**. La lógica contable PGC de `clasificar_fila` y las
reglas existentes (`R_OBS`, `R_EXCEPCIONES`) no cambian de comportamiento; se añaden niveles
de cascada, entradas y salidas nuevas.

---

## Alcance por fases

El spec cubre el roadmap completo. El plan de implementación lo secuencia 1 → 2 → 3 para no
romper el motor de una empresa antes de escalar a lote.

### Fase 1 — Fiabilidad de clasificación (por movimiento)

**F1.1 — Match por importe contra el libro mayor** *(prioridad alta)*
Nuevo nivel en la cascada de `clasificar_fila`, **después** del match por nombre (nivel 4) y
**antes** de las excepciones al 41000000 (nivel 5):
- Para un pago sin match por nombre, buscar en el mayor un tercero (400/410/430) con un apunte
  del **mismo importe** (tolerancia ±0,02 €).
- Si el match es **único** → asignar ese tercero, confianza **ALTA**, método `MATCH_IMPORTE`.
- Si hay **varios** candidatos → NO asignar; sigue la cascada (acabará en REVISAR o genérico).
- Requiere cargar del mayor, además del nombre, la lista de importes de apuntes por tercero.

**F1.2 — `criterios.json` por empresa** *(prioridad alta)*
La skill busca `criterios*.json` en la subcarpeta de la empresa. Si no existe, funciona igual
(con menos autonomía). Esquema (todos los campos opcionales):
```json
{
  "cliente": "Elegance Group Enterprise S.L.",
  "cuenta_ventas_tpv": "43000000",
  "cuenta_ingreso_efectivo": "43000000",
  "cuenta_prl": "41000009",
  "compras_tarjeta_sueltas": "55100000",
  "politica_proveedor_nuevo": "crear_410",
  "politica_seguros_domiciliados": "crear_410",
  "alias_terceros": { "YOIGO": "41000015", "FACEBK": "41000008", "FLOCUSS": "40000001" },
  "match_por_importe": true
}
```
- `alias_terceros`: se consulta **antes** del matcher por nombre; si el texto contiene la clave,
  se asigna la cuenta directamente (ALTA, método `ALIAS`).
- Las `cuenta_*` sustituyen los destinos por defecto de las reglas correspondientes.
- `match_por_importe` activa/desactiva F1.1 (por defecto `true`).

**F1.3 — Tokenización robusta** *(media)*
`toks()` separa camelCase y secuencias `letras+MAYÚSCULAS` pegadas antes de tokenizar
(`"TotalEnergiesClientesSAU"` → `TOTAL ENERGIES CLIENTES SAU`).

**F1.4 — Regla TPV / "liquidación remesa de comercios"** *(media)*
Nueva entrada en `R_OBS` para `LIQUIDACI[OO]N\s+REMESA\s+DE\s+COMERCIOS` (cobros de TPV) →
cuenta de `criterios.cuenta_ventas_tpv` (por defecto `43000000`, confianza MEDIA por depender de
cómo registre ventas el cliente).

**F1.5 — Flexibilizar reglas de nómina y SS** *(baja)*
- Nómina: permitir texto entre `NOMINA` y el mes (`NOMINA liquidacion de junio 2025`).
- SS: aceptar `TESORERIA … SEGURID` truncado en el PDF.

**F1.6 — PGC indexado también por subcuenta de 8 dígitos** *(baja, prerrequisito de F3.2)*
`cargar_pgc` devuelve el índice actual (3 díg) **y** un índice por subcuenta de 8 díg, para poder
comprobar si una cuenta completa (`55100000`) existe sin falsos negativos.

**F1.7 — Par retrocedido (+/−) → 555** *(baja)*
Detectar pares de movimientos de igual importe y signo opuesto muy próximos en fecha (ingreso por
error + retrocesión) → `55500000` Partidas pendientes de aplicación (netean a 0).

### Fase 2 — Entrada PDF + salida Cegid

**F2.1 — Soporte de extractos PDF** *(alta)*
Rama `.pdf` en `cargar_extracto` → `cargar_extracto_pdf(path, banco)`. Se parte del parser
**BBVA ya escrito y validado** en la retrospectiva (clustering de líneas por salto vertical > 18 pt,
regex tolerante de la línea numérica, `header`/`obs` arriba/abajo, split en `|`). Estructura por
banco para poder añadir Santander/CaixaBank/etc. cuando aparezcan; si un PDF no casa ningún parser
conocido, avisar claramente en vez de devolver basura. Dependencia nueva: `pdfplumber`.

**F2.2 — Check de continuidad de saldo** *(media)*
Tras cargar cualquier extracto (PDF o xlsx): comprobar `saldo[i] == saldo[i+1] + importe[i]`
(orden descendente, tolerancia 0,02). Si hay roturas, **avisar** y listar las filas sospechosas
(probable parse incompleto). No bloquea la ejecución.

**F2.3 — Hoja `CEGID` de 4 columnas** *(alta)*
`generar_excel` añade una hoja `CEGID` con **exactamente** `Fecha │ Concepto │ Importe │
Contrapartida`, en ese orden, sin columnas auxiliares intercaladas, lista para el selector de
importación de Cegid. Las hojas de trabajo (Historico, RESUMEN, REVISAR, ACCIONES_CEGID) se
mantienen aparte.

### Fase 3 — Batch multi-empresa + acciones + informes

**F3.1 — Recorrido de carpeta madre** *(alta)*
`procesar_lote(carpeta_madre)` recorre cada subcarpeta = una empresa. Por empresa autodetecta por
nombre de fichero: `*mayor*` → mayor, `*plan*`/`*cuentas*` → plan, el resto de `.xlsx/.csv/.pdf` →
extractos. Varios extractos = varios bancos → **un Excel por empresa** con columna `BANCO`. Si falta
mayor o plan, la empresa se salta con aviso (no aborta el lote).

**F3.2 — Acciones para Cegid (hoja `ACCIONES_CEGID` por empresa)** *(alta)*
Tres bloques, en orden de "haz esto antes de importar":
1. **Terceros a crear**: los pagos/cobros sin match (hoy → 41000000), agrupados por nombre
   normalizado. Por grupo: **código correlativo libre** por bloque (proveedor `400xxxxx`, acreedor
   `410xxxxx`, cliente `430xxxxx`, 8 dígitos = máximo usado en mayor+plan + 1; si el bloque está
   vacío, arranca en `40000001`/`41000001`/`43000001`), nombre extraído, tipo, y NIF `(pendiente)`
   salvo que el nombre cuadre con un tercero conocido.
2. **Subcuentas a crear**: contrapartidas que una regla propone y que **no existen** en el plan
   (usa el índice de 8 díg de F1.6).
3. **Avisos**: movimientos sin NIF/nombre identificable, roturas de saldo (F2.2), formatos raros.

**F3.3 — Informe global (`INFORME_YYYYMMDD.md`)** *(alta)*
Una fila por empresa: nº movimientos, % ALTA / MEDIA, nº terceros a crear, nº subcuentas, nº avisos,
y un "orden de trabajo" (qué crear antes de importar). Se guarda en la carpeta madre.

**F3.4 — Ronda final única de dudas** *(media)*
Tras procesar todo el lote, presentar **una sola** tanda con las dudas de **alto impacto**
consolidadas de todas las empresas: movimientos REVISAR/BAJA cuyo `|importe|` esté en el top, o que
se repitan (mismo proveedor/concepto varias veces). El resto queda documentado en cada `REVISAR` e
`INFORME`. Tras las respuestas, re-generar los Excel afectados.

---

## Arquitectura

- **Un solo `conciliar.py`** que crece con funciones nuevas; el CLI (`__main__`) admite dos modos:
  - una empresa (comportamiento actual, rutas en CONFIGURACIÓN),
  - lote: `procesar_lote(CARPETA_MADRE)`.
- Piezas con una responsabilidad clara y testeable por separado:
  `cargar_extracto*` (I/O + parse), `cargar_lm`/`cargar_pgc` (índices), `cargar_criterios`,
  `clasificar_fila` (cascada), `detectar_acciones` (terceros/subcuentas/avisos), `generar_excel`
  (con hojas CEGID y ACCIONES_CEGID), `generar_informe`, `procesar_lote` (orquestación).
- Sin cambios en la semántica contable existente; los niveles nuevos de la cascada se insertan en
  el orden indicado (alias → nombre → **importe** → excepciones → genérico).

## Datos / privacidad

- El repo es público: `criterios.json`, extractos, mayor y plan **viven en la carpeta de la empresa**
  (fuera del repo) y están cubiertos por `.gitignore` (`*.xlsx`, `*.csv`, `*.pdf`, `datos/`,
  `clientes/`). Los `criterios*.json` también se ignoran (contienen nombres/cuentas de clientes).

## Manejo de errores

- Extracto PDF que no casa ningún parser conocido → aviso explícito, esa empresa a REVISAR manual.
- Roturas de saldo → aviso + filas sospechosas listadas; no se bloquea.
- Empresa sin mayor o sin plan → se salta con aviso; el lote continúa.
- Match por importe ambiguo → nunca se adivina; va a REVISAR.

## Testing (self-check ampliado)

`python conciliar.py` sin ficheros sigue imprimiendo `self-check OK`, ahora cubriendo:
- Match por importe: único → ALTA; múltiple → no asigna.
- Siguiente código libre por bloque (400/410/430), incl. bloque vacío.
- `toks()` separa camelCase/pegados.
- Alias de `criterios.json` gana al matcher por nombre.
- PGC: existencia por subcuenta de 8 dígitos.
- Parser PDF: sobre un fragmento BBVA sintético incrustado (o marcado `skip` si no hay `pdfplumber`),
  y `validar_saldo` detecta una rotura introducida a propósito.

## Fuera de alcance (YAGNI)

- Generar fichero importable de **terceros** para Cegid (se crean a mano; solo checklist).
- Parsers de bancos distintos de BBVA hasta que aparezca un extracto real de ese banco.
- OCR de PDFs escaneados (se asume PDF con texto).
