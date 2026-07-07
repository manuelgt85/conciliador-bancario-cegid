# Autonomía del Conciliador Bancario Cegid — Plan de Implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reducir el ~15% de movimientos que hoy caen a criterio manual y escalar la skill a lote multi-empresa con informes y acciones previas a importar en Cegid.

**Architecture:** Todo aditivo sobre `conciliar.py` (un solo fichero que crece). Se insertan niveles nuevos en la cascada de `clasificar_fila` (alias → nombre → importe → excepciones → genérico), se añaden entradas (PDF, criterios por cliente) y salidas (hojas CEGID y ACCIONES_CEGID, informe global) y un orquestador de lote `procesar_lote`. La semántica contable existente no cambia.

**Tech Stack:** Python 3, openpyxl, pandas, python-dateutil, pdfplumber (nuevo). Tests = asserts dentro de `_self_check()` en `conciliar.py`, sin framework; se corren con `python3 conciliar.py` (sin ficheros de datos → imprime `self-check OK`).

## Global Constraints

- **Sin romper comportamiento existente:** las reglas `R_OBS`/`R_EXCEPCIONES` y el orden contable actual se mantienen. Los cambios son aditivos.
- **Test = self-check en fichero:** cada tarea añade `assert`s a `_self_check()` en `conciliar.py`. Comando único: `python3 conciliar.py` (sin `extracto.xlsx`/`mayor.xlsx`/`plan_cuentas.xlsx` presentes) → debe imprimir `self-check OK`. No introducir pytest.
- **Privacidad / repo público:** ningún dato de cliente al repo. `.gitignore` ya cubre `*.xlsx *.csv *.pdf criterios*.json datos/ clientes/`.
- **Match por importe nunca adivina:** si hay más de un candidato, no se asigna.
- **Tolerancia de importes:** ±0,02 € en comparaciones de saldo e importe.
- **Código de tercero:** 8 dígitos, correlativo libre por bloque (proveedor 400, acreedor 410, cliente 430).
- **Commits frecuentes y push:** cada tarea termina con `git commit`. Al final del plan, `git push origin main` para actualizar el repositorio en GitHub (en el PC/servidor donde esté clonado, `git pull`).

---

## FASE 1 — Fiabilidad de clasificación

### Task 1: PGC indexado también por subcuenta de 8 dígitos

**Files:**
- Modify: `conciliar.py` (`cargar_pgc`, `_self_check`)

**Interfaces:**
- Produces: `cargar_pgc(path)` devuelve un dict que contiene **tanto** las claves de 3 dígitos (código, actual) **como** las de 8 dígitos (subcuenta, `row[8]`), ambas → nombre. Consumido por Task 10 para saber si una subcuenta completa existe.

- [ ] **Step 1: Añadir el assert que falla en `_self_check()`**

Insertar dentro de `_self_check()`, antes de `print("self-check OK")`:

```python
    # Task 1: cargar_pgc indexa por 3 y por 8 dígitos
    from openpyxl import Workbook as _WB
    import tempfile, os as _os
    _wb = _WB(); _ws = _wb.create_sheet("Datos"); del _wb[_wb.sheetnames[0]]
    # fila 6 en adelante: col índice 4 = código 3 díg, col 8 = subcuenta 8 díg, col 9 = nombre
    for _ in range(5): _ws.append([])
    _ws.append([None,None,None,None,"572",None,None,None,"57200000","BANCOS"])
    _tmp = _os.path.join(tempfile.gettempdir(), "_pgc_test.xlsx"); _wb.save(_tmp)
    _pgc = cargar_pgc(_tmp); _os.remove(_tmp)
    assert _pgc.get("572") == "BANCOS"
    assert _pgc.get("57200000") == "BANCOS"
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `AssertionError` en `_pgc.get("57200000")` (hoy solo indexa 3 díg).

- [ ] **Step 3: Modificar `cargar_pgc`**

Reemplazar el cuerpo del bucle de `cargar_pgc`:

```python
    for row in ws.iter_rows(min_row=6, values_only=True):
        if row[4] and row[9]:
            pgc[str(row[4]).strip()] = str(row[9]).strip()
        # Task 1: indexar también la subcuenta de 8 dígitos (columna 8)
        if len(row) > 8 and row[8] and row[9]:
            pgc[str(row[8]).strip()] = str(row[9]).strip()
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(pgc): indexar plan de cuentas por subcuenta de 8 dígitos (F1.6)"
```

---

### Task 2: Tokenización que separa camelCase y mayúsculas pegadas

**Files:**
- Modify: `conciliar.py` (`toks`, `_self_check`)

**Interfaces:**
- Produces: `toks(t)` separa `camelCase` y transiciones `minúscula→MAYÚSCULA` antes de tokenizar. Mismo tipo de retorno (`set[str]`).

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 2: toks separa camelCase/pegados
    assert "CLIENTES" in toks("TotalEnergiesClientesSAU")
    assert "ENERGIES" in toks("TotalEnergiesClientesSAU")
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `AssertionError` (hoy `CLIENTESSAU` queda pegado).

- [ ] **Step 3: Modificar `toks`**

Insertar el desglose camelCase justo tras la normalización de acentos, antes de pasar a mayúsculas:

```python
def toks(t):
    if not t: return set()
    import unicodedata
    t = unicodedata.normalize('NFD', str(t)).encode('ascii','ignore').decode('ascii')
    # Task 2: separar minúscula→MAYÚSCULA pegadas (camelCase) antes de subir todo a mayúsculas
    t = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', t)
    t = re.sub(r'[^A-Z0-9\s]',' ', t.upper())
    return {x for x in t.split() if len(x)>2 and x.lower() not in STOPWORDS}
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(match): tokenizar separando camelCase/pegados (F1.3)"
```

---

### Task 3: Criterios por cliente + nivel ALIAS en la cascada

**Files:**
- Modify: `conciliar.py` (nuevo `cargar_criterios`, `clasificar_fila`, `conciliar`, `_self_check`)

**Interfaces:**
- Produces:
  - `cargar_criterios(carpeta)` → dict (vacío si no hay `criterios*.json`). Lee el primer `criterios*.json` de la carpeta.
  - `clasificar_fila(row, terceros, criterios=None)` — nuevo tercer parámetro opcional; si `criterios['alias_terceros']` tiene una clave contenida en concepto/obs/beneficiario, devuelve esa cuenta con método `ALIAS`, confianza `ALTA`, **antes** que el matcher por nombre.
  - `conciliar(df, terceros, pgc, criterios=None, api_key='', lote=30)` — pasa `criterios` a `clasificar_fila`.

- [ ] **Step 1: Añadir los asserts que fallan**

En `_self_check()`:

```python
    # Task 3: alias de criterios gana al matcher por nombre
    _crit = {'alias_terceros': {'FACEBK': '41000008'}}
    _r = clasificar_fila(fila(obs='PAGO FACEBK ADS'), {}, _crit)
    assert _r[0] == '41000008' and _r[4] == 'ALIAS'
    # cargar_criterios devuelve {} si no hay fichero
    assert cargar_criterios(tempfile.gettempdir() + '/no_existe_dir_xyz') == {}
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `TypeError` (clasificar_fila no acepta 3º argumento) o `NameError` (cargar_criterios).

- [ ] **Step 3: Añadir `cargar_criterios` y el nivel ALIAS**

Añadir la función (junto a los otros `cargar_*`):

```python
def cargar_criterios(carpeta):
    """Lee el primer criterios*.json de la carpeta. {} si no hay."""
    import glob
    try:
        matches = sorted(glob.glob(os.path.join(carpeta, 'criterios*.json')))
    except Exception:
        return {}
    if not matches:
        return {}
    try:
        with open(matches[0], encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"criterios.json ilegible: {e}")
        return {}
```

Modificar la firma y el primer nivel de `clasificar_fila`:

```python
def clasificar_fila(row, terceros, criterios=None):
    criterios = criterios or {}
    conc   = str(row['CONCEPTO']).upper()
    obs    = str(row['OBSERVACIONES']).upper()
    obs_r  = str(row['OBSERVACIONES'])
    benef  = str(row['BENEFICIARIO']).upper()
    benef_r= str(row['BENEFICIARIO'])
    imp    = float(row['IMPORTE'])
    texto_obs = f"{obs} {conc}"
    # 0. Alias marca→cuenta (criterios por cliente): gana a todo lo demás
    texto_alias = f"{conc} {obs} {benef}"
    for marca, cuenta in (criterios.get('alias_terceros') or {}).items():
        if marca.upper() in texto_alias:
            nom = terceros.get(cuenta, '')
            return cuenta, nom, 'ALTA', f"Alias criterios: «{marca}» → {cuenta}", 'ALIAS'
    # 1. Empresa propia
    # ... (resto igual)
```

Modificar `conciliar` para aceptar y propagar `criterios`:

```python
def conciliar(df, terceros, pgc, criterios=None, api_key='', lote=30):
    init_cache_lm(terceros)
    resultados = []
    for idx, row in df.iterrows():
        cta, desc, conf, just, met = clasificar_fila(row, terceros, criterios)
        resultados.append({
            'idx': idx, 'CUENTA': cta, 'DESCRIPCION': desc,
            'CONFIANZA': conf, 'JUSTIFICACION': just, 'METODO': met
        })
    # ... (resto del cuerpo igual, incluida la rama api_key)
```

Actualizar la llamada en `__main__`: `res = conciliar(df, terceros, pgc, None, API_KEY)`.

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(criterios): criterios.json por cliente + nivel ALIAS en la cascada (F1.2)"
```

---

### Task 4: Match por importe contra el libro mayor

**Files:**
- Modify: `conciliar.py` (`cargar_lm`, `clasificar_fila`, `conciliar`, `_self_check`)

**Interfaces:**
- Produces:
  - `cargar_lm(path)` sigue devolviendo `terceros: {cod: nombre}` (sin cambios de firma) y **además** rellena un módulo-global `_importes_lm: {cod: set(importes_abs)}`. Se expone `init_importes_lm(mapa)` para tests.
  - `match_importe(imp)` → `cod` único o `None`. Consumido por `clasificar_fila`.
  - `clasificar_fila` inserta el nivel importe entre nombre (4) y excepciones (5).

- [ ] **Step 1: Añadir los asserts que fallan**

En `_self_check()`:

```python
    # Task 4: match por importe único → ALTA; múltiple → no asigna
    init_importes_lm({'40000001': {125.84}, '40000002': {99.0}})
    _r = clasificar_fila(fila(benef='PROV RARO', imp=-125.84), {'40000001': 'FLOCCUS'})
    assert _r[0] == '40000001' and _r[4] == 'MATCH_IMPORTE'
    init_importes_lm({'40000001': {125.84}, '40000002': {125.84}})
    _r = clasificar_fila(fila(benef='PROV RARO', imp=-125.84), {'40000001': 'A', '40000002': 'B'})
    assert _r[4] != 'MATCH_IMPORTE'   # ambiguo → no adivina
    init_importes_lm({})              # limpiar para no afectar otros asserts
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: init_importes_lm`.

- [ ] **Step 3: Implementar carga de importes y nivel de cascada**

Añadir global y helpers cerca de `_cache_lm`:

```python
_importes_lm = {}
def init_importes_lm(mapa):
    global _importes_lm
    _importes_lm = {c: {round(abs(x), 2) for x in s} for c, s in mapa.items()}

def match_importe(imp):
    """Devuelve el cod de tercero cuyo LM tiene un apunte del mismo importe, si es único."""
    objetivo = round(abs(float(imp)), 2)
    candidatos = [c for c, s in _importes_lm.items()
                  if any(abs(v - objetivo) <= 0.02 for v in s)]
    return candidatos[0] if len(candidatos) == 1 else None
```

Hacer que `cargar_lm` acumule importes por tercero. La estructura del mayor lista, tras la fila de cabecera de cada tercero, sus apuntes con importe en las columnas de debe/haber. Acumular el valor absoluto de cualquier celda numérica de las filas de apunte que cuelgan del tercero actual:

```python
def cargar_lm(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    terceros = {}
    importes = {}
    cod = None
    for row in ws.iter_rows(min_row=5, values_only=True):
        a = str(row[0]).strip() if row[0] else ''
        b = row[1]
        if a and b is None and a not in ['','None']:
            if not any(x in a for x in ['Saldo','Total','periodo','ejercicio']):
                p = a.split(' ',1); c = p[0].strip()
                nom = p[1].strip() if len(p)>1 else a
                if c[:2] in ('40','41','43'):
                    cod = c; terceros[cod] = nom; importes.setdefault(cod, set())
                else:
                    cod = None
        elif cod:
            # fila de apunte del tercero actual: guardar importes numéricos > 0
            for cell in row[1:]:
                if isinstance(cell, (int, float)) and abs(cell) > 0:
                    importes[cod].add(round(abs(float(cell)), 2))
    wb.close()
    init_importes_lm(importes)
    return terceros
```

Insertar el nivel en `clasificar_fila`, justo **después** del bloque `# 4. Libro Mayor` (match por nombre) y **antes** de `# 5. Excepciones`:

```python
    # 4b. Match por importe contra el LM (solo si criterios no lo desactiva)
    if criterios.get('match_por_importe', True):
        cod_imp = match_importe(imp)
        if cod_imp:
            return cod_imp, terceros.get(cod_imp, ''), 'ALTA', \
                   f"Match por importe: {abs(imp):.2f} → {cod_imp}", 'MATCH_IMPORTE'
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(match): nivel de match por importe contra el libro mayor (F1.1)"
```

---

### Task 5: Overrides de cuentas por criterios + regla TPV + nómina/SS flexibles

**Files:**
- Modify: `conciliar.py` (`R_OBS`, `clasificar_fila`, `_self_check`)

**Interfaces:**
- Consumes: `criterios` de Task 3.
- Produces: regla TPV (`LIQUIDACION REMESA DE COMERCIOS` → `criterios.cuenta_ventas_tpv` o `43000000`); patrones de nómina/SS más tolerantes; override de cuenta PRL vía `criterios.cuenta_prl`.

- [ ] **Step 1: Añadir los asserts que fallan**

En `_self_check()`:

```python
    # Task 5: TPV usa cuenta de criterios; nómina con texto intermedio; SS truncada
    _r = clasificar_fila(fila(conc='LIQUIDACION REMESA DE COMERCIOS', imp=1500.0),
                         {}, {'cuenta_ventas_tpv': '70500000'})
    assert _r[0] == '70500000'
    assert clasificar_fila(fila(obs='NOMINA LIQUIDACION DE JUNIO 2025'), {})[0] == '46500000'
    assert clasificar_fila(fila(obs='TESORERIA GENERAL DE LA SEGURID'), {})[0] == '47600000'
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `AssertionError` (TPV cae a genérico; nómina con "LIQUIDACION DE" no casa; SS truncada no casa).

- [ ] **Step 3: Flexibilizar patrones y añadir TPV con override**

Sustituir en `R_OBS` los dos patrones afectados y añadir el de TPV:

```python
    # Nóminas: permitir texto entre NOMINA y el mes
    (r'NOMIN[A-Z]*\b.{0,25}(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)',
     '46500000','REMUNERACIONES PENDIENTES DE PAGO','ALTA'),
    # Cotización SS: aceptar 'SEGURID' truncado
    (r'TGSS|TESORERIA\s+GENERAL.{0,20}SEGURID|COTIZACION\s+SS|CUOTA\s+SS|SEG\.?\s*SOCIAL',
     '47600000','ORGANISMOS SEGURIDAD SOCIAL, ACREEDORES','ALTA'),
    # TPV / liquidación remesa de comercios (cuenta destino por criterios)
    (r'LIQUIDACI[OO]N\s+REMESA\s+DE\s+COMERCIOS',
     '43000000','VENTAS TPV (segun cliente)','MEDIA'),
```

En `clasificar_fila`, dentro del bucle `for pat, cta, desc, conf in R_OBS:`, aplicar los overrides de criterios antes de devolver:

```python
    for pat, cta, desc, conf in R_OBS:
        if re.search(pat, texto_obs, re.IGNORECASE):
            # overrides por criterios de cliente
            if cta == '43000000' and criterios.get('cuenta_ventas_tpv'):
                cta = criterios['cuenta_ventas_tpv']
            if cta == '41000001' and criterios.get('cuenta_prl'):
                cta = criterios['cuenta_prl']
            fuente = obs_r.strip() or conc[:50]
            return cta, desc, conf, f"Detectado en: «{fuente}»", 'REGLA_OBS'
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(reglas): TPV con cuenta por criterios; nómina/SS tolerantes; PRL por cliente (F1.4/F1.5)"
```

---

### Task 6: Par retrocedido (+/−) → 555

**Files:**
- Modify: `conciliar.py` (nuevo `marcar_retrocesiones`, `conciliar`, `_self_check`)

**Interfaces:**
- Produces: `marcar_retrocesiones(df, resultados)` — post-pass que, para cada par de movimientos con importe opuesto (±0,02) y fechas próximas (≤5 días), reasigna ambos a `55500000` (método `RETROCESION`). Llamado dentro de `conciliar` tras la clasificación por filas.

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 6: par +/- retrocedido → 555
    import pandas as _pd
    _df = _pd.DataFrame([
        {'F_CONTABLE':'01/03/2025','F_VALOR':'','CODIGO':'','CONCEPTO':'CONFIRMING',
         'BENEFICIARIO':'','OBSERVACIONES':'','IMPORTE':300.0,'SALDO':0.0},
        {'F_CONTABLE':'02/03/2025','F_VALOR':'','CODIGO':'','CONCEPTO':'RETROCESION CONFIRMING',
         'BENEFICIARIO':'','OBSERVACIONES':'','IMPORTE':-300.0,'SALDO':0.0},
    ])
    _res = [{'idx':0,'CUENTA':'X','DESCRIPCION':'','CONFIANZA':'BAJA','JUSTIFICACION':'','METODO':'x'},
            {'idx':1,'CUENTA':'Y','DESCRIPCION':'','CONFIANZA':'BAJA','JUSTIFICACION':'','METODO':'y'}]
    marcar_retrocesiones(_df, _res)
    assert _res[0]['CUENTA'] == '55500000' and _res[1]['CUENTA'] == '55500000'
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: marcar_retrocesiones`.

- [ ] **Step 3: Implementar el post-pass**

Añadir la función y llamarla en `conciliar` justo antes de `resultados.sort(...)`:

```python
def marcar_retrocesiones(df, resultados):
    """Empareja movimientos de importe opuesto y fecha próxima → 555 (netean a 0)."""
    from datetime import datetime as _dt
    def _fecha(s):
        for fmt in ('%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try: return _dt.strptime(str(s).strip()[:19], fmt)
            except Exception: pass
        return None
    res_map = {r['idx']: r for r in resultados}
    usados = set()
    filas = list(df.iterrows())
    for a in range(len(filas)):
        ia, ra = filas[a]
        if ia in usados: continue
        fa = _fecha(ra['F_CONTABLE']); impa = round(float(ra['IMPORTE']), 2)
        if impa == 0: continue
        for b in range(a+1, len(filas)):
            ib, rb = filas[b]
            if ib in usados: continue
            if round(float(rb['IMPORTE']), 2) != -impa: continue
            fb = _fecha(rb['F_CONTABLE'])
            if fa and fb and abs((fb - fa).days) > 5: continue
            for idx in (ia, ib):
                r = res_map.get(idx)
                if r:
                    r.update({'CUENTA':'55500000',
                              'DESCRIPCION':'PARTIDAS PENDIENTES DE APLICACION',
                              'CONFIANZA':'MEDIA',
                              'JUSTIFICACION':'Par +/- retrocedido (netea a 0)',
                              'METODO':'RETROCESION'})
            usados.add(ia); usados.add(ib)
            break
```

En `conciliar`, antes de `resultados.sort(key=lambda x: x['idx'])`:

```python
    marcar_retrocesiones(df, resultados)
    resultados.sort(key=lambda x: x['idx'])
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(reglas): detectar par retrocedido +/- y llevar a 555 (F1.7)"
```

---

## FASE 2 — Entrada PDF + salida Cegid

### Task 7: Check de continuidad de saldo

**Files:**
- Modify: `conciliar.py` (nuevo `validar_saldo`, `_self_check`)

**Interfaces:**
- Produces: `validar_saldo(df)` → `list[int]` de índices sospechosos donde `saldo[i] != saldo[i+1] + importe[i]` (orden descendente, tolerancia 0,02). Lista vacía = parse íntegro. Consumido por Task 8 y Task 10 (avisos).

- [ ] **Step 1: Añadir los asserts que fallan**

En `_self_check()`:

```python
    # Task 7: validar_saldo detecta roturas
    import pandas as _pd
    _ok = _pd.DataFrame([{'IMPORTE':-10.0,'SALDO':90.0},{'IMPORTE':-5.0,'SALDO':100.0}])
    assert validar_saldo(_ok) == []          # 90 == 100 + (-10)
    _bad = _pd.DataFrame([{'IMPORTE':-10.0,'SALDO':80.0},{'IMPORTE':-5.0,'SALDO':100.0}])
    assert validar_saldo(_bad) == [0]        # 80 != 100 + (-10)
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: validar_saldo`.

- [ ] **Step 3: Implementar `validar_saldo`**

```python
def validar_saldo(df):
    """Índices i donde saldo[i] != saldo[i+1] + importe[i] (orden descendente, tol 0,02)."""
    roturas = []
    for i in range(len(df) - 1):
        esperado = float(df['SALDO'].iloc[i+1]) + float(df['IMPORTE'].iloc[i])
        if abs(esperado - float(df['SALDO'].iloc[i])) > 0.02:
            roturas.append(i)
    return roturas
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(integridad): check de continuidad de saldo (F2.2)"
```

---

### Task 8: Soporte de extractos PDF (parser BBVA)

**Files:**
- Modify: `conciliar.py` (nuevos `_agrupar_lineas_pdf`, `_parsear_movimientos_pdf`, `cargar_extracto_pdf`; rama `.pdf` en `cargar_extracto`; `_self_check`), `requirements.txt`

**Interfaces:**
- Consumes: `validar_saldo` (Task 7).
- Produces:
  - `_parsear_movimientos_pdf(lineas)` — recibe `list[(page:int, top:float, text:str)]` y devuelve `list[dict]` con las claves estándar del extracto. **Testeable sin pdfplumber.**
  - `cargar_extracto_pdf(path)` — abre el PDF con pdfplumber, extrae líneas y delega en `_parsear_movimientos_pdf`. Devuelve `pd.DataFrame`.
  - `cargar_extracto(path, ...)` enruta `.pdf` a `cargar_extracto_pdf`.

- [ ] **Step 1: Añadir `pdfplumber` a requirements**

Editar `requirements.txt`:

```
openpyxl
pandas
python-dateutil
pdfplumber
# anthropic  # opcional: solo si se define API_KEY para refinar acreedores genéricos
```

Instalar: `python3 -m pip install pdfplumber`

- [ ] **Step 2: Añadir el assert que falla (parser sobre líneas sintéticas, sin PDF real)**

En `_self_check()`:

```python
    # Task 8: parser PDF sobre líneas sintéticas BBVA (concepto|beneficiario arriba, obs abajo)
    _lineas = [
        (0, 10.0, "PAGO NOMINA | JUAN PEREZ"),
        (0, 20.0, "05/06/2025 05/06/2025 1234 -1.200,00 EUR 3.800,00 EUR"),
        (0, 30.0, "TRANSFERENCIA NOMINA JUNIO"),
        (0, 60.0, "COMPRA TARJETA | MERCADONA"),           # gap 30 > 18 → nuevo movimiento
        (0, 70.0, "06/06/2025 06/06/2025 5678 -45,20 EUR 3.754,80 EUR"),
    ]
    _movs = _parsear_movimientos_pdf(_lineas)
    assert len(_movs) == 2
    assert _movs[0]['IMPORTE'] == -1200.0 and _movs[0]['BENEFICIARIO'] == 'JUAN PEREZ'
    assert _movs[0]['CONCEPTO'] == 'PAGO NOMINA'
    assert _movs[1]['IMPORTE'] == -45.2 and _movs[1]['CONCEPTO'] == 'COMPRA TARJETA'
```

- [ ] **Step 3: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: _parsear_movimientos_pdf`.

- [ ] **Step 4: Implementar parser y ruteo**

Añadir cerca de `cargar_extracto` (usa el `_smart_float` existente y un normalizador local):

```python
_NUM_PDF = re.compile(
    r'^(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d{4,5})'
    r'(?:\s+(.+?))??\s+(-?[\d\.]*,\d{1,2})\s+EUR\s+(-?[\d\.]*,\d{1,2})\s+EUR\s*$')
_SKIP_PDF = ('Movimientos','Titular ','Cuenta ES','Divisa ','Banco BANCO','F. CONTABLE','F.CONTABLE')

def _norm_pdf(t):
    import unicodedata
    if not t: return ''
    t = unicodedata.normalize('NFD', str(t)).encode('ascii','ignore').decode('ascii')
    return re.sub(r'\s+', ' ', t).strip()

def _parsear_movimientos_pdf(lineas):
    """lineas: list[(page, top, text)] ya filtradas. Devuelve list[dict] estándar."""
    # clustering por salto vertical (>18 = nuevo movimiento); corta también al cambiar de página
    movs, cur, pp, pt = [], [], None, None
    for pi, t, txt in lineas:
        if cur and (pi != pp or pt is None or t - pt > 18):
            movs.append(cur); cur = []
        cur.append((t, txt)); pp, pt = pi, t
    if cur: movs.append(cur)
    rows = []
    for mv in movs:
        ni = m = None
        for i, (_, txt) in enumerate(mv):
            mm = _NUM_PDF.match(txt)
            if mm: ni, m = i, mm; break
        if ni is None: continue
        fc, fv, cod, cmid, imp, saldo = m.groups()
        head = _norm_pdf(' '.join(txt for _, txt in mv[:ni]) + (' ' + cmid if cmid else ''))
        obs = _norm_pdf(' '.join(txt for _, txt in mv[ni+1:]))
        conc, benef = (head.split('|', 1) + [''])[:2] if '|' in head else (head, '')
        rows.append({'F_CONTABLE': fc, 'F_VALOR': fv, 'CODIGO': cod,
                     'CONCEPTO': _norm_pdf(conc), 'BENEFICIARIO': _norm_pdf(benef),
                     'OBSERVACIONES': obs, 'IMPORTE': _smart_float(imp),
                     'SALDO': _smart_float(saldo)})
    return rows

def cargar_extracto_pdf(path):
    import pdfplumber
    from collections import defaultdict
    lineas = []
    with pdfplumber.open(path) as pdf:
        for pi, pg in enumerate(pdf.pages):
            grp = defaultdict(list)
            for w in pg.extract_words(use_text_flow=False, keep_blank_chars=False):
                grp[round(w['top'])].append(w)
            for t in sorted(grp):
                txt = ' '.join(w['text'] for w in sorted(grp[t], key=lambda x: x['x0']))
                if re.match(r'^\d+/\d+$', txt): continue
                if any(txt.startswith(s) for s in _SKIP_PDF): continue
                lineas.append((pi, t, txt))
    rows = _parsear_movimientos_pdf(lineas)
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=['F_CONTABLE','F_VALOR','CODIGO','CONCEPTO','BENEFICIARIO',
                 'OBSERVACIONES','IMPORTE','SALDO'])
    roturas = validar_saldo(df)
    if roturas:
        print(f"AVISO: {len(roturas)} posible(s) rotura(s) de saldo en el PDF (filas {roturas[:10]}...)")
    return df
```

En `cargar_extracto`, al principio, enrutar PDF:

```python
def cargar_extracto(path, fecha_desde=None, fecha_hasta=None):
    if path.lower().endswith('.pdf'):
        return cargar_extracto_pdf(path)   # el filtro de fechas se aplica aguas arriba si hace falta
    import unicodedata, shutil
    # ... (resto igual)
```

- [ ] **Step 5: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 6: Commit**

```bash
git add conciliar.py requirements.txt
git commit -m "feat(pdf): soporte de extractos PDF con parser BBVA testeable (F2.1)"
```

---

### Task 9: Hoja CEGID de 4 columnas

**Files:**
- Modify: `conciliar.py` (`generar_excel`, `_self_check`)

**Interfaces:**
- Produces: `generar_excel` añade una hoja `CEGID` con exactamente `Fecha | Concepto | Importe | Contrapartida`. Nueva helper `_hoja_cegid(wb, df, resultados)` para poder testearla aislada.

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 9: hoja CEGID con 4 columnas exactas
    import pandas as _pd
    from openpyxl import Workbook as _WB2
    _wb2 = _WB2()
    _df2 = _pd.DataFrame([{'F_CONTABLE':'05/06/2025','CONCEPTO':'PAGO','IMPORTE':-10.0,
                           'BENEFICIARIO':'','OBSERVACIONES':'','SALDO':0.0}])
    _res2 = [{'idx':0,'CUENTA':'41000001','DESCRIPCION':'','CONFIANZA':'ALTA',
              'JUSTIFICACION':'','METODO':'x'}]
    _hoja_cegid(_wb2, _df2, _res2)
    _ws2 = _wb2['CEGID']
    assert [c.value for c in _ws2[1]] == ['Fecha','Concepto','Importe','Contrapartida']
    assert [c.value for c in _ws2[2]] == ['05/06/2025','PAGO',-10.0,'41000001']
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: _hoja_cegid`.

- [ ] **Step 3: Implementar `_hoja_cegid` y llamarla en `generar_excel`**

```python
def _hoja_cegid(wb, df, resultados):
    if 'CEGID' in wb.sheetnames: del wb['CEGID']
    cg = wb.create_sheet('CEGID')
    cg.append(['Fecha', 'Concepto', 'Importe', 'Contrapartida'])
    res_map = {r['idx']: r for r in resultados}
    for idx in range(len(df)):
        r = res_map.get(idx)
        if not r: continue
        row = df.iloc[idx]
        cg.append([row['F_CONTABLE'], row['CONCEPTO'], row['IMPORTE'], r['CUENTA']])
    for col, w in zip('ABCD', (14, 50, 14, 16)):
        cg.column_dimensions[col].width = w
```

En `generar_excel`, tras crear la hoja `REVISAR` y antes de `wb.save(output_path)`:

```python
    _hoja_cegid(wb, df, resultados)
    wb.save(output_path)
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(cegid): hoja CEGID con 4 columnas para el selector de importación (F2.3)"
```

---

## FASE 3 — Batch multi-empresa + acciones + informes

### Task 10: Detección de acciones (terceros a crear, subcuentas, avisos)

**Files:**
- Modify: `conciliar.py` (nuevos `siguiente_codigo_libre`, `detectar_acciones`, `_self_check`)

**Interfaces:**
- Consumes: `terceros` (Task 4), `pgc` con índice 8 díg (Task 1), `validar_saldo` (Task 7).
- Produces:
  - `siguiente_codigo_libre(terceros, bloque)` — `bloque` in `('400','410','430')` → siguiente código de 8 díg libre (string).
  - `detectar_acciones(df, resultados, terceros, pgc)` → `dict` con listas `terceros_crear`, `subcuentas_crear`, `avisos`.

- [ ] **Step 1: Añadir los asserts que fallan**

En `_self_check()`:

```python
    # Task 10: siguiente código libre por bloque
    assert siguiente_codigo_libre({'40000001':'A','40000005':'B'}, '400') == '40000006'
    assert siguiente_codigo_libre({}, '410') == '41000001'
    # detectar_acciones agrupa genéricos y propone código; detecta subcuenta inexistente
    import pandas as _pd
    _df3 = _pd.DataFrame([
        {'F_CONTABLE':'01/06/2025','CONCEPTO':'PAGO','BENEFICIARIO':'FERRETERIA LOPEZ',
         'OBSERVACIONES':'FERRETERIA LOPEZ','IMPORTE':-80.0,'SALDO':0.0},
    ])
    _res3 = [{'idx':0,'CUENTA':'41000000','DESCRIPCION':'','CONFIANZA':'BAJA',
              'JUSTIFICACION':'','METODO':'ACREEDOR_GENERICO'}]
    _acc = detectar_acciones(_df3, _res3, {}, {})
    assert len(_acc['terceros_crear']) == 1
    assert _acc['terceros_crear'][0]['codigo'] == '41000001'
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: siguiente_codigo_libre`.

- [ ] **Step 3: Implementar helpers**

```python
def siguiente_codigo_libre(terceros, bloque):
    """bloque: '400'|'410'|'430'. Siguiente subcuenta de 8 díg libre en ese bloque."""
    usados = [int(c) for c in terceros
              if c.isdigit() and len(c) == 8 and c.startswith(bloque)]
    base = int(bloque + '00000')  # p.ej. 40000000
    return str((max(usados) + 1) if usados else base + 1)

def detectar_acciones(df, resultados, terceros, pgc):
    """Terceros a crear (genéricos agrupados), subcuentas inexistentes y avisos."""
    pgc = pgc or {}
    acc = {'terceros_crear': [], 'subcuentas_crear': [], 'avisos': []}
    # 1) terceros a crear: agrupar los ACREEDOR_GENERICO por nombre normalizado
    grupos = {}
    for r in resultados:
        if r['METODO'] != 'ACREEDOR_GENERICO':
            continue
        row = df.iloc[r['idx']]
        nombre = (str(row['BENEFICIARIO']).strip() or str(row['OBSERVACIONES']).strip()
                  or str(row['CONCEPTO']).strip())[:50]
        clave = ' '.join(sorted(toks(nombre))) or nombre.upper()
        grupos.setdefault(clave, {'nombre': nombre, 'n': 0, 'importe': 0.0})
        grupos[clave]['n'] += 1
        grupos[clave]['importe'] += float(row['IMPORTE'])
    reservados = dict(terceros)
    for g in grupos.values():
        bloque = '410' if g['importe'] < 0 else '430'  # pagos→acreedor, cobros→cliente
        cod = siguiente_codigo_libre(reservados, bloque)
        reservados[cod] = g['nombre']
        acc['terceros_crear'].append({
            'codigo': cod, 'nombre': g['nombre'],
            'tipo': 'acreedor' if bloque == '410' else 'cliente',
            'nif': '(pendiente)', 'n_movimientos': g['n']})
    # 2) subcuentas de contrapartida que no existen en el plan
    vistos = set()
    for r in resultados:
        cta = r['CUENTA']
        if cta in ('PENDIENTE',) or not str(cta).isdigit():
            continue
        if cta in vistos or cta in pgc:
            continue
        vistos.add(cta)
        acc['subcuentas_crear'].append({'codigo': cta, 'nombre': r['DESCRIPCION']})
    # 3) avisos: movimientos sin nombre/NIF y roturas de saldo
    for r in resultados:
        row = df.iloc[r['idx']]
        if r['METODO'] == 'ACREEDOR_GENERICO' and not str(row['BENEFICIARIO']).strip() \
           and not str(row['OBSERVACIONES']).strip():
            acc['avisos'].append({'idx': r['idx'], 'motivo': 'Sin beneficiario ni observaciones',
                                  'importe': float(row['IMPORTE'])})
    for i in validar_saldo(df):
        acc['avisos'].append({'idx': i, 'motivo': 'Posible rotura de saldo', 'importe': None})
    return acc
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(acciones): detectar terceros a crear, subcuentas y avisos (F3.2)"
```

---

### Task 11: Hoja ACCIONES_CEGID

**Files:**
- Modify: `conciliar.py` (`generar_excel` acepta `acciones=None`; nueva `_hoja_acciones`; `_self_check`)

**Interfaces:**
- Consumes: salida de `detectar_acciones` (Task 10).
- Produces: `generar_excel(df, resultados, extracto_path, output_path, acciones=None)`; si `acciones`, escribe hoja `ACCIONES_CEGID` con 3 bloques.

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 11: hoja ACCIONES_CEGID con los 3 bloques
    from openpyxl import Workbook as _WB3
    _wb3 = _WB3()
    _acc3 = {'terceros_crear':[{'codigo':'41000001','nombre':'FERRETERIA LOPEZ',
              'tipo':'acreedor','nif':'(pendiente)','n_movimientos':2}],
             'subcuentas_crear':[{'codigo':'62800000','nombre':'SUMINISTROS'}],
             'avisos':[{'idx':3,'motivo':'Sin beneficiario','importe':-9.0}]}
    _hoja_acciones(_wb3, _acc3)
    _wsa = _wb3['ACCIONES_CEGID']
    _textos = [str(c.value) for r in _wsa.iter_rows() for c in r if c.value]
    assert any('FERRETERIA LOPEZ' in t for t in _textos)
    assert any('62800000' in t for t in _textos)
    assert any('TERCEROS A CREAR' in t.upper() for t in _textos)
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: _hoja_acciones`.

- [ ] **Step 3: Implementar `_hoja_acciones` y engancharla**

```python
def _hoja_acciones(wb, acciones):
    if 'ACCIONES_CEGID' in wb.sheetnames: del wb['ACCIONES_CEGID']
    aw = wb.create_sheet('ACCIONES_CEGID')
    aw.append(['ANTES DE IMPORTAR EN CEGID — hacer en este orden'])
    aw.append([])
    aw.append(['1) TERCEROS A CREAR'])
    aw.append(['Codigo', 'Nombre', 'Tipo', 'NIF', 'Nº movs'])
    for t in acciones.get('terceros_crear', []):
        aw.append([t['codigo'], t['nombre'], t['tipo'], t['nif'], t['n_movimientos']])
    aw.append([])
    aw.append(['2) SUBCUENTAS A CREAR (no existen en el plan)'])
    aw.append(['Codigo', 'Nombre'])
    for s in acciones.get('subcuentas_crear', []):
        aw.append([s['codigo'], s['nombre']])
    aw.append([])
    aw.append(['3) AVISOS (revisar antes de importar)'])
    aw.append(['Fila', 'Motivo', 'Importe'])
    for a in acciones.get('avisos', []):
        aw.append([a['idx'], a['motivo'], a['importe']])
    for col, w in zip('ABCDE', (16, 40, 12, 16, 10)):
        aw.column_dimensions[col].width = w
```

Cambiar la firma de `generar_excel` a `def generar_excel(df, resultados, extracto_path, output_path, acciones=None):` y, tras `_hoja_cegid(wb, df, resultados)`:

```python
    if acciones:
        _hoja_acciones(wb, acciones)
    wb.save(output_path)
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(acciones): hoja ACCIONES_CEGID con terceros/subcuentas/avisos (F3.2)"
```

---

### Task 12: Orquestador de lote `procesar_lote`

**Files:**
- Modify: `conciliar.py` (nuevos `_detectar_ficheros_empresa`, `procesar_lote`; `_self_check`)

**Interfaces:**
- Consumes: `cargar_extracto`, `cargar_lm`, `cargar_pgc`, `cargar_criterios`, `conciliar`, `detectar_acciones`, `generar_excel`.
- Produces:
  - `_detectar_ficheros_empresa(carpeta)` → `dict(mayor, plan, extractos:list)`.
  - `procesar_lote(carpeta_madre)` → `list[dict]` resumen por empresa (`empresa, n_mov, pct_alta, pct_media, n_terceros, n_avisos, salida`). Genera un Excel por empresa (columna `BANCO`) y el informe global (Task 13).

- [ ] **Step 1: Añadir el assert que falla (solo `_detectar_ficheros_empresa`, sin I/O de Excel)**

En `_self_check()`:

```python
    # Task 12: detección de roles de fichero por nombre
    _fich = _detectar_ficheros_empresa.__wrapped__ if hasattr(_detectar_ficheros_empresa,'__wrapped__') else _detectar_ficheros_empresa
    _res12 = _clasificar_nombres(['LIBRO MAYOR 2025.xlsx','PLAN CONTABLE.xlsx',
                                  'BBVA movimientos.pdf','CaixaBank.csv'])
    assert _res12['mayor'] == 'LIBRO MAYOR 2025.xlsx'
    assert _res12['plan'] == 'PLAN CONTABLE.xlsx'
    assert set(_res12['extractos']) == {'BBVA movimientos.pdf','CaixaBank.csv'}
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: _clasificar_nombres`.

- [ ] **Step 3: Implementar clasificación de nombres, detección y orquestador**

```python
def _clasificar_nombres(nombres):
    """Reparte una lista de nombres de fichero en mayor / plan / extractos."""
    roles = {'mayor': None, 'plan': None, 'extractos': []}
    for n in nombres:
        low = n.lower()
        ext = low.rsplit('.', 1)[-1] if '.' in low else ''
        if ext not in ('xlsx', 'xls', 'csv', 'pdf'):
            continue
        if 'mayor' in low and roles['mayor'] is None:
            roles['mayor'] = n
        elif ('plan' in low or 'cuenta' in low) and roles['plan'] is None:
            roles['plan'] = n
        else:
            roles['extractos'].append(n)
    return roles

def _detectar_ficheros_empresa(carpeta):
    roles = _clasificar_nombres(sorted(os.listdir(carpeta)))
    j = lambda x: os.path.join(carpeta, x) if x else None
    return {'mayor': j(roles['mayor']), 'plan': j(roles['plan']),
            'extractos': [j(e) for e in roles['extractos']]}

def procesar_lote(carpeta_madre):
    resumen = []
    for nombre in sorted(os.listdir(carpeta_madre)):
        sub = os.path.join(carpeta_madre, nombre)
        if not os.path.isdir(sub):
            continue
        fich = _detectar_ficheros_empresa(sub)
        if not fich['mayor'] or not fich['plan'] or not fich['extractos']:
            print(f"SALTO {nombre}: falta mayor/plan/extracto")
            continue
        terceros = cargar_lm(fich['mayor'])
        pgc = cargar_pgc(fich['plan'])
        criterios = cargar_criterios(sub)
        # unir todos los bancos con columna BANCO
        dfs = []
        for ext in fich['extractos']:
            d = cargar_extracto(ext)
            d['BANCO'] = os.path.splitext(os.path.basename(ext))[0]
            dfs.append(d)
        df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        if df.empty:
            print(f"SALTO {nombre}: extracto vacío")
            continue
        res = conciliar(df, terceros, pgc, criterios)
        acciones = detectar_acciones(df, res, terceros, pgc)
        salida = os.path.join(sub, f"{nombre}_conciliado.xlsx")
        generar_excel(df, res, fich['extractos'][0], salida, acciones)
        total = len(res)
        pct = lambda c: sum(1 for r in res if r['CONFIANZA'] == c) / total * 100 if total else 0
        resumen.append({'empresa': nombre, 'n_mov': total,
                        'pct_alta': round(pct('ALTA'), 1), 'pct_media': round(pct('MEDIA'), 1),
                        'n_terceros': len(acciones['terceros_crear']),
                        'n_avisos': len(acciones['avisos']), 'salida': salida})
        print(f"OK {nombre}: {total} movs → {salida}")
    generar_informe(carpeta_madre, resumen)
    return resumen
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK` (nota: `generar_informe` se define en Task 13; hasta entonces el self-check de esta tarea no la invoca, pero el `import`/def de `procesar_lote` sí la referencia — definir un stub `def generar_informe(*a, **k): pass` en esta tarea y sustituirlo en Task 13).

Añadir el stub temporal antes de `procesar_lote`:

```python
def generar_informe(carpeta_madre, resumen):  # stub, se completa en Task 13
    pass
```

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(batch): procesar_lote recorre empresas y genera un Excel por empresa (F3.1)"
```

---

### Task 13: Informe global

**Files:**
- Modify: `conciliar.py` (`generar_informe` real; `_self_check`)

**Interfaces:**
- Consumes: `resumen` de `procesar_lote`.
- Produces: `generar_informe(carpeta_madre, resumen)` escribe `INFORME_YYYYMMDD.md` (fecha vía `datetime.now()`); devuelve la ruta.

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 13: informe global se escribe y contiene la fila de empresa
    import tempfile as _tf, os as _os2
    _dir = _tf.mkdtemp()
    _ruta = generar_informe(_dir, [{'empresa':'AMAYA','n_mov':100,'pct_alta':80.0,
        'pct_media':15.0,'n_terceros':3,'n_avisos':2,'salida':'x.xlsx'}])
    _txt = open(_ruta, encoding='utf-8').read()
    assert 'AMAYA' in _txt and '80.0' in _txt
    _os2.remove(_ruta)
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `AssertionError` (el stub devuelve `None` → `open(None)` lanza antes; ajustar: el stub falla el assert). Confirmar fallo.

- [ ] **Step 3: Sustituir el stub por la implementación real**

```python
def generar_informe(carpeta_madre, resumen):
    fecha = datetime.now().strftime('%Y%m%d')
    ruta = os.path.join(carpeta_madre, f"INFORME_{fecha}.md")
    lineas = [f"# Informe de conciliación — {datetime.now().strftime('%d/%m/%Y %H:%M')}", "",
              "| Empresa | Movs | % ALTA | % MEDIA | Terceros a crear | Avisos |",
              "|---|---:|---:|---:|---:|---:|"]
    for r in resumen:
        lineas.append(f"| {r['empresa']} | {r['n_mov']} | {r['pct_alta']} | "
                      f"{r['pct_media']} | {r['n_terceros']} | {r['n_avisos']} |")
    lineas += ["", "## Orden de trabajo", ""]
    for r in resumen:
        if r['n_terceros'] or r['n_avisos']:
            lineas.append(f"- **{r['empresa']}**: crear {r['n_terceros']} tercero(s), "
                          f"revisar {r['n_avisos']} aviso(s) — ver hoja ACCIONES_CEGID de "
                          f"`{os.path.basename(r['salida'])}`")
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lineas) + '\n')
    return ruta
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(informe): informe global INFORME_YYYYMMDD.md por lote (F3.3)"
```

---

### Task 14: Selección de dudas de alto impacto + modo lote en `__main__`

**Files:**
- Modify: `conciliar.py` (nuevo `dudas_alto_impacto`; `__main__` admite `CARPETA_MADRE`; `_self_check`)

**Interfaces:**
- Produces:
  - `dudas_alto_impacto(df, resultados, top=15, min_repeticiones=3)` → `list[dict]` de movimientos REVISAR/BAJA de alto impacto (top por `|importe|` ∪ conceptos/proveedores repetidos ≥ `min_repeticiones`). La skill (SKILL.md) usa esta lista para la ronda final única.
  - `__main__`: si `CARPETA_MADRE` está definida y existe, ejecuta `procesar_lote`; si no, el flujo de una empresa; si no hay ficheros, self-check.

- [ ] **Step 1: Añadir el assert que falla**

En `_self_check()`:

```python
    # Task 14: dudas de alto impacto (por importe y por recurrencia)
    import pandas as _pd
    _rows = [{'F_CONTABLE':'','CONCEPTO':f'C{i}','BENEFICIARIO':'','OBSERVACIONES':'',
              'IMPORTE':-(i+1)*100.0,'SALDO':0.0} for i in range(20)]
    _df14 = _pd.DataFrame(_rows)
    _res14 = [{'idx':i,'CUENTA':'41000000','DESCRIPCION':'','CONFIANZA':'REVISAR',
               'JUSTIFICACION':'','METODO':'x'} for i in range(20)]
    _dudas = dudas_alto_impacto(_df14, _res14, top=5)
    assert len(_dudas) == 5
    assert _dudas[0]['idx'] == 19          # el de mayor importe primero
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: `python3 conciliar.py`
Expected: `NameError: dudas_alto_impacto`.

- [ ] **Step 3: Implementar selección y modo lote**

```python
def dudas_alto_impacto(df, resultados, top=15, min_repeticiones=3):
    dudosos = [r for r in resultados if r['CONFIANZA'] in ('REVISAR', 'BAJA')]
    def imp(r): return abs(float(df.iloc[r['idx']]['IMPORTE']))
    por_importe = sorted(dudosos, key=imp, reverse=True)[:top]
    # recurrentes: mismo concepto normalizado repetido
    from collections import Counter
    claves = Counter(str(df.iloc[r['idx']]['CONCEPTO']).upper().strip() for r in dudosos)
    recurrentes = [r for r in dudosos
                   if claves[str(df.iloc[r['idx']]['CONCEPTO']).upper().strip()] >= min_repeticiones]
    seleccion, vistos = [], set()
    for r in por_importe + recurrentes:
        if r['idx'] in vistos: continue
        vistos.add(r['idx'])
        row = df.iloc[r['idx']]
        seleccion.append({'idx': r['idx'], 'concepto': row['CONCEPTO'],
                          'importe': float(row['IMPORTE']), 'cuenta_propuesta': r['CUENTA']})
    return seleccion
```

En `__main__`, añadir la rama de lote (una constante `CARPETA_MADRE = None` en CONFIGURACIÓN):

```python
    if CARPETA_MADRE and os.path.isdir(CARPETA_MADRE):
        procesar_lote(CARPETA_MADRE)
    elif not all(os.path.exists(p) for p in (EXTRACTO_PATH, LM_PATH, PGC_PATH)):
        _self_check()
    else:
        # ... (flujo de una empresa actual)
```

Y añadir `CARPETA_MADRE = None` en el bloque CONFIGURACIÓN.

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 5: Commit**

```bash
git add conciliar.py
git commit -m "feat(batch): dudas de alto impacto para la ronda final + modo lote en CLI (F3.4)"
```

---

### Task 15: Documentar en SKILL.md/README y actualizar el repo remoto

**Files:**
- Modify: `SKILL.md`, `README.md`

**Interfaces:** (documentación; sin código nuevo)

- [ ] **Step 1: Actualizar `SKILL.md`**

Añadir al final de la sección técnica (tras la descripción de `conciliar.py`) un bloque nuevo:

```markdown
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
```

- [ ] **Step 2: Actualizar `README.md`**

Añadir bajo "Uso" una línea sobre el modo lote y el `criterios.json`, y añadir `pdfplumber`
a la lista de dependencias mencionada.

- [ ] **Step 3: Verificar el self-check completo una última vez**

Run: `python3 conciliar.py`
Expected: `self-check OK`

- [ ] **Step 4: Commit y push (actualizar el repo de GitHub)**

```bash
git add SKILL.md README.md
git commit -m "docs: modo lote, criterios.json, PDF y cascada actualizada (F3)"
git push origin main
```

- [ ] **Step 5: Recordatorio de despliegue**

Donde la skill esté clonada (PC/servidor Windows):

```bash
cd <ruta>/.claude/skills/conciliador-bancario-cegid
git pull
python -m pip install -r requirements.txt   # trae pdfplumber
python conciliar.py                          # self-check OK
```

---

## Self-Review (cobertura del spec)

- **F1.1** match por importe → Task 4 ✓ · **F1.2** criterios.json + ALIAS → Task 3 ✓ ·
  **F1.3** tokenización → Task 2 ✓ · **F1.4** TPV → Task 5 ✓ · **F1.5** nómina/SS → Task 5 ✓ ·
  **F1.6** PGC 8 díg → Task 1 ✓ · **F1.7** retrocesión 555 → Task 6 ✓
- **F2.1** PDF → Task 8 ✓ · **F2.2** check saldo → Task 7 ✓ · **F2.3** hoja CEGID → Task 9 ✓
- **F3.1** procesar_lote → Task 12 ✓ · **F3.2** acciones (terceros/subcuentas/avisos) → Tasks 10–11 ✓ ·
  **F3.3** informe global → Task 13 ✓ · **F3.4** ronda final alto impacto → Task 14 ✓
- **Docs + push repo** → Task 15 ✓

Consistencia de tipos: `clasificar_fila(row, terceros, criterios=None)` usado igual en Tasks 3–6;
`conciliar(df, terceros, pgc, criterios=None, api_key='', lote=30)` en Tasks 3/4/6/12;
`generar_excel(df, resultados, extracto_path, output_path, acciones=None)` en Tasks 9/11/12;
`detectar_acciones(df, resultados, terceros, pgc)` en Tasks 10/12. Sin placeholders pendientes.
```
