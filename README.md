# Conciliador Bancario Cegid — EMETE Asesoría

Skill de Claude Code que **concilia extractos bancarios** contra el libro mayor y el
plan de cuentas de Cegid Diez: asigna a cada movimiento su **cuenta de contrapartida**
(PGC 2007) y genera un Excel con la columna `CUENTA_CONTRAPARTIDA` lista para importar,
más una hoja `REVISAR` con lo que necesita criterio manual.

Detecta el formato del banco automáticamente (BBVA, CaixaBank, Santander, Sabadell…),
cruza con los terceros del libro mayor (400/410/430) y aplica una cascada de reglas
contables.

## Instalación

Se instala clonando este repo dentro de la carpeta de skills de Claude Code.
Pasos detallados (Windows y macOS) en [INSTALL.md](INSTALL.md).

Resumen (macOS):
```
cd ~/.claude/skills
git clone git@github.com:manuelgt85/conciliador-bancario-cegid.git
cd conciliador-bancario-cegid
python3 -m pip install -r requirements.txt
python3 conciliar.py          # debe decir: self-check OK
```

## Uso

Abre Claude Code y escribe algo como *"concilia el extracto de CaixaBank de [cliente]"*.
Claude te pedirá los **tres ficheros**: extracto bancario, libro mayor y plan de cuentas.

**La documentación de cada empresa vive en una carpeta que tú le das** en cada
conciliación — no se guarda en este repo. El `.gitignore` está configurado para no
subir nunca `.xlsx`, `.csv`, `.pdf` ni imágenes.

## Contenido

- `SKILL.md` — el procedimiento que ejecuta Claude Code.
- `conciliar.py` — toda la lógica (parser, auto-detección de formato, matching con el
  libro mayor, reglas de clasificación PGC). Incluye self-check.
- `requirements.txt` — dependencias Python (openpyxl, pandas, python-dateutil).

**Este repositorio no contiene datos de clientes.**
