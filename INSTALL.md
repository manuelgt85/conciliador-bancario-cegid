# Instalar la skill "conciliador-bancario-cegid"

Se instala **clonando este repositorio** dentro de la carpeta de skills de Claude Code.
Sirve igual para el PC Windows y el servidor Windows.

## Windows (PC y servidor)

**1. Python** (si no lo tienes):
```
python --version
```
Si da error: `winget install Python.Python.3` (o desde https://www.python.org/downloads/
marcando **"Add python.exe to PATH"**).

**2. Clonar el repo** en la carpeta de skills:
```
mkdir C:\Users\<usuario>\.claude\skills
cd C:\Users\<usuario>\.claude\skills
git clone https://github.com/manuelgt85/conciliador-bancario-cegid.git
```
(el `mkdir` solo hace falta si la carpeta no existe todavía).

**3. Instalar dependencias**:
```
cd conciliador-bancario-cegid
python -m pip install -r requirements.txt
```

**4. Comprobar**:
```
python conciliar.py
```
Debe imprimir `self-check OK`. Ya está: abre Claude Code y pídele *"concilia el extracto de [banco] de [cliente]"*.

## Actualizar a la última versión
```
cd C:\Users\<usuario>\.claude\skills\conciliador-bancario-cegid
git pull
```

## macOS
```
cd ~/.claude/skills
git clone git@github.com:manuelgt85/conciliador-bancario-cegid.git
cd conciliador-bancario-cegid
python3 -m pip install -r requirements.txt
python3 conciliar.py
```
