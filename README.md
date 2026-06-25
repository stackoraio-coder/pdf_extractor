# Extractor PDF escaneado a filas para Excel

App FastAPI para probar la extracción de PDFs escaneados y convertir cada PDF en una fila revisable/copiar-pegar hacia Excel.

## Qué hace

- Permite subir uno o varios PDFs.
- Extrae los campos principales: nombre, cédula, ingreso, valor crédito, seguro proporcional, extraprima y fecha de nacimiento.
- Agrega cada PDF como una fila hacia abajo en una tabla web editable.
- Permite copiar todas las filas en formato tabulado para pegarlas directamente en Excel.
- Muestra las fechas en formato `MM-DD-YY`, igual al formato encontrado en las otras hojas de la plantilla.
- Mantiene el endpoint anterior para generar un Excel diligenciado directamente.

## Instalar

```bash
cd pdf_excel_extractor_app
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

También debes tener Tesseract OCR instalado en el sistema.

En macOS:

```bash
brew install tesseract
```

En Ubuntu/Debian:

```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr
```

## Ejecutar

```bash
uvicorn app:app --reload
```

Luego abre:

```text
http://127.0.0.1:8000
```

Preview con una fila de ejemplo:

```text
http://127.0.0.1:8000/preview
```

## Uso recomendado

1. Abre la app.
2. Selecciona uno o varios PDFs.
3. Haz clic en **Extraer PDFs**.
4. Revisa o corrige cualquier celda amarilla/editable.
5. Haz clic en **Copiar filas para Excel**.
6. Pega en la hoja correspondiente del Excel.

## Endpoints útiles

- `GET /` interfaz principal.
- `GET /preview` interfaz con una fila de ejemplo.
- `POST /extract-json` recibe un PDF y devuelve JSON.
- `POST /extract-multiple-json` recibe varios PDFs y devuelve registros.
- `POST /extract` recibe PDF + Excel y descarga un Excel diligenciado.

## Notas de negocio implementadas

- El valor de seguro/mes se toma desde `Seguro Cartera proporcional`, no desde la primera cuota de amortización.
- Si se detecta alguna enfermedad marcada en `SI`, la columna EXTRAPRIMA queda como `EXTRAPRIMA INCLUIDA`.
- Las fechas para copia/preview se muestran como `MM-DD-YY`.
- Al generar Excel directamente, las fechas se escriben como fechas reales de Excel con formato `mm-dd-yy`.
