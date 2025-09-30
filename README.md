# Gestor de Exámenes Ocupacionales

Herramienta Python para procesar archivos Excel de exámenes médicos ocupacionales. Consolida registros de pacientes por CUIL, ordena alfabéticamente y genera una matriz formateada con pacientes y sus estudios asignados.

## Características

- ✓ Consolida registros duplicados de pacientes usando CUIL como identificador único
- ✓ Ordena pacientes alfabéticamente por nombre (A-Z)
- ✓ Genera matriz profesional con pacientes y exámenes asignados
- ✓ Extracción automática de datos de empresa (CUIT, domicilio, contacto)
- ✓ Formato Excel profesional con bordes, alineación y rotación de texto
- ✓ Ajuste automático de ancho de columnas
- ✓ Procesamiento por lotes de múltiples archivos




## Instalación
```bash
git clone https://github.com/plazagustavo/consolidador-examenes-medicos.git
cd consolidador-examenes-medicos
```
## Requisitos
```bash
pip install pandas openpyxl xlwings
```
Si no están instaladas, el script funcionará con formato básico usando solo pandas.
Instalación

## Ejecución
```bash
python main.py
```

---

Creado por [Gustavo Plaza](https://github.com/plazagustavo)

