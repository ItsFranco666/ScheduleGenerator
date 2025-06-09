# Generador de Horarios de Laboratorio

Este script automatiza la generación de horarios para los laboratorios de industrial a partir de un reporte de ocupación en Excel.

## Requisitos

- Python 3.x
- pandas
- openpyxl

Para instalar las dependencias necesarias, ejecuta:
```bash
pip install -r requirements.txt
```
O en su defecto usar el archivo requirements.txt
```bash
pip install pandas openpyxl
```

## Clonar repositorio

1. Abre una terminal o línea de comandos
2. Navega al directorio donde deseas clonar el repositorio
3. Ejecuta el siguiente comando:
```bash
git clone https://github.com/tu-usuario/ScheduleGenerator.git
```
4. Navega al directorio del proyecto:
```bash
cd ScheduleGenerator
```

## Estructura del Proyecto

- `generar_horarios.py`: Script principal
- `reporte_ocupacion.xlsx`: Archivo de entrada con los datos de ocupación (se debe proporcionar)
- `HORARIO_LABORATORIOS.xlsx`: Archivo de salida generado

## Configuración de Laboratorios

El script utiliza un mapeo mediante diccionarios de nombres de laboratorios para convertir los nombres del reporte de ocupación a los nombres deseados en el horario final. Para modificar este mapeo:

1. Abre el archivo `generar_horarios.py`
2. Localiza el diccionario `mapeo_laboratorios` en la clase `GeneradorHorariosLaboratorio`
3. Modifica o agrega entradas siguiendo el formato:
```python
'NOMBRE_LABORATORIO_ORIGEN': 'NOMBRE_LABORATORIO_SALIDA'
```

Ejemplo de mapeo actual:
```python
self.mapeo_laboratorios = {
    'LABORATORIO GEIO CAP(25)': 'GEIO (321) TECHNE',
    'SALA DE SOFTWARE DE TECNOLOGIA E INGENIERIA DE PRODUCCION A CAP(17)': 'Sala de Software A - 16 EST - 416- TECHNE',
    'SALA DE SOFTWARE DE TECNOLOGIA E NGENIERIA DE PRODUCCION B CAP(25)': 'Sala de Software B - 24 EST - 417 TECHNE',
    'LABORATORIO HAS CAP(22)': 'HAS-200 (317) TECHNE',
    'LABORATORIO FMS CAP(18)': 'FMS-200 (320) TECHNE',
    'LABORATORIO DE PROCESOS DE TRANSFORMACIÓN MECÁNICA': 'LABORATORIO DE PROCESOS DE TRANSFORMACIÓN BLOQUE 1-102'
}
```

## Formato del Archivo de Entrada

El archivo `reporte_ocupacion.xlsx` debe contener las siguientes columnas:
- Periodo
- Día
- Hora
- Asignatura
- Grupo
- Proyecto
- Salón
- Área
- Edificio
- Sede
- Inscritos
- Docente

## Ejecución del Script

1. Asegúrate de tener Python 3.x instalado en tu sistema
2. Instala las dependencias necesarias:
```bash
pip install -r requirements.txt
```
```bash
pip install pandas openpyxl
```
3. Coloca el archivo `reporte_ocupacion.xlsx` en el directorio raíz del proyecto
4. Abre una terminal o línea de comandos
5. Navega al directorio del proyecto si no estás en él:
```bash
cd ScheduleGenerator
```
6. Ejecuta el script:
```bash
python generar_horarios.py
```
7. El script generará automáticamente el archivo `HORARIO_LABORATORIOS.xlsx` en el mismo directorio

## Características del Horario Generado

- Organización por días de la semana (Lunes a Sábado)
- Franjas horarias de 6AM a 10PM
- Separación visual entre días
- Colores alternos para mejor legibilidad
- Información detallada de asignaturas, grupos e inscritos
- Formato profesional con bordes y alineación

## Personalización

### Franjas Horarias
Las franjas horarias están definidas en la variable `franjas_horarias`. Pueden ser modificados.

### Días de la Semana
Los días de la semana están definidos en la variable `dias`. Pueden ser modificados.

## Solución de Problemas

Si encuentras algún error:
1. Verifica que el archivo de entrada tenga el formato correcto
2. Asegúrate de que los nombres de los laboratorios en el mapeo coincidan exactamente con los del reporte
3. Revisa que el archivo de entrada esté en el directorio correcto
4. Verifica que todas las dependencias estén instaladas

## Notas Importantes

- El script solo procesa laboratorios ubicados en el edificio TECHNE
- Las clases de dos horas se muestran con formato especial
- El número de inscritos se incluye junto al grupo