# AcSePrU

Aplicativo creado para automatizar el ingreso de actividades semanales las cuales son extraídas de un archivo Google Sheets y gracias a la librería Selenium son agregadas dentro de la página creada por la universidad Piloto de Colombia para el reporte de actividades realizadas durante la práctica profesional

## Ejecución:

Al ser un aplicativo de único propósito se debe contar con las dependencias de Python y agregar en el archivo **encuesta.json** la siguiente información:
- El nombre y la carrera del estudiante
- Ruta y credenciales de conexión con el archivo Google Sheets (Debe seguir la plantilla establecida)
- Url de página de universidad (por seguridad no es entregada)
- Ubicación del WebDriver de Chrome

## Dependencias:
- Selenium v3.141.0
- ChromeDriver v91.0.4472.101
- google-api-python-client v2.14.0
- google-auth-httplib2 v0.1.0
- google-auth-oauthlib v0.4.4
