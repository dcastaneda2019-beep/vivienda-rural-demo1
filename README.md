# Demostración mapa de vivienda rural (datos ficticios)

Esta aplicación es una demostración técnica. Todos los registros y cifras del mapa inicial son ficticios.
No cargar datos oficiales, personales, confidenciales o internos.

## Variables de entorno requeridas
- `APP_SECRET_KEY`: valor aleatorio largo, privado.
- `VIEWER_PASSWORD`: contraseña para consulta.
- `ADMIN_PASSWORD`: contraseña distinta para administración.
- `COOKIE_SECURE`: `true` cuando se use HTTPS.

## Inicio local
Instalar dependencias con `pip install -r requirements.txt`, configurar las variables de entorno y ejecutar `python app.py`.

## Despliegue
El `Procfile` usa Gunicorn y la app escucha en el puerto proporcionado por la plataforma.
Los servicios gratuitos pueden suspenderse o tener almacenamiento efímero; esta demo no garantiza persistencia de las cargas.
Antes de usar datos oficiales, se requiere autorización institucional, revisión de seguridad y almacenamiento persistente aprobado.
