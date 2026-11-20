# Sabina Calendar

Agenda de reservas de Sabina Nails Valencia.

## Catálogo y horario aprobados
- 21 servicios del catálogo PDF confirmado.
- Manicura con color: 90 min; Semi Nude: 75 min; pedicura completa + semipermanente: 90 min.
- Diseños complejos: reservables independientemente, 15 min, desde 5 €.
- Lunes a viernes: 09:30–21:30. Sábados y domingos: cerrado.
- Zona horaria: Europe/Madrid.

## Archivos
- `index.html`: página pública de prototipo, con catálogo completo y horario actualizado.
- `admin.html`: dashboard visual de prototipo, con catálogo completo.
- `Code.gs`: backend inicial de Google Apps Script. Crea las seis pestañas y carga servicios, horarios y configuración; ofrece consulta de catálogo, disponibilidad y creación de reservas con LockService.

## Estado real
El backend aún NO está desplegado ni conectado a la hoja de Google de sabinaleonp@gmail.com. Las dos páginas siguen usando localStorage y no están conectadas al backend. No aceptar reservas reales ni compartir la web como agenda operativa hasta completar autenticación administrativa, operaciones privadas del dashboard, integración frontend/backend, privacidad RGPD y pruebas de concurrencia.

## Preparación de Google Sheets
1. En la cuenta propietaria, crear una hoja de cálculo llamada `Sabina Calendar`.
2. Abrir Extensiones → Apps Script y copiar `Code.gs` al editor.
3. Ejecutar `setupSabinaCalendar()` y autorizar los permisos.
4. Verificar las pestañas CLIENTAS, SERVICIOS, RESERVAS, HORARIOS, BLOQUEOS y CONFIGURACION.
5. Desplegar como aplicación web únicamente después de completar la seguridad y las pruebas. Mantener la hoja privada.

El repositorio GitHub es privado; no hay sitio público activo ni dominio conectado todavía.
