# Sabina Calendar

Primera maqueta de la agenda de Sabina Nails Valencia.

## Archivos
- `index.html`: interfaz pública de selección de servicio, fecha, hora y datos de contacto.
- `admin.html`: interfaz de administración para servicios, duración, bloqueos, reservas y clientas.

## Estado de esta entrega
Esta es una **maqueta frontend/prototipo local**. Utiliza `localStorage` para demostrar flujos dentro del mismo navegador. No es todavía un sistema de reservas apto para producción: no sincroniza usuarios/dispositivos, no ofrece autenticación segura y no garantiza exclusión mutua de reservas.

## Siguiente fase necesaria
Conectar una API y base de datos (por ejemplo Supabase o Firebase) con:
- autenticación privada para Sabina;
- tablas de servicios, horarios, excepciones/bloqueos, reservas y clientas;
- comprobación atómica de disponibilidad en servidor;
- política de nuevas clientas;
- consentimiento y política de privacidad RGPD;
- notificaciones por email y/o WhatsApp.

Los servicios iniciales son los seis que se han podido verificar en la ficha pública de Fresha. Completar el catálogo de 21 servicios cuando se disponga del listado completo.
