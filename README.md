# CRM VENTA BOOKING - Apps Script + Google Sheets

Aplicacion de automatizacion para el CRM de ventas de **ARTES BUHO**.

## Autor
- **RUBEN COTON**

## Objetivo del proyecto
- Auditar emails de contactos en todas las pestanas del CRM.
- Marcar estado del email como `BIEN`, `MAL` o `CORREGIDO`.
- Colorear toda la fila segun el estado del email.
- Ejecutar autocompletado IA de celdas vacias.
- Mantener `Merge status` siempre como ultima columna y sin tocar su contenido.

## Menu en la hoja (CRM Venta Booking)
El menu superior muestra solo:
1. `Auditar contactos (IA)`
2. `Autocompletado de celdas (IA)`

## Reglas funcionales clave
- `CORREO REVISADO` usa solo: `BIEN`, `MAL`, `CORREGIDO`.
- Auditoria siempre sobrescribe el estado previo (no conserva valor anterior).
- Si email valido -> `BIEN` (fila verde).
- Si email invalido y sin alternativa -> `MAL` (fila roja).
- Si email invalido y se encuentra alternativa valida -> `CORREGIDO` (fila azul).
- `Merge status`:
  - siempre se mueve al final de la hoja,
  - nunca se rellena ni modifica su contenido automaticamente.

## Ventanas IA (modales)
### Auditoria IA
- Modal bloqueante con progreso 0% -> 100%.
- Muestra en vivo:
  - hoja/fila en proceso,
  - contadores BIEN/MAL/CORREGIDO,
  - cambios recientes.

### Autocompletado IA
- Modal bloqueante con progreso 0% -> 100%.
- Rellena celdas vacias.
- Si no encuentra dato web: `IA NO ENCUENTRA`.

## Rendimiento al abrir la hoja
- `onOpen` aplica solo ajustes visuales ligeros en la hoja activa.
- Objetivo: apertura rapida (<= 5 segundos, segun tamano/carga de la hoja y red).
- La carga pesada se ejecuta solo al pulsar botones IA.
- El limite de busquedas web se controla por pestana completa (no se reinicia en cada bloque), para evitar saturacion de API.

## Archivos principales
- `Code.js`: menu y triggers de apertura/edicion.
- `EMAIL_REVIEW_IA.gs`: logica de auditoria, autocompletado, panel y estados.
- `UI_AUDIT_PANEL.html`: interfaz modal futurista de auditoria.
- `UI_AUTOCOMPLETE_PANEL.html`: interfaz modal futurista de autocompletado.

## Sincronizacion
- Apps Script: `clasp push`
- GitHub: `git push origin main`
- Auditoria CLI (`email:review`): si la cuenta tecnica no tiene permiso de escritura, ejecuta fallback automatico a `DryRun` para no bloquear el flujo.

## Nota operativa
- El modo automatico de auditoria queda desactivado por politica de seguridad y control.
- La auditoria se lanza manualmente desde el menu de la hoja.

## CIERRE CLOUD 2026-04-08
- Estado: sincronizado para migracion a nuevo PC/sistema.
- Preparado para retomar desde GitHub.
- Ultima revision: 2026-04-08 15:26:05 +02:00

## CIERRE MIGRACION CLOUD

- Fecha: 2026-04-08
- Estado: preparado para retomar desde nuevo sistema


<!-- MIGRACION_CLOUD_START -->
## ESTADO MIGRACION CLOUD
- Revisado: 2026-04-08
- Repo listo para continuar en otro sistema.
- Estado Git al cerrar: sincronizado en GitHub.
<!-- MIGRACION_CLOUD_END -->
