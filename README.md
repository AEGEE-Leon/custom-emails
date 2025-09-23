# Send Personalized Emails

Script en Google Apps Script para enviar correos personalizados a partir de una hoja de cálculo de Google Sheets.

## Requisitos
- Una hoja de cálculo en Google Sheets con columnas como:  
  `Nombre`, `Email`, `TipoCarnet`, `FechaRecogida`, `Ciudad`, `País`, `Teléfono`, `TipoMiembro`, `Comentarios`.
- Tener activado Google Apps Script en tu cuenta de Google.
- Autorizar el acceso a Gmail y Google Sheets la primera vez que ejecutes el script.

## Uso

1. Abre tu hoja de cálculo en Google Sheets.  
   Copia el **ID del archivo** desde la URL:
```

[https://docs.google.com/spreadsheets/d/](https://docs.google.com/spreadsheets/d/)\<AQUI\_VA\_EL\_ID>/edit

````

2. Entra en **Extensiones > Apps Script** y pega el contenido de `sendPersonalizedEmails.gs` (ajustando el correo, recomendable hacer un borrador y luego copiar el elemento HTML del correo, sin firma).

3. Edita las variables principales:
```javascript
const fileId = "ID_DE_TU_SHEET";
const sheetName = "Hoja 1";
const draftSubject = "Información de tu pedido";
const boardName = "Tu Nombre";
const boardSpot = "Tu Cargo";
````

4. Guarda el proyecto y pulsa **▶ Ejecutar** en la función `sendPersonalizedEmails`.

Los correos se enviarán automáticamente a cada fila de la hoja, generando un mensaje personalizado con los datos de la persona.

## Notas

* El script incluye una firma automática (`createSignature`).
