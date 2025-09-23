function sendPersonalizedEmails() {
    // Si el archivo es https://docs.google.com/spreadsheets/d/13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo/edit?gid=0#gid=0 
    // el id es -> 13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo
    const fileId = "13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo";
    const sheetName = "Hoja 1";
    const draftSubject = "Información de tu pedido";
    const boardName = "Mateo González Alonso";
    const boardSpot = "IT Responsible";


    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        throw new Error(`La hoja '${sheetName}' no existe en el archivo`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // encabezados
    const firmaHTML = createSignature(boardName, boardSpot);

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Crear objeto con nombres de columna para acceso dinámico
        let rowData = {};
        for (let j = 0; j < headers.length; j++) {
            rowData[headers[j]] = formatValue(row[j]);
        }

        // Ejemplo de columnas conocidas
        const email = rowData["Email"];

        // Cuerpo en Markdown
        let bodyMD =
            `
<div>
  <p>Hola <b>${rowData["Nombre"]}</b>,</p>

  <p>Te informamos que tu <b>carnet Erasmus (${rowData["TipoCarnet"]})</b> ya está listo para ser recogido en la oficina.</p>

  <p><b>Detalles de la recogida:</b></p>
  <ul>
    <li><b>Fecha y hora:</b> ${rowData["FechaRecogida"]}</li>
    <li><b>Ciudad:</b> ${rowData["Ciudad"]}</li>
    <li><b>País:</b> ${rowData["País"]}</li>
    <li><b>Teléfono registrado:</b> ${rowData["Teléfono"]}</li>
    <li><b>Tipo de miembro:</b> ${rowData["TipoMiembro"]}</li>
  </ul>

  <p><b>Comentarios importantes:</b><br>
  ${rowData["Comentarios"]}</p>

  <p>Recuerda traer <b>tu documento de identidad y el comprobante de pago</b>.</p>

  <p><span style="color:green;">¡Gracias por tu puntualidad y por ser parte de la comunidad Erasmus!</span></p>
</div>

`;

        let bodyHTML = bodyMD + firmaHTML;

        MailApp.sendEmail({
            to: email,
            subject: draftSubject,
            htmlBody: bodyHTML
        });

        Utilities.sleep(1000);
    }
}

/**
 * Formatea valores automáticamente:
 * - Fechas: dd/mm/yyyy
 * - Fechas con hora: dd/mm/yyyy HH:MM
 * - Horas: HH:MM
 * - Texto: tal cual
 */
function formatValue(value) {
    if (value instanceof Date) {
        const dd = String(value.getDate()).padStart(2, '0');
        const mm = String(value.getMonth() + 1).padStart(2, '0');
        const yyyy = value.getFullYear();
        const hh = String(value.getHours()).padStart(2, '0');
        const min = String(value.getMinutes()).padStart(2, '0');

        if (value.getHours() !== 0 || value.getMinutes() !== 0) {
            return `${dd}/${mm}/${yyyy} ${hh}:${min}`;
        }
        return `${dd}/${mm}/${yyyy}`;
    }
    return value; // texto u otro
}
















function createSignature(nombre, cargo) {
    return `<tr style="height:23.4pt"><td width="224" valign="top" style="width:168pt;border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black;padding:0in 5.4pt;height:23.4pt"><p style="margin-bottom:0.0001pt;line-height:normal"><font size="1"></font><img width="200" height="114" src="https://ci3.googleusercontent.com/mail-sig/AIorK4wfgJIwrItIeK8S8-YjI3ycEn62tgAabzOa1nDafyePRknFRzMNczTFcJR7PGQ3sFDyrxG3bLc" class="CToWUd a6T" data-bit="iit" tabindex="0"><div class="a6S" dir="ltr" style="opacity: 0.01; left: 167.2px; top: 971.5px;"><span data-is-tooltip-wrapper="true" class="a5q" jsaction="JIbuQc:.CLIENT"><button class="VYBDae-JX-I VYBDae-JX-I-ql-ay5-ays CgzRE" jscontroller="PIVayb" jsaction="click:h5M12e;clickmod:h5M12e;pointerdown:FEiYhc;pointerup:mF5Elf;pointerenter:EX0mI;pointerleave:vpvbp;pointercancel:xyn4sd;contextmenu:xexox;focus:h06R8; blur:zjh6rb;mlnRJb:fLiPzd;" data-idom-class="CgzRE" data-use-native-focus-logic="true" jsname="hRZeKc" aria-label="Descargar el archivo adjunto " data-tooltip-enabled="true" data-tooltip-id="tt-c15" data-tooltip-classes="AZPksf" id="" jslog="91252; u014N:cOuCgd,Kr2w4b,xr6bB; 4:WyIjbXNnLWE6cjkyMTA5NjEyNTU5MTExMjU2NjUiXQ..; 43:WyJpbWFnZS9qcGVnIl0."><span class="OiePBf-zPjgPe VYBDae-JX-UHGRz"></span><span class="bHC-Q" jscontroller="LBaJxb" jsname="m9ZlFb" soy-skip="" ssk="6:RWVI5c"></span><span class="VYBDae-JX-ank-Rtc0Jf" jsname="S5tZuc" aria-hidden="true"><span class="notranslate bzc-ank" aria-hidden="true"><svg viewBox="0 -960 960 960" height="20" width="20" focusable="false" class=" aoH"><path d="M480-336L288-528l51-51L444-474V-816h72v342L621-579l51,51L480-336ZM263.72-192Q234-192 213-213.15T192-264v-72h72v72H696v-72h72v72q0,29.7-21.16,50.85T695.96-192H263.72Z"></path></svg></span></span><div class="VYBDae-JX-ano"></div></button><div class="ne2Ple-oshW8e-J9" id="tt-c15" role="tooltip" aria-hidden="true">Descargar</div></span></div><font size="4"><br></font></p></td><td width="537" valign="top" style="width:402.4pt;border:none;padding:0in 5.4pt;height:23.4pt"><p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)"><font size="4"><br></font></span></p><p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)">·<span style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;"><font size="6">&nbsp; </font></span></span><span style="line-height:14.4px;font-family:Symbol;color:rgb(0,0,0)"><span style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;"><b><font size="6"><span style="background-color:rgb(255,255,255)"><span class="il"><span class="il"> ${nombre}</font></b></span></span></p><p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><font face="Symbol" style="color:rgb(160,197,20)">·</font><span style="color:rgb(160,197,20);font-family:&quot;Times New Roman&quot;;font-stretch:normal;line-height:normal">&nbsp; &nbsp; &nbsp;</span><span style="font-stretch:normal;line-height:normal"><font face="Open Sans, sans-serif" color="#000000">${cargo}&nbsp;</font></span></p><p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span style="font-family:Symbol;color:rgb(147,25,145)">·<span style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp; &nbsp; &nbsp;</span></span><span style="font-family:&quot;Open Sans&quot;,sans-serif">AEGEE-León | European Students’ Forum</span></p><p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span style="font-family:Symbol;color:rgb(251,180,0)">·<span style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp; &nbsp; &nbsp;</span></span><font face="Open Sans, sans-serif">Mobile: +34&nbsp;</font>&nbsp;623 23 35 34&nbsp;</p><p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span style="font-family:Symbol;color:rgb(20,104,197)">·<span style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp; &nbsp; &nbsp;</span></span><a href="http://www.aegeeleon.org/" style="color:rgb(17,85,204)" target="_blank" data-saferedirecturl="https://www.google.com/url?q=http://www.aegeeleon.org/&amp;source=gmail&amp;ust=1758710272631000&amp;usg=AOvVaw3H7XL2JYiDKgnKXUoBrk7X"><span style="font-family:&quot;Open Sans&quot;,sans-serif">www.aegeeleon.org</span></a></p></td></tr>`
}