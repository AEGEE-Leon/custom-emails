function sendPersonalizedEmailsRenovaciones() {

    const fileId = getFileIdFromUrl(
        "https://docs.google.com/spreadsheets/d/14JCRkvHxWSGnY5c30s50h9bog8rJu_AuQESyQDr7G-4/edit?gid=1820408239#gid=1820408239"
    );
    const sheetName = "Importes";
    const draftSubject = "⚠ IMPORTANTE: Periodo de renovaciones 2026";
    const boardName = "David Mediavilla Aller";
    const boardPronouns = "(he/him)"
    const boardSpot = "Secretary";

    const columnaRenovacion = "Renovado";
    const columnaImporte = "Importe renovación";
    const columnaNombre = "Nombre.";
    const columnaApellidos = "Apellidos.";
    // const columnaEmail = "falso e-mail falso";
    const columnaEmail = "E-mail.";


    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        throw new Error(`La hoja '${sheetName}' no existe en el archivo`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    let total = 0;
    let yaRenovado = 0;
    let cuotaPendiente = 0;
    let correosEnviados = 0;

    function logCaso(nombreCompleto, email, renovado, importe, decision) {
        const estadoRenov = renovado ? "✅ Renovado" : "⏳ No renovado";
        const estadoCuota = (importe === "TBC") ? "❓ Cuota por determinar (TBC)" : `💶 Cuota: ${importe}€`;
        console.log(`👤 ${nombreCompleto} | ${estadoRenov} | ${estadoCuota} | 📧 ${email} | ➜ ${decision}`);
    }

    console.log("=== INICIO: Envío de correos de renovaciones 2026 ===");

    const firmaHTML = `<tbody>
    <tr style="height:23.4pt">
        <td width="224" valign="top"
            style="width:168pt;border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black;padding:0in 5.4pt;height:23.4pt">
            <p style="margin-bottom:0.0001pt;line-height:normal">
                <font size="1"></font><img width="200" height="114"
                    src="https://ci3.googleusercontent.com/mail-sig/AIorK4wfgJIwrItIeK8S8-YjI3ycEn62tgAabzOa1nDafyePRknFRzMNczTFcJR7PGQ3sFDyrxG3bLc"
                    class="CToWUd a6T" data-bit="iit" tabindex="0">
            <div class="a6S" dir="ltr" style="opacity: 0.01; left: 167.188px; top: 835px;"><span
                    data-is-tooltip-wrapper="true" class="a5q" jsaction="JIbuQc:.CLIENT"><button
                        class="VYBDae-JX-I VYBDae-JX-I-ql-ay5-ays CgzRE" jscontroller="PIVayb"
                        jsaction="click:h5M12e;clickmod:h5M12e;pointerdown:FEiYhc;pointerup:mF5Elf;pointerenter:EX0mI;pointerleave:vpvbp;pointercancel:xyn4sd;contextmenu:xexox;focus:h06R8; blur:zjh6rb;mlnRJb:fLiPzd;"
                        data-idom-class="CgzRE" data-use-native-focus-logic="true" jsname="hRZeKc"
                        aria-label="Descargar el archivo adjunto " data-tooltip-enabled="true" data-tooltip-id="tt-c20"
                        data-tooltip-classes="AZPksf" id=""
                        jslog="91252; u014N:cOuCgd,Kr2w4b,xr6bB; 4:WyIjbXNnLWY6MTg1MjUxNDI1MzMxNjkzMDUyMSJd; 43:WyJpbWFnZS9qcGVnIl0."><span
                            class="XjoK4b VYBDae-JX-UHGRz"></span><span class="UTNHae" jscontroller="LBaJxb"
                            jsname="m9ZlFb" soy-skip="" ssk="6:RWVI5c"></span><span class="VYBDae-JX-ank-Rtc0Jf"
                            jsname="S5tZuc" aria-hidden="true"><span class="notranslate bzc-ank" aria-hidden="true"><svg
                                    viewBox="0 -960 960 960" height="20" width="20" focusable="false" class=" aoH">
                                    <path
                                        d="M480-336L288-528l51-51L444-474V-816h72v342L621-579l51,51L480-336ZM263.72-192Q234-192 213-213.15T192-264v-72h72v72H696v-72h72v72q0,29.7-21.16,50.85T695.96-192H263.72Z">
                                    </path>
                                </svg></span></span>
                        <div class="VYBDae-JX-ano"></div>
                    </button>
                    <div class="ne2Ple-oshW8e-J9" id="tt-c20" role="tooltip" aria-hidden="true">Descargar</div>
                </span></div>
            <font size="4"><br></font>
            </p>
        </td>
        <td width="537" valign="top" style="width:402.4pt;border:none;padding:0in 5.4pt;height:23.4pt">
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span
                    style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)">
                    <font size="4"><br></font>
                </span></p>
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span
                    style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)">·<span
                        style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">
                        <font size="6">&nbsp;&nbsp; </font>
                    </span></span><span style="line-height:24px;font-family:&quot;Bebas Neue Regular&quot;">
                    <font size="6">${boardName}&nbsp;</font>
                    <font size="4">${boardPronouns}<br></font>
                </span></p>
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                    style="font-family:Symbol;color:rgb(160,197,20)">·<span
                        style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span
                    style="font-family:&quot;Open Sans&quot;,sans-serif">${boardSpot}</span></p>
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                    style="font-family:Symbol;color:rgb(147,25,145)">·<span
                        style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span
                    style="font-family:&quot;Open Sans&quot;,sans-serif">AEGEE-León | European Students’ Forum</span>
            </p>
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                    style="font-family:Symbol;color:rgb(251,180,0)">·<span
                        style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span>
                <font face="Open Sans, sans-serif">Mobile: +34 623 23 35 34&nbsp;</font>
            </p>
            <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                    style="font-family:Symbol;color:rgb(20,104,197)">·<span
                        style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;
                        &nbsp; &nbsp; &nbsp;</span></span><a href="http://www.aegeeleon.org/"
                    style="color:rgb(17,85,204)" target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=http://www.aegeeleon.org/&amp;source=gmail&amp;ust=1766781668359000&amp;usg=AOvVaw0AXZ4KR5SFsEz_6tvkffpZ"><span
                        style="font-family:&quot;Open Sans&quot;,sans-serif">www.aegeeleon.org</span></a></p>
        </td>
    </tr>
</tbody>`;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        let rowData = {};
        for (let j = 0; j < headers.length; j++) {
            rowData[headers[j]] = row[j];
        }

        total++;

        const nombreCompleto = `${rowData[columnaNombre]} ${rowData[columnaApellidos]}`;
        const renovado = (rowData[columnaRenovacion] === true || rowData[columnaRenovacion] === "TRUE");
        const importe = rowData[columnaImporte];
        const email = rowData[columnaEmail];

        // MISMA LÓGICA: si renovado => continue
        if (rowData[columnaRenovacion] === true || rowData[columnaRenovacion] === "TRUE") {
            yaRenovado++;
            logCaso(nombreCompleto, email, renovado, importe, "No se envía correo (ya ha renovado)");
            continue;
        }

        // MISMA LÓGICA: si TBC => continue
        if (rowData[columnaImporte] === "TBC") {
            cuotaPendiente++;
            logCaso(nombreCompleto, email, renovado, importe, "No se envía correo (cuota pendiente de calcular)");
            continue;
        }

        // Aquí llegamos solo si se va a enviar (o preparar) correo
        correosEnviados++;
        logCaso(nombreCompleto, email, renovado, importe, "Se envía correo de renovación");

        let bodyHTML = `
<div style="font-family:verdana,sans-serif">

🚨 <b>¡IMPORTANTE!</b> 📯  
✅ PERIODO DE RENOVACIÓN DE SOCI@S <b>ANTERIORES AL 1 DE SEPTIEMBRE DE 2025</b><br><br>

¡Buenos días socios, y felices fiestas!<br><br>

Os escribimos desde la Junta de AEGEE-León para recordaros que 
<b>ahora es momento de renovar</b>, aunque el año haya sido duro…
¡¡Porque tenemos muchas cosas que os van a encantar para este nuevo año!!<br><br>

Abajo os explicamos cómo se harán las renovaciones, y ya sabéis que para cualquier duda
nos tenéis en redes y en el email a un click.<br><br>

📥 <b>Deadline: 19 de enero de 2026</b><br><br>

Para renovar, solo tienes que <b>pagar la fee correspondiente</b> y rellenar el form para
<b>actualizar tus datos</b>.<br><br>

<b>PARA PAGAR LA FEE:</b><br>
– Te recordamos que, en base a la fecha a la que te uniste,
<b>tu fee es de ${rowData[columnaImporte]} €</b><br><br>

Concepto: <b><i>RENOVACIÓN ${rowData[columnaNombre]} ${rowData[columnaApellidos]}</i></b><br>
Beneficiario: ASOC DE LOS ESTADOS GENERALES DE LOS ESTUDIANTES DE EUROPA LEON<br>
Nº cuenta (Santander):<br>
<b>ES12 0049 1421 6828 1004 7499</b><br><br>

⚠ Una vez pagada, envía el justificante a
<a href="mailto:renovaciones@aegeeleon.org">renovaciones@aegeeleon.org</a><br><br>

🚨 Las renovaciones fuera de plazo pagarán la fee de socio nuevo (25€)<br><br>

<b>PARA ACTUALIZAR TUS DATOS:</b><br>
Por favor, aunque no hayan cambiado tus datos, rellena el formulario para aceptar
la política de privacidad de datos.<br><br>

👇👇👇<br>
<a href="https://forms.gle/FDmyjNExevg4Yjqj8" target="_blank">
https://forms.gle/FDmyjNExevg4Yjqj8
</a>

</div>
${firmaHTML}
`;

        if (rowData[columnaImporte] === 0 || rowData[columnaImporte] === "0") {
            bodyHTML = `
<div style="font-family:verdana,sans-serif">

🚨 <b>¡IMPORTANTE!</b> 📯  
✅ PERIODO DE RENOVACIÓN DE SOCI@S <b>ANTERIORES AL 1 DE SEPTIEMBRE DE 2025</b><br><br>

¡Buenos días socios, y felices fiestas!<br><br>

Os escribimos desde la Junta de AEGEE-León para recordaros que 
<b>ahora es momento de renovar</b>, aunque el año haya sido duro…
¡¡Porque tenemos muchas cosas que os van a encantar para este nuevo año!!<br><br>

Abajo os explicamos cómo se harán las renovaciones, y ya sabéis que para cualquier duda
nos tenéis en redes y en el email a un click.<br><br>

📥 <b>Deadline: 19 de enero de 2026</b><br><br>

Para renovar, solo tienes que y rellenar el form para <b>actualizar tus datos</b>.<br><br>

🚨 Las renovaciones fuera de plazo pagarán la fee de socio nuevo (25€)<br><br>

<b>PARA ACTUALIZAR TUS DATOS:</b><br>
Por favor, aunque no hayan cambiado tus datos, rellena el formulario para aceptar
la política de privacidad de datos.<br><br>

👇👇👇<br>
<a href="https://forms.gle/FDmyjNExevg4Yjqj8" target="_blank">
https://forms.gle/FDmyjNExevg4Yjqj8
</a>

</div>
${firmaHTML}
`;
        }

        MailApp.sendEmail({
            to: email,
            subject: draftSubject,
            htmlBody: bodyHTML
        });


        console.log(`📨 Correo enviado a: ${nombreCompleto} | 📧 ${email}`);
        Utilities.sleep(1000);

    }

    console.log("=== RESUMEN EJECUCIÓN ===");
    console.log(`Filas revisadas: ${total}`);
    console.log(`✅ Ya renovados : ${yaRenovado}`);
    console.log(`⏳ Pendientes renovar : ${total - yaRenovado}`);
    console.log(`❓ Cuota pendiente TBC : ${cuotaPendiente}`);
    console.log(`📨 Correos preparados/enviados: ${correosEnviados}`);
    console.log("=== FIN ===");
}

function getFileIdFromUrl(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
        throw new Error("No se pudo extraer el fileId de la URL.");
    }
    return match[1];
}
