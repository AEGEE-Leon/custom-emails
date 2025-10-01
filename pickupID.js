function sendPersonalizedEmails() {
    // Si el archivo es https://docs.google.com/spreadsheets/d/13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo/edit?gid=0#gid=0 
    // el id es -> 13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo
    const fileId = "13qOWs22eWI366HDaDCBOJbOo52l50WPaKqnmnxTnboo";
    const sheetName = "Hoja 1";
    const draftSubject = "Recoge tu carnet";
    const boardName = "BOARD OF AEGEE-León";
    const boardSpot = "";


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

        // Si "Pendiente tarjeta" es FALSE, saltar al siguiente registro
        if (rowData["Pendiente tarjeta"] === false || rowData["Pendiente tarjeta"] === "FALSE") {
            continue; // no enviar correo
        }

        const email = rowData["E-mail"];
        let bodyMD = `
  <div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">(<img data-emoji="🇬🇧" class="an1" alt="🇬🇧"
            aria-label="🇬🇧" draggable="false" src="https://fonts.gstatic.com/s/e/notoemoji/16.0/1f1ec_1f1e7/72.png"
            loading="lazy">&nbsp;below)&nbsp;<br></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Hola!</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Hemos visto que nos trajiste tu&nbsp;foto para
        hacerte el carnet de socio Erasmus de AEGEE-León, pero que no has venido aún a por él!</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Ven en cuanto puedas, que <b>a partir de esta
            semana ya no vale con solo la tarjeta de descuentos</b>, tenéis que tener las dos tarjetas a la vez!</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Podéis pasaros por la oficina de AEGEE-León el
        lunes, martes y jueves de 12:30-14:30 :)</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i>No tardéis por favor</i></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i>&lt;3</i></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i><br></i></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i>--</i></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Hi!</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">We've noticed that you gave us a photo so that we
        could make your AEGEE Erasmus ID card, but you never came to the office to pick it up!</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">Please come to collect it as soon as possible,
        because <b>starting this week, only using the discount card won't be enough, you need the discount card and the
            AEGEE card!!</b></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default">You can come collect it at the AEGEE office on
        Monday, Tuesday or Thursday from 12:30-14:30 :)</div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i>Please do not take too long</i></div>
    <div style="font-family:verdana,sans-serif" class="gmail_default"><i>&lt;3</i></div><br clear="all">
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

    return `<tr style="height:23.4pt">
    <td width="224" valign="top"
        style="width:168pt;border-top:none;border-bottom:none;border-left:none;border-right:1pt solid black;padding:0in 5.4pt;height:23.4pt">
        <p style="margin-bottom:0.0001pt;line-height:normal">
            <font size="1"></font><img width="200" height="114"
                src="https://ci3.googleusercontent.com/mail-sig/AIorK4wfgJIwrItIeK8S8-YjI3ycEn62tgAabzOa1nDafyePRknFRzMNczTFcJR7PGQ3sFDyrxG3bLc"
                data-os="https://docs.google.com/uc?export=download&amp;id=1Q0qxC8YE4C_RpdqHtFDrQirLZtO02reE&amp;revid=0B_bUH7YT_GA_aGxnSDAxcjloa2FMUDY2WXU1bEluZ29Hdm5FPQ">
            <font size="4"><br></font>
        </p>
    </td>
    <td width="537" valign="top" sty  le="width:402.4pt;border:none;padding:0in 5.4pt;height:23.4pt">
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span
                style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)">
                <font size="4"><br></font>
            </span></p>
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:11.52px"><span
                style="line-height:14.4px;font-family:Symbol;color:rgb(197,28,19)">·<span
                    style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">
                    <font size="6">&nbsp; &nbsp;</font>
                </span></span><span style="line-height:24px;font-family:&quot;Bebas Neue Regular&quot;">
                <font size="6">${nombre}</font>
            </span></p>
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                style="font-family:Symbol;color:rgb(160,197,20)">·<span
                    style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span
                style="font-family:&quot;Open Sans&quot;,sans-serif">${cargo}</span></p>
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                style="font-family:Symbol;color:rgb(147,25,145)">·<span
                    style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span><span
                style="font-family:&quot;Open Sans&quot;,sans-serif">AEGEE-León | European Students’ Forum</span></p>
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                style="font-family:Symbol;color:rgb(251,180,0)">·<span
                    style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span>
            <font face="Open Sans, sans-serif">Mobile: +34 623 23 35 34&nbsp;</font>
        </p>
        <p style="margin:0in 0in 0.0001pt 0.25in;line-height:normal"><span
                style="font-family:Symbol;color:rgb(20,104,197)">·<span
                    style="font-stretch:normal;line-height:normal;font-family:&quot;Times New Roman&quot;">&nbsp; &nbsp;
                    &nbsp; &nbsp;</span></span><a href="http://www.aegeeleon.org/" style="color:rgb(17,85,204)"
                target="_blank"><span style="font-family:&quot;Open Sans&quot;,sans-serif">www.aegeeleon.org</span></a>
        </p>
    </td>
</tr>`
}