# Mini README — “Pick Up Your Card” Script(Google Apps Script)

## What it does

This script iterates over a Google Sheets file and ** sends a bilingual(Spanish / English) HTML email ** to each member ** whose card is still pending ** (`Pendiente tarjeta` = TRUE).
It uses the spreadsheet ** URL ** and extracts the ** fileId ** automatically.

---

## Requirements

    * A ** Google Sheet ** you have access to, with a sheet named ** `Lista Erasmus 24/25` ** (the names can be changed, but make sure they match).
* The sheet must have this structure:

  * ** Row 1 **: anything(title / notes / etc.)
    * ** Row 2 **: column headers(this script reads headers from`data[1]`)
        * ** Row 3 +**: data rows(this script starts reading from`i = 2`)
            * The header row(row 2) must contain the following columns(with the exact names used in the script or change them but make sure they match):

  * `Pendiente tarjeta`
    * `E-mail`

![](img / sheet_example.png)

---

## Quick setup(things you can change)

Inside `sendPersonalizedEmails()` you must adjust:

* ** Spreadsheet **: the Sheet URL passed to`getFileIdFromUrl("...")`
    * ** sheetName **: `"Lista Erasmus 24/25"`
        * ** Email subject **: `draftSubject`
            * ** Signature **: `boardName`, `boardSpot`
                * ** Column names **: update`"Pendiente tarjeta"` / `"E-mail"` in the code if your headers differ(keeping them as constants is recommended to reduce errors)

---

## Sending logic(important)

For each row:

1. If `Pendiente tarjeta` is `FALSE` → ** no email is sent ** (card already collected / not pending).
2. Otherwise(`Pendiente tarjeta` is`TRUE`) → sends an email to the address in `E-mail`.
3. The email includes:

   * A Spanish message + an English message(bilingual)
    * Office opening times(Mon / Tue / Thu 12: 30–14: 30)
        * The HTML signature appended at the end(`createSignature()`)

---

## How to run it

1. Open ** Google Apps Script ** (Extensions → Apps Script) in your project.
2. Paste the code and save it.
3. Run the function:

    * `sendPersonalizedEmails`
4. The first time, Google will ask for permissions:

   * access to Google Sheets
    * permission to send emails(MailApp)

---

## Viewing results / debugging

    * Go to ** Executions ** in the left panel to see:

  * execution errors(missing sheet, missing headers, invalid emails, etc.)
    * runtime details(duration, quota issues, failures)

        > Note: this script doesn’t currently log per - recipient actions.If you want, you can add `console.log()` lines to record who was emailed or skipped.

---

## Recommendations before “production”

* Do a ** test run ** with 2–3 rows and your own email address.
* Make sure:

  * `E-mail` has no empty cells(MailApp will fail otherwise).
  * `Pendiente tarjeta` contains real boolean values(`TRUE` / `FALSE`) and not free text.
* Keep Gmail / Apps Script daily sending limits in mind if the list is large. (Sending limit is 100 / day for personal`@gmail.com` accounts)

    ---

## Helper functions

    * `getFileIdFromUrl(url)`: extracts the ** fileId ** from a Google Sheets URL.
* `formatValue(value)`: formats Date cells into `dd/mm/yyyy` or`dd/mm/yyyy HH:MM`(other values unchanged).
* `createSignature(nombre, cargo)`: generates the HTML signature appended to each email.
