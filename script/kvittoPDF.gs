/*******************************************************
 * KVITTОPDF.GS
 *
 * Enkel PDF-layout för Domarkvitto
 *
 * Används av:
 *   createReceiptForRow_(sheet, row)
 *
 * Kräver från Code.gs:
 *   CONFIG
 *   COL
 *   rowToObject_()
 *******************************************************/


/*******************************************************
 * SKAPA KVITTO FÖR EN RAD
 *******************************************************/

function createReceiptForRow_(sheet, row) {

  if (!sheet) {
    throw new Error("Ingen tabell angavs.");
  }

  if (row < 2) {
    throw new Error("Ogiltig datarad.");
  }


  // Hämta hela raden
  const values = sheet
    .getRange(
      row,
      1,
      1,
      COLUMNS.length
    )
    .getValues()[0];


  // Gör om raden till ett objekt
  const data = rowToObject_(values);


  if (!data.kvittonummer) {
    throw new Error(
      "Raden saknar kvittonummer."
    );
  }
  const existingPdfUrl = sheet
    .getRange(row, COL.KVITTO_PDF)
    .getValue();

  if (existingPdfUrl) {
    try {
      const match = String(existingPdfUrl).match(/[-\w]{25,}/);

      if (match) {
        const oldFile = DriveApp.getFileById(match[0]);
        oldFile.setTrashed(true);

        Logger.log(
          "Gammalt kvitto borttaget: " + oldFile.getName()
        );
      }
    } catch (err) {
      Logger.log(
        "Kunde inte ta bort gammalt kvitto: " +
        err.message
      );
    }
  }

  /*****************************************************
   * SKAPA TILLFÄLLIGT GOOGLE-DOKUMENT
   *****************************************************/

  const docName =
    "TEMP Domarkvitto #" +
    data.kvittonummer;

  const doc =
    DocumentApp.create(docName);


  const body =
    doc.getBody();


  /*****************************************************
   * SIDINSTÄLLNINGAR
   *****************************************************/

  body
    .setMarginTop(28)
    .setMarginBottom(28)
    .setMarginLeft(42)
    .setMarginRight(42);


  /*****************************************************
   * GRUNDINSTÄLLNINGAR
   *****************************************************/

  const normalStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: "Arial",
    [DocumentApp.Attribute.FONT_SIZE]: 9
  };


  /*****************************************************
   * LOGO + RUBRIK
   *****************************************************/

  /*const header =
    body.appendParagraph("");

  header
    .setAlignment(
      DocumentApp.HorizontalAlignment.LEFT
    )
    .setSpacingAfter(2);


  try {

    const logoFile =
      DriveApp.getFileById(
        CONFIG.LOGO_FILE_ID
      );

    const logo =
      header
        .appendInlineImage(
          logoFile.getBlob()
        );

    // Liten logga
    logo.setWidth(55);
    logo.setHeight(55);

  } catch (error) {

    console.error(
      "Kunde inte lägga in loggan: " +
      error.message
    );

  }*/


  /*****************************************************
   * TITEL
   *****************************************************/

  const title =
    body.appendParagraph(
      "DOMARKVITTO"
    );

  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  title
    .setFontFamily("Arial")
    .setFontSize(22)
    .setBold(true)
    .setSpacingAfter(0);

  try {
  const logoFile = DriveApp.getFileById(
    CONFIG.LOGO_FILE_ID
  );

  console.log("Loggans namn: " + logoFile.getName());
  console.log("Filtyp: " + logoFile.getMimeType());
  console.log("Filstorlek: " + logoFile.getSize());

  const logoParagraph = body.appendParagraph("");

  logoParagraph.setAlignment(
    DocumentApp.HorizontalAlignment.LEFT
  );

  const logo = logoParagraph.appendInlineImage(
    logoFile.getBlob()
  );

  logo.setWidth(55);
  logo.setHeight(55);

  logoParagraph.setSpacingAfter(2);

} catch (error) {
  console.error(
    "Kunde inte lägga in loggan: " + error.message
  );
}




  const club =
    body.appendParagraph(
      CONFIG.CLUB_NAME
    );

  club
    .setFontFamily("Arial")
    .setFontSize(11)
    .setBold(true)
    .setSpacingAfter(0);


  const receipt =
    body.appendParagraph(
      "Kvitto #" +
      data.kvittonummer
    );

  receipt
    .setFontFamily("Arial")
    .setFontSize(9)
    .setBold(true)
    .setSpacingAfter(0);


  const submitted =
    body.appendParagraph(
      "Inskickat: " +
      formatDate_(data.inskickat)
    );

  submitted
    .setFontFamily("Arial")
    .setFontSize(8)
    .setForegroundColor("#555555")
    .setSpacingAfter(8);


  /*****************************************************
   * LINJE
   *****************************************************/

  addLine_(body);


  /*****************************************************
   * MATCH
   *****************************************************/

  addHeading_(body, "MATCH");

const match = body.appendParagraph(
  safe_(data.hemmalag) +
  " – " +
  safe_(data.bortalag)
);

match
  .setFontFamily("Arial")
  .setFontSize(13)
  .setForegroundColor("#555555")
  .setBold(true)
  .setSpacingAfter(2);
  

// Hämta exakt visad text från kalkylarket
const headers = sheet
  .getRange(1, 1, 1, sheet.getLastColumn())
  .getDisplayValues()[0];

const spelformCol = headers.findIndex(
  h => h.trim().toLowerCase() === "spelform"
) + 1;

const matchdatumCol = headers.findIndex(
  h => h.trim().toLowerCase() === "matchdatum"
) + 1;

const spelformText = spelformCol > 0
  ? sheet.getRange(row, spelformCol).getDisplayValue()
  : safe_(data.spelform);

const matchdatumText = matchdatumCol > 0
  ? sheet.getRange(row, matchdatumCol).getDisplayValue()
  : formatMatchDate_(data.matchdatum);

const spelform = body.appendParagraph(
  "Spelform: " + spelformText
);

spelform
  .setFontFamily("Arial")
  .setFontSize(9)
  .setForegroundColor("#555555")
  .setSpacingAfter(2);

const matchdatum = body.appendParagraph(
  "Matchdatum: " + matchdatumText
);

matchdatum
  .setFontFamily("Arial")
  .setFontSize(9)
  .setForegroundColor("#555555")
  .setSpacingAfter(7);


  /*****************************************************
   * DOMARE
   *****************************************************/

  addHeading_(
    body,
    "DOMARE"
  );


  addField_(
    body,
    "Namn",
    data.domarensNamn
  );

  addField_(
    body,
    "Personnummer",
    data.personnummer
  );

  addField_(
    body,
    "E-post",
    data.mailadress
  );

  addField_(
    body,
    "Adress",
    safe_(data.adress) +
    ", " +
    safe_(data.postnummer) +
    " " +
    safe_(data.postort)
  );


  /*****************************************************
   * ANSVARIG LEDARE
   *****************************************************/

  addHeading_(
    body,
    "ANSVARIG LEDARE"
  );

addField_(
    body,
    "Namn",
    data.ansvarigLedare
  );


  addField_(
    body,
    "Lag",
    data.lag || "-"
  );

  

  addField_(
    body,
    "Telefon",
    data.mobil
  );


  /*****************************************************
   * SWISH
   *****************************************************/

  addHeading_(
    body,
    "SWISH"
  );


  addField_(
    body,
    "Telefon",
    data.swish
  );

  addField_(
    body,
    "Registrerad i domarens namn",
    data.swishRegistrerad
  );


  if (
    data.swishRegistrerad === "Nej" &&
    data.telefonAgare
  ) {

    addField_(
      body,
      "Swish registrerad på",
      data.telefonAgare
    );

  }


  /*****************************************************
   * ERSÄTTNING
   *****************************************************/

  addHeading_(
    body,
    "ERSÄTTNING"
  );


  const amount =
    body.appendParagraph(
      formatAmount_(data.domararvode)
    );

  amount
    .setFontFamily("Arial")
    .setFontSize(20)
    .setBold(true)
    .setAlignment(
      DocumentApp.HorizontalAlignment.LEFT
    )
    .setSpacingAfter(8);




  /*****************************************************
   * INFORMATION
   *****************************************************/

  addLine_(body);

  addHeading_(
    body,
    "INFORMATION"
  );


  addInfo_(
    body,
    "Ansvarig ledare ansvarar för att informationen är korrekt ifylld."
  );

  addInfo_(
    body,
    "Ersättningsnivåerna ovan gäller såväl tränings- som tävlingsmatcher."
  );

  addInfo_(
    body,
    "Uppgifterna i sammanställningen ovan lagras och behandlas vidare av Kareby IS, org.nr: 853300-1742 för att uppfylla lagstadgade krav."
  );

  addInfo_(
    body,
    "Utbetalning har skett via Swish"
  );


  /*****************************************************
   * FOOTER
   *****************************************************/

  const footer =
    body.appendParagraph(
      CONFIG.CLUB_NAME +
      " • Org.nr " +
      CONFIG.ORG_NUMBER
    );

  footer
    .setFontFamily("Arial")
    .setFontSize(7)
    .setForegroundColor("#666666")
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    )
    .setSpacingBefore(8);


  /*****************************************************
   * SPARA DOKUMENTET
   *****************************************************/

  doc.saveAndClose();


  /*****************************************************
   * SKAPA PDF
   *****************************************************/

  const tempFile =
    DriveApp.getFileById(
      doc.getId()
    );


  const pdfBlob =
    tempFile
      .getBlob()
      .getAs(MimeType.PDF);


  pdfBlob.setName(
    "Domarkvitto-" +
    data.kvittonummer +
    ".pdf"
  );


  /*****************************************************
   * SPARA PDF I DRIVE-MAPPEN
   *****************************************************/

  const folder =
    DriveApp.getFolderById(
      CONFIG.DRIVE_FOLDER_ID
    );


  const pdfFile =
    folder.createFile(
      pdfBlob
    );


  /*****************************************************
   * TA BORT TEMP-DOKUMENTET
   *****************************************************/

  tempFile.setTrashed(true);


  /*****************************************************
   * LÄNK TILL PDF I GOOGLE SHEETS
   *****************************************************/

  sheet
    .getRange(
      row,
      COL.KVITTO_PDF
    )
    .setFormula(
      '=HYPERLINK("' +
      pdfFile.getUrl() +
      '";"Öppna kvitto")'
    );


  /*****************************************************
   * SKICKA PDF VIA E-POST
   *****************************************************/

  sendReceiptEmail_(
    data,
    pdfBlob
  );


  /*****************************************************
   * UPPDATERA STATUS
   *****************************************************/

  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue("Godkänd");


  sheet
    .getRange(
      row,
      COL.KVITTO_SKICKAT
    )
    .setValue(new Date());


  /*****************************************************
   * AVMARKERA "SKAPA KVITTO"
   *****************************************************/

  sheet
    .getRange(
      row,
      COL.SKAPA_KVITTO
    )
    .setValue(false);


  return pdfFile;

}


/*******************************************************
 * RUBRIK
 *******************************************************/

function addHeading_(body, text) {

  const p =
    body.appendParagraph(
      text
    );

  p
    .setFontFamily("Arial")
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor("#555555")
    .setSpacingBefore(5)
    .setSpacingAfter(2);

}


/*******************************************************
 * FÄLT
 *******************************************************/

function addField_(body, label, value) {

  const p =
    body.appendParagraph("");


  p
    .setFontFamily("Arial")
    .setFontSize(8.5)
    .setSpacingAfter(1);


  const labelText =
    p.appendText(
      label + ": "
    );

  labelText
    .setBold(true);


  p.appendText(
    safe_(value)
  );

}


/*******************************************************
 * INFORMATIONSTEXT
 *******************************************************/

function addInfo_(body, text) {

  const p =
    body.appendParagraph(
      text
    );

  p
    .setFontFamily("Arial")
    .setFontSize(7.5)
    .setForegroundColor("#555555")
    .setSpacingAfter(2);

}


/*******************************************************
 * HORISONTELL LINJE
 *******************************************************/

function addLine_(body) {

  const p =
    body.appendParagraph(
      "────────────────────────────────────────"
    );

  p
    .setFontFamily("Arial")
    .setFontSize(7)
    .setForegroundColor("#999999")
    .setSpacingBefore(4)
    .setSpacingAfter(4);

}


/*******************************************************
 * SÄKERT VÄRDE
 *******************************************************/

function safe_(value) {

  if (
    value === undefined ||
    value === null
  ) {
    return "";
  }

  return String(value);

}


/*******************************************************
 * DATUM
 *******************************************************/


function formatMatchDate_(value) {
  if (!value) {
    return "";
  }

  // Om värdet är ett datumobjekt från Google Sheets
  if (
    Object.prototype.toString.call(value) ===
      "[object Date]" &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  // Om datumet är text, exempelvis:
  // 2026-10-10
  // 2026-10-10T00:00:00.000Z
  const text = String(value).trim();

  const match = text.match(
    /^(\d{4})-(\d{2})-(\d{2})/
  );

  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}`;
  }

  return text;
}



/*******************************************************
 * DATUM + TID
 *******************************************************/

function formatDate_(value) {

  if (!value) {
    return "";
  }


  if (
    Object.prototype.toString.call(value) ===
    "[object Date]" &&
    !isNaN(value)
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm"
    );

  }


  return String(value);

}


/*******************************************************
 * TELEFONNUMMER
 *
 * Skyddar mot exempelvis:
 * 7.01234567E8
 *******************************************************/

function formatPhone_(value) {

  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return "";
  }


  let phone =
    String(value)
      .trim();


  // Om Google redan gjort om numret till E-notation
  if (
    phone.toUpperCase().includes("E+")
  ) {

    const number =
      Number(phone);

    if (!isNaN(number)) {

      phone =
        String(
          Math.round(number)
        );

    }

  }


  // Ta bort mellanslag, bindestreck osv.
  let digits =
    phone.replace(
      /\D/g,
      ""
    );


  // Svenskt nummer med 46 framför
  if (
    digits.startsWith("46") &&
    digits.length === 11
  ) {

    digits =
      "0" +
      digits.substring(2);

  }


  // Vanligt svenskt mobilnummer
  if (
    digits.length === 10 &&
    digits.startsWith("07")
  ) {

    return (
      digits.substring(0, 3) +
      "-" +
      digits.substring(3, 6) +
      " " +
      digits.substring(6)
    );

  }


  // Annat nummer – visa det som det är
  return phone;

}


/*******************************************************
 * BELOPP
 *******************************************************/

function formatAmount_(value) {

  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return "0 kr";
  }


  const number =
    Number(
      String(value)
        .replace(",", ".")
        .replace(/[^\d.-]/g, "")
    );


  if (isNaN(number)) {
    return String(value) + " kr";
  }


  return (
    new Intl.NumberFormat(
      "sv-SE",
      {
        maximumFractionDigits: 0
      }
    ).format(number) +
    " kr"
  );

}


/*******************************************************
 * SKICKA PDF VIA E-POST
 *******************************************************/

function sendReceiptEmail_(
  data,
  pdfBlob
) {

  if (
    !data.mailadress
  ) {

    throw new Error(
      "Domarens mailadress saknas."
    );

  }


  const subject =
    "Domarkvitto #" +
    data.kvittonummer +
    " – " +
    CONFIG.CLUB_NAME;


  const body =

    "Hej " +
    safe_(data.domarensNamn) +
    "!\n\n" +

    "Här kommer ditt domarkvitto.\n\n" +

    "Kvittonummer: #" +
    data.kvittonummer +
    "\n" +

    "Match: " +
    safe_(data.hemmalag) +
    " – " +
    safe_(data.bortalag) +
    "\n" +

    "Matchdatum: " +
    formatMatchDate_(data.matchdatum) +
    "\n" +

    "Domararvode: " +
    formatAmount_(data.domararvode) +
    "\n\n" +

    "Utbetalning har skett via Swish.\n\n" +

    "Med vänlig hälsning\n" +
    CONFIG.CLUB_NAME;


  MailApp.sendEmail({

    to:
      data.mailadress,

    subject:
      subject,

    body:
      body,

    attachments: [
      pdfBlob
    ]

  });

}