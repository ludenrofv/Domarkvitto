
/*******************************************************
 * DOMARKVITTO – KAREBY IS
 *
 * Huvudscript:
 * - Tar emot formulär
 * - Sparar till Google Sheet
 * - Skickar bekräftelse
 * - Skickar admin-notis
 * - Hanterar "Skapa kvitto"
 *
 * PDF-layouten finns i KvittoPDF.gs
 *******************************************************/


/*******************************************************
 * KONFIGURATION
 *******************************************************/

const CONFIG = {

  SHEET_NAME: "Domarkvitton",

  ADMIN_EMAIL: "olle.ostlund@kareby.com",

  // ID till Drive-mappen där PDF-kvitton ska sparas
  DRIVE_FOLDER_ID: "1s5O3SOivo1_eotcnB0YDMuAGVxWva0Nw",

  // ID till Kareby IS-loggan
  LOGO_FILE_ID: "1LLbDYRRqucNzX3jJuaLiZFa5vMEBUqiS",

  CLUB_NAME: "Kareby IS",

  ORG_NUMBER: "853300-1742"

};


/*******************************************************
 * KOLUMNER
 *******************************************************/

const COLUMNS = [

  "Kvittonummer",
  "Inskickat",
  "Matchdatum",
  "Hemmalag",
  "Bortalag",
  "Spelform",
  "Domararvode i kr",
  "Lag i Kareby IS som domaren tillhör",
  "Ansvarig ledare",
  "Mobiltelefonnummer till ledare",
  "Domarens namn",
  "Domarens personnummer",
  "Domarens mailadress",
  "Domarens adress",
  "Domarens postnummer",
  "Domarens postort",
  "Telefonnummer att Swisha till",
  "Är telefonnumret registrerat i domarens namn?",
  "Om svaret är Nej – vem står telefonnumret på?",
  "Övrigt",
  "Status",
  "Skapa kvitto",
  "Kvitto PDF",
  "Kvitto skickat"

];


/*******************************************************
 * KOLUMNINDEX
 *******************************************************/

const COL = {

  KVITTONUMMER: 1,
  INSKICKAT: 2,
  MATCHDATUM: 3,
  HEMMALAG: 4,
  BORTALAG: 5,
  SPELFORM: 6,
  DOMARARVODE: 7,
  LAG: 8,
  ANSVARIG_LEDARE: 9,
  MOBIL: 10,
  DOMARENS_NAMN: 11,
  PERSONNUMMER: 12,
  MAILADRESS: 13,
  ADRESS: 14,
  POSTNUMMER: 15,
  POSTORT: 16,
  SWISH: 17,
  SWISH_REGISTRERAD: 18,
  TELEFON_AGARE: 19,
  OVRIGT: 20,
  STATUS: 21,
  SKAPA_KVITTO: 22,
  KVITTO_PDF: 23,
  KVITTO_SKICKAT: 24

};


/*******************************************************
 * STATUSAR
 *******************************************************/

const STATUS_VALUES = [
  "Ny",
  "Godkänd",
  "Betald",
  "Makulerad"
];


/*******************************************************
 * MENY
 *******************************************************/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("Domarkvitto")
    .addItem(
      "Skapa kvitto för markerad rad",
      "createReceiptForActiveRow"
    )
    .addItem(
      "Skapa kvitton för markerade rader",
      "createReceiptsForSelectedRows"
    )
    .addSeparator()
    .addItem(
      "Kör setup igen",
      "setup"
    )
    .addToUi();

}


/*******************************************************
 * SETUP
 *******************************************************/

function setup() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.SHEET_NAME
      );

  }


  /*****************************************************
   * RUBRIKER
   *****************************************************/

  sheet
    .getRange(
      1,
      1,
      1,
      COLUMNS.length
    )
    .setValues([
      COLUMNS
    ]);


  sheet
    .getRange(
      1,
      1,
      1,
      COLUMNS.length
    )
    .setFontWeight("bold")
    .setBackground("#111111")
    .setFontColor("#ffffff")
    .setVerticalAlignment("middle");


  sheet.setFrozenRows(1);


  /*****************************************************
   * STATUS DROPDOWN
   *****************************************************/

  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        STATUS_VALUES,
        true
      )
      .setAllowInvalid(false)
      .build();


  sheet
    .getRange(
      2,
      COL.STATUS,
      1000,
      1
    )
    .setDataValidation(
      statusRule
    );


  /*****************************************************
   * DATUMFORMAT
   *****************************************************/

  sheet
    .getRange(
      2,
      COL.INSKICKAT,
      1000,
      1
    )
    .setNumberFormat(
      "yyyy-mm-dd hh:mm"
    );


  sheet
    .getRange(
      2,
      COL.KVITTO_SKICKAT,
      1000,
      1
    )
    .setNumberFormat(
      "yyyy-mm-dd hh:mm"
    );


  /*****************************************************
   * KOLUMNBREDDER
   *****************************************************/

  sheet.setColumnWidth(
    COL.KVITTONUMMER,
    110
  );

  sheet.setColumnWidth(
    COL.INSKICKAT,
    150
  );

  sheet.setColumnWidth(
    COL.MATCHDATUM,
    110
  );

  sheet.setColumnWidth(
    COL.HEMMALAG,
    180
  );

  sheet.setColumnWidth(
    COL.BORTALAG,
    180
  );

  sheet.setColumnWidth(
    COL.SPELFORM,
    100
  );

  sheet.setColumnWidth(
    COL.DOMARARVODE,
    130
  );

  sheet.setColumnWidth(
    COL.LAG,
    180
  );

  sheet.setColumnWidth(
    COL.ANSVARIG_LEDARE,
    180
  );

  sheet.setColumnWidth(
    COL.MOBIL,
    150
  );

  sheet.setColumnWidth(
    COL.DOMARENS_NAMN,
    180
  );

  sheet.setColumnWidth(
    COL.PERSONNUMMER,
    150
  );

  sheet.setColumnWidth(
    COL.MAILADRESS,
    220
  );

  sheet.setColumnWidth(
    COL.ADRESS,
    220
  );

  sheet.setColumnWidth(
    COL.POSTNUMMER,
    100
  );

  sheet.setColumnWidth(
    COL.POSTORT,
    150
  );

  sheet.setColumnWidth(
    COL.SWISH,
    150
  );

  sheet.setColumnWidth(
    COL.SWISH_REGISTRERAD,
    180
  );

  sheet.setColumnWidth(
    COL.TELEFON_AGARE,
    200
  );

  sheet.setColumnWidth(
    COL.OVRIGT,
    250
  );

  sheet.setColumnWidth(
    COL.STATUS,
    110
  );

  sheet.setColumnWidth(
    COL.SKAPA_KVITTO,
    110
  );

  sheet.setColumnWidth(
    COL.KVITTO_PDF,
    120
  );

  sheet.setColumnWidth(
    COL.KVITTO_SKICKAT,
    150
  );


  /*****************************************************
   * ONEDIT-TRIGGER
   *****************************************************/

  createOnEditTrigger_();


  Logger.log(
    "Setup klar."
  );

}


/*******************************************************
 * SKAPA ONEDIT-TRIGGER
 *******************************************************/

function createOnEditTrigger_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const triggers =
    ScriptApp.getProjectTriggers();


  triggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        "handleReceiptCheckboxEdit"
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );

      }

    }
  );


  ScriptApp
    .newTrigger(
      "handleReceiptCheckboxEdit"
    )
    .forSpreadsheet(ss)
    .onEdit()
    .create();

}


/*******************************************************
 * DOPOST
 *******************************************************/

function doPost(e) {

  try {

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      throw new Error(
        "Ingen data mottogs från formuläret."
      );

    }


    const data =
      JSON.parse(
        e.postData.contents
      );


    validateFormData_(data);


    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();


    const sheet =
      ss.getSheetByName(
        CONFIG.SHEET_NAME
      );


    if (!sheet) {

      throw new Error(
        'Hittar inte fliken "' +
        CONFIG.SHEET_NAME +
        '". Kör setup().'
      );

    }


    /***************************************************
     * NYTT KVITTONUMMER
     ***************************************************/

    const receiptNumber =
      getNextReceiptNumber_(
        sheet
      );


    /***************************************************
     * FÖRSTA LEDIGA RAD
     ***************************************************/

    const nextRow =
      getNextDataRow_(
        sheet
      );


    /***************************************************
     * DATA
     ***************************************************/

    const row = [

      receiptNumber,

      new Date(),

      data.matchdatum || "",

      data.hemmalag || "",

      data.bortalag || "",

      data.spelform || "",

      data.domararvode || "",

      data.lag || "",

      data.ansvarigLedare || "",

      data.mobil || "",

      data.domarensNamn || "",

      data.personnummer || "",

      data.mailadress || "",

      data.adress || "",

      data.postnummer || "",

      data.postort || "",

      data.swish || "",

      data.swishRegistrerad || "",

      data.telefonAgare || "",

      data.ovrigt || "",

      "Ny",

      false,

      "",

      ""

    ];


    /***************************************************
     * SKRIV TILL SHEET
     ***************************************************/

    sheet
      .getRange(
        nextRow,
        1,
        1,
        row.length
      )
      .setValues([
        row
      ]);


    /***************************************************
     * CHECKBOX ENDAST PÅ DENNA RAD
     ***************************************************/

    sheet
      .getRange(
        nextRow,
        COL.SKAPA_KVITTO
      )
      .insertCheckboxes();


    /***************************************************
     * BEKRÄFTELSEMEJL
     ***************************************************/

    sendConfirmationEmail_(
      data,
      receiptNumber
    );


    /***************************************************
     * ADMIN-MEJL
     ***************************************************/

    sendAdminNotification_(
      data,
      receiptNumber
    );


    /***************************************************
     * SVAR TILL FORMULÄRET
     ***************************************************/

    return jsonResponse_({

      success: true,

      receiptNumber:
        receiptNumber,

      row:
        nextRow,

      message:
        "Domarkvittot har registrerats."

    });


  } catch (error) {

    console.error(
      error
    );


    return jsonResponse_({

      success: false,

      message:
        error.message ||
        "Ett okänt fel uppstod."

    });

  }

}


/*******************************************************
 * VALIDERING
 *******************************************************/

function validateFormData_(data) {

  const requiredFields = [

    [
      "matchdatum",
      "Matchdatum"
    ],

    [
      "hemmalag",
      "Hemmalag"
    ],

    [
      "bortalag",
      "Bortalag"
    ],

    [
      "spelform",
      "Spelform"
    ],

    [
      "domararvode",
      "Domararvode"
    ],

    [
      "ansvarigLedare",
      "Ansvarig ledare"
    ],

    [
      "mobil",
      "Mobiltelefonnummer till ledare"
    ],

    [
      "domarensNamn",
      "Domarens namn"
    ],

    [
      "personnummer",
      "Domarens personnummer"
    ],

    [
      "mailadress",
      "Domarens mailadress"
    ],

    [
      "adress",
      "Domarens adress"
    ],

    [
      "postnummer",
      "Domarens postnummer"
    ],

    [
      "postort",
      "Domarens postort"
    ],

    [
      "swish",
      "Telefonnummer att Swisha till"
    ],

    [
      "swishRegistrerad",
      "Är telefonnumret registrerat i domarens namn?"
    ]

  ];


  requiredFields.forEach(
    function(field) {

      const key =
        field[0];

      const label =
        field[1];


      if (
        data[key] === undefined ||
        data[key] === null ||
        String(data[key]).trim() === ""
      ) {

        throw new Error(
          'Fältet "' +
          label +
          '" saknas.'
        );

      }

    }
  );


  if (
    data.swishRegistrerad === "Nej" &&
    (
      !data.telefonAgare ||
      String(
        data.telefonAgare
      ).trim() === ""
    )
  ) {

    throw new Error(
      "Ange vem telefonnumret står på."
    );

  }

}


/*******************************************************
 * NÄSTA KVITTONUMMER
 *******************************************************/

function getNextReceiptNumber_(sheet) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return 1001;

  }


  const values =
    sheet
      .getRange(
        2,
        COL.KVITTONUMMER,
        lastRow - 1,
        1
      )
      .getValues();


  let maxNumber =
    1000;


  values.forEach(
    function(row) {

      const value =
        Number(
          row[0]
        );


      if (
        !isNaN(value) &&
        value > maxNumber
      ) {

        maxNumber =
          value;

      }

    }
  );


  return maxNumber + 1;

}


/*******************************************************
 * FÖRSTA LEDIGA DATARAD
 *******************************************************/

function getNextDataRow_(sheet) {

  const lastRow =
    Math.max(
      sheet.getLastRow(),
      2
    );


  const values =
    sheet
      .getRange(
        2,
        COL.KVITTONUMMER,
        lastRow - 1,
        1
      )
      .getValues();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      values[i][0] === "" ||
      values[i][0] === null
    ) {

      return i + 2;

    }

  }


  return lastRow + 1;

}


/*******************************************************
 * BEKRÄFTELSEMEJL
 *******************************************************/

function sendConfirmationEmail_(
  data,
  receiptNumber
) {

  const subject =
    "Domarkvitto mottaget – " +
    CONFIG.CLUB_NAME +
    " – #" +
    receiptNumber;


  const htmlBody = `

    <div style="font-family:Arial,sans-serif">

      <h2>Domarkvitto mottaget</h2>

      <p>
        Hej ${escapeHtml_(
          data.domarensNamn
        )}!
      </p>

      <p>
        Ditt domarkvitto har tagits emot av
        ${escapeHtml_(
          CONFIG.CLUB_NAME
        )}.
      </p>

      <p>
        <strong>Kvittonummer:</strong>
        ${receiptNumber}
      </p>

      <p>
        <strong>Match:</strong><br>
        ${escapeHtml_(
          data.hemmalag
        )}
        –
        ${escapeHtml_(
          data.bortalag
        )}
      </p>

      <p>
        <strong>Matchdatum:</strong>
        ${escapeHtml_(
          data.matchdatum
        )}
      </p>

      <p>
        <strong>Domararvode:</strong>
        ${escapeHtml_(
          data.domararvode
        )} kr
      </p>

      <hr>

      <p>
        Utbetalning sker via Swish senast fem
        arbetsdagar efter att domarkvittot har
        skickats in.
      </p>

      <p>
        Med vänlig hälsning<br>
        ${escapeHtml_(
          CONFIG.CLUB_NAME
        )}
      </p>

    </div>

  `;


  MailApp.sendEmail({

    to:
      data.mailadress,

    subject:
      subject,

    htmlBody:
      htmlBody

  });

}


/*******************************************************
 * ADMIN-MEJL
 *******************************************************/

function sendAdminNotification_(
  data,
  receiptNumber
) {

  const subject =
    "Nytt domarkvitto #" +
    receiptNumber;


  const body = `

Nytt domarkvitto har skickats in.

Kvittonummer: ${receiptNumber}

Matchdatum: ${data.matchdatum}

Hemmalag: ${data.hemmalag}

Bortalag: ${data.bortalag}

Spelform: ${data.spelform}

Domararvode: ${data.domararvode} kr

Lag: ${data.lag || "-"}

Domare: ${data.domarensNamn}

Personnummer: ${data.personnummer}

E-post: ${data.mailadress}

Adress: ${data.adress}

Postnummer: ${data.postnummer}

Postort: ${data.postort}

Ansvarig ledare: ${data.ansvarigLedare}

Telefon ledare: ${data.mobil}

Swish: ${data.swish}

Swish registrerad i domarens namn:
${data.swishRegistrerad}

Telefonägare:
${data.telefonAgare || "-"}

Övrigt:
${data.ovrigt || "-"}

  `;


  MailApp.sendEmail({

    to:
      CONFIG.ADMIN_EMAIL,

    subject:
      subject,

    body:
      body

  });

}


/*******************************************************
 * CHECKBOX
 *******************************************************/

function handleReceiptCheckboxEdit(e) {

  try {

    if (
      !e ||
      !e.range
    ) {

      return;

    }


    const range =
      e.range;

    const sheet =
      range.getSheet();


    if (
      sheet.getName() !==
      CONFIG.SHEET_NAME
    ) {

      return;

    }


    if (
      range.getColumn() !==
      COL.SKAPA_KVITTO
    ) {

      return;

    }


    if (
      range.getRow() < 2
    ) {

      return;

    }


    if (
      e.value !== "TRUE"
    ) {

      return;

    }


    createReceiptForRow_(
      sheet,
      range.getRow()
    );


  } catch (error) {

    console.error(
      "Fel när kvitto skapades:",
      error
    );

  }

}


/*******************************************************
 * SKAPA KVITTO FÖR AKTIV RAD
 *******************************************************/

function createReceiptForActiveRow() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  if (
    sheet.getName() !==
    CONFIG.SHEET_NAME
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Öppna fliken " +
        CONFIG.SHEET_NAME +
        " först."
      );

    return;

  }


  const activeRange =
    sheet.getActiveRange();


  if (!activeRange) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Markera en datarad först."
      );

    return;

  }


  const row =
    activeRange.getRow();


  if (
    row < 2
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Markera en datarad först."
      );

    return;

  }


  createReceiptForRow_(
    sheet,
    row
  );


  SpreadsheetApp
    .getUi()
    .alert(
      "Kvitto skapat."
    );

}


/*******************************************************
 * SKAPA FLERA KVITTON
 *******************************************************/

function createReceiptsForSelectedRows() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  if (
    sheet.getName() !==
    CONFIG.SHEET_NAME
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Öppna fliken " +
        CONFIG.SHEET_NAME +
        " först."
      );

    return;

  }


  const range =
    sheet.getActiveRange();


  if (!range) {

    return;

  }


  const startRow =
    Math.max(
      range.getRow(),
      2
    );


  const endRow =
    range.getLastRow();


  let count =
    0;


  for (
    let row = startRow;
    row <= endRow;
    row++
  ) {

    try {

      createReceiptForRow_(
        sheet,
        row
      );

      count++;

    } catch (error) {

      console.error(
        "Fel på rad " +
        row +
        ": " +
        error.message
      );

    }

  }


  SpreadsheetApp
    .getUi()
    .alert(
      count +
      " kvitto(n) skapades."
    );

}


/*******************************************************
 * RAD → OBJEKT
 *******************************************************/

function rowToObject_(
  values
) {

  return {

    kvittonummer:
      values[
        COL.KVITTONUMMER - 1
      ],

    inskickat:
      values[
        COL.INSKICKAT - 1
      ],

    matchdatum:
      values[
        COL.MATCHDATUM - 1
      ],

    hemmalag:
      values[
        COL.HEMMALAG - 1
      ],

    bortalag:
      values[
        COL.BORTALAG - 1
      ],

    spelform:
      values[
        COL.SPELFORM - 1
      ],

    domararvode:
      values[
        COL.DOMARARVODE - 1
      ],

    lag:
      values[
        COL.LAG - 1
      ],

    ansvarigLedare:
      values[
        COL.ANSVARIG_LEDARE - 1
      ],

    mobil:
      values[
        COL.MOBIL - 1
      ],

    domarensNamn:
      values[
        COL.DOMARENS_NAMN - 1
      ],

    personnummer:
      values[
        COL.PERSONNUMMER - 1
      ],

    mailadress:
      values[
        COL.MAILADRESS - 1
      ],

    adress:
      values[
        COL.ADRESS - 1
      ],

    postnummer:
      values[
        COL.POSTNUMMER - 1
      ],

    postort:
      values[
        COL.POSTORT - 1
      ],

    swish:
      values[
        COL.SWISH - 1
      ],

    swishRegistrerad:
      values[
        COL.SWISH_REGISTRERAD - 1
      ],

    telefonAgare:
      values[
        COL.TELEFON_AGARE - 1
      ],

    ovrigt:
      values[
        COL.OVRIGT - 1
      ],

    status:
      values[
        COL.STATUS - 1
      ],

    skapaKvitto:
      values[
        COL.SKAPA_KVITTO - 1
      ],

    kvittoPdf:
      values[
        COL.KVITTO_PDF - 1
      ],

    kvittoSkickat:
      values[
        COL.KVITTO_SKICKAT - 1
      ]

  };

}


/*******************************************************
 * ESCAPE HTML
 *******************************************************/

function escapeHtml_(
  value
) {

  if (
    value === undefined ||
    value === null
  ) {

    return "";

  }


  return String(value)

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}


/*******************************************************
 * JSON-SVAR
 *******************************************************/

function jsonResponse_(
  data
) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}