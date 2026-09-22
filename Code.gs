const SPREADSHEET_ID = "1DmlkRPvPPF9V7x_BqiSbf3y7Hs1ncPo8";
const SHEET_NAME = "SEP";

const MONTH_ROW = 4;
const DATE_ROW = 5;
const LEC_ROW = 6;
const STUDENT_START_ROW = 7;

const FIRST_ATTENDANCE_COLUMN = 5; // E
const YEAR = 2026;


/* ================================
   WEB APP
================================ */

function doGet(e) {

  // API mode used by the Vercel frontend.
  if (e && e.parameter && e.parameter.action) {
    return handleApiGet(e);
  }

  // Normal Apps Script web-app mode.
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Attendance Entry")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/* ================================
   SHEET
================================ */

function getSheet() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      SHEET_NAME +
      '" not found.'
    );
  }

  return sheet;
}


/* ================================
   STUDENTS
================================ */

function getStudents() {

  const sheet = getSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < STUDENT_START_ROW) {
    return [];
  }

  const count =
    lastRow -
    STUDENT_START_ROW +
    1;

  const data =
    sheet
      .getRange(
        STUDENT_START_ROW,
        1,
        count,
        3
      )
      .getDisplayValues();

  const students = [];

  for (let i = 0; i < data.length; i++) {

    if (
      data[i][0] === "" &&
      data[i][1] === "" &&
      data[i][2] === ""
    ) {
      continue;
    }

    students.push({

      row:
        STUDENT_START_ROW + i,

      sno:
        data[i][0],

      registration:
        data[i][1],

      name:
        data[i][2]

    });
  }

  return students;
}


/* ================================
   FIND DATE COLUMN
================================ */

function findDateColumn(
  sheet,
  month,
  day
) {

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastColumn <
    FIRST_ATTENDANCE_COLUMN
  ) {
    return -1;
  }

  const count =
    lastColumn -
    FIRST_ATTENDANCE_COLUMN +
    1;

  const months =
    sheet
      .getRange(
        MONTH_ROW,
        FIRST_ATTENDANCE_COLUMN,
        1,
        count
      )
      .getDisplayValues()[0];

  const dates =
    sheet
      .getRange(
        DATE_ROW,
        FIRST_ATTENDANCE_COLUMN,
        1,
        count
      )
      .getDisplayValues()[0];

  for (let i = 0; i < count; i++) {

    if (
      Number(months[i]) === month &&
      Number(dates[i]) === day
    ) {

      return (
        FIRST_ATTENDANCE_COLUMN + i
      );
    }
  }

  return -1;
}


/* ================================
   FIND ATT COLUMN
================================ */

function findAttColumn(sheet) {

  const lastColumn =
    sheet.getLastColumn();

  for (
    let col = FIRST_ATTENDANCE_COLUMN;
    col <= lastColumn;
    col++
  ) {

    for (
      let row = 1;
      row <= 6;
      row++
    ) {

      const value =
        String(
          sheet
            .getRange(row, col)
            .getDisplayValue()
        )
        .trim()
        .toLowerCase();

      if (value === "att") {
        return col;
      }
    }
  }

  return lastColumn + 1;
}


/* ================================
   GET BLANK ATTENDANCE COLUMN
================================ */

function getBlankAttendanceColumn(sheet) {

  const attColumn =
    findAttColumn(sheet);

  /*
   * First use an already blank column
   * before Att / Att %.
   */

  for (
    let col = FIRST_ATTENDANCE_COLUMN;
    col < attColumn;
    col++
  ) {

    const month =
      String(
        sheet
          .getRange(MONTH_ROW, col)
          .getDisplayValue()
      ).trim();

    const date =
      String(
        sheet
          .getRange(DATE_ROW, col)
          .getDisplayValue()
      ).trim();

    if (
      month === "" &&
      date === ""
    ) {
      return col;
    }
  }

  /*
   * If there is no blank column,
   * insert a new one before Att.
   */

  sheet.insertColumnBefore(
    attColumn
  );

  return attColumn;
}


/* ================================
   CREATE DATE COLUMN
================================ */

function createDateColumn(
  sheet,
  month,
  day
) {

  const existing =
    findDateColumn(
      sheet,
      month,
      day
    );

  if (existing !== -1) {
    return existing;
  }

  const column =
    getBlankAttendanceColumn(
      sheet
    );

  /*
   * IMPORTANT:
   * We only write values.
   * No copyTo() because your sheet
   * contains merged cells.
   */

  sheet
    .getRange(
      MONTH_ROW,
      column
    )
    .setValue(month);

  sheet
    .getRange(
      DATE_ROW,
      column
    )
    .setValue(day);

  /*
   * Determine lecture number.
   */

  let lecture = 1;

  for (
    let col = FIRST_ATTENDANCE_COLUMN;
    col < column;
    col++
  ) {

    const m =
      String(
        sheet
          .getRange(MONTH_ROW, col)
          .getDisplayValue()
      ).trim();

    const d =
      String(
        sheet
          .getRange(DATE_ROW, col)
          .getDisplayValue()
      ).trim();

    if (
      m !== "" &&
      d !== ""
    ) {
      lecture++;
    }
  }

  sheet
    .getRange(
      LEC_ROW,
      column
    )
    .setValue(lecture);

  return column;
}


/* ================================
   LOAD DATE
================================ */

function getAttendance(dateString) {

  const sheet =
    getSheet();

  const students =
    getStudents();

  if (students.length === 0) {
    throw new Error(
      "No students found."
    );
  }

  const selectedDate =
    new Date(
      dateString +
      "T00:00:00"
    );

  const month =
    selectedDate.getMonth() + 1;

  const day =
    selectedDate.getDate();

  /*
   * DO NOT CREATE COLUMN WHILE
   * JUST VIEWING A DATE.
   */

  const column =
    findDateColumn(
      sheet,
      month,
      day
    );

  const result = [];

  if (column === -1) {

    /*
     * New/today/future date.
     * Show students with blank status.
     */

    students.forEach(
      function(student) {

        result.push({

          row:
            student.row,

          sno:
            student.sno,

          registration:
            student.registration,

          name:
            student.name,

          status:
            ""

        });

      }
    );

  } else {

    /*
     * Existing date.
     */

    const attendance =
      sheet
        .getRange(
          STUDENT_START_ROW,
          column,
          students.length,
          1
        )
        .getDisplayValues();

    students.forEach(
      function(student, index) {

        let status =
          String(
            attendance[index][0] || ""
          )
          .trim()
          .toUpperCase();

        if (
          status !== "P" &&
          status !== "A"
        ) {
          status = "";
        }

        result.push({

          row:
            student.row,

          sno:
            student.sno,

          registration:
            student.registration,

          name:
            student.name,

          status:
            status

        });

      }
    );
  }

  return {

    total:
      result.length,

    students:
      result

  };
}


/* ================================
   SAVE ONE STUDENT
================================ */

function saveOneAttendance(
  dateString,
  row,
  status
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const sheet =
      getSheet();

    status =
      String(status)
        .trim()
        .toUpperCase();

    if (
      status !== "P" &&
      status !== "A"
    ) {

      throw new Error(
        "Invalid attendance status."
      );
    }

    const selectedDate =
      new Date(
        dateString +
        "T00:00:00"
      );

    const month =
      selectedDate.getMonth() + 1;

    const day =
      selectedDate.getDate();

    /*
     * Find or create date column.
     */

    let column =
      findDateColumn(
        sheet,
        month,
        day
      );

    if (column === -1) {

      column =
        createDateColumn(
          sheet,
          month,
          day
        );
    }

    /*
     * Save immediately.
     */

    sheet
      .getRange(
        Number(row),
        column
      )
      .setValue(status);

    SpreadsheetApp.flush();

    return {
      success: true,
      status: status
    };

  } finally {

    lock.releaseLock();

  }
}


/* ================================
   SAVE ALL
   Used by Mark All buttons
================================ */

function saveAllAttendance(
  dateString,
  attendanceData
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const sheet =
      getSheet();

    const selectedDate =
      new Date(
        dateString +
        "T00:00:00"
      );

    const month =
      selectedDate.getMonth() + 1;

    const day =
      selectedDate.getDate();

    let column =
      findDateColumn(
        sheet,
        month,
        day
      );

    if (column === -1) {

      column =
        createDateColumn(
          sheet,
          month,
          day
        );
    }

    let present = 0;
    let absent = 0;

    attendanceData.forEach(
      function(item) {

        const status =
          String(item.status)
            .trim()
            .toUpperCase();

        if (
          status !== "P" &&
          status !== "A"
        ) {
          throw new Error(
            "Invalid attendance for " +
            item.name
          );
        }

        sheet
          .getRange(
            Number(item.row),
            column
          )
          .setValue(status);

        if (status === "P") {
          present++;
        } else {
          absent++;
        }

      }
    );

    SpreadsheetApp.flush();

    return {

      success: true,

      present: present,

      absent: absent,

      message:
        "Attendance saved successfully."

    };

  } finally {

    lock.releaseLock();

  }
}


/* ================================
   TEST
================================ */

function testConnection() {

  const sheet =
    getSheet();

  const students =
    getStudents();

  return {

    sheet:
      sheet.getName(),

    studentCount:
      students.length,

    firstStudent:
      students.length
        ? students[0]
        : null,

    message:
      "Connection successful."

  };
}

/* ================================
   VERCEL API
================================ */

const VERCEL_API_TOKEN = "7fQC2Sx7d6RHKm-73gAE2FZjpPUebFIdcy1YxUg8slo";

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyApiToken(e, body) {
  const supplied =
    (e && e.parameter && e.parameter.token) ||
    (body && body.token) ||
    "";

  if (supplied !== VERCEL_API_TOKEN) {
    throw new Error("Unauthorized API request.");
  }
}

function handleApiGet(e) {

  try {

    verifyApiToken(e, null);

    const action =
      String(e.parameter.action || "").trim();

    if (action === "get") {

      const dateString =
        String(e.parameter.date || "").trim();

      if (!dateString) {
        throw new Error("Date is required.");
      }

      return jsonResponse(getAttendance(dateString));
    }

    if (action === "saveOne") {

      const dateString =
        String(e.parameter.date || "").trim();

      const row =
        Number(e.parameter.row);

      const status =
        String(e.parameter.status || "").trim();

      return jsonResponse(
        saveOneAttendance(
          dateString,
          row,
          status
        )
      );
    }

    return jsonResponse({
      success: false,
      error: "Unknown API action."
    });

  } catch (error) {

    return jsonResponse({
      success: false,
      error: error.message
    });

  }
}

function doPost(e) {

  try {

    const body =
      JSON.parse(
        (e && e.postData && e.postData.contents)
          ? e.postData.contents
          : "{}"
      );

    verifyApiToken(e, body);

    const action =
      String(body.action || "").trim();

    if (action === "saveOne") {

      return jsonResponse(
        saveOneAttendance(
          body.date,
          Number(body.row),
          body.status
        )
      );
    }

    if (action === "saveAll") {

      return jsonResponse(
        saveAllAttendance(
          body.date,
          body.attendanceData || []
        )
      );
    }

    throw new Error("Unknown API action.");

  } catch (error) {

    return jsonResponse({
      success: false,
      error: error.message
    });

  }
}
