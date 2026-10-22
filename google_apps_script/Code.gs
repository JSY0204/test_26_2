/**
 * INU 벽돌깨기 게임 시트 연동 Google Apps Script
 * 현재 구글 스프레드시트 헤더 컬럼: timestamp | name | finishtime (A1, B1, C1)
 */

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()[0];
}

/**
 * 1. Read (조회): doGet
 */
function doGet(e) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return createJsonResponse({ success: true, data: [] });
    }

    // 1행은 헤더: ['timestamp', 'name', 'finishtime']
    const headers = data[0];
    const rows = data.slice(1);
    
    const records = rows.map((row, index) => {
      const obj = { rowIndex: index + 2 };
      headers.forEach((h, colIdx) => {
        const key = String(h).trim();
        if (key) {
          obj[key] = row[colIdx];
        }
      });
      return obj;
    });

    return createJsonResponse({
      success: true,
      data: records
    });
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: err.toString()
    });
  }
}

/**
 * 2. Create, Update, Delete: doPost
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet();
    
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
    }

    const action = body.action || "create";

    // 1) CREATE (신규 행 등록)
    // 컬럼 순서: A(timestamp), B(name), C(finishtime)
    if (action === "create") {
      const timestamp = body.timestamp || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      const name = body.name || "";
      const finishtime = body.finishtime || "";

      // 텍스트 형식으로 정확히 A, B, C 열에 추가
      sheet.appendRow([timestamp, name, finishtime]);

      return createJsonResponse({
        success: true,
        message: "구글 시트에 성공적으로 등록되었습니다.",
        data: {
          timestamp,
          name,
          finishtime,
          rowIndex: sheet.getLastRow()
        }
      });
    }

    // 2) UPDATE (행 수정)
    else if (action === "update") {
      const data = sheet.getDataRange().getValues();
      const targetRowIndex = Number(body.rowIndex);

      if (!targetRowIndex || targetRowIndex <= 1 || targetRowIndex > data.length) {
        return createJsonResponse({ success: false, error: "수정할 행(rowIndex)을 찾을 수 없습니다." });
      }

      const currentRow = data[targetRowIndex - 1];
      const updatedTimestamp = body.timestamp !== undefined ? body.timestamp : currentRow[0];
      const updatedName = body.name !== undefined ? body.name : currentRow[1];
      const updatedFinishTime = body.finishtime !== undefined ? body.finishtime : currentRow[2];

      sheet.getRange(targetRowIndex, 1, 1, 3).setValues([[
        updatedTimestamp,
        updatedName,
        updatedFinishTime
      ]]);

      return createJsonResponse({
        success: true,
        message: "성공적으로 수정되었습니다.",
        data: {
          rowIndex: targetRowIndex,
          timestamp: updatedTimestamp,
          name: updatedName,
          finishtime: updatedFinishTime
        }
      });
    }

    // 3) DELETE (행 삭제)
    else if (action === "delete") {
      const data = sheet.getDataRange().getValues();
      const targetRowIndex = Number(body.rowIndex);

      if (!targetRowIndex || targetRowIndex <= 1 || targetRowIndex > data.length) {
        return createJsonResponse({ success: false, error: "삭제할 행(rowIndex)을 찾을 수 없습니다." });
      }

      sheet.deleteRow(targetRowIndex);

      return createJsonResponse({
        success: true,
        message: "성공적으로 삭제되었습니다.",
        deletedRowIndex: targetRowIndex
      });
    }

    return createJsonResponse({ success: false, error: "지원하지 않는 action: " + action });

  } catch (err) {
    return createJsonResponse({
      success: false,
      error: err.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
