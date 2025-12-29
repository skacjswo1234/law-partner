// 웹앱 엔드포인트 응답
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var message = params.msg || "OK";
  return ContentService.createTextOutput("pong: " + message);
}

// 웹에서 폼 데이터를 받아서 시트에 추가하는 함수
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var postData = (e && e.parameter) ? e.parameter : {};
    
    // 헤더 배열 정의
    var expectedHeaders = [
      "제출일시",
      "이름",
      "전화번호",
      "직업",
      "채무금액",
      "상담가능시간",
      "연체여부"
    ];
    
    // 시트 헤더 확인 및 자동 설정
    ensureSheetHeaders(sheet, expectedHeaders);
    
    // 현재 시간
    var timestamp = new Date();
    
    // 데이터 배열 (시트 헤더 순서에 맞춤: 제출일시, 이름, 전화번호, 직업, 채무금액, 상담가능시간, 연체여부)
    var rowData = [
      timestamp,
      postData.name || "",
      postData.phone || "",
      postData.job || "",
      postData.debt || "",
      postData.consultation_time || "",
      postData.overdue || ""
    ];
    
    // 시트에 데이터 추가
    sheet.appendRow(rowData);
    
    // 이메일 전송
    sendEmailNotification(rowData);
    
    // 성공 응답
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "상담 신청이 완료되었습니다."
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log("오류 발생: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 이메일 알림 전송 함수
function sendEmailNotification(rowData) {
  try {
    var email = "bbong1019@gmail.com";
    var subject = "[법무법인 파트너] 새 문의가 접수되었습니다 [나이스]";
    
    var headers = ["제출일시", "이름", "전화번호", "직업", "채무금액", "상담가능시간", "연체여부"];
    
    var bodyLines = [];
    bodyLines.push("새로운 상담 신청이 접수되었습니다.");
    bodyLines.push("");
    bodyLines.push("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    bodyLines.push("");
    
    for (var i = 0; i < headers.length && i < rowData.length; i++) {
      if (rowData[i]) {
        bodyLines.push(headers[i] + ": " + rowData[i]);
      }
    }
    
    bodyLines.push("");
    bodyLines.push("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    bodyLines.push("");
    bodyLines.push("구글 시트에서 확인: " + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    
    var htmlBody = bodyLines.join("<br>");
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("이메일 전송 완료: " + email);
    
  } catch (error) {
    Logger.log("이메일 전송 오류: " + error.toString());
  }
}

// 폼 제출 시 자동 실행 (트리거 설정 필요)
function onFormSubmit(e) {
  try {
    var email = "bbong1019@gmail.com";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // e가 없거나 undefined인 경우 체크
    if (!e) {
      Logger.log("이벤트 객체가 없습니다. 시트의 마지막 행을 읽어옵니다.");
    }
    
    // 시트의 헤더 행 확인 (일반적으로 1행)
    var headerRow = 1;
    var lastRow = sheet.getLastRow();
    
    // 데이터가 없으면 종료
    if (lastRow <= headerRow) {
      Logger.log("전송할 데이터가 없습니다.");
      return;
    }
    
    // 새로 추가된 행의 데이터 가져오기
    var rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 이메일 제목
    var subject = "[법무법인 파트너] 새 문의가 접수되었습니다";
    
    // 이메일 본문 작성
    var bodyLines = [];
    bodyLines.push("새로운 상담 신청이 접수되었습니다.");
    bodyLines.push("");
    bodyLines.push("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    bodyLines.push("");
    
    // 시트의 헤더와 데이터 매칭하여 이메일 본문 작성
    var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var value = rowData[i] || "";
      
      if (header && value) {
        // 타임스탬프는 "제출일시"로 표시
        if (header.toString().toLowerCase().includes("timestamp") || 
            header.toString().includes("제출") || 
            header.toString().includes("타임스탬프")) {
          bodyLines.push("제출일시: " + value);
        } else {
          bodyLines.push(header + ": " + value);
        }
      }
    }
    
    bodyLines.push("");
    bodyLines.push("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    bodyLines.push("");
    bodyLines.push("구글 시트에서 확인: " + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    
    // HTML 형식으로 이메일 전송
    var htmlBody = bodyLines.join("<br>");
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("이메일 전송 완료: " + email);
    
  } catch (error) {
    Logger.log("오류 발생: " + error.toString());
    // 오류가 발생해도 폼 제출은 정상적으로 처리되도록 함
  }
}

// 시트 헤더 확인 및 자동 설정 함수
function ensureSheetHeaders(sheet, expectedHeaders) {
  try {
    var lastRow = sheet.getLastRow();
    var needsHeader = false;
    
    // 시트가 비어있거나 헤더가 없는 경우
    if (lastRow === 0) {
      needsHeader = true;
    } else {
      // 1행의 헤더 확인
      var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 헤더가 없거나 개수가 맞지 않거나 내용이 다른 경우
      if (existingHeaders.length !== expectedHeaders.length) {
        needsHeader = true;
      } else {
        // 각 헤더가 일치하는지 확인
        for (var i = 0; i < expectedHeaders.length; i++) {
          if (existingHeaders[i] !== expectedHeaders[i]) {
            needsHeader = true;
            break;
          }
        }
      }
    }
    
    // 헤더가 필요한 경우 설정
    if (needsHeader) {
      // 기존 헤더가 있으면 덮어쓰기, 없으면 새로 생성
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      
      // 헤더 행 스타일 설정
      var headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("#ffffff");
      
      // 열 너비 자동 조정
      for (var i = 1; i <= expectedHeaders.length; i++) {
        sheet.autoResizeColumn(i);
      }
      
      Logger.log("시트 헤더가 자동으로 설정되었습니다.");
    }
    
  } catch (error) {
    Logger.log("헤더 설정 오류: " + error.toString());
    // 오류가 발생해도 계속 진행
  }
}

// 시트 헤더 수동 설정 함수 (한 번만 실행하면 됩니다)
function setupSheetHeaders() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 헤더 배열 (구글 폼은 첫 번째 열에 타임스탬프를 자동으로 넣습니다)
    var headers = [
      "제출일시",
      "이름",
      "전화번호",
      "직업",
      "채무금액",
      "상담가능시간",
      "연체여부"
    ];
    
    ensureSheetHeaders(sheet, headers);
    
    Logger.log("시트 헤더가 성공적으로 설정되었습니다.");
    return "시트 헤더 설정 완료!";
    
  } catch (error) {
    Logger.log("오류 발생: " + error.toString());
    return "오류: " + error.toString();
  }
}

