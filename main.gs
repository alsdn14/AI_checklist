const GEMINI_API_KEY = "재미나이 API 키 넣는 공간";

function generateTestCases_WithResultDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const inputValues = sheet.getRange(`A2:R${lastRow}`).getValues();
  const inputText = inputValues
    .map(row => row.join(" ").trim())
    .filter(line => line.length > 0)
    .join("\n");

  if (!inputText) {
    SpreadsheetApp.getUi().alert("⚠️ A2 ~ R 셀에 기획 문서를 입력하세요.");
    return;
  }

  const prompt = `
"${inputText}" 는 게임 기획 문서입니다. 이 기획문서에 대한 게임 테스트 케이스를 작성해 줘.

형식: 중분류 | 소분류 | 테스트 내용

조건:
- 중분류는 해당 기획 문서의 대표 주제이며, 반복적으로 쓰지 말고 필요한 경우에만 작성해줘
- 소분류는 테스트 대상 요소 (예: 버튼, 기능, 조건) 명확히 작성 해줘
- 형식: "중분류 | 소분류 | 테스트 내용" 을 반드시 지켜줘. 각 필드는 절대 생략하거나 합치지 말고 줄마다 정확히 3개의 필드로 작성해줘
- 예시는 쓰지 말고 바로 테스트 케이스 목록부터 작성해줘
- 테스트 내용은 "~되는지 확인"과 같이, 테스터가 실제 수행할 행동과 확인 요소가 명확히 드러나도록 작성해줘
- 한 줄에는 반드시 하나의 테스트만 작성해줘
- 테스트 내용에 "그리고", "및", "또는", "혹은", "같이", "하면서", "진행하고" 등의 접속사(또는 쉼표 등)를 포함하여 2개 이상의 테스트를 한 줄에 작성하지 마세요.
- 예: "A를 진행하고, B가 되는지 확인" → 이렇게 작성하지 말고, 아래처럼 나눠 작성하세요:
   - A를 진행할 수 있는지 확인
   - B가 되는지 확인
- 각 테스트 케이스는 개별적인 목적을 가져야 하며, 상세하게 작성해줘
- 동적 테스트 진행 가능한 테스트 케이스로 작성해줘
- 내용을 보고 유사 게임들에게서 발생했던 과거 버그 발생 사건들을 참고해서 예외 사항 항목들만 따로 분리해서 같이 써줘
- 최소 5줄 이상 작성해줘 (가능하면 10줄 이상)
`.trim();

  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    Logger.log("🔵 Gemini 응답 원문: " + response.getContentText());
    const result = json.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!result) throw new Error("Gemini 응답 없음");

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`TC_${timestamp}`);

    // ✅ 헤더 작성
    const headerRange = newSheet.getRange("B1:E1");
    headerRange.setValues([["중분류", "소분류", "테스트 내용", "확인 결과"]]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f1f3f4");
    headerRange.setHorizontalAlignment("center");
    newSheet.setRowHeight(1, 28);
    newSheet.setFrozenRows(1);

    // ✅ Gemini 응답 줄 나누기 + 필터링
    const lines = result.split("\n").filter(line => {
      const trimmed = line.trim();
      return (
        trimmed.length > 0 &&
        trimmed.split("|").length >= 3 &&
        !trimmed.includes("---") &&
        !trimmed.toLowerCase().includes("중분류")
      );
    });

    let rowIndex = 2;

    // ✅ 줄별 파싱 및 쓰기
    lines.forEach((line, index) => {
      const parts = line.split("|").map(p => p.trim());

      if (parts.length < 3) {
        Logger.log(`⚠️ [무시됨] ${index + 1}행 - 필드 부족: ${line}`);
        return;
      }

      const 중분류 = parts[0];
      const 소분류 = parts[1];
      const 테스트내용 = parts.slice(2).join(" ");

      if (!중분류 || !소분류 || !테스트내용) {
        Logger.log(`⚠️ [무시됨] ${index + 1}행 - 필드 비어있음: ${line}`);
        return;
      }

      newSheet.getRange(rowIndex, 2, 1, 3).setValues([[중분류, 소분류, 테스트내용]]);
      newSheet.getRange(rowIndex, 5).setValue("N/T");
      rowIndex++;
    });

    // ✅ 드롭다운 생성
    const resultRange = newSheet.getRange(2, 5, rowIndex - 2, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pass", "Fail", "N/T", "N/A"], true)
      .setAllowInvalid(false)
      .build();
    resultRange.setDataValidation(rule);
    resultRange.setHorizontalAlignment("center");

    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert("✅ 체크리스트 시트 생성 완료!");

  } catch (e) {
    Logger.log("❌ 오류 발생: " + e.message);
    SpreadsheetApp.getUi().alert("❌ 오류: " + e.message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⏵ 체크리스트 생성")
    .addItem("클릭해서 실행", "generateTestCases_WithResultDropdown")
    .addToUi();
}
