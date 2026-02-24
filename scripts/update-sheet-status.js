import { google } from "googleapis";
import path from "path";
import os from "os";
import { createRequire } from "module";
import { fileURLToPath } from "url";

// 인증 모듈 경로 설정
const DEFAULT_AUTH_MODULE_PATH = path.join(
  os.homedir(),
  "Documents",
  "github_cloud",
  "module_auth",
  "auth.js",
);

// 홈 경로 확장 함수
function expandHomePath(inputPath) {
  if (!inputPath) return inputPath;
  if (inputPath === "~") return os.homedir();
  if (inputPath.startsWith("~/")) {
    return path.join(os.homedir(), inputPath.slice(2));
  }
  return inputPath;
}

// 인증 클라이언트 가져오기
async function getAuthClient(authModulePath) {
  const resolvedPath = expandHomePath(authModulePath || DEFAULT_AUTH_MODULE_PATH);
  let getCredentials;
  try {
    // ESM에서 CommonJS 모듈을 동적으로 로드하기 위해 createRequire 사용
    const require = createRequire(import.meta.url);
    ({ getCredentials } = require(resolvedPath));
  } catch (error) {
    throw new Error(`인증 모듈을 불러오지 못했습니다: ${resolvedPath} (${error.message})`);
  }

  if (typeof getCredentials !== "function") {
    throw new Error("인증 모듈에 getCredentials 함수가 없습니다.");
  }

  return getCredentials();
}

const SPREADSHEET_ID = "1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs";

// 색상 매핑
function getColor(colorName) {
  const colors = {
    "진한 회색 1": { red: 0.8, green: 0.8, blue: 0.8 },
    주황색: { red: 1.0, green: 0.6, blue: 0.0 },
  };
  return colors[colorName] || { red: 1, green: 1, blue: 1 };
}

// 입금완료 텍스트 생성
function getDepositCompleteText() {
  const today = new Date();
  const year = String(today.getFullYear()).slice(-2);
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `입금완료_${year}${month}${day}`;
}

// 행 색상 업데이트 요청 생성
function updateRowColor(sheetId, rowIndex, color) {
  return {
    updateCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: rowIndex - 1,
        endRowIndex: rowIndex,
        startColumnIndex: 1, // B열부터 시작
        endColumnIndex: 16, // P열까지
      },
      rows: [
        {
          values: Array(15).fill({
            userEnteredFormat: { backgroundColor: getColor(color) },
          }),
        },
      ],
      fields: "userEnteredFormat.backgroundColor",
    },
  };
}

// 셀 값 업데이트 요청 생성
function updateCellValue(sheetId, rowIndex, columnIndex, newValue) {
  return {
    updateCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: rowIndex - 1,
        endRowIndex: rowIndex,
        startColumnIndex: columnIndex,
        endColumnIndex: columnIndex + 1,
      },
      rows: [
        {
          values: [
            {
              userEnteredValue: { stringValue: newValue },
            },
          ],
        },
      ],
      fields: "userEnteredValue.stringValue",
    },
  };
}

// 시트 처리 함수
async function processSheet(sheets, sheetId, sheetName) {
  console.log(`Processing sheet: ${sheetName} (ID: ${sheetId})`);

  const rangeName = `${sheetName}!A1:P1000`; // P열까지만 범위 지정

  let values;
  try {
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: rangeName,
    });
    values = result.data.values || [];
  } catch (error) {
    console.error(`${sheetName} 데이터 읽기 실패: ${error.message}`);
    return;
  }

  if (!values || values.length === 0) {
    console.log(`${sheetName}에서 데이터를 찾을 수 없습니다.`);
    return;
  }

  console.log(`${sheetName}의 총 행 수: ${values.length}`);

  const requests = [];
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    if (row.length >= 16) {
      // P열이 존재하는 경우
      const status = (row[15] || "").trim();
      const statusLower = status.toLowerCase();

      if (statusLower === "입금요청") {
        requests.push(updateRowColor(sheetId, rowIndex + 1, "진한 회색 1"));
        requests.push(
          updateCellValue(sheetId, rowIndex + 1, 15, getDepositCompleteText())
        );
        console.log(`Row ${rowIndex + 1}: 입금요청 처리`);
      } else if (statusLower === "입금오류") {
        requests.push(updateRowColor(sheetId, rowIndex + 1, "주황색"));
        console.log(`Row ${rowIndex + 1}: 입금오류 처리`);
      }
    }
  }

  // 원본 시트 업데이트
  if (requests.length > 0) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: requests,
        },
      });
      console.log(
        `${sheetName} 업데이트 완료 (${Math.floor(requests.length / 2)} 행 처리됨)`
      );
    } catch (error) {
      console.error(`${sheetName} 업데이트 실패: ${error.message}`);
    }
  } else {
    console.log(`${sheetName}에서 업데이트할 내용이 없습니다.`);
  }
}

// 메인 함수
async function updateSheets() {
  let retryCount = 0;
  const maxRetries = 5;

  while (retryCount < maxRetries) {
    try {
      const auth = await getAuthClient();
      const sheets = google.sheets({ version: "v4", auth });

      // 스프레드시트 메타데이터 가져오기
      const sheetMetadata = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID,
      });

      const sheetList = sheetMetadata.data.sheets || [];
      const sheetMap = {};
      for (const sheet of sheetList) {
        const title = sheet.properties.title;
        const sheetId = sheet.properties.sheetId;
        sheetMap[title] = sheetId;
      }

      // 각 시트 처리
      for (const [sheetName, sheetId] of Object.entries(sheetMap)) {
        // '완료'가 포함된 시트명은 제외
        if (!sheetName.includes("완료")) {
          await processSheet(sheets, sheetId, sheetName);
        } else {
          console.log(`'${sheetName}' 시트는 '완료'가 포함되어 제외됩니다.`);
        }
      }

      // 성공적으로 완료되면 루프 종료
      break;
    } catch (error) {
      if (error.code === 502 || (error.response && error.response.status === 502)) {
        retryCount++;
        if (retryCount < maxRetries) {
          console.log("서버 오류가 발생했습니다. 30초 후에 다시 시도합니다.");
          await new Promise((resolve) => setTimeout(resolve, 30000));
          continue;
        }
      }
      console.error(`오류가 발생했습니다: ${error.message}`);
      if (error.stack) {
        console.error(error.stack);
      }
      process.exit(1);
    }
  }
}

// 메인 실행
async function main() {
  try {
    await updateSheets();
  } catch (error) {
    console.error(`작업 중 오류가 발생했습니다: ${error.message}`);
    if (error.stack) {
      console.error(error.stack);
    }
    process.exit(1);
  }
}

// ESM에서 직접 실행 여부 확인
const __filename = fileURLToPath(import.meta.url);
if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  main();
}

export { updateSheets, main };
