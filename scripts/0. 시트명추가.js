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

/**
 * 스프레드시트의 모든 시트명을 가져오는 함수
 * @param {string} spreadsheetId - 스프레드시트 ID
 * @param {string} authModulePath - 인증 모듈 경로 (선택사항)
 * @returns {Promise<string[]>} 시트명 배열
 */
async function getSheetNames(spreadsheetId, authModulePath) {
  console.log("스프레드시트에서 시트명 가져오기 시작");
  try {
    const auth = await getAuthClient(authModulePath);
    const sheets = google.sheets({ version: "v4", auth });

    // 스프레드시트 정보 가져오기
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId,
    });

    // 시트 정보에서 시트명만 추출
    const sheetList = spreadsheet.data.sheets || [];
    const sheetNames = sheetList.map((sheet) => sheet.properties.title);

    // '완료'가 포함된 시트명 제외
    const filteredSheetNames = sheetNames.filter((name) => !name.includes("완료"));

    console.log(`시트명 가져오기 성공: ${filteredSheetNames.length} 시트`);
    return filteredSheetNames;
  } catch (error) {
    console.error(`시트명 가져오기 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 시트명을 이용해 QUERY 식 생성
 * @param {string} spreadsheetId - 스프레드시트 ID
 * @param {string[]} sheetNames - 시트명 배열
 * @returns {string} 생성된 QUERY 식
 */
function createQueryFormula(spreadsheetId, sheetNames) {
  let formula = "=QUERY({";

  for (let i = 0; i < sheetNames.length; i++) {
    const name = sheetNames[i];
    // 시트명에 공백이나 특수문자가 있는 경우 작은따옴표로 감싸기
    let formattedName;
    if (name.includes(" ") || name.includes("(") || name.includes(")")) {
      formattedName = `'${name}'`;
    } else {
      formattedName = name;
    }

    const importLine = `IMPORTRANGE("https://docs.google.com/spreadsheets/d/${spreadsheetId}", "${formattedName}!A1:P1000")`;

    // 마지막 줄이 아니면 세미콜론 추가
    if (i < sheetNames.length - 1) {
      formula += importLine + ";";
    } else {
      formula += importLine;
    }
  }

  formula += '}, "select * where Col16 = \'입금요청\'", 0)';
  return formula;
}

/**
 * 스프레드시트의 A2 셀에 생성된 식 입력
 * @param {string} targetSpreadsheetId - 타겟 스프레드시트 ID
 * @param {string} formula - 입력할 식
 * @param {string} authModulePath - 인증 모듈 경로 (선택사항)
 * @returns {Promise<object>} 업데이트 결과
 */
async function writeToSpreadsheet(targetSpreadsheetId, formula, authModulePath) {
  console.log(`스프레드시트 ${targetSpreadsheetId}에 식 입력 시작`);
  try {
    const auth = await getAuthClient(authModulePath);
    const sheets = google.sheets({ version: "v4", auth });

    // A2 셀에 쓰기
    const result = await sheets.spreadsheets.values.update({
      spreadsheetId: targetSpreadsheetId,
      range: "A2",
      valueInputOption: "USER_ENTERED", // 'USER_ENTERED'로 설정하여 수식으로 인식되도록 함
      resource: {
        values: [[formula]],
      },
    });

    console.log(`식 입력 성공: ${result.data.updatedCells} 셀 업데이트됨`);
    return result.data;
  } catch (error) {
    console.error(`식 입력 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 메인 함수
 */
async function main() {
  // 소스 스프레드시트 ID (시트명을 가져올 스프레드시트)
  const SOURCE_SPREADSHEET_ID = "1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs";

  // 타겟 스프레드시트 ID (식을 쓸 스프레드시트)
  const TARGET_SPREADSHEET_ID = "1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c";

  try {
    // 시트명 가져오기
    const sheetNames = await getSheetNames(SOURCE_SPREADSHEET_ID);

    // QUERY 식 생성
    const formula = createQueryFormula(SOURCE_SPREADSHEET_ID, sheetNames);

    // 생성된 식 출력 (확인용)
    console.log("생성된 QUERY 식:");
    console.log(formula);

    // 타겟 스프레드시트에 식 쓰기
    await writeToSpreadsheet(TARGET_SPREADSHEET_ID, formula);
    console.log(
      `QUERY 식이 스프레드시트 ${TARGET_SPREADSHEET_ID}의 A2 셀에 성공적으로 입력되었습니다.`
    );
  } catch (error) {
    console.error(`작업 중 오류가 발생했습니다: ${error.message}`);
    console.error(error.stack);
    process.exit(1);
  }
}

// ESM에서 직접 실행 여부 확인
const __filename = fileURLToPath(import.meta.url);
if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  main();
}

export {
  getSheetNames,
  createQueryFormula,
  writeToSpreadsheet,
  main,
};
