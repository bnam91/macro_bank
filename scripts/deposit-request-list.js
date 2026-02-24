/**
 * 입금요청 현황판 스프레드시트의 모든 시트를 순회하며,
 * P열 값이 '입금요청'인 행을 찾아 출력합니다.
 * 시트명에 '완료_'가 포함된 시트는 제외합니다.
 * 시트: https://docs.google.com/spreadsheets/d/1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs
 *
 * update-sheet-status.js 의 인증·시트 순회 방식을 참고했습니다.
 */
import { google } from "googleapis";
import path from "path";
import os from "os";
import { createRequire } from "module";
import { fileURLToPath } from "url";

const DEFAULT_AUTH_MODULE_PATH = path.join(
  os.homedir(),
  "Documents",
  "github_cloud",
  "module_auth",
  "auth.js",
);

function expandHomePath(inputPath) {
  if (!inputPath) return inputPath;
  if (inputPath === "~") return os.homedir();
  if (inputPath.startsWith("~/")) {
    return path.join(os.homedir(), inputPath.slice(2));
  }
  return inputPath;
}

async function getAuthClient(authModulePath) {
  const resolvedPath = expandHomePath(authModulePath || DEFAULT_AUTH_MODULE_PATH);
  let getCredentials;
  try {
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
const P_COLUMN_INDEX = 15; // P열 (0-based)
const STATUS_VALUE = "입금요청";

/**
 * P열이 '입금요청'인 행만 필터링
 */
function filterDepositRequestRows(rows) {
  if (!rows || rows.length === 0) return [];
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const status = (row[P_COLUMN_INDEX] || "").trim();
    if (status === STATUS_VALUE) {
      result.push({ rowIndex: i + 1, row });
    }
  }
  return result;
}

/**
 * 한 시트에서 P열 '입금요청' 행 조회 후 출력하고, 항목 목록 반환
 * @returns {Array<{ sheetName: string, rowIndex: number, row: string[] }>}
 */
async function processSheet(sheets, sheetName) {
  const rangeName = `${sheetName}!A1:P1000`;

  let values;
  try {
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: rangeName,
    });
    values = result.data.values || [];
  } catch (error) {
    console.error(`  ${sheetName} 데이터 읽기 실패: ${error.message}`);
    return [];
  }

  if (!values || values.length === 0) return [];

  const matched = filterDepositRequestRows(values);
  if (matched.length === 0) return [];

  const header = values[0] || [];
  console.log(`\n[시트: ${sheetName}] P열 '입금요청' 행 ${matched.length}개`);
  const items = [];
  for (const { rowIndex, row } of matched) {
    items.push({ sheetName, rowIndex, row });
    console.log(`  --- 행 ${rowIndex} ---`);
    header.forEach((colName, i) => {
      if (colName != null && colName !== "" && row[i] != null && row[i] !== "") {
        console.log(`    ${colName}: ${row[i]}`);
      }
    });
    console.log(`    (전체 행): ${JSON.stringify(row)}`);
    console.log("");
  }
  return items;
}

/**
 * 모든 시트 순회하여 P열 '입금요청' 행 출력 (시트명에 '완료_' 포함 시 제외)
 * @returns {Promise<{ spreadsheetId: string, sheetUrl: string, items: Array<{ sheetName: string, rowIndex: number, row: string[] }> }>}
 */
export async function runDepositRequestList() {
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: "v4", auth });

  const sheetMetadata = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const sheetList = sheetMetadata.data.sheets || [];
  const sheetMap = {};
  for (const s of sheetList) {
    const title = s.properties.title;
    sheetMap[title] = s.properties.sheetId;
  }

  console.log("\n📋 입금요청 현황 — 모든 시트에서 P열 '입금요청' 행 조회 (시트명 '완료_' 제외)\n");
  console.log(`대상 스프레드시트: ${SPREADSHEET_ID}`);
  console.log(`시트 수: ${Object.keys(sheetMap).length}개\n`);

  const allItems = [];
  for (const [sheetName] of Object.entries(sheetMap)) {
    if (sheetName.includes("완료_")) continue;
    const items = await processSheet(sheets, sheetName);
    allItems.push(...items);
  }

  console.log("\n✅ 조회 완료");
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
  return { spreadsheetId: SPREADSHEET_ID, sheetUrl, items: allItems };
}

const __filename = fileURLToPath(import.meta.url);
if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(__filename)) {
  runDepositRequestList().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
