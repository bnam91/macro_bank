import path from "path";
import os from "os";
import { createRequire } from "module";
import { google } from "googleapis";
import { preprocessAccountInfo } from "../config/bank-config.js";

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

function extractSpreadsheetId(sheetUrlOrId) {
  if (!sheetUrlOrId) return "";
  const match = String(sheetUrlOrId).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  return String(sheetUrlOrId);
}

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

async function fetchSheetValues({ sheetUrl, sheetName, authModulePath, range }) {
  const spreadsheetId = extractSpreadsheetId(sheetUrl);
  if (!spreadsheetId) {
    throw new Error("스프레드시트 ID를 확인할 수 없습니다.");
  }

  const auth = await getAuthClient(authModulePath);
  const sheets = google.sheets({ version: "v4", auth });
  const targetRange = range || `${sheetName}!A:Z`;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: targetRange,
    majorDimension: "ROWS",
  });

  return res.data.values || [];
}

// 시트에 값 쓰기 함수
async function updateSheetValue({ sheetUrl, sheetName, authModulePath, rowIndex, columnIndex, value }) {
  const spreadsheetId = extractSpreadsheetId(sheetUrl);
  if (!spreadsheetId) {
    throw new Error("스프레드시트 ID를 확인할 수 없습니다.");
  }

  const auth = await getAuthClient(authModulePath);
  const sheets = google.sheets({ version: "v4", auth });
  
  // 열 인덱스를 알파벳으로 변환 (0=A, 1=B, ..., 16=Q)
  const columnLetter = String.fromCharCode(65 + columnIndex); // A=65
  const range = `${sheetName}!${columnLetter}${rowIndex + 1}`; // +1은 시트 행 번호가 1부터 시작하므로

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: range,
    valueInputOption: 'USER_ENTERED',
    resource: {
      values: [[value]]
    }
  });
}

// 주민번호 검증 함수 (하이픈 제외 13자리 확인)
function validateResidentNumber(residentNumber) {
  if (!residentNumber || residentNumber.trim() === "") {
    return true; // 빈 값은 유효함
  }
  
  // 하이픈 제거 후 숫자만 추출
  const digitsOnly = residentNumber.replace(/[^0-9]/g, "");
  
  // 13자리인지 확인
  return digitsOnly.length === 13;
}

async function buildTransferData(rows, columnMapping, sheetConfig) {
  const processedData = [];
  const errorsToWrite = []; // Q열에 쓸 오류 정보 저장

  // 기본 열 매핑 (columnMapping이 없을 경우 사용)
  const defaultMapping = {
    productName: 4,    // E열: 제품
    customerName: 5,   // F열: 이름
    accountInfo: 8,    // I열: 계좌번호
    amount: 10         // K열: 금액
  };

  const mapping = columnMapping || defaultMapping;
  const STATUS_COLUMN_INDEX = 16; // Q열: 상태 (인덱스 16, A=0부터 시작)
  const RESIDENT_NUMBER_COLUMN_INDEX = 9; // J열: 주민번호 (인덱스 9, A=0부터 시작)
  const MAX_ROWS = 10; // 최대 10개 행만 읽기

  // 1단계: Q열이 빈 값인 행들 중에서 J열(주민번호) 검증
  for (let i = 0; i < rows.length; i++) {
    if (i === 0) continue; // 헤더 제외

    const row = rows[i];
    // 최소 컬럼 수 확인
    const minColumns = Math.max(
      mapping.productName,
      mapping.customerName,
      mapping.accountInfo,
      mapping.amount,
      STATUS_COLUMN_INDEX,
      RESIDENT_NUMBER_COLUMN_INDEX
    ) + 1;
    
    if (!Array.isArray(row) || row.length < minColumns) continue;

    // Q열(상태) 확인 - 빈값인 경우만 검증 진행
    const status = (row[STATUS_COLUMN_INDEX] || "").trim();
    if (status !== "") {
      // 상태가 비어있지 않으면 건너뛰기
      continue;
    }

    // J열(주민번호) 검증
    const residentNumber = (row[RESIDENT_NUMBER_COLUMN_INDEX] || "").trim();
    if (residentNumber !== "" && !validateResidentNumber(residentNumber)) {
      // 주민번호가 입력되어 있는데 13자리가 아니면 Q열에 오류 기록
      errorsToWrite.push({
        rowIndex: i,
        columnIndex: STATUS_COLUMN_INDEX,
        value: '주민번호 오류'
      });
      console.log(`행 ${i + 1}: 주민번호 오류 감지 - "${residentNumber}" (하이픈 제외 ${residentNumber.replace(/[^0-9]/g, "").length}자리)`);
    }
  }

  // 2단계: Q열에 오류 기록
  if (errorsToWrite.length > 0 && sheetConfig) {
    try {
      for (const error of errorsToWrite) {
        await updateSheetValue({
          sheetUrl: sheetConfig.sheetUrl,
          sheetName: sheetConfig.sheetName,
          authModulePath: sheetConfig.authModulePath,
          rowIndex: error.rowIndex,
          columnIndex: error.columnIndex,
          value: error.value
        });
      }
      console.log(`✅ ${errorsToWrite.length}개 행의 Q열에 '주민번호 오류' 기록 완료`);
    } catch (error) {
      console.error(`❌ Q열 오류 기록 실패: ${error.message}`);
    }
  }

  // 3단계: Q열이 빈 값인 행 중 최대 10개만 반환
  for (let i = 0; i < rows.length; i++) {
    if (i === 0) continue; // 헤더 제외

    // 최대 10개까지만 처리
    if (processedData.length >= MAX_ROWS) {
      console.log(`최대 ${MAX_ROWS}개 행까지만 처리합니다.`);
      break;
    }

    const row = rows[i];
    // 최소 컬럼 수 확인
    const minColumns = Math.max(
      mapping.productName,
      mapping.customerName,
      mapping.accountInfo,
      mapping.amount,
      STATUS_COLUMN_INDEX
    ) + 1;
    
    if (!Array.isArray(row) || row.length < minColumns) continue;

    // Q열(상태) 확인 - 빈값인 경우만 처리
    const status = (row[STATUS_COLUMN_INDEX] || "").trim();
    if (status !== "") {
      // 상태가 비어있지 않으면 건너뛰기
      continue;
    }

    const productName = row[mapping.productName] || "";
    const customerName = row[mapping.customerName] || "";
    const accountInfo = row[mapping.accountInfo] || "";
    const amount = row[mapping.amount] || 0;

    const { bankName, accountNumber } = preprocessAccountInfo(accountInfo);
    const nameProduct = `${customerName}${productName}`;

    processedData.push({
      bank: bankName,
      accountNumber,
      nameProduct,
      productName,
      amount,
    });
  }

  return processedData;
}

async function loadSheetTransferData({ sheetUrl, sheetName, authModulePath, columnMapping }) {
  const rows = await fetchSheetValues({ sheetUrl, sheetName, authModulePath });
  const sheetConfig = { sheetUrl, sheetName, authModulePath };
  return buildTransferData(rows, columnMapping, sheetConfig);
}

export {
  loadSheetTransferData,
  extractSpreadsheetId,
};
