/**
 * 실시간 입금요청내역 시트에서 Q열(상태)이 "이체완료"인 행을 찾아 출력합니다.
 * 시트: https://docs.google.com/spreadsheets/d/1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c
 * 매칭된 입금요청 현황 행은 P열 → 입금완료_날짜, 본 시트 Q열 → 비움.
 */
import path from "path";
import { fileURLToPath } from "url";
import config from "../config/config.js";
import { fetchSheetValues, updateSheetValue, updateRowColor } from "../modules/google-sheet-module.js";
import { runDepositRequestList } from "./deposit-request-list.js";

const STATUS_COLUMN_INDEX = 16; // Q열 (0-based)
const P_COLUMN_INDEX = 15;      // P열 (0-based)
const STATUS_VALUE = "이체완료";

// 매칭용 열 인덱스 (이름, 계좌, 주민, 금액)
const COL_NAME = 5, COL_ACCOUNT = 8, COL_RESIDENT = 9, COL_AMOUNT = 10;

/** 금액 셀을 숫자만 붙인 문자열로 정규화 (쉼표 제거) */
function normalizeAmount(v) {
  return String(v || "").replace(/[^0-9]/g, "");
}

/** 같은 행인지 판단하기 위한 키 생성 */
function rowKey(row) {
  const name = (row[COL_NAME] || "").trim();
  const account = (row[COL_ACCOUNT] || "").trim();
  const resident = (row[COL_RESIDENT] || "").trim();
  const amount = normalizeAmount(row[COL_AMOUNT]);
  return `${name}|${account}|${resident}|${amount}`;
}

function getDepositCompleteText() {
  const today = new Date();
  const y = String(today.getFullYear()).slice(-2);
  const m = String(today.getMonth() + 1).padStart(2, "0");
  const d = String(today.getDate()).padStart(2, "0");
  return `입금완료_${y}${m}${d}`;
}

/**
 * Q열이 "이체완료"인 행만 필터링하여 반환
 */
function filterTransferCompleteRows(rows) {
  if (!rows || rows.length === 0) return [];
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const status = (row[STATUS_COLUMN_INDEX] || "").trim();
    if (status === STATUS_VALUE) {
      result.push({ rowIndex: i + 1, row }); // 시트 행 번호는 1부터
    }
  }
  return result;
}

/**
 * 이체완료 목록 조회 및 출력.
 * @returns {Promise<{ sheetUrl: string, sheetName: string, completed: Array<{ rowIndex: number, row: string[] }> }>}
 */
export async function runTransferCompleteList() {
  const sheet = config.sheets.find(
    (s) => s.url && s.url.includes("1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c")
  ) || config.sheets[0];

  const empty = { sheetUrl: "", sheetName: "", completed: [] };
  if (!sheet) {
    console.log("❌ 입금요청내역 시트 설정을 찾을 수 없습니다.");
    return empty;
  }

  const sheetName = sheet.sheetName || "시트1";
  const rangeName = `${sheetName}!A:Q`;

  try {
    console.log("\n📋 이체완료 업데이트 — 시트에서 Q열 '이체완료' 행 조회 중...\n");
    const rows = await fetchSheetValues({
      sheetUrl: sheet.url,
      sheetName,
      range: rangeName,
    });

    const completed = filterTransferCompleteRows(rows);
    if (completed.length === 0) {
      console.log("Q열이 '이체완료'인 행이 없습니다.");
      return { sheetUrl: sheet.url, sheetName, completed: [] };
    }

    console.log(`Q열 '이체완료' 행 수: ${completed.length}개\n`);
    const header = rows[0] || [];
    completed.forEach(({ rowIndex, row }) => {
      console.log(`--- 행 ${rowIndex} ---`);
      header.forEach((colName, i) => {
        if (colName != null && colName !== "" && row[i] != null && row[i] !== "") {
          console.log(`  ${colName}: ${row[i]}`);
        }
      });
      console.log(`  (전체 행): ${JSON.stringify(row)}`);
      console.log("");
    });
    return { sheetUrl: sheet.url, sheetName, completed };
  } catch (error) {
    console.error(`이체완료 목록 조회 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 이체완료 목록과 입금요청 현황을 매칭해, 같은 행이면
 * - 입금요청 현황: 해당 시트 해당 행 P열 → '입금완료_날짜'
 * - 실시간 입금요청내역: 해당 행 Q열 → '' (이체완료 제거)
 */
async function applyMatchedUpdates(completedData, depositData) {
  if (!completedData.sheetUrl || completedData.completed.length === 0 || !depositData.items.length) {
    return;
  }

  const depositByKey = new Map();
  for (const it of depositData.items) {
    const key = rowKey(it.row);
    if (!depositByKey.has(key)) depositByKey.set(key, it);
  }

  const dateText = getDepositCompleteText();
  let updateCount = 0;

  for (const { rowIndex, row } of completedData.completed) {
    const key = rowKey(row);
    const deposit = depositByKey.get(key);
    if (!deposit) continue;

    try {
      // 입금요청 현황: P열 → 입금완료_날짜
      await updateSheetValue({
        sheetUrl: depositData.sheetUrl,
        sheetName: deposit.sheetName,
        rowIndex: deposit.rowIndex - 1,
        columnIndex: P_COLUMN_INDEX,
        value: dateText,
      });
      // 입금요청 현황: 해당 행 B~P열 배경색 회색
      await updateRowColor({
        sheetUrl: depositData.sheetUrl,
        sheetName: deposit.sheetName,
        rowIndex: deposit.rowIndex,
        colorName: "진한 회색 1",
      });
      // 실시간 입금요청내역: Q열 비우기 (이체완료 제거)
      await updateSheetValue({
        sheetUrl: completedData.sheetUrl,
        sheetName: completedData.sheetName,
        rowIndex: rowIndex - 1,
        columnIndex: STATUS_COLUMN_INDEX,
        value: "",
      });
      updateCount++;
      console.log(`  매칭 반영: ${(row[COL_NAME] || "").trim()} — [${deposit.sheetName}] 행 ${deposit.rowIndex} P열 → ${dateText}, 입금요청내역 Q열 비움`);
    } catch (e) {
      console.error(`  매칭 반영 실패 (${(row[COL_NAME] || "").trim()}): ${e.message}`);
    }
  }

  if (updateCount > 0) {
    console.log(`\n✅ 매칭 반영 완료: ${updateCount}건`);
  }
}

/** 이체완료 목록 + 입금요청 목록 출력 후, 같은 행이면 입금요청 현황 P열·입금요청내역 Q열 반영 */
async function runAll() {
  const completedData = await runTransferCompleteList();
  const depositData = await runDepositRequestList();
  await applyMatchedUpdates(completedData, depositData);
}

// 이 파일을 직접 실행할 때 (node scripts/transfer-complete-list.js)
const __filename = fileURLToPath(import.meta.url);
if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(__filename)) {
  runAll().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
