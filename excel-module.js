const XLSX = require('xlsx');
const { preprocessAccountInfo } = require('./bank-config');

/**
 * 엑셀 데이터를 읽고 전처리된 배열을 반환합니다.
 * @param {string} excelPath - 엑셀 파일 경로
 * @returns {Array<{bank: string, accountNumber: string, nameProduct: string, productName: string, amount: number}>}
 */
function loadExcelData(excelPath) {
  try {
    const workbook = XLSX.readFile(excelPath);
    const sheetName = workbook.SheetNames[0]; // 첫 번째 시트 사용
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const processedData = [];
    for (let i = 0; i < data.length; i++) {
      if (i === 0) continue; // 헤더 제외

      const row = data[i];
      if (row.length < 10) continue; // 데이터가 충분하지 않으면 스킵

      const productName = row[3] || ''; // 제품명
      const customerName = row[4] || ''; // 이름
      const accountInfo = row[7] || ''; // 은행+계좌번호
      const amount = row[9] || 0; // 금액

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
  } catch (error) {
    console.error(`엑셀 파일 읽기 오류: ${error.message}`);
    return [];
  }
}

module.exports = {
  loadExcelData,
};

