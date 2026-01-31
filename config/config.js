import path from 'path';
import os from 'os';

// 경로 확장 함수
function expandPath(inputPath) {
  if (!inputPath) return inputPath;
  if (inputPath === '~') return os.homedir();
  if (inputPath.startsWith('~/')) {
    return path.join(os.homedir(), inputPath.slice(2));
  }
  return inputPath;
}

// 설정 객체
const config = {
  // 사용자 데이터 디렉토리 경로
  userDataPath: expandPath('~/Documents/github_cloud/user_data'),
  
  // Google 시트 목록
  sheets: [
    {
      name: '테스트(수정)_입금요청내역',
      url: 'https://docs.google.com/spreadsheets/d/1EKRG7IF9UA7tCRfgg_qykAVnK00vonCKe2LJLAwiUvI/edit?gid=0#gid=0',
      sheetName: '시트1',
      // 열 매핑 (인덱스 기준, 0부터 시작)
      columnMapping: {
        productName: 4,    // E열: 제품
        customerName: 5,   // F열: 이름
        accountInfo: 8,    // I열: 계좌번호
        amount: 10         // K열: 금액
      }
    },
    {
      name: '고야_입금요청내역',
      url: 'https://docs.google.com/spreadsheets/d/1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c/edit?gid=0#gid=0',
      sheetName: '시트1',
      // 열 매핑 (인덱스 기준, 0부터 시작)
      columnMapping: {
        productName: 4,    // E열: 제품
        customerName: 5,   // F열: 이름
        accountInfo: 8,    // I열: 계좌번호
        amount: 10         // K열: 금액
      }
    }
  ]
};

export default config;
