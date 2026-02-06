import { handleUserInput, waitForEnter } from './user-input-module.js';
import { executeCertLogin } from './cert-module.js';
import { loadSheetTransferData, updateSheetValue, getSheetValue, fetchSheetValues } from "./google-sheet-module.js";
import { switchToFrame } from './utils/frame-utils.js';
import { 
  clickTransferMenu, 
  clickMultiTransferButton, 
  adjustScroll, 
  inputTransferInfo, 
  enterPassword, 
  clickTransferButton 
} from './utils/transfer-actions.js';
import { 
  checkAndCloseMainPagePopup, 
  checkAndHandleDevicePopup, 
  handleVoicePhishingPopup 
} from './utils/popup-handlers.js';

// 메인 이체 프로세스 실행
async function executeTransferProcess(page, sheetConfig, autoTransfer = false) {
  try {
    console.log("\n=== 이체 프로세스 시작 ===");

    // 0. 구글 시트 데이터 먼저 로드 및 확인 (로그인 전에 확인)
    console.log("\n📋 시트 데이터 확인 중...");
    const processedData = await loadSheetTransferData(sheetConfig);
    console.log("\n전처리된 데이터:");
    if (processedData.length === 0) {
      console.log("❌ 가져온 이체 데이터가 없습니다. 시트 내용을 확인해주세요.");
      console.log("   이체 프로세스를 시작하지 않습니다.");
      return false;
    }
    
    processedData.forEach((data, idx) => {
      console.log(`${idx + 1}. 은행: ${data.bank}, 계좌번호: ${data.accountNumber}, 이름.제품명: ${data.nameProduct}, 제품명: ${data.productName}, 금액: ${data.amount}`);
    });
    console.log(`\n✅ 총 ${processedData.length}개의 이체 데이터를 확인했습니다.`);
    console.log("   로그인 후 이체 프로세스를 진행합니다.\n");

    // 1. 공인인증서 로그인 프로세스 (cert-module 사용)
    const certLoginSuccess = await executeCertLogin(page);
    if (!certLoginSuccess) {
      console.log("공인인증서 로그인 실패. 수동으로 진행해주세요.");
      await page.waitForTimeout(10000);
      return false;
    }

    // 2. 메인 페이지 팝업 확인 및 처리 (최대 10초)
    await checkAndCloseMainPagePopup(page);

    // 3. 프레임 전환 및 이체 메뉴 클릭
    const transferMenuClicked = await clickTransferMenu(page);
    if (!transferMenuClicked) {
      console.log("이체 메뉴 클릭 실패");
      return false;
    }

    // 4. 다계좌 이체 버튼 클릭 전 사용자 확인 (현재 미사용)
    // const frame = await switchToFrame(page, "hanaMainframe");
    // const shouldContinue = await handleUserInput(
    //   frame,
    //   "이체 메뉴가 클릭되었습니다. 다계좌이체 버튼을 클릭하고 다음 프로세스를 진행할까요? (y/d/n): ",
    //   page
    // );
    // if (!shouldContinue) {
    //   console.log("사용자가 다계좌이체 진행을 중단했습니다.");
    //   return false;
    // }

    // 5. 다계좌 이체 버튼 클릭
    const multiTransferClicked = await clickMultiTransferButton(page);
    if (!multiTransferClicked) {
      console.log("다계좌이체 버튼 클릭 실패");
      return false;
    }

    // 6. 단말기 미지정 PC이용 안내 팝업 확인 및 처리
    const devicePopupDetected = await checkAndHandleDevicePopup(page);
    if (devicePopupDetected) {
      console.log("\n단말기 미지정 PC이용 안내 팝업이 감지되어 닫았습니다.");
      
      // 사용자에게 계속 진행할지 묻기
      const frame = await switchToFrame(page, "hanaMainframe");
      const shouldContinue = await handleUserInput(
        frame,
        "\n팝업을 닫았습니다. 계속 진행할까요? (y/d/n): ",
        page
      );
      
      if (!shouldContinue) {
        console.log("사용자가 진행을 중단했습니다.");
        return false;
      }
    }

    // 7. 스크롤 조정
    await adjustScroll(page);

    // 9. 이체 정보 입력 (최대 10개)
    for (let i = 0; i < Math.min(processedData.length, 10); i++) {
      await inputTransferInfo(page, processedData[i], i);
      await page.waitForTimeout(500);
    }

    // 10. 비밀번호 입력
    await enterPassword(page);

    // 11. 자동 이체 진행 여부에 따라 처리
    if (autoTransfer) {
      await clickTransferButton(page);
      await handleVoicePhishingPopup(page);
      console.log("✅ 이체가 완료되었습니다.");
    } else {
      console.log("\n✅ 모든 입력이 완료되었습니다.");
      console.log("📌 브라우저에서 '다계좌이체진행' 버튼을 수동으로 클릭해주세요.");
      console.log("   (자동 이체 진행이 비활성화되어 있습니다.)");
    }

    // 12. 이체 완료 후 시트에 상태 업데이트
    console.log("\n📝 이체 완료 처리를 위해 엔터를 눌러주세요...");
    await waitForEnter("이체가 완료되었으면 엔터를 눌러주세요: ");
    
    console.log("\n🔄 시트에 '이체완료' 상태를 기록하는 중...");
    const STATUS_COLUMN_INDEX = 16; // Q열 (인덱스 16)
    const columnMapping = sheetConfig.columnMapping || {
      productName: 4,    // E열: 제품
      customerName: 5,   // F열: 이름
      accountInfo: 8,    // I열: 계좌번호
      amount: 10         // K열: 금액
    };
    
    for (const data of processedData) {
      if (data.rowIndex !== undefined) {
        try {
          // rowIndex는 헤더를 제외한 인덱스 (0부터 시작)
          // 시트에서는 헤더가 1행이므로, 데이터는 2행부터 시작
          // 따라서 시트 행 번호는 rowIndex + 2가 되어야 함
          const expectedSheetRow = data.rowIndex + 2; // 예상 시트 행 번호
          const actualRowIndex = data.rowIndex + 1; // updateSheetValue에 전달할 인덱스
          
          // 입력한 데이터 정보 출력 (이체 정보 입력 시와 동일한 형식)
          console.log(`\n  📋 검수 대상: ${data.nameProduct}`);
          console.log(`     은행: ${data.bank}`);
          console.log(`     계좌번호: ${data.accountNumber}`);
          console.log(`     금액: ${data.amount.toLocaleString()}`);
          console.log(`     이름.제품명: ${data.nameProduct}`);
          console.log(`     제품명: ${data.productName || '(없음)'}`);
          console.log(`     예상 행: ${expectedSheetRow} (Q${expectedSheetRow})`);
          
          // 기록 전 검수: 시트의 해당 행에서 여러 컬럼 확인
          const sheetRow = await fetchSheetValues({
            sheetUrl: sheetConfig.sheetUrl,
            sheetName: sheetConfig.sheetName,
            authModulePath: sheetConfig.authModulePath,
            range: `${sheetConfig.sheetName}!A${actualRowIndex + 1}:Q${actualRowIndex + 1}`
          });
          
          if (sheetRow.length > 0 && sheetRow[0].length > 0) {
            const row = sheetRow[0];
            const sheetCustomerName = (row[columnMapping.customerName] || '').toString().trim();
            const sheetProductName = (row[columnMapping.productName] || '').toString().trim();
            const sheetAccountInfo = (row[columnMapping.accountInfo] || '').toString().trim();
            const sheetAmount = row[columnMapping.amount] || '';
            
            // 계좌번호에서 숫자만 추출하여 비교
            const sheetAccountNumber = sheetAccountInfo.replace(/[^0-9]/g, '');
            const inputAccountNumber = data.accountNumber.replace(/[^0-9]/g, '');
            
            // 금액 비교 (쉼표 제거)
            const sheetAmountNum = parseFloat(String(sheetAmount).replace(/[^0-9.]/g, '')) || 0;
            const inputAmountNum = parseFloat(String(data.amount).replace(/[^0-9.]/g, '')) || 0;
            
            // 검수 결과
            const nameMatch = sheetCustomerName.includes(data.customerName) || data.customerName.includes(sheetCustomerName);
            const productMatch = !data.productName || sheetProductName.includes(data.productName) || data.productName.includes(sheetProductName);
            const accountMatch = sheetAccountNumber.includes(inputAccountNumber) || inputAccountNumber.includes(sheetAccountNumber);
            const amountMatch = Math.abs(sheetAmountNum - inputAmountNum) < 1; // 1원 이하 차이는 허용
            
            console.log(`     시트 데이터:`);
            console.log(`       이름: "${sheetCustomerName}" ${nameMatch ? '✅' : '❌'}`);
            console.log(`       제품: "${sheetProductName}" ${productMatch ? '✅' : '❌'}`);
            console.log(`       계좌: "${sheetAccountInfo}" ${accountMatch ? '✅' : '❌'}`);
            console.log(`       금액: ${sheetAmountNum.toLocaleString()} ${amountMatch ? '✅' : '❌'}`);
            
            const isValidRow = nameMatch && productMatch && accountMatch && amountMatch;
            
            if (!isValidRow) {
              console.log(`  ⚠️ 경고: 행 ${expectedSheetRow}의 데이터가 입력한 정보와 일치하지 않습니다!`);
              console.log(`     입력한 행이 맞는지 확인해주세요.`);
            } else {
              console.log(`  ✅ 행 ${expectedSheetRow} 데이터 검수 통과`);
            }
          } else {
            console.log(`  ⚠️ 경고: 행 ${expectedSheetRow}의 데이터를 읽을 수 없습니다.`);
          }
          
          // 이체완료 기록
          await updateSheetValue({
            sheetUrl: sheetConfig.sheetUrl,
            sheetName: sheetConfig.sheetName,
            authModulePath: sheetConfig.authModulePath,
            rowIndex: actualRowIndex,
            columnIndex: STATUS_COLUMN_INDEX,
            value: '이체완료'
          });
          
          // 기록 후 검수: 실제로 올바른 셀에 기록되었는지 확인
          await new Promise(resolve => setTimeout(resolve, 500)); // 잠시 대기
          const recordedValue = await getSheetValue({
            sheetUrl: sheetConfig.sheetUrl,
            sheetName: sheetConfig.sheetName,
            authModulePath: sheetConfig.authModulePath,
            rowIndex: actualRowIndex,
            columnIndex: STATUS_COLUMN_INDEX
          });
          
          if (recordedValue === '이체완료') {
            console.log(`  ✅ 행 ${expectedSheetRow} (Q${expectedSheetRow}): '이체완료' 기록 완료`);
          } else {
            console.error(`  ❌ 검수 실패: 행 ${expectedSheetRow} (Q${expectedSheetRow})에 '이체완료'가 기록되지 않았습니다. (실제 값: "${recordedValue}")`);
          }
        } catch (error) {
          console.error(`  ❌ 행 ${data.rowIndex + 2}: ${data.nameProduct} - 상태 기록 실패: ${error.message}`);
        }
      }
    }
    
    console.log("\n✅ 모든 이체 완료 상태 기록 및 검수가 완료되었습니다.");

    // 13. 다음 이체 진행 여부 확인
    console.log("\n");
    await waitForEnter("다음 이체도 진행할까요? (기능구현예정) - 엔터를 눌러주세요: ");

    console.log("\n=== 이체 프로세스 완료 ===");
    return true;
  } catch (error) {
    console.error(`이체 프로세스 실행 중 오류: ${error.message}`);
    return false;
  }
}

export {
  executeTransferProcess,
  clickTransferMenu,
  clickMultiTransferButton,
  inputTransferInfo,
  enterPassword,
  clickTransferButton,
  handleVoicePhishingPopup
};
