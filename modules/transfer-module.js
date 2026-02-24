import { handleUserInput, waitForEnter, selectNextAction } from './user-input-module.js';
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
    
    // 시트 전체 데이터 읽기 (매칭을 위해)
    console.log("📊 시트 전체 데이터를 읽어서 매칭 중...");
    const allSheetRows = await fetchSheetValues({
      sheetUrl: sheetConfig.sheetUrl,
      sheetName: sheetConfig.sheetName,
      authModulePath: sheetConfig.authModulePath
    });
    
    for (const data of processedData) {
      try {
        // 입력한 데이터 정보 출력
        console.log(`\n  📋 검수 대상: ${data.nameProduct}`);
        console.log(`     은행: ${data.bank}`);
        console.log(`     계좌번호: ${data.accountNumber}`);
        console.log(`     금액: ${data.amount.toLocaleString()}`);
        console.log(`     이름.제품명: ${data.nameProduct}`);
        console.log(`     제품명: ${data.productName || '(없음)'}`);
        
        // 시트에서 일치하는 행 찾기 (제품명, 이름, 계좌번호, 금액으로 매칭)
        let matchedRowIndex = -1;
        const inputCustomerName = String(data.customerName || '').trim();
        const inputProductName = String(data.productName || '').trim();
        const inputAccountNumber = (data.accountNumber || '').replace(/[^0-9]/g, '');
        const inputAmountNum = parseFloat(String(data.amount || '').replace(/[^0-9.]/g, '')) || 0;
        
        // 디버깅: 입력 데이터 확인
        console.log(`     🔍 매칭 시도 - 이름: "${inputCustomerName}", 제품: "${inputProductName}", 계좌: "${inputAccountNumber}", 금액: ${inputAmountNum.toLocaleString()}`);
        
        for (let i = 1; i < allSheetRows.length; i++) { // 헤더 제외 (i=0)
          const row = allSheetRows[i];
          if (!Array.isArray(row) || row.length === 0) continue;
          
          // 최소 컬럼 수 확인
          const maxColumnIndex = Math.max(
            columnMapping.customerName || 0,
            columnMapping.productName || 0,
            columnMapping.accountInfo || 0,
            columnMapping.amount || 0
          );
          
          if (row.length <= maxColumnIndex) continue;
          
          const sheetCustomerName = (row[columnMapping.customerName] != null ? String(row[columnMapping.customerName]) : '').trim();
          const sheetProductName = (row[columnMapping.productName] != null ? String(row[columnMapping.productName]) : '').trim();
          const sheetAccountInfo = (row[columnMapping.accountInfo] != null ? String(row[columnMapping.accountInfo]) : '').trim();
          const sheetAmount = row[columnMapping.amount] != null ? row[columnMapping.amount] : '';
          
          // 계좌번호에서 숫자만 추출하여 비교
          const sheetAccountNumber = (sheetAccountInfo || '').replace(/[^0-9]/g, '');
          
          // 금액 비교 (쉼표 제거)
          const sheetAmountNum = parseFloat(String(sheetAmount || '').replace(/[^0-9.]/g, '')) || 0;
          
          // 매칭 확인 (모든 조건이 일치해야 함)
          const nameMatch = (sheetCustomerName && inputCustomerName) && 
            (sheetCustomerName.includes(inputCustomerName) || inputCustomerName.includes(sheetCustomerName));
          const productMatch = !inputProductName || !sheetProductName || 
            (sheetProductName.includes(inputProductName) || inputProductName.includes(sheetProductName));
          const accountMatch = (sheetAccountNumber && inputAccountNumber) && 
            (sheetAccountNumber.includes(inputAccountNumber) || inputAccountNumber.includes(sheetAccountNumber));
          const amountMatch = Math.abs(sheetAmountNum - inputAmountNum) < 1; // 1원 이하 차이는 허용
          
          // 디버깅: 매칭 결과 확인
          if (i <= 5) { // 처음 5개 행만 디버깅 출력
            console.log(`     [행 ${i + 1}] 이름: "${sheetCustomerName}" ${nameMatch ? '✅' : '❌'} | 제품: "${sheetProductName}" ${productMatch ? '✅' : '❌'} | 계좌: ${accountMatch ? '✅' : '❌'} | 금액: ${amountMatch ? '✅' : '❌'}`);
          }
          
          if (nameMatch && productMatch && accountMatch && amountMatch) {
            matchedRowIndex = i; // 시트 행 인덱스 (헤더 포함, 0부터 시작)
            const sheetRowNumber = i + 1; // 실제 시트 행 번호 (1부터 시작)
            
            console.log(`  ✅ 매칭된 행 발견: 행 ${sheetRowNumber} (Q${sheetRowNumber})`);
            console.log(`     이름: "${sheetCustomerName}" ✅`);
            console.log(`     제품: "${sheetProductName}" ✅`);
            console.log(`     계좌: "${sheetAccountInfo}" ✅`);
            console.log(`     금액: ${sheetAmountNum.toLocaleString()} ✅`);
            break;
          }
        }
        
        if (matchedRowIndex === -1) {
          console.error(`  ❌ 일치하는 행을 찾을 수 없습니다!`);
          console.error(`     이름: "${inputCustomerName}", 제품: "${inputProductName}", 계좌: "${inputAccountNumber}", 금액: ${inputAmountNum.toLocaleString()}`);
          continue;
        }
        
        // 이체완료 기록
        // updateSheetValue는 rowIndex에 +1을 하므로, matchedRowIndex를 그대로 전달하면 됨
        // matchedRowIndex는 헤더 포함 인덱스 (0부터 시작), 시트 행 번호는 matchedRowIndex + 1
        await updateSheetValue({
          sheetUrl: sheetConfig.sheetUrl,
          sheetName: sheetConfig.sheetName,
          authModulePath: sheetConfig.authModulePath,
          rowIndex: matchedRowIndex, // 헤더 포함 인덱스 (0부터 시작)
          columnIndex: STATUS_COLUMN_INDEX,
          value: '이체완료'
        });
        
        // 기록 후 검수: 실제로 올바른 셀에 기록되었는지 확인
        await new Promise(resolve => setTimeout(resolve, 500)); // 잠시 대기
        const recordedValue = await getSheetValue({
          sheetUrl: sheetConfig.sheetUrl,
          sheetName: sheetConfig.sheetName,
          authModulePath: sheetConfig.authModulePath,
          rowIndex: matchedRowIndex, // getSheetValue도 rowIndex에 +1을 하므로 동일하게 전달
          columnIndex: STATUS_COLUMN_INDEX
        });
        
        const sheetRowNumber = matchedRowIndex + 1;
        if (recordedValue === '이체완료') {
          console.log(`  ✅ 행 ${sheetRowNumber} (Q${sheetRowNumber}): '이체완료' 기록 완료`);
        } else {
          console.error(`  ❌ 검수 실패: 행 ${sheetRowNumber} (Q${sheetRowNumber})에 '이체완료'가 기록되지 않았습니다. (실제 값: "${recordedValue}")`);
        }
      } catch (error) {
        console.error(`  ❌ ${data.nameProduct} - 상태 기록 실패: ${error.message}`);
      }
    }
    
    console.log("\n✅ 모든 이체 완료 상태 기록 및 검수가 완료되었습니다.");

    // 13. 다음 작업 선택
    console.log("\n");
    const selectedAction = await selectNextAction();
    
    // 선택한 작업에 따른 처리 (기능은 나중에 구현 예정)
    if (selectedAction === 1) {
      console.log("다음 이체 진행하기 기능은 추후 구현 예정입니다.");
    } else if (selectedAction === 2) {
      console.log("이체완료 업데이트 기능은 추후 구현 예정입니다.");
    }

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
