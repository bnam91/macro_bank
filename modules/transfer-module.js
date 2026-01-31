import { handleUserInput, waitForEnter } from './user-input-module.js';
import { executeCertLogin } from './cert-module.js';
import { loadSheetTransferData, updateSheetValue } from "./google-sheet-module.js";
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

    // 8. 구글 시트 데이터 로드
    const processedData = await loadSheetTransferData(sheetConfig);
    console.log("\n전처리된 데이터:");
    processedData.forEach((data, idx) => {
      console.log(`${idx + 1}. 은행: ${data.bank}, 계좌번호: ${data.accountNumber}, 이름.제품명: ${data.nameProduct}, 제품명: ${data.productName}, 금액: ${data.amount}`);
    });

    if (processedData.length === 0) {
      console.log("가져온 이체 데이터가 없습니다. 시트 내용을 확인해주세요.");
      return false;
    }

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
    
    for (const data of processedData) {
      if (data.rowIndex !== undefined) {
        try {
          // rowIndex는 헤더를 제외한 인덱스이므로, 시트에서는 +2 (헤더 + 1부터 시작)
          await updateSheetValue({
            sheetUrl: sheetConfig.sheetUrl,
            sheetName: sheetConfig.sheetName,
            authModulePath: sheetConfig.authModulePath,
            rowIndex: data.rowIndex + 1, // 헤더 다음 행부터 시작하므로 +1
            columnIndex: STATUS_COLUMN_INDEX,
            value: '이체완료'
          });
          console.log(`  ✅ 행 ${data.rowIndex + 2}: ${data.nameProduct} - '이체완료' 기록 완료`);
        } catch (error) {
          console.error(`  ❌ 행 ${data.rowIndex + 2}: ${data.nameProduct} - 상태 기록 실패: ${error.message}`);
        }
      }
    }
    
    console.log("\n✅ 모든 이체 완료 상태 기록이 완료되었습니다.");

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
