import { handleCertPopup } from './cert-popup/index.js';

// 공동/금융인증서 로그인 메뉴 클릭
async function clickCertMenu(page) {
  try {
    await page.waitForXPath("//a[contains(text(), '공동/금융인증서 로그인')]", { timeout: 30000 });
    const [element] = await page.$x("//a[contains(text(), '공동/금융인증서 로그인')]");
    if (element) {
      await element.click();
      console.log("공동/금융인증서 로그인 메뉴 클릭 성공");
      await page.waitForTimeout(3000);
      return true;
    }
    return false;
  } catch (error) {
    console.log(`공동/금융인증서 로그인 메뉴 클릭 실패: ${error.message}`);
    return false;
  }
}

// 공동인증서 로그인 버튼 클릭
async function clickCertLoginButton(page) {
  try {
    await page.waitForSelector("#certLogin", { timeout: 30000 });
    
    let clicked = false;
    for (let attempt = 0; attempt < 60; attempt++) {
      try {
        await page.evaluate(() => {
          const element = document.querySelector("#certLogin");
          if (element) element.click();
        });
        console.log("공동인증서 로그인 버튼 클릭 성공");
        clicked = true;
        break;
      } catch (error) {
        if (attempt < 59) {
          console.log(`클릭 실패, 1초 후 재시도합니다. (${attempt + 1}/60)`);
          await page.waitForTimeout(1000);
        } else {
          throw error;
        }
      }
    }
    return clicked;
  } catch (error) {
    console.log(`certLogin 버튼을 찾을 수 없습니다: ${error.message}`);
    return false;
  }
}

// 공인인증서 팝업 처리는 cert-popup 모듈에서 수행

// 공인인증서 로그인 전체 프로세스
async function executeCertLogin(page) {
  try {
    console.log("\n=== 공인인증서 로그인 프로세스 시작 ===");
    
    // 1. 공동/금융인증서 로그인 메뉴 클릭
    const menuClicked = await clickCertMenu(page);
    if (!menuClicked) {
      console.log("공동/금융인증서 로그인 메뉴 클릭 실패");
      return false;
    }
    
    // 2. 공동인증서 로그인 버튼 클릭
    const loginClicked = await clickCertLoginButton(page);
    if (!loginClicked) {
      console.log("로그인 버튼 클릭 실패");
      return false;
    }

    // 3. 공인인증서 팝업 처리 (로컬디스크 선택, 확장매체 선택, 비밀번호 입력, 확인 버튼 클릭)
    const popupHandled = await handleCertPopup(page);
    if (!popupHandled) {
      console.log("공인인증서 팝업 처리 실패. 수동으로 진행해주세요.");
      await page.waitForTimeout(10000);
      return false;
    } else {
      console.log("공인인증서 로그인 완료");
      await page.waitForTimeout(3000); // 로그인 완료 후 대기
    }
    
    console.log("\n=== 공인인증서 로그인 프로세스 완료 ===");
    return true;
  } catch (error) {
    console.error(`공인인증서 로그인 프로세스 실행 중 오류: ${error.message}`);
    return false;
  }
}

export {
  clickCertMenu,
  clickCertLoginButton,
  handleCertPopup,
  executeCertLogin
};

