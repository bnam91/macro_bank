// 이체 관련 액션 함수들

import { bankOptions } from '../../config/bank-config.js';
import { switchToFrame, focusAndTypeWithKeyboard } from './frame-utils.js';

// 이체 메뉴 클릭
export async function clickTransferMenu(page) {
  try {
    const frame = await switchToFrame(page, "hanaMainframe");
    
    // "이체" 메뉴 링크 찾기 및 클릭
    let transferMenuClicked = false;
    
    try {
      // 방법 1: CSS 선택자로 찾기
      await frame.waitForSelector("a[title='이체']", { timeout: 10000 });
      const element = await frame.$("a[title='이체']");
      if (element) {
        // 마우스 이벤트로 호버 효과 발생 (서브메뉴가 나타나도록)
        await element.hover();
        await page.waitForTimeout(300);
        await element.click();
        console.log("'이체' 메뉴 링크를 클릭했습니다.");
        transferMenuClicked = true;
      }
    } catch (cssError) {
      // 방법 2: evaluate를 사용하여 JavaScript로 찾고 클릭
      const clicked = await frame.evaluate(() => {
        const links = Array.from(document.querySelectorAll("a[title='이체']"));
        const transferLink = links.find(link => link.textContent.trim() === '이체');
        if (transferLink) {
          // 호버 이벤트 발생
          transferLink.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
          transferLink.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
          transferLink.click();
          return true;
        }
        return false;
      });
      
      if (clicked) {
        console.log("'이체' 메뉴 링크를 클릭했습니다. (evaluate 사용)");
        transferMenuClicked = true;
      }
    }
    
    if (!transferMenuClicked) {
      console.log("'이체' 메뉴를 찾을 수 없습니다.");
      return false;
    }
    
    // 서브메뉴가 나타날 때까지 대기
    await page.waitForTimeout(500);
    
    // 서브메뉴가 표시되는지 확인 (submenu의 display가 none이 아닌지 확인)
    const submenuVisible = await frame.evaluate(() => {
      const transferMenu = document.querySelector("a[title='이체']");
      if (!transferMenu) return false;
      
      // 부모 li 요소 찾기
      let parentLi = transferMenu.closest('li');
      if (!parentLi) return false;
      
      // 서브메뉴 찾기
      const submenu = parentLi.querySelector('.submenu');
      if (!submenu) return false;
      
      // display 스타일 확인
      const style = window.getComputedStyle(submenu);
      return style.display !== 'none';
    });
    
    if (submenuVisible) {
      console.log("서브메뉴가 표시되었습니다.");
    } else {
      console.log("서브메뉴가 아직 표시되지 않았습니다. 강제로 표시 시도...");
      // 서브메뉴를 강제로 표시
      await frame.evaluate(() => {
        const transferMenu = document.querySelector("a[title='이체']");
        if (transferMenu) {
          let parentLi = transferMenu.closest('li');
          if (parentLi) {
            const submenu = parentLi.querySelector('.submenu');
            if (submenu) {
              submenu.style.display = 'block';
            }
          }
        }
      });
      await page.waitForTimeout(500);
    }
    
    return true;
  } catch (error) {
    console.log(`'이체' 메뉴 링크 클릭 중 오류 발생: ${error.message}`);
    return false;
  }
}

// 다계좌 이체 버튼 클릭
export async function clickMultiTransferButton(page) {
  try {
    const frame = await switchToFrame(page, "hanaMainframe");
    
    // "다계좌이체" 링크 찾기 및 클릭
    let multiTransferClicked = false;
    
    try {
      // 방법 1: title 속성으로 찾기
      await frame.waitForSelector("a[title='다계좌이체']", { timeout: 10000 });
      const element = await frame.$("a[title='다계좌이체']");
      if (element) {
        // 요소가 보이는지 확인
        const isVisible = await frame.evaluate((el) => {
          return el && el.offsetParent !== null;
        }, element);
        
        if (isVisible) {
          await element.click();
          console.log("'다계좌이체' 링크를 클릭했습니다.");
          multiTransferClicked = true;
        } else {
          console.log("'다계좌이체' 링크를 찾았지만 보이지 않습니다.");
        }
      }
    } catch (cssError) {
      console.log(`CSS 선택자로 찾기 실패: ${cssError.message}`);
    }
    
    // 방법 2: onclick 속성으로 찾기 (다계좌이체의 고유한 onclick)
    if (!multiTransferClicked) {
      try {
        const clicked = await frame.evaluate(() => {
          const links = Array.from(document.querySelectorAll("a[onclick*='wpdep416_01t_01']"));
          const multiTransferLink = links.find(link => 
            link.getAttribute('onclick') && 
            link.getAttribute('onclick').includes('wpdep416_01t_01') &&
            (link.textContent.trim() === '다계좌이체' || link.getAttribute('title') === '다계좌이체')
          );
          if (multiTransferLink) {
            multiTransferLink.click();
            return true;
          }
          return false;
        });
        
        if (clicked) {
          console.log("'다계좌이체' 링크를 클릭했습니다. (onclick 속성으로 찾기)");
          multiTransferClicked = true;
        }
      } catch (evalError) {
        console.log(`evaluate로 찾기 실패: ${evalError.message}`);
      }
    }
    
    // 방법 3: 텍스트 내용으로 찾기
    if (!multiTransferClicked) {
      try {
        const clicked = await frame.evaluate(() => {
          const links = Array.from(document.querySelectorAll("a"));
          const multiTransferLink = links.find(link => 
            link.textContent.trim() === '다계좌이체' &&
            link.getAttribute('onclick') && 
            link.getAttribute('onclick').includes('wpdep416_01t_01')
          );
          if (multiTransferLink) {
            // 부모 요소가 보이는지 확인
            let parent = multiTransferLink;
            while (parent && parent !== document.body) {
              const style = window.getComputedStyle(parent);
              if (style.display === 'none' || style.visibility === 'hidden') {
                // 부모의 display를 block으로 변경
                parent.style.display = 'block';
                parent.style.visibility = 'visible';
              }
              parent = parent.parentElement;
            }
            multiTransferLink.click();
            return true;
          }
          return false;
        });
        
        if (clicked) {
          console.log("'다계좌이체' 링크를 클릭했습니다. (텍스트 내용으로 찾기)");
          multiTransferClicked = true;
        }
      } catch (evalError) {
        console.log(`텍스트로 찾기 실패: ${evalError.message}`);
      }
    }
    
    if (!multiTransferClicked) {
      console.log("'다계좌이체' 링크를 찾을 수 없습니다.");
      return false;
    }
    
    await page.waitForTimeout(3000);
    console.log("다계좌 이체 버튼 클릭 완료");
    return true;
  } catch (error) {
    console.log(`다계좌 이체 버튼 클릭 중 오류: ${error.message}`);
    return false;
  }
}

// 스크롤 조정
export async function adjustScroll(page, amount = 1000) {
  try {
    await page.evaluate((scrollAmount) => {
      window.scrollBy(0, scrollAmount);
    }, -amount);
    await page.waitForTimeout(1000);
    await page.evaluate((scrollAmount) => {
      window.scrollBy(0, scrollAmount / 2);
    }, amount);
  } catch (error) {
    console.log(`스크롤 조정 오류: ${error.message}`);
  }
}

// 이체 정보 입력
export async function inputTransferInfo(page, data, index) {
  const { bank, accountNumber, nameProduct, productName, amount } = data;

  try {
    const frame = await switchToFrame(page, "hanaMainframe");

    // 은행 선택
    const bankElementId = `#rcvBnkCd${index}`;
    const optionValue = bankOptions[bank] || "";
    if (optionValue) {
      await frame.waitForSelector(bankElementId, { timeout: 5000 });
      await frame.select(bankElementId, optionValue);
      console.log(`${nameProduct} 은행 선택 성공: ${bank}`);
    } else {
      console.log(`${nameProduct} 은행을 찾을 수 없습니다: ${bank}`);
    }

    // 계좌번호 입력
    const accountElementId = `#rcvAcctNo${index}`;
    await focusAndTypeWithKeyboard(frame, page, accountElementId, accountNumber, { waitBeforeType: 1000 });
    console.log(`${nameProduct} 계좌번호 입력 성공: ${accountNumber}`);

    // 금액 입력
    const amountElementId = `#trnsAmt${index}`;
    const sanitizedAmount = String(amount ?? '').replace(/[^\d]/g, '');
    await focusAndTypeWithKeyboard(frame, page, amountElementId, sanitizedAmount, { delay: 1 });
    console.log(`${nameProduct} 금액 입력 성공: ${amount}`);

    // 이름.제품명 입력
    const nameProductElementId = `#wdrwPsbkMarkCtt${index}`;
    await frame.evaluate((id, value) => {
      const element = document.querySelector(id);
      if (element) {
        element.value = value;
        element.dispatchEvent(new Event('input', { bubbles: true }));
      }
    }, nameProductElementId, nameProduct);
    console.log(`${nameProduct} 이름.제품명 입력 성공: ${nameProduct}`);

    // 제품명 입력
    const productNameElementId = `#rcvPsbkMarkCtt${index}`;
    await frame.evaluate((id, value) => {
      const element = document.querySelector(id);
      if (element) {
        element.value = value;
        element.dispatchEvent(new Event('input', { bubbles: true }));
      }
    }, productNameElementId, productName);
    console.log(`${nameProduct} 제품명 입력 성공: ${productName}`);

    return true;
  } catch (error) {
    console.log(`오류 발생 ${nameProduct}: ${error.message}`);
    return false;
  }
}

// 비밀번호 입력
export async function enterPassword(page) {
  try {
    const frame = await switchToFrame(page, "hanaMainframe");
    
    await frame.evaluate(() => {
      window.scrollTo(0, 0);
    });
    await page.waitForTimeout(500);

    const passwordElementId = "#paymAcctPw";
    await focusAndTypeWithKeyboard(frame, page, passwordElementId, "3724");
    console.log("비밀번호 입력 성공");
    return true;
  } catch (error) {
    console.log(`비밀번호 입력 오류: ${error.message}`);
    return false;
  }
}

// 다계좌이체진행 버튼 클릭
export async function clickTransferButton(page) {
  try {
    await page.waitForTimeout(2000);
    
    const frame = await switchToFrame(page, "hanaMainframe");
    
    await frame.evaluate(() => {
      window.scrollBy(0, 500);
    });
    await page.waitForTimeout(500);

    await frame.waitForXPath("//a[contains(text(), '다계좌이체진행')]", { timeout: 10000 });
    const [element] = await frame.$x("//a[contains(text(), '다계좌이체진행')]");
    if (element) {
      await element.click();
      console.log("다계좌이체진행 버튼 클릭 성공");
      return true;
    }
    return false;
  } catch (error) {
    console.log(`다계좌이체진행 버튼 클릭 오류: ${error.message}`);
    return false;
  }
}
