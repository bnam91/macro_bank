const { bankOptions } = require('./bank-config');
const { handleUserInput } = require('./user-input-module');
const { executeCertLogin } = require('./cert-module');
const { loadExcelData } = require('./excel-module');

// 프레임 전환
async function switchToFrame(page, frameId) {
  try {
    // 방법 1: page.frames()를 사용하여 프레임 목록에서 찾기 (frameset의 frame 요소에 접근)
    const frames = page.frames();
    for (const frame of frames) {
      if (frame.name() === frameId) {
        console.log(`프레임 전환 성공: ${frameId}`);
        return frame;
      }
    }
    
    // 방법 2: iframe 요소로 찾기 시도
    try {
      await page.waitForSelector(`iframe[name="${frameId}"]`, { timeout: 2000 });
    const frameElement = await page.$(`iframe[name="${frameId}"]`);
    if (frameElement) {
      const frame = await frameElement.contentFrame();
        if (frame) {
          console.log(`프레임 전환 성공 (iframe): ${frameId}`);
          return frame;
        }
      }
    } catch (iframeError) {
      // iframe을 찾지 못한 경우 무시 (frameset의 frame일 수 있음)
    }
    
    // 방법 3: frame 요소로 찾기 시도 (frameset의 frame)
    try {
      await page.waitForSelector(`frame[name="${frameId}"]`, { timeout: 2000 });
      // frame 요소는 직접 contentFrame()을 호출할 수 없으므로 frames()에서 찾은 것을 사용
      // 위에서 이미 찾았으므로 여기서는 로그만 출력
      console.log(`frame 요소 발견: ${frameId}, frames()에서 찾기 시도`);
    } catch (frameError) {
      // frame 요소를 찾지 못한 경우
    }
    
    // 프레임을 찾지 못한 경우 메인 페이지 반환
    console.log(`프레임을 찾지 못했습니다: ${frameId}, 메인 페이지 사용`);
    return page;
  } catch (error) {
    console.log(`프레임 전환 오류: ${error.message}`);
    return page;
  }
}

const modifierKey = process.platform === 'darwin' ? 'Meta' : 'Control';

async function focusAndTypeWithKeyboard(frame, page, selector, text, options = {}) {
  const { delay = 5, waitBeforeType = 0 } = options;
  await frame.waitForSelector(selector, { timeout: 5000 });
  const focused = await frame.evaluate((sel) => {
    const element = document.querySelector(sel);
    if (!element) return false;
    if (typeof element.focus === 'function') element.focus();
    if (typeof element.select === 'function') element.select();
    return true;
  }, selector);

  if (!focused) {
    throw new Error(`Element not found or focus failed: ${selector}`);
  }

  await page.keyboard.down(modifierKey);
  await page.keyboard.press('a');
  await page.keyboard.up(modifierKey);
  await page.keyboard.press('Backspace');

  const value = String(text ?? '');
  if (value.length > 0) {
    if (waitBeforeType > 0) {
      await page.waitForTimeout(waitBeforeType);
    }
    await page.keyboard.type(value, { delay });
  }
  await page.waitForTimeout(50);
}

// 이체 메뉴 클릭
async function clickTransferMenu(page) {
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
async function clickMultiTransferButton(page) {
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
async function adjustScroll(page, amount = 1000) {
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
async function inputTransferInfo(page, data, index) {
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
async function enterPassword(page) {
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
async function clickTransferButton(page) {
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

// 메인 페이지 팝업 확인 및 처리 (최대 10초)
async function checkAndCloseMainPagePopup(page) {
  try {
    console.log("메인 페이지 팝업 확인 중... (최대 10초)");
    
    const maxWaitTime = 10; // 최대 10초
    
    for (let i = 0; i < maxWaitTime; i++) {
      await page.waitForTimeout(1000); // 1초마다 확인
      
      // 팝업 컨테이너 선택자 (제공된 구조에 맞춤)
      const popupSelectors = [
        '#opbLayerMessage0',
        'div.pop_ty11',  // 제공된 팝업 구조
        '.popup',
        '.modal',
        '.dialog',
        '[role="dialog"]',
        '.layer_popup',
        '.popup_layer'
      ];
      
      // 확인/닫기 버튼 선택자 (제공된 구조에 맞춤)
      const buttonSelectors = [
        'a#opbLayerMessage0_OK',  // 확인 버튼
        'a#opbLayerMessage0_Close',  // 닫기 버튼
        '#opbLayerMessage0_OK',
        '#opbLayerMessage0_Close',
        '.pop_close a',  // 닫기 버튼 (클래스 기반)
        'a[id*="Close"]',  // ID에 Close가 포함된 버튼
        'a[id*="OK"]',  // ID에 OK가 포함된 버튼
        '.btn_ex01 a',  // 확인 버튼 영역
        '.close',
        '.btn-close',
        '.popup_close',
        '.layer_close'
      ];
      
      // 팝업 찾기
      let popupFound = false;
      const contexts = [page, ...page.frames()];
      for (const context of contexts) {
        for (const popupSelector of popupSelectors) {
          try {
            const popup = await context.$(popupSelector);
          if (popup) {
              const isVisible = await context.evaluate((el) => {
              if (!el) return false;
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, popup);
            
            if (isVisible) {
                console.log(`팝업 발견: ${popupSelector}`);
              popupFound = true;
              
              // 확인/닫기 버튼 찾기 및 클릭
              let closed = false;
              
              // 먼저 확인/닫기 버튼 시도 (팝업 내부에서 탐색)
              for (const buttonSelector of buttonSelectors) {
                try {
                    const button = await popup.$(buttonSelector);
                  if (button) {
                      const isButtonVisible = await context.evaluate((el) => {
                      if (!el) return false;
                      const style = window.getComputedStyle(el);
                      return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                    }, button);
                    
                    if (isButtonVisible) {
                      await button.click();
                      console.log(`팝업 버튼 클릭 성공: ${buttonSelector}`);
                      closed = true;
                      await page.waitForTimeout(500);
                      
                      // 팝업이 사라졌는지 확인
                      const stillVisible = await context.evaluate((el) => {
                        if (!el) return false;
                        const style = window.getComputedStyle(el);
                        return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                      }, popup);
                      
                      if (!stillVisible) {
                        console.log("팝업이 닫혔습니다.");
                        return true;
                      }
                      break;
                    }
                  }
                } catch (e) {
                  // 이 선택자로 찾지 못하면 다음 선택자 시도
                  continue;
                }
              }
              
              // 버튼을 찾지 못한 경우 ESC 키 시도
                if (!closed) {
                try {
                  await page.keyboard.press('Escape');
                  console.log("ESC 키로 팝업 닫기 시도");
                  await page.waitForTimeout(500);
                  
                  // 팝업이 사라졌는지 확인
                  const stillVisible = await context.evaluate((el) => {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                  }, popup);
                  
                  if (!stillVisible) {
                    console.log("ESC 키로 팝업 닫기 성공");
                    return true;
                  }
                } catch (e) {
                  console.log(`ESC 키로 팝업 닫기 실패: ${e.message}`);
                }
              }
              
              // 여전히 닫히지 않은 경우 팝업을 숨기기 시도
              if (!closed) {
                try {
                  await context.evaluate((el) => {
                    if (el) {
                      el.style.display = 'none';
                      el.style.visibility = 'hidden';
                    }
                  }, popup);
                  console.log("팝업을 강제로 숨김");
                  await page.waitForTimeout(500);
                  return true;
                } catch (e) {
                  console.log(`팝업 숨기기 실패: ${e.message}`);
                }
              }
              
              // 팝업을 처리했으면 종료
              if (closed) {
                return true;
              }
            }
          }
          } catch (e) {
            // 이 선택자로 팝업을 찾지 못함
            continue;
          }
        }
      }
      
      // 팝업이 발견되지 않았고 아직 시간이 남았다면 계속 확인
      if (!popupFound && i < maxWaitTime - 1) {
        // 다음 확인을 위해 계속
      }
    }
    
    console.log("10초 동안 팝업이 나타나지 않았습니다. 이체 메뉴 클릭을 진행합니다.");
    return true; // 팝업이 없어도 정상 진행
  } catch (error) {
    console.log(`메인 페이지 팝업 확인 중 오류: ${error.message}`);
    // 오류가 발생해도 계속 진행
    return true;
  }
}

// 보이스피싱 예방 팝업 처리
async function handleVoicePhishingPopup(page) {
  try {
    const frame = await switchToFrame(page, "hanaMainframe");
    
    for (let i = 0; i < 60; i++) {
      await page.waitForTimeout(1000);
      console.log(`팝업 감지 중... ${i + 1}초 경과`);

      const popup1 = await frame.$("#voicePhishingPopup1");
      const popup2 = await frame.$("#lonFrdInfoPop");

      if (popup1) {
        const isVisible = await frame.evaluate((el) => {
          return el && el.offsetParent !== null;
        }, popup1);
        
        if (isVisible) {
          console.log("보이스피싱 예방 팝업(voicePhishingPopup1) 감지됨");
          const [noButton] = await frame.$x("//a[contains(@onclick, 'pbk.transfer.common.lonFrdInfoPopN()')]");
          if (noButton) {
            await noButton.click();
            console.log("'아니요' 버튼 클릭 성공");
            return true;
          }
        }
      }

      if (popup2) {
        const isVisible = await frame.evaluate((el) => {
          return el && el.offsetParent !== null;
        }, popup2);
        
        if (isVisible) {
          console.log("보이스피싱 예방 팝업(lonFrdInfoPop) 감지됨");
          const [noButton] = await frame.$x("//a[contains(@onclick, 'pbk.transfer.common.lonFrdInfoPopN()')]");
          if (noButton) {
            await noButton.click();
            console.log("'아니요' 버튼 클릭 성공");
            return true;
          }
        }
      }
    }

    console.log("60초 동안 보이스피싱 예방 팝업이 나타나지 않음");
    return false;
  } catch (error) {
    console.log(`팝업 처리 중 오류 발생: ${error.message}`);
    return false;
  }
}

// 메인 이체 프로세스 실행
async function executeTransferProcess(page, excelPath, autoTransfer = false) {
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

    // 7. 스크롤 조정
    await adjustScroll(page);

    // 8. 엑셀 데이터 로드
    const processedData = loadExcelData(excelPath);
    console.log("\n전처리된 데이터:");
    processedData.forEach((data, idx) => {
      console.log(`${idx + 1}. 은행: ${data.bank}, 계좌번호: ${data.accountNumber}, 이름.제품명: ${data.nameProduct}, 제품명: ${data.productName}, 금액: ${data.amount}`);
    });

    // 8. 이체 정보 입력 (최대 10개)
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
      console.log("이체가 완료되었습니다.");
    } else {
      console.log("이체가 취소되었습니다. 필요시 수동으로 다계좌이체진행 버튼을 클릭하세요.");
    }

    console.log("\n=== 이체 프로세스 완료 ===");
    return true;
  } catch (error) {
    console.error(`이체 프로세스 실행 중 오류: ${error.message}`);
    return false;
  }
}

module.exports = {
  executeTransferProcess,
  clickTransferMenu,
  clickMultiTransferButton,
  inputTransferInfo,
  enterPassword,
  clickTransferButton,
  handleVoicePhishingPopup
};

