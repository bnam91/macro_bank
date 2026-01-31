// 팝업 처리 함수들

import { switchToFrame } from './frame-utils.js';

// 메인 페이지 팝업 확인 및 처리 (최대 15초)
export async function checkAndCloseMainPagePopup(page) {
  try {
    console.log("메인 페이지 팝업 확인 중... (최대 15초)");
    
    const maxWaitTime = 15; // 최대 15초
    
    for (let i = 0; i < maxWaitTime; i++) {
      const remainingTime = maxWaitTime - i;
      // 같은 줄에 카운트다운 표시 (업데이트)
      process.stdout.write(`\r⏳ 팝업 확인 중... 남은 시간: ${remainingTime}초`);
      
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
                // 카운트다운 줄 정리 후 팝업 발견 메시지 출력
                process.stdout.write('\r' + ' '.repeat(50) + '\r');
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
                          // 카운트다운 줄 정리
                          process.stdout.write('\r' + ' '.repeat(50) + '\r');
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
                      // 카운트다운 줄 정리
                      process.stdout.write('\r' + ' '.repeat(50) + '\r');
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
                    // 카운트다운 줄 정리
                    process.stdout.write('\r' + ' '.repeat(50) + '\r');
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
    
    // 카운트다운 줄 정리
    process.stdout.write('\r' + ' '.repeat(50) + '\r');
    console.log("15초 동안 팝업이 나타나지 않았습니다. 이체 메뉴 클릭을 진행합니다.");
    return true; // 팝업이 없어도 정상 진행
  } catch (error) {
    // 카운트다운 줄 정리
    process.stdout.write('\r' + ' '.repeat(50) + '\r');
    console.log(`메인 페이지 팝업 확인 중 오류: ${error.message}`);
    // 오류가 발생해도 계속 진행
    return true;
  }
}

// 단말기 미지정 PC이용 안내 팝업 감지 및 처리
export async function checkAndHandleDevicePopup(page) {
  try {
    console.log("\n단말기 미지정 PC이용 안내 팝업 확인 중...");
    
    const maxWaitTime = 5; // 최대 5초 동안 확인
    
    for (let i = 0; i < maxWaitTime; i++) {
      await page.waitForTimeout(1000);
      
      // 팝업 선택자들 (메인 페이지와 프레임 모두 확인)
      const popupSelectors = [
        'div.pop_ty01.pop_ty05.unpc_1',
        '.pop_ty01.pop_ty05.unpc_1',
        '.unpc_1',
        'div[class*="unpc_1"]'
      ];
      
      // 모든 컨텍스트에서 팝업 찾기 (메인 페이지 + 모든 프레임)
      const contexts = [page, ...page.frames()];
      
      for (const context of contexts) {
        for (const popupSelector of popupSelectors) {
          try {
            const popup = await context.$(popupSelector);
            if (popup) {
              const isVisible = await context.evaluate((el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                return style.display !== 'none' && 
                       style.visibility !== 'hidden' && 
                       el.offsetParent !== null;
              }, popup);
              
              if (isVisible) {
                // 팝업 제목 확인
                const title = await context.evaluate((el) => {
                  const h4 = el.querySelector('h4');
                  return h4 ? h4.textContent.trim() : '';
                }, popup);
                
                if (title.includes('단말기 미지정 PC이용 안내')) {
                  console.log("✅ 단말기 미지정 PC이용 안내 팝업 감지됨");
                  
                  // 팝업 닫기 시도
                  let closed = false;
                  
                  // 방법 1: 닫기 버튼 클릭 (onclick으로 닫기)
                  try {
                    const closeButton = await popup.$('a[onclick*="closeLayer_fnc(\'commonInfoPopup\')"]');
                    if (closeButton) {
                      await closeButton.click();
                      await page.waitForTimeout(500);
                      closed = true;
                      console.log("팝업 닫기 버튼 클릭 성공");
                    }
                  } catch (e) {
                    // 닫기 버튼을 찾지 못함
                  }
                  
                  // 방법 2: JavaScript로 직접 닫기 함수 호출
                  if (!closed) {
                    try {
                      await context.evaluate(() => {
                        if (typeof opb !== 'undefined' && 
                            opb.common && 
                            opb.common.layerpopup && 
                            opb.common.layerpopup.closeLayer_fnc) {
                          opb.common.layerpopup.closeLayer_fnc('commonInfoPopup');
                        }
                      });
                      await page.waitForTimeout(500);
                      closed = true;
                      console.log("JavaScript로 팝업 닫기 성공");
                    } catch (e) {
                      console.log(`JavaScript로 팝업 닫기 실패: ${e.message}`);
                    }
                  }
                  
                  // 방법 3: 팝업을 강제로 숨기기
                  if (!closed) {
                    try {
                      await context.evaluate((el) => {
                        if (el) {
                          el.style.display = 'none';
                          el.style.visibility = 'hidden';
                        }
                      }, popup);
                      await page.waitForTimeout(500);
                      closed = true;
                      console.log("팝업을 강제로 숨김");
                    } catch (e) {
                      console.log(`팝업 숨기기 실패: ${e.message}`);
                    }
                  }
                  
                  // 팝업이 닫혔는지 확인
                  if (closed) {
                    await page.waitForTimeout(500);
                    const stillVisible = await context.evaluate((el) => {
                      if (!el) return false;
                      const style = window.getComputedStyle(el);
                      return style.display !== 'none' && 
                             style.visibility !== 'hidden' && 
                             el.offsetParent !== null;
                    }, popup);
                    
                    if (!stillVisible) {
                      console.log("✅ 팝업이 닫혔습니다.");
                      return true; // 팝업이 감지되고 닫혔음
                    }
                  }
                  
                  // 팝업이 여전히 보이면 사용자에게 확인
                  return true; // 팝업 감지됨
                }
              }
            }
          } catch (e) {
            // 이 선택자로 팝업을 찾지 못함
            continue;
          }
        }
      }
    }
    
    // 팝업이 나타나지 않음
    return false;
  } catch (error) {
    console.log(`단말기 미지정 팝업 확인 중 오류: ${error.message}`);
    return false;
  }
}

// 보이스피싱 예방 팝업 처리
export async function handleVoicePhishingPopup(page) {
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
