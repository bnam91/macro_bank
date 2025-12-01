const { debugPage } = require('./debug-utils');

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

// 공인인증서 팝업 처리
async function handleCertPopup(page) {
  try {
    console.log("공인인증서 팝업 대기 중...");
    
    // 팝업이 iframe 안에 있으므로 iframe을 찾아야 함
    let popupFrame = null;
    let targetPage = page; // 기본값은 메인 페이지
    
    // delfino4htmlIframe 찾기
    try {
      await page.waitForSelector("#delfino4htmlIframe, iframe[name='delfino4htmlIframe']", { timeout: 30000 });
      console.log("인증서 팝업 iframe 발견");
      
      // iframe 요소 가져오기
      const iframeElement = await page.$("#delfino4htmlIframe");
      if (iframeElement) {
        popupFrame = await iframeElement.contentFrame();
        if (popupFrame) {
          console.log("iframe 내부 프레임 접근 성공");
          targetPage = popupFrame;
          // iframe 내부가 로드될 때까지 대기
          await page.waitForTimeout(2000);
        }
      }
    } catch (error) {
      console.log(`iframe 찾기 실패: ${error.message}`);
      // 다른 방법으로 iframe 찾기 시도
      const frames = page.frames();
      for (const frame of frames) {
        if (frame.name() === 'delfino4htmlIframe' || frame.url().includes('delfinoG10')) {
          popupFrame = frame;
          targetPage = frame;
          console.log("프레임 목록에서 인증서 iframe 발견");
          break;
        }
      }
    }
    
    // iframe을 찾지 못한 경우 페이지 소스 저장 및 프레임 정보 출력
    if (!popupFrame) {
      try {
        await debugPage(page, 'page-debug.html');
      } catch (error) {
        console.log(`페이지 디버깅 정보 저장 실패: ${error.message}`);
      }
      console.log("iframe을 찾지 못했습니다. 메인 페이지에서 시도합니다...");
    }
    
    // 팝업 컨테이너가 나타날 때까지 대기
    try {
      await targetPage.waitForSelector("#w2ui-popup_0, .w2ui-popup, #selectDialogBody", { timeout: 10000 });
      console.log("팝업 컨테이너 발견");
    } catch (error) {
      console.log("팝업 컨테이너를 찾지 못했습니다.");
    }
    
    // 1. 로컬디스크 버튼 클릭 (이미 선택되어 있을 수도 있음)
    try {
      // 로컬디스크 버튼이 실제로 나타날 때까지 대기 (waitForSelector 사용)
      let localDiskButton = null;
      const selectors = [
        "#w2ui-popup_0 .localDiskButton",  // 성공한 선택자
        "#selectDialogBody .localDiskButton"  // 대체 선택자
      ];
      
      // 각 선택자로 버튼이 나타날 때까지 대기 시도 (waitForSelector 사용)
      for (let i = 0; i < selectors.length; i++) {
        const selector = selectors[i];
        console.log(`[로컬디스크] 선택자 시도 ${i + 1}/${selectors.length}: ${selector}`);
        try {
          // waitForSelector로 버튼이 나타날 때까지 대기
          await targetPage.waitForSelector(selector, { timeout: 10000, visible: true });
          localDiskButton = await targetPage.$(selector);
          if (localDiskButton) {
            // 버튼이 실제로 보이고 클릭 가능한지 확인
            const isVisible = await targetPage.evaluate((el) => {
              return el && el.offsetParent !== null && !el.disabled;
            }, localDiskButton);
            if (isVisible) {
              console.log(`✅ [로컬디스크] 선택자 성공: ${selector}`);
              break;
            } else {
              console.log(`⚠️ [로컬디스크] 선택자 발견했지만 보이지 않음: ${selector}`);
            }
          } else {
            console.log(`❌ [로컬디스크] 선택자로 요소를 찾지 못함: ${selector}`);
          }
        } catch (e) {
          // 이 선택자로 찾지 못하면 다음 선택자 시도
          console.log(`❌ [로컬디스크] 선택자 타임아웃 또는 오류: ${selector} - ${e.message}`);
          continue;
        }
      }
      
      if (localDiskButton) {
        // JavaScript로 직접 클릭 (더 안정적)
        await targetPage.evaluate((button) => {
          if (button) button.click();
        }, localDiskButton);
        console.log("로컬디스크 버튼 클릭 성공");
        await page.waitForTimeout(500);
      } else {
        console.log("로컬디스크 버튼을 찾을 수 없습니다.");
      }
    } catch (error) {
      console.log("로컬디스크 버튼 클릭 실패: " + error.message);
    }
    
    // 2. 확장매체(외장하드) 버튼 클릭 - "Seagate Backup Plus Drive (D:)" 체크
    try {
      // Seagate Backup Plus Drive를 찾아서 체크박스를 체크
      let extensionElement = null;
      let checkbox = null;
      const selectors = [
        "li.certStore1",  // certStore1 클래스를 가진 li 요소
        "li[aria-label*='Seagate']",  // Seagate가 포함된 aria-label을 가진 li
        "li[aria-label*='Backup']",  // Backup이 포함된 aria-label을 가진 li
        "#certStorePopupBody li.certStore1"  // certStorePopupBody 내부의 certStore1
      ];
      
      // 각 선택자로 요소가 나타날 때까지 대기 시도
      for (let i = 0; i < selectors.length; i++) {
        const selector = selectors[i];
        console.log(`[확장매체] 선택자 시도 ${i + 1}/${selectors.length}: ${selector}`);
        try {
          await targetPage.waitForSelector(selector, { timeout: 5000, visible: true });
          extensionElement = await targetPage.$(selector);
          if (extensionElement) {
            // 요소가 실제로 보이는지 확인
            const isVisible = await targetPage.evaluate((el) => {
              return el && el.offsetParent !== null;
            }, extensionElement);
            if (isVisible) {
              console.log(`✅ [확장매체] 선택자 성공: ${selector}`);
              
              // li 내부의 체크박스 찾기
              checkbox = await extensionElement.$("input[type='checkbox']");
              if (checkbox) {
                console.log(`✅ [확장매체] 체크박스 발견`);
                // 체크박스 발견 후 1초 대기
                await page.waitForTimeout(1000);
                break;
              } else {
                console.log(`⚠️ [확장매체] 요소는 찾았지만 체크박스를 찾지 못함: ${selector}`);
              }
            } else {
              console.log(`⚠️ [확장매체] 선택자 발견했지만 보이지 않음: ${selector}`);
            }
          } else {
            console.log(`❌ [확장매체] 선택자로 요소를 찾지 못함: ${selector}`);
          }
        } catch (e) {
          // 이 선택자로 찾지 못하면 다음 선택자 시도
          console.log(`❌ [확장매체] 선택자 타임아웃 또는 오류: ${selector} - ${e.message}`);
          continue;
        }
      }
      
      if (extensionElement && checkbox) {
        // 체크박스가 체크되어 있지 않으면 체크
        const isChecked = await targetPage.evaluate((cb) => {
          return cb && cb.checked;
        }, checkbox);
        
        if (!isChecked) {
          // 체크박스를 직접 체크
          await targetPage.evaluate((cb) => {
            if (cb) {
              cb.checked = true;
              // change 이벤트 발생
              cb.dispatchEvent(new Event('change', { bubbles: true }));
              cb.dispatchEvent(new Event('click', { bubbles: true }));
            }
          }, checkbox);
          
          // li 요소도 클릭 (UI 업데이트를 위해)
          await targetPage.evaluate((el) => {
            if (el) el.click();
          }, extensionElement);
          
          console.log("Seagate Backup Plus Drive (D:) 체크박스 체크 성공");
        } else {
          console.log("Seagate Backup Plus Drive (D:)는 이미 체크되어 있습니다.");
        }
        
        await page.waitForTimeout(1000);
      } else {
        console.log("Seagate Backup Plus Drive 요소 또는 체크박스를 찾을 수 없습니다.");
      }
    } catch (error) {
      console.log("확장매체 버튼 클릭 실패: " + error.message);
    }
    
    // 3. 인증서 리스트에 '신현빈'이 나타나는지 확인 및 선택
    try {
      // 인증서 리스트가 업데이트될 때까지 대기 (확장매체 선택 후)
      console.log("인증서 리스트 업데이트 대기 중...");
      await page.waitForTimeout(2000); // 인증서 리스트 로드 대기
      
      // '신현빈'이 포함된 인증서를 찾기
      let shinCert = null;
      const certSelectors = [
        "#grid_certificateInfos_0_rec_0",  // 첫 번째 인증서
        "tr[id^='grid_certificateInfos_0_rec_']",  // 모든 인증서 행
      ];
      
      // 인증서 리스트에서 '신현빈' 찾기
      try {
        // 인증서 테이블이 나타날 때까지 대기
        await targetPage.waitForSelector("#grid_certificateInfos_0_rec_0, tr[id^='grid_certificateInfos_0_rec_']", { timeout: 10000 });
        
        // 모든 인증서 행을 가져와서 '신현빈'이 포함된 것 찾기
        const certRows = await targetPage.$$("tr[id^='grid_certificateInfos_0_rec_']");
        
        for (const row of certRows) {
          const text = await targetPage.evaluate((el) => {
            return el ? el.textContent || el.innerText : '';
          }, row);
          
          if (text.includes('신현빈')) {
            shinCert = row;
            console.log("✅ 인증서 리스트에서 '신현빈' 발견!");
            break;
          }
        }
        
        // '신현빈'을 찾지 못한 경우 첫 번째 인증서 사용
        if (!shinCert && certRows.length > 0) {
          shinCert = certRows[0];
          console.log("⚠️ '신현빈'을 찾지 못했습니다. 첫 번째 인증서를 사용합니다.");
        }
      } catch (e) {
        console.log(`인증서 리스트 찾기 실패: ${e.message}`);
      }
      
      // 인증서 선택
      if (shinCert) {
        const isSelected = await targetPage.evaluate((el) => {
          return el && el.classList.contains('w2ui-selected');
        }, shinCert);
        
        if (!isSelected) {
          await targetPage.evaluate((el) => {
            if (el) el.click();
          }, shinCert);
          console.log("인증서 선택 성공");
          await page.waitForTimeout(500);
        } else {
          console.log("인증서가 이미 선택되어 있습니다.");
        }
      } else {
        console.log("인증서를 찾을 수 없습니다.");
      }
    } catch (error) {
      console.log("인증서 선택 실패 또는 이미 선택됨: " + error.message);
    }
    
    // 4. 비밀번호 입력
    try {
      // 비밀번호 입력 필드가 나타날 때까지 대기 (waitForSelector 사용)
      let passwordInput = null;
      const passwordSelector = "input[name='selectDialogPasswordInput']";  // 성공한 선택자만 사용
      
      try {
        await targetPage.waitForSelector(passwordSelector, { timeout: 10000, visible: true });
        passwordInput = await targetPage.$(passwordSelector);
        if (passwordInput) {
          // 입력 필드가 실제로 보이고 활성화되어 있는지 확인
          const isVisible = await targetPage.evaluate((el) => {
            return el && el.offsetParent !== null && !el.disabled;
          }, passwordInput);
          if (isVisible) {
            console.log(`✅ [비밀번호] 선택자 성공: ${passwordSelector}`);
          } else {
            console.log(`⚠️ [비밀번호] 선택자 발견했지만 보이지 않음: ${passwordSelector}`);
          }
        }
      } catch (e) {
        console.log(`❌ [비밀번호] 선택자 타임아웃 또는 오류: ${passwordSelector} - ${e.message}`);
      }
      
      let passwordInputSuccess = false; // 비밀번호 입력 성공 여부
      const password = "@gusqls120";
      
      if (passwordInput) {
        console.log(`비밀번호 입력 시도: ${password.replace(/./g, '*')}`);
        
        // 여러 방법을 순차적으로 시도
        const methods = [
          {
            name: "방법 1/5: focus + click + select + Backspace + keyboard.type",
            execute: async () => {
        await targetPage.evaluate((input) => {
          if (input) {
            input.focus();
            input.click();
                  input.select();
          }
        }, passwordInput);
              await page.waitForTimeout(300);
              await page.keyboard.press('Backspace');
        await page.waitForTimeout(100);
              for (const char of password) {
                await page.keyboard.type(char);
                await page.waitForTimeout(70);
              }
              await page.waitForTimeout(500);
            }
          },
          {
            name: "방법 2/5: 더 긴 대기 시간 후 입력",
            execute: async () => {
              await targetPage.evaluate((input) => {
                if (input) {
                  input.focus();
                  input.click();
                }
              }, passwordInput);
              await page.waitForTimeout(500);
              // Ctrl+A 대신 Meta+A 사용 (Windows에서는 Control, Mac에서는 Meta)
              await page.keyboard.down('Control');
              await page.keyboard.press('a');
              await page.keyboard.up('Control');
              await page.waitForTimeout(200);
              for (const char of password) {
                await page.keyboard.type(char);
                await page.waitForTimeout(70);
              }
              await page.waitForTimeout(500);
            }
          },
          {
            name: "방법 3/5: 여러 번 클릭 후 입력",
            execute: async () => {
              await targetPage.evaluate((input) => {
                if (input) {
                  input.click();
                  input.click();
                  input.focus();
                }
              }, passwordInput);
              await page.waitForTimeout(400);
              await page.keyboard.press('End');
              await page.waitForTimeout(100);
              for (let i = 0; i < 20; i++) {
                await page.keyboard.press('Backspace');
              }
              await page.waitForTimeout(200);
              for (const char of password) {
                await page.keyboard.type(char);
                await page.waitForTimeout(70);
              }
              await page.waitForTimeout(500);
            }
          },
          {
            name: "방법 4/5: input.value 직접 설정 + 이벤트 발생",
            execute: async () => {
        await targetPage.evaluate((input, pwd) => {
          if (input) {
                  input.focus();
                  input.click();
            input.value = pwd;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
                  input.dispatchEvent(new Event('keyup', { bubbles: true }));
          }
        }, passwordInput, password);
        await page.waitForTimeout(500);
            }
          },
          {
            name: "방법 5/5: 키보드 직접 입력 (필드 없이)",
            execute: async () => {
              // 필드에 포커스만 주고 직접 타이핑
              await targetPage.evaluate((input) => {
                if (input) {
                  input.focus();
                  input.click();
                }
              }, passwordInput);
              await page.waitForTimeout(300);
              for (const char of password) {
                await page.keyboard.type(char);
                await page.waitForTimeout(70);
              }
              await page.waitForTimeout(500);
            }
          }
        ];
        
        // 각 방법을 순차적으로 시도
        for (let i = 0; i < methods.length; i++) {
          const method = methods[i];
          const methodNum = i + 1;
          console.log(`[비밀번호] 방법 ${methodNum}/${methods.length} 시도 중...`);
          
          try {
            await method.execute();
            
            // 비밀번호가 실제로 입력되었는지 확인
            const inputValue = await targetPage.evaluate((input) => {
              return input ? input.value : '';
            }, passwordInput);
            
            // 암호화된 필드는 값이 다를 수 있으므로 길이로 확인
            if (inputValue && inputValue.length >= password.length) {
              console.log(`✅ 비밀번호 입력 성공! (방법 ${methodNum}/${methods.length}, 입력된 길이: ${inputValue.length})`);
              passwordInputSuccess = true;
              break; // 성공하면 다음 방법 시도하지 않음
            } else {
              console.log(`   ⚠️ 방법 ${methodNum} 실패 (입력된 길이: ${inputValue ? inputValue.length : 0}, 예상: ${password.length})`);
              if (i < methods.length - 1) {
                await page.waitForTimeout(300);
              }
            }
          } catch (error) {
            console.log(`   ❌ 방법 ${methodNum} 오류: ${error.message}`);
            if (i < methods.length - 1) {
              await page.waitForTimeout(300);
            }
          }
        }
        
        if (!passwordInputSuccess) {
          console.log(`❌ 모든 비밀번호 입력 방법 실패 (${methods.length}가지 방법 시도)`);
        } else {
          console.log(`✅ 비밀번호 입력 완료!`);
        }
      } else {
        // 비밀번호 입력 필드를 찾지 못한 경우 직접 타이핑
        console.log("비밀번호 입력 필드를 찾지 못했습니다. 키보드로 직접 입력 시도...");
        for (const char of password) {
          await page.keyboard.type(char);
          await page.waitForTimeout(70);
        }
        console.log("키보드로 비밀번호 입력 성공");
        passwordInputSuccess = true; // 직접 타이핑은 성공으로 간주
      }
      
      // 비밀번호 입력 실패 시 확인 버튼 클릭하지 않고 종료
      if (!passwordInputSuccess) {
        console.log("❌ 비밀번호 입력이 실패하여 프로세스를 중단합니다.");
        return false;
      }
    } catch (error) {
      console.log("비밀번호 입력 실패: " + error.message);
      return false;
    }
    
    // 6. 확인 버튼 클릭
    try {
      // 확인 버튼이 나타날 때까지 대기 (waitForSelector 사용)
      let okButton = null;
      const okButtonSelectors = [
        "#delfino-section .okButton.okButtonMsg",  // delfino-section 내부의 확인 버튼 (가장 정확)
        "button.okButton.okButtonMsg",  // 기본 선택자
        ".okButtonBlock button.okButton.okButtonMsg"  // okButtonBlock 내부의 확인 버튼
      ];
      
      for (let i = 0; i < okButtonSelectors.length; i++) {
        const selector = okButtonSelectors[i];
        console.log(`[확인버튼] 선택자 시도 ${i + 1}/${okButtonSelectors.length}: ${selector}`);
        try {
          await targetPage.waitForSelector(selector, { timeout: 5000, visible: true });
          okButton = await targetPage.$(selector);
          if (okButton) {
            const isVisible = await targetPage.evaluate((el) => {
              return el && el.offsetParent !== null && !el.disabled;
            }, okButton);
            if (isVisible) {
              console.log(`✅ [확인버튼] 선택자 성공: ${selector}`);
              break;
            } else {
              console.log(`⚠️ [확인버튼] 선택자 발견했지만 보이지 않거나 비활성화됨: ${selector}`);
            }
          }
        } catch (e) {
          console.log(`❌ [확인버튼] 선택자 타임아웃 또는 오류: ${selector} - ${e.message}`);
          continue;
        }
      }
      
      if (okButton) {
        // 확인 버튼이 비활성화되어 있는지 확인
        const buttonInfo = await targetPage.evaluate((button) => {
          return {
            disabled: button.disabled
          };
        }, okButton);
        
        // 확인 버튼이 비활성화되어 있으면 활성화 대기
        if (buttonInfo.disabled) {
          console.log("⚠️ 확인 버튼이 비활성화되어 있습니다. 활성화될 때까지 대기...");
          await targetPage.waitForFunction(
            (selector) => {
              const btn = document.querySelector(selector);
              return btn && !btn.disabled;
            },
            { timeout: 10000 },
            ".okButton, button.okButton, #w2ui-popup_0 .okButton, #selectDialogBody .okButton"
          );
          console.log("✅ 확인 버튼이 활성화되었습니다.");
        }
        
        // 팝업이 사라졌는지 확인하는 헬퍼 함수
        const checkPopupDisappeared = async () => {
          try {
            // delfino-section이 사라졌는지 확인 (최대 3초 대기)
            for (let i = 0; i < 6; i++) {
              await page.waitForTimeout(500);
              
              try {
                const delfinoSection = await targetPage.$("#delfino-section");
                if (!delfinoSection) {
                  return true;
                } else {
                  const isVisible = await targetPage.evaluate((el) => {
                    return el && el.offsetParent !== null;
                  }, delfinoSection);
                  if (!isVisible) {
                    return true;
                  }
                }
              } catch (e) {
                // iframe이 detached된 경우 성공으로 간주
                if (e.message.includes('detached Frame')) {
                  return true;
                }
              }
            }
            
            // iframe이 사라졌는지도 확인
            try {
              const iframe = await page.$("#delfino4htmlIframe");
              if (!iframe) {
                return true;
              } else {
                const isVisible = await page.evaluate((el) => {
                  return el && el.offsetParent !== null;
                }, iframe);
                if (!isVisible) {
                  return true;
                }
              }
            } catch (e) {
              return true;
            }
            
            return false;
          } catch (e) {
            // 오류 발생 시 성공으로 간주 (iframe이 detached된 경우 등)
            if (e.message.includes('detached Frame')) {
              return true;
            }
            return false;
          }
        };
        
        // 확인 버튼 클릭 (여러 방법 시도)
        console.log("⏳ 확인 버튼 클릭 전 1초 대기...");
        await page.waitForTimeout(1000);
        console.log("🖱️ 확인 버튼 클릭 시도");
        
        // 여러 방법 정의
        const clickMethods = [
          {
            name: "방법 1/3: Puppeteer 직접 클릭",
            execute: async () => {
              await okButton.click();
              await page.waitForTimeout(1000);
            }
          },
          {
            name: "방법 2/3: JavaScript 이벤트 발생",
            execute: async () => {
        await targetPage.evaluate((button) => {
                if (button) {
                  button.focus();
                  button.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
                  button.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
                  button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                  if (typeof button.click === 'function') {
                    button.click();
                  }
                }
        }, okButton);
              await page.waitForTimeout(1000);
            }
          },
          {
            name: "방법 3/3: JavaScript 직접 클릭",
            execute: async () => {
              await targetPage.evaluate((button) => {
                if (button) {
                  button.click();
                }
              }, okButton);
              await page.waitForTimeout(1000);
            }
          }
        ];
        
        let clickSuccess = false;
        
        // 각 방법을 순차적으로 시도
        for (let i = 0; i < clickMethods.length; i++) {
          const method = clickMethods[i];
          const methodNum = i + 1;
          
          console.log(`[확인버튼] ${method.name} 시도 중...`);
          
          try {
            await method.execute();
            console.log(`✅ ${method.name} 완료`);
            
            // 팝업이 사라졌는지 확인
            const disappeared = await checkPopupDisappeared();
            
            if (disappeared) {
              console.log(`✅ 확인 버튼 클릭 성공! (${method.name})`);
              clickSuccess = true;
              break; // 성공하면 다음 방법 시도하지 않음
            } else {
              console.log(`⚠️ ${method.name} 시도했지만 팝업이 아직 존재합니다.`);
            }
          } catch (e) {
            console.log(`❌ ${method.name} 오류: ${e.message}`);
            
            // iframe이 detached된 경우 성공으로 간주
            if (e.message.includes('detached Frame')) {
              console.log(`✅ 확인 버튼 클릭 성공! (iframe이 사라짐 - ${method.name})`);
              clickSuccess = true;
              break;
            }
          }
        }
        
        if (!clickSuccess) {
          console.log("❌ 모든 확인 버튼 클릭 방법 실패");
          return false;
        }
        
        await page.waitForTimeout(2000);
        return true;
      } else {
        console.log("확인 버튼을 찾을 수 없습니다.");
        return false;
      }
    } catch (error) {
      console.log("확인 버튼 클릭 실패: " + error.message);
      return false;
    }
    
  } catch (error) {
    console.log(`공인인증서 팝업 처리 중 오류: ${error.message}`);
    return false;
  }
}

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

module.exports = {
  clickCertMenu,
  clickCertLoginButton,
  handleCertPopup,
  executeCertLogin
};

