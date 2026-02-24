import { debugPage } from '../../utils/debug-utils.js';

async function findPopupFrame(page) {
  let popupFrame = null;
  let targetPage = page;

  try {
    await page.waitForSelector("#delfino4htmlIframe, iframe[name='delfino4htmlIframe']", { timeout: 30000 });
    console.log("인증서 팝업 iframe 발견");

    const iframeElement = await page.$("#delfino4htmlIframe");
    if (iframeElement) {
      popupFrame = await iframeElement.contentFrame();
      if (popupFrame) {
        console.log("iframe 내부 프레임 접근 성공");
        targetPage = popupFrame;
        await page.waitForTimeout(2000);
      }
    }
  } catch (error) {
    console.log(`iframe 찾기 실패: ${error.message}`);
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

  if (!popupFrame) {
    try {
      await debugPage(page, 'page-debug.html');
    } catch (error) {
      console.log(`페이지 디버깅 정보 저장 실패: ${error.message}`);
    }
    console.log("iframe을 찾지 못했습니다. 메인 페이지에서 시도합니다...");
  }

  return { popupFrame, targetPage };
}

async function waitForPopupContainer(targetPage) {
  try {
    await targetPage.waitForSelector("#w2ui-popup_0, .w2ui-popup, #selectDialogBody", { timeout: 10000 });
    console.log("팝업 컨테이너 발견");
  } catch (error) {
    console.log("팝업 컨테이너를 찾지 못했습니다.");
  }
}

async function clickLocalDisk(targetPage, page) {
  try {
    let localDiskButton = null;
    const selectors = [
      "#w2ui-popup_0 .localDiskButton",
      "#selectDialogBody .localDiskButton"
    ];

    for (let i = 0; i < selectors.length; i++) {
      const selector = selectors[i];
      console.log(`[로컬디스크] 선택자 시도 ${i + 1}/${selectors.length}: ${selector}`);
      try {
        await targetPage.waitForSelector(selector, { timeout: 10000, visible: true });
        localDiskButton = await targetPage.$(selector);
        if (localDiskButton) {
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
        console.log(`❌ [로컬디스크] 선택자 타임아웃 또는 오류: ${selector} - ${e.message}`);
        continue;
      }
    }

    if (localDiskButton) {
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
}

async function selectCertStore(targetPage, page, certStoreKeywords, certStoreIndex = 0) {
  try {
    let extensionElement = null;
    let checkbox = null;

    if (certStoreKeywords.length > 0) {
      console.log(`설정된 저장매체 키워드로 검색: ${certStoreKeywords.join(', ')}`);
      try {
        await targetPage.waitForSelector("#certStorePopupBody li", { timeout: 5000, visible: true });
        const storeElements = await targetPage.$$("#certStorePopupBody li");
        const availableStores = [];

        for (const element of storeElements) {
          const labelText = await targetPage.evaluate((el) => {
            const aria = (el.getAttribute('aria-label') || '').trim();
            const text = (el.innerText || '').trim();
            if (aria && text && aria !== text) {
              return `${aria} ${text}`.trim();
            }
            return (aria || text).trim();
          }, element);
          if (labelText) {
            if (!availableStores.includes(labelText)) {
              availableStores.push(labelText);
            }
          }

          const normalized = labelText.toLowerCase();
          const matched = certStoreKeywords.some((keyword) =>
            normalized.includes(keyword.toLowerCase())
          );

          if (matched) {
            console.log(`✅ [확장매체] 설정 키워드로 발견: ${labelText}`);
            const foundCheckbox = await element.$("input[type='checkbox']");
            if (foundCheckbox) {
              extensionElement = element;
              checkbox = foundCheckbox;
              console.log(`✅ [확장매체] 체크박스 발견`);
              break;
            } else {
              console.log(`⚠️ [확장매체] 체크박스를 찾지 못함: ${labelText}`);
            }
          }
        }
        if (availableStores.length > 0) {
          console.log(`로컬디스크/저장매체 목록(${availableStores.length}):`);
          availableStores.forEach((store) => console.log(`- ${store}`));
        } else {
          console.log("로컬디스크/저장매체 목록을 찾지 못했습니다.");
        }
      } catch (e) {
        console.log(`❌ [확장매체] 설정 키워드 검색 실패: ${e.message}`);
      }
    }

    // config의 certStoreIndex 사용 (0=로컬 디스크 (C:), 1=DVD (D:), 2=Home on Mac (Z:) 등)
    const index = Math.max(0, certStoreIndex);
    const selectors = [
      `li.certStore${index}`,
      `#certStorePopupBody li.certStore${index}`,
      "li.certStore1",
      "li[aria-label*='Seagate']",
      "li[aria-label*='Backup']",
      "#certStorePopupBody li.certStore1"
    ];

    for (let i = 0; i < selectors.length && !extensionElement; i++) {
      const selector = selectors[i];
      console.log(`[확장매체] 선택자 시도 ${i + 1}/${selectors.length}: ${selector}`);
      try {
        await targetPage.waitForSelector(selector, { timeout: 5000, visible: true });
        extensionElement = await targetPage.$(selector);
        if (extensionElement) {
          const isVisible = await targetPage.evaluate((el) => {
            return el && el.offsetParent !== null;
          }, extensionElement);
          if (isVisible) {
            console.log(`✅ [확장매체] 선택자 성공: ${selector}`);
            checkbox = await extensionElement.$("input[type='checkbox']");
            if (checkbox) {
              console.log(`✅ [확장매체] 체크박스 발견`);
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
        console.log(`❌ [확장매체] 선택자 타임아웃 또는 오류: ${selector} - ${e.message}`);
        continue;
      }
    }

    if (extensionElement && checkbox) {
      const isChecked = await targetPage.evaluate((cb) => {
        return cb && cb.checked;
      }, checkbox);

      if (!isChecked) {
        await targetPage.evaluate((cb) => {
          if (cb) {
            cb.checked = true;
            cb.dispatchEvent(new Event('change', { bubbles: true }));
            cb.dispatchEvent(new Event('click', { bubbles: true }));
          }
        }, checkbox);

        await targetPage.evaluate((el) => {
          if (el) el.click();
        }, extensionElement);

        console.log("확장매체 체크박스 체크 성공");
      } else {
        console.log("확장매체가 이미 체크되어 있습니다.");
      }

      await page.waitForTimeout(1000);
    } else {
      console.log("확장매체 요소 또는 체크박스를 찾을 수 없습니다.");
    }
  } catch (error) {
    console.log("확장매체 버튼 클릭 실패: " + error.message);
  }
}

async function selectCertificate(targetPage, page, certOwnerName) {
  try {
    console.log("인증서 리스트 업데이트 대기 중...");
    await page.waitForTimeout(2000);

    let selectedCert = null;
    const certSelectors = [
      "#grid_certificateInfos_0_rec_0",
      "tr[id^='grid_certificateInfos_0_rec_']",
    ];

    try {
      await targetPage.waitForSelector("#grid_certificateInfos_0_rec_0, tr[id^='grid_certificateInfos_0_rec_']", { timeout: 10000 });
      const certRows = await targetPage.$$("tr[id^='grid_certificateInfos_0_rec_']");

      for (const row of certRows) {
        const text = await targetPage.evaluate((el) => {
          return el ? el.textContent || el.innerText : '';
        }, row);

        if (text.includes(certOwnerName)) {
          selectedCert = row;
          console.log(`✅ 인증서 리스트에서 '${certOwnerName}' 발견!`);
          break;
        }
      }

      if (!selectedCert && certRows.length > 0) {
        selectedCert = certRows[0];
        console.log(`⚠️ '${certOwnerName}'을 찾지 못했습니다. 첫 번째 인증서를 사용합니다.`);
      }
    } catch (e) {
      console.log(`인증서 리스트 찾기 실패: ${e.message}`);
    }

    if (selectedCert) {
      const isSelected = await targetPage.evaluate((el) => {
        return el && el.classList.contains('w2ui-selected');
      }, selectedCert);

      if (!isSelected) {
        await targetPage.evaluate((el) => {
          if (el) el.click();
        }, selectedCert);
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
}

async function handlePasswordInput(targetPage, page, password) {
  let passwordInputSuccess = false;
  let manualPopupHandled = false;

  try {
    let passwordInput = null;
    const passwordSelector = "input[name='selectDialogPasswordInput']";

    try {
      await targetPage.waitForSelector(passwordSelector, { timeout: 10000, visible: true });
      passwordInput = await targetPage.$(passwordSelector);
      if (passwordInput) {
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

    if (passwordInput) {
      let waitForPopupInContext = async () => false;
      let waitForPopupGoneInContext = async () => false;
      const handleManualPopupIfAppeared = async (stageLabel) => {
        if (passwordInputSuccess) return true;
        const appearedInTarget = await waitForPopupInContext(targetPage, 2000);
        const appearedInPage = !appearedInTarget && targetPage !== page
          ? await waitForPopupInContext(page, 2000)
          : false;
        if (!(appearedInTarget || appearedInPage)) return false;

        console.log(`⚠️ ${stageLabel} 가상키보드/간편비밀번호 팝업 감지됨. 수동으로 비밀번호 입력 후 팝업을 닫아주세요.`);
        await waitForPopupGoneInContext(targetPage, 180000);
        if (targetPage !== page) {
          await waitForPopupGoneInContext(page, 180000);
        }
        console.log("✅ 수동 입력 팝업이 닫혔습니다. 자동 입력을 계속 진행합니다.");
        manualPopupHandled = true;
        const manualInputValue = await targetPage.evaluate((input) => {
          return input ? input.value : '';
        }, passwordInput);
        if (manualInputValue && manualInputValue.length > 0) {
          passwordInputSuccess = true;
          return true;
        }
        return false;
      };

      try {
        const manualPopupSelectors = [
          "#w2ui-popup_1",
          "#keyboardDialogBody",
          ".easyPassword",
          "#w2ui-lock-transparent"
        ];
        const isVisibleInContext = async (ctx) => {
          try {
            return await ctx.evaluate((selectors) => {
              const isVisible = (el) => {
                if (!el) return false;
                if (el.offsetParent !== null) return true;
                const style = window.getComputedStyle(el);
                return el.getClientRects().length > 0 && style.display !== 'none' && style.visibility !== 'hidden';
              };
              const visible = selectors.some((selector) => {
                const el = document.querySelector(selector);
                if (selector === '#w2ui-lock-transparent') {
                  if (!el) return false;
                  const style = window.getComputedStyle(el);
                  const opacity = parseFloat(style.opacity || '0');
                  return opacity > 0.01;
                }
                return isVisible(el);
              });
              if (!visible) return false;
              const hasPopup = ['#w2ui-popup_1', '#keyboardDialogBody', '.easyPassword']
                .some((selector) => isVisible(document.querySelector(selector)));
              return hasPopup || (() => {
                const lockEl = document.querySelector('#w2ui-lock-transparent');
                if (!lockEl) return false;
                const style = window.getComputedStyle(lockEl);
                const opacity = parseFloat(style.opacity || '0');
                return opacity > 0.01;
              })();
            }, manualPopupSelectors);
          } catch {
            return false;
          }
        };

        waitForPopupInContext = async (ctx, timeoutMs) => {
          try {
            await ctx.waitForFunction(
              (selectors) => {
                const isVisible = (el) => {
                  if (!el) return false;
                  if (el.offsetParent !== null) return true;
                  const style = window.getComputedStyle(el);
                  return el.getClientRects().length > 0 && style.display !== 'none' && style.visibility !== 'hidden';
                };
                const visible = selectors.some((selector) => {
                  const el = document.querySelector(selector);
                  if (selector === '#w2ui-lock-transparent') {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const opacity = parseFloat(style.opacity || '0');
                    return opacity > 0.01;
                  }
                  return isVisible(el);
                });
                if (!visible) return false;
                const hasPopup = ['#w2ui-popup_1', '#keyboardDialogBody', '.easyPassword']
                  .some((selector) => isVisible(document.querySelector(selector)));
                return hasPopup || (() => {
                  const lockEl = document.querySelector('#w2ui-lock-transparent');
                  if (!lockEl) return false;
                  const style = window.getComputedStyle(lockEl);
                  const opacity = parseFloat(style.opacity || '0');
                  return opacity > 0.01;
                })();
              },
              { timeout: timeoutMs },
              manualPopupSelectors
            );
            return true;
          } catch {
            return false;
          }
        };

        waitForPopupGoneInContext = async (ctx, timeoutMs) => {
          try {
            await ctx.waitForFunction(
              (selectors) => {
                const isVisible = (el) => {
                  if (!el) return false;
                  if (el.offsetParent !== null) return true;
                  const style = window.getComputedStyle(el);
                  return el.getClientRects().length > 0 && style.display !== 'none' && style.visibility !== 'hidden';
                };
                const visible = selectors.some((selector) => {
                  const el = document.querySelector(selector);
                  if (selector === '#w2ui-lock-transparent') {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const opacity = parseFloat(style.opacity || '0');
                    return opacity > 0.01;
                  }
                  return isVisible(el);
                });
                if (!visible) return true;
                const hasPopup = ['#w2ui-popup_1', '#keyboardDialogBody', '.easyPassword']
                  .some((selector) => isVisible(document.querySelector(selector)));
                if (hasPopup) return false;
                const lockEl = document.querySelector('#w2ui-lock-transparent');
                if (!lockEl) return true;
                const style = window.getComputedStyle(lockEl);
                const opacity = parseFloat(style.opacity || '0');
                return opacity <= 0.01;
              },
              { timeout: timeoutMs },
              manualPopupSelectors
            );
            return true;
          } catch {
            return false;
          }
        };

        let isManualPopupVisible = await isVisibleInContext(targetPage);
        if (!isManualPopupVisible && targetPage !== page) {
          isManualPopupVisible = await isVisibleInContext(page);
        }

        if (!isManualPopupVisible) {
          const appearedInTarget = await waitForPopupInContext(targetPage, 2000);
          const appearedInPage = !appearedInTarget && targetPage !== page
            ? await waitForPopupInContext(page, 2000)
            : false;
          isManualPopupVisible = appearedInTarget || appearedInPage;
        }

        if (isManualPopupVisible) {
          console.log("⚠️ 가상키보드/간편비밀번호 팝업 감지됨. 수동으로 비밀번호 입력 후 팝업을 닫아주세요.");
          await waitForPopupGoneInContext(targetPage, 180000);
          if (targetPage !== page) {
            await waitForPopupGoneInContext(page, 180000);
          }
          console.log("✅ 수동 입력 팝업이 닫혔습니다. 자동 입력을 계속 진행합니다.");
          manualPopupHandled = true;
          const manualInputValue = await targetPage.evaluate((input) => {
            return input ? input.value : '';
          }, passwordInput);
          if (manualInputValue && manualInputValue.length > 0) {
            passwordInputSuccess = true;
          }
        }
      } catch (error) {
        console.log(`⚠️ 수동 입력 팝업 감지/대기 중 오류: ${error.message}`);
      }

      console.log(`비밀번호 입력 시도: ${password.replace(/./g, '*')}`);

      await handleManualPopupIfAppeared("입력 직후");

      if (passwordInputSuccess) {
        console.log("✅ 수동 입력 결과가 감지되어 자동 입력을 건너뜁니다.");
      }

      const methods = passwordInputSuccess ? [] : [
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
            if (await handleManualPopupIfAppeared("방법 1/5 입력 직후")) return;
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
            if (await handleManualPopupIfAppeared("방법 2/5 입력 직후")) return;
            await page.waitForTimeout(500);
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
            if (await handleManualPopupIfAppeared("방법 3/5 입력 직후")) return;
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
            if (await handleManualPopupIfAppeared("방법 4/5 입력 직후")) return;
            await page.waitForTimeout(500);
          }
        },
        {
          name: "방법 5/5: 키보드 직접 입력 (필드 없이)",
          execute: async () => {
            await targetPage.evaluate((input) => {
              if (input) {
                input.focus();
                input.click();
              }
            }, passwordInput);
            if (await handleManualPopupIfAppeared("방법 5/5 입력 직후")) return;
            await page.waitForTimeout(300);
            for (const char of password) {
              await page.keyboard.type(char);
              await page.waitForTimeout(70);
            }
            await page.waitForTimeout(500);
          }
        }
      ];

      for (let i = 0; i < methods.length; i++) {
        const method = methods[i];
        const methodNum = i + 1;
        console.log(`[비밀번호] 방법 ${methodNum}/${methods.length} 시도 중...`);

        try {
          await method.execute();

          const inputValue = await targetPage.evaluate((input) => {
            return input ? input.value : '';
          }, passwordInput);

          if (inputValue && inputValue.length >= password.length) {
            console.log(`✅ 비밀번호 입력 성공! (방법 ${methodNum}/${methods.length}, 입력된 길이: ${inputValue.length})`);
            passwordInputSuccess = true;
            break;
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
        if (manualPopupHandled) {
          console.log("✅ 수동 입력으로 로그인 진행됨. 확인 버튼 클릭을 생략합니다.");
        }
      }
    } else {
      console.log("비밀번호 입력 필드를 찾지 못했습니다. 키보드로 직접 입력 시도...");
      for (const char of password) {
        await page.keyboard.type(char);
        await page.waitForTimeout(70);
      }
      console.log("키보드로 비밀번호 입력 성공");
      passwordInputSuccess = true;
    }
  } catch (error) {
    console.log("비밀번호 입력 실패: " + error.message);
    return { success: false, manualPopupHandled };
  }

  return { success: passwordInputSuccess, manualPopupHandled };
}

async function clickOkButton(targetPage, page) {
  try {
    const isPopupVisible = async () => {
      try {
        const checkInContext = async (ctx) => {
          const el = await ctx.$("#delfino-section");
          if (!el) return false;
          return await ctx.evaluate((node) => node && node.offsetParent !== null, el);
        };
        if (typeof targetPage?.isDetached === 'function' && targetPage.isDetached()) {
          return false;
        }
        if (await checkInContext(targetPage)) return true;
        if (targetPage !== page && await checkInContext(page)) return true;
        return false;
      } catch (e) {
        if (e.message.includes('detached Frame')) {
          return false;
        }
        return false;
      }
    };

    if (typeof targetPage?.isDetached === 'function' && targetPage.isDetached()) {
      console.log("⚠️ 인증서 iframe이 분리되었습니다. 메인 페이지로 진행합니다.");
      targetPage = page;
    }

    const popupVisibleBeforeClick = await isPopupVisible();
    if (!popupVisibleBeforeClick) {
      console.log("✅ 인증서 팝업이 이미 닫혀있습니다. 확인 버튼 클릭을 생략합니다.");
      return true;
    }

    let okButton = null;
    const okButtonSelectors = [
      "#delfino-section .okButton.okButtonMsg",
      "button.okButton.okButtonMsg",
      ".okButtonBlock button.okButton.okButtonMsg"
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
      const buttonInfo = await targetPage.evaluate((button) => {
        return { disabled: button.disabled };
      }, okButton);

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

      const checkPopupDisappeared = async () => {
        try {
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
              if (e.message.includes('detached Frame')) {
                return true;
              }
            }
          }

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
          if (e.message.includes('detached Frame')) {
            return true;
          }
          return false;
        }
      };

      console.log("⏳ 확인 버튼 클릭 전 1초 대기...");
      await page.waitForTimeout(1000);
      console.log("🖱️ 확인 버튼 클릭 시도");

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

      for (let i = 0; i < clickMethods.length; i++) {
        const method = clickMethods[i];
        const methodNum = i + 1;

        console.log(`[확인버튼] ${method.name} 시도 중...`);

        try {
          await method.execute();
          console.log(`✅ ${method.name} 완료`);

          const disappeared = await checkPopupDisappeared();

          if (disappeared) {
            console.log(`✅ 확인 버튼 클릭 성공! (${method.name})`);
            clickSuccess = true;
            break;
          } else {
            console.log(`⚠️ ${method.name} 시도했지만 팝업이 아직 존재합니다.`);
          }
        } catch (e) {
          console.log(`❌ ${method.name} 오류: ${e.message}`);

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
      const popupVisible = await isPopupVisible();
      if (!popupVisible) {
        console.log("✅ 인증서 팝업이 이미 닫혀있습니다. 확인 버튼 클릭을 생략합니다.");
        return true;
      }
      console.log("확인 버튼을 찾을 수 없습니다.");
      return false;
    }
  } catch (error) {
    console.log("확인 버튼 클릭 실패: " + error.message);
    return false;
  }
}

export {
  findPopupFrame,
  waitForPopupContainer,
  clickLocalDisk,
  selectCertStore,
  selectCertificate,
  handlePasswordInput,
  clickOkButton
};
