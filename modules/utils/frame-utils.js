// 프레임 전환 및 키보드 입력 유틸리티

const modifierKey = process.platform === 'darwin' ? 'Meta' : 'Control';

// 프레임 전환
export async function switchToFrame(page, frameId) {
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

// 포커스 및 키보드 입력
export async function focusAndTypeWithKeyboard(frame, page, selector, text, options = {}) {
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
