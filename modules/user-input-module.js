import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// readline 인터페이스는 외부에서 전달받음 (중복 생성 방지)
let rl = null;

/**
 * readline 인터페이스 설정 (외부에서 호출)
 * @param {readline.Interface} readlineInterface - readline 인터페이스
 */
function setReadlineInterface(readlineInterface) {
  rl = readlineInterface;
}

/**
 * 사용자에게 다음 프로세스 진행 여부를 묻는 함수
 * @param {string} message - 사용자에게 보여줄 메시지 (기본값: "다음 프로세스를 진행할까요?")
 * @returns {Promise<boolean>} - true: 진행, false: 중단
 */
async function askToContinue(message = "다음 프로세스를 진행할까요? (y/d/n): ") {
  if (!rl) {
    throw new Error('readline 인터페이스가 설정되지 않았습니다. setReadlineInterface를 먼저 호출하세요.');
  }
  
  return new Promise((resolve) => {
    rl.question(message, (answer) => {
      // 첫 글자만 확인하여 중복 입력 방지
      const firstChar = answer.toLowerCase().trim().charAt(0);
      
      if (firstChar === 'y') {
        resolve({ action: 'continue' });
      } else if (firstChar === 'd') {
        resolve({ action: 'debug' });
      } else {
        resolve({ action: 'skip' });
      }
    });
  });
}

/**
 * 페이지의 HTML을 파일로 저장하는 함수
 * @param {Page|Frame} page - Puppeteer 페이지 객체 또는 iframe 프레임 객체
 * @param {string} filename - 저장할 파일명 (기본값: debug-{timestamp}.html)
 * @returns {Promise<boolean>} - 저장 성공 여부
 * @description iframe을 전달하면 iframe 내부의 HTML을 저장합니다.
 */
async function savePageHTML(page, filename = null, fallbackPage = null) {
  try {
    if (!filename) {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
      filename = `debug-${timestamp}.html`;
    }
    
    // 파일 경로를 프로젝트 루트로 설정
    const filePath = path.join(__dirname, '..', filename);
    
    let content;
    let pageType;
    
    try {
      // 페이지의 HTML 소스 가져오기 (iframe인 경우 iframe 내부의 HTML을 가져옴)
      content = await page.content();
      pageType = page.url ? 'iframe' : 'page';
    } catch (error) {
      // iframe이 detached된 경우 fallback 페이지 사용
      if (error.message.includes('detached Frame') && fallbackPage) {
        console.log(`⚠️ iframe이 detached되어 메인 페이지를 저장합니다.`);
        content = await fallbackPage.content();
        pageType = 'page (fallback)';
      } else {
        throw error;
      }
    }
    
    // 파일로 저장
    fs.writeFileSync(filePath, content, 'utf-8');
    
    console.log(`✅ 디버그 HTML 저장 완료 (${pageType}): ${filename}`);
    return true;
  } catch (error) {
    console.error(`❌ 디버그 HTML 저장 실패: ${error.message}`);
    return false;
  }
}

/**
 * 사용자 입력을 받아서 진행/디버그/중단을 처리하는 함수
 * @param {Page|Frame} page - Puppeteer 페이지 객체 또는 iframe 프레임 객체
 * @param {string} message - 사용자에게 보여줄 메시지
 * @param {Page} fallbackPage - iframe이 detached된 경우 사용할 메인 페이지 객체 (선택사항)
 * @returns {Promise<boolean>} - true: 진행, false: 중단
 */
async function handleUserInput(page, message = "다음 프로세스를 진행할까요? (y/d/n): ", fallbackPage = null) {
  while (true) {
    const result = await askToContinue(message);
    
    if (result.action === 'continue') {
      return true;
    } else if (result.action === 'debug') {
      await savePageHTML(page, null, fallbackPage);
      // 다시 질문
      continue;
    } else {
      console.log("프로세스를 중단합니다.");
      return false;
    }
  }
}

/**
 * 사용자에게 엔터 입력을 기다리는 함수
 * @param {string} message - 사용자에게 보여줄 메시지
 * @returns {Promise<void>}
 */
function waitForEnter(message = "계속하려면 엔터를 누르세요...") {
  if (!rl) {
    throw new Error('readline 인터페이스가 설정되지 않았습니다. setReadlineInterface를 먼저 호출하세요.');
  }
  
  return new Promise((resolve) => {
    rl.question(message, () => {
      resolve();
    });
  });
}

/**
 * readline 인터페이스 종료
 */
function closeInput() {
  if (rl) {
    rl.close();
  }
}

export {
  setReadlineInterface,
  askToContinue,
  savePageHTML,
  handleUserInput,
  waitForEnter,
  closeInput
};

