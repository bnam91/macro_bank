import puppeteer from 'puppeteer';
import fs from 'fs';
import { promises as fsPromises } from 'fs';
import path from 'path';
import readline from 'readline';
import os from 'os';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { executeTransferProcess } from './modules/transfer-module.js';
import { setReadlineInterface } from './modules/user-input-module.js';
import config from './config/config.js';
import { main as addSheetNames } from './scripts/add-sheet-names.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// npm run dev 로 실행 시 true (크롬 창 유지 모드)
const isDevMode = process.env.npm_lifecycle_event === 'dev';

// readline 인터페이스 생성 (단일 인스턴스)
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// user-input-module에 readline 인터페이스 전달
setReadlineInterface(rl);

// 사용자 입력을 Promise로 변환하는 헬퍼 함수
function question(prompt) {
  return new Promise((resolve) => {
    rl.question(prompt, resolve);
  });
}

// 프로필 이름에 google_ 접두사 추가 (없으면 추가)
function addGooglePrefix(profileName) {
  if (!profileName) return profileName;
  if (profileName.startsWith('google_')) {
    return profileName;
  }
  return `google_${profileName}`;
}

// 프로필 이름에서 google_ 접두사 제거 (표시용)
function removeGooglePrefix(profileName) {
  if (!profileName) return profileName;
  if (profileName.startsWith('google_')) {
    return profileName.substring(7); // 'google_'.length = 7
  }
  return profileName;
}

// config.js에서 경로 읽기
function readPathFromFile() {
  return config.userDataPath;
}

// 사용자에게 시트 선택하도록 함
async function selectSheet() {
  const sheets = config.sheets;
  
  if (sheets.length === 0) {
    console.log("\n사용 가능한 시트가 없습니다.");
    return null;
  }
  
  console.log("\n사용 가능한 시트 목록:");
  sheets.forEach((sheet, idx) => {
    console.log(`${idx + 1}. ${sheet.name}`);
  });
  
  while (true) {
    try {
      const choiceStr = await question("\n사용할 시트 번호를 선택하세요: ");
      const choice = parseInt(choiceStr);
      
      if (1 <= choice && choice <= sheets.length) {
        const selectedSheet = sheets[choice - 1];
        console.log(`\n선택된 시트: ${selectedSheet.name}`);
        return {
          sheetUrl: selectedSheet.url,
          sheetName: selectedSheet.sheetName,
          columnMapping: selectedSheet.columnMapping || {
            productName: 4,
            customerName: 5,
            accountInfo: 8,
            amount: 10
          }
        };
      } else {
        console.log("유효하지 않은 번호입니다. 다시 선택해주세요.");
      }
    } catch (e) {
      console.log("숫자를 입력해주세요.");
    }
  }
}

// 사용 가능한 프로필 목록을 가져옴
async function getAvailableProfiles(userDataParent) {
  const profiles = [];
  
  try {
    await fsPromises.access(userDataParent);
  } catch {
    await fsPromises.mkdir(userDataParent, { recursive: true });
    return profiles;
  }
  
  try {
    const items = await fsPromises.readdir(userDataParent);
    for (const item of items) {
      const itemPath = path.join(userDataParent, item);
      try {
        const stats = await fsPromises.stat(itemPath);
        if (stats.isDirectory()) {
          const defaultPath = path.join(itemPath, 'Default');
          let hasDefault = false;
          try {
            await fsPromises.access(defaultPath);
            hasDefault = true;
          } catch {}
          
          let hasProfile = false;
          if (!hasDefault) {
            const subItems = await fsPromises.readdir(itemPath);
            for (const subItem of subItems) {
              const subItemPath = path.join(itemPath, subItem);
              try {
                const subStats = await fsPromises.stat(subItemPath);
                if (subStats.isDirectory() && subItem.startsWith('Profile')) {
                  hasProfile = true;
                  break;
                }
              } catch {}
            }
          }
          
          if (hasDefault || hasProfile) {
            // google_로 시작하는 프로필만 추가
            if (item.startsWith('google_')) {
              profiles.push(item);
            }
          }
        }
      } catch {}
    }
  } catch (e) {
    console.log(`프로필 목록 읽기 중 오류: ${e.message}`);
  }
  
  return profiles;
}

// 사용자에게 프로필을 선택하도록 함
async function selectProfile(userDataParent) {
  const profiles = await getAvailableProfiles(userDataParent);
  
  if (profiles.length === 0) {
    console.log("\n사용 가능한 프로필이 없습니다.");
    const createNew = (await question("새 프로필을 생성하시겠습니까? (y/n): ")).toLowerCase();
    if (createNew === 'y') {
      while (true) {
        const name = await question("새 프로필 이름을 입력하세요: ");
        if (!name) {
          console.log("프로필 이름을 입력해주세요.");
          continue;
        }
        
        if (/[\\/:*?"<>|]/.test(name)) {
          console.log("프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? \" < > |");
          continue;
        }
        
        // google_ 접두사 추가
        const profileNameWithPrefix = addGooglePrefix(name);
        const newProfilePath = path.join(userDataParent, profileNameWithPrefix);
        
        // 접두사가 추가된 이름으로 프로필 존재 여부 확인
        try {
          await fsPromises.access(newProfilePath);
          console.log(`'${profileNameWithPrefix}' 프로필이 이미 존재합니다.`);
          continue;
        } catch {}
        
        try {
          await fsPromises.mkdir(newProfilePath, { recursive: true });
          await fsPromises.mkdir(path.join(newProfilePath, 'Default'), { recursive: true });
          console.log(`'${profileNameWithPrefix}' 프로필이 생성되었습니다.`);
          return profileNameWithPrefix;
        } catch (e) {
          console.log(`프로필 생성 중 오류가 발생했습니다: ${e.message}`);
          const retry = (await question("다시 시도하시겠습니까? (y/n): ")).toLowerCase();
          if (retry !== 'y') {
            return null;
          }
        }
      }
    }
    return null;
  }
  
  console.log("\n사용 가능한 프로필 목록:");
  profiles.forEach((profile, idx) => {
    // 표시할 때는 google_ 접두사 제거
    const displayName = removeGooglePrefix(profile);
    console.log(`${idx + 1}. ${displayName}`);
  });
  console.log(`${profiles.length + 1}. 새 프로필 생성`);
  
  while (true) {
    try {
      const choiceStr = await question("\n사용할 프로필 번호를 선택하세요: ");
      const choice = parseInt(choiceStr);
      
      if (1 <= choice && choice <= profiles.length) {
        const selectedProfile = profiles[choice - 1];
        const displayName = removeGooglePrefix(selectedProfile);
        console.log(`\n선택된 프로필: ${displayName}`);
        return selectedProfile; // 실제 프로필 이름(접두사 포함) 반환
      } else if (choice === profiles.length + 1) {
        // 새 프로필 생성
        while (true) {
          const name = await question("새 프로필 이름을 입력하세요: ");
          if (!name) {
            console.log("프로필 이름을 입력해주세요.");
            continue;
          }
          
          if (/[\\/:*?"<>|]/.test(name)) {
            console.log("프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? \" < > |");
            continue;
          }
          
          // google_ 접두사 추가
          const profileNameWithPrefix = addGooglePrefix(name);
          const newProfilePath = path.join(userDataParent, profileNameWithPrefix);
          
          // 접두사가 추가된 이름으로 다시 확인
          try {
            await fsPromises.access(newProfilePath);
            console.log(`'${profileNameWithPrefix}' 프로필이 이미 존재합니다.`);
            continue;
          } catch {}
          
          try {
            await fsPromises.mkdir(newProfilePath, { recursive: true });
            await fsPromises.mkdir(path.join(newProfilePath, 'Default'), { recursive: true });
            console.log(`'${profileNameWithPrefix}' 프로필이 생성되었습니다.`);
            return profileNameWithPrefix;
          } catch (e) {
            console.log(`프로필 생성 중 오류가 발생했습니다: ${e.message}`);
            const retry = (await question("다시 시도하시겠습니까? (y/n): ")).toLowerCase();
            if (retry !== 'y') {
              break;
            }
          }
        }
      } else {
        console.log("유효하지 않은 번호입니다. 다시 선택해주세요.");
      }
    } catch (e) {
      console.log("숫자를 입력해주세요.");
    }
  }
}

let browserInstance = null;

async function openCoupang() {
  let browser;
  
  try {
    // 입금요청 현황판 열기 여부 확인
    const openStatusBoard = (await question("입금요청 현황판도 함께 열까요? (y/n): ")).toLowerCase().trim();
    const statusBoardUrl = 'https://docs.google.com/spreadsheets/d/1CK2UXTy7HKjBe2T0ovm5hfzAAKZxZAR_ev3cbTPOMPs/edit?gid=1565864271#gid=1565864271';
    const depositRequestSheetUrl = 'https://docs.google.com/spreadsheets/d/1NOP5_s0gNUCWaGIgMo5WZmtqBbok_5a4XdpNVwu8n5c/edit?gid=0#gid=0';
    
    // 시트명 추가 스크립트 먼저 실행
    console.log("📋 시트명 추가 스크립트 실행 중...\n");
    try {
      await addSheetNames();
      console.log("\n✅ 시트명 추가 완료\n");
    } catch (error) {
      console.error(`⚠️ 시트명 추가 중 오류 발생: ${error.message}`);
      console.log("계속 진행합니다...\n");
    }
    
    // 사용자 프로필 경로 설정 (config.txt에서 읽기)
    const userDataParent = readPathFromFile();
    
    // 프로필 선택 (하드코딩: 첫 번째 프로필 자동 선택)
    const profiles = await getAvailableProfiles(userDataParent);
    if (profiles.length === 0) {
      console.log("사용 가능한 프로필이 없습니다. 프로그램을 종료합니다.");
      rl.close();
      return;
    }
    const selectedProfile = profiles[0]; // 첫 번째 프로필 자동 선택
    const displayName = removeGooglePrefix(selectedProfile);
    console.log(`\n선택된 프로필: ${displayName} (자동 선택)`);
    
    const userDataDir = path.join(userDataParent, selectedProfile);
    
    // 프로필 디렉토리가 없으면 생성
    try {
      await fsPromises.access(userDataDir);
    } catch {
      await fsPromises.mkdir(userDataDir, { recursive: true });
      await fsPromises.mkdir(path.join(userDataDir, 'Default'), { recursive: true });
    }
    
    // Chrome 경로
    const chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
    
    // 브라우저 실행 옵션
    const options = {
      headless: false,
      defaultViewport: null,
      userDataDir: userDataDir,
      args: [
        '--window-size=1440,1080', // 가로 1440px, 세로는 충분히 큰 값
        '--no-sandbox',
        '--disable-blink-features=AutomationControlled',
        // 캐시 크기 제한 (100MB로 제한)
        '--disk-cache-size=104857600',
        // 메모리 캐시 크기 제한 (50MB로 제한)
        '--media-cache-size=52428800',
        // 백그라운드 네트워킹 비활성화 (불필요한 데이터 저장 방지)
        '--disable-background-networking',
        // 서비스 워커 비활성화 (캐시 누적 방지)
        '--disable-background-timer-throttling',
      ],
      ignoreHTTPSErrors: true,
    };
    
    // Chrome이 있으면 사용
    if (fs.existsSync(chromePath)) {
      options.executablePath = chromePath;
    }

    browser = await puppeteer.launch(options);
    browserInstance = browser;
    console.log('✅ 크롬이 열렸습니다. 종료하려면 Ctrl+C를 누르세요.\n');
    if (isDevMode) {
      console.log('🔧 dev 모드: Ctrl+C 시 크롬 창은 유지됩니다.\n');
    }

    // 첫 번째 페이지 사용
    const pages = await browser.pages();
    const page = pages[0];
  
    // 창 크기 설정 (세로는 최대화)
    await page.goto('about:blank'); // 먼저 페이지를 로드해야 screen 정보 접근 가능
    const screenHeight = await page.evaluate(() => window.screen.availHeight);
    await page.setViewport({ width: 1440, height: screenHeight || 1080 });

    // 구글로 이동
    await page.goto('https://www.google.com');
    
    // 입금요청 현황판·입금요청내역 열기 (사용자가 요청한 경우)
    if (openStatusBoard === 'y') {
      const depositRequestPage = await browser.newPage();
      await depositRequestPage.goto(depositRequestSheetUrl);
      const statusBoardPage = await browser.newPage();
      await statusBoardPage.goto(statusBoardUrl);
      console.log('입금요청내역 시트와 입금요청 현황판을 열었습니다.\n');
      
      // 이체프로세스 진행 여부 확인
      const proceedTransfer = (await question("이체프로세스를 진행할까요? (y/n): ")).toLowerCase().trim();
      if (proceedTransfer !== 'y') {
        console.log('이체프로세스를 건너뜁니다. 브라우저는 열려있습니다.\n');
        // 브라우저 종료 감지 (dev 모드가 아니면 프로세스 종료)
        browser.on('disconnected', () => {
          console.log('브라우저가 닫혔습니다.');
          if (!isDevMode) process.exit(0);
        });
        // 무한 대기
        await new Promise(() => {});
        return;
      }
    }

    // 새 탭 열기 - 한은 로그인 페이지
    const newPage = await browser.newPage();
    await newPage.goto('https://www.kebhana.com/common/login.do');
    console.log('한은 로그인 페이지로 이동했습니다.\n');

    // 다계좌이체진행 자동 처리 (개발 중이므로 n으로 설정)
    const autoTransfer = false;
    console.log("🟠다계좌이체진행(자동): 자동으로 n으로 처리합니다. (개발 중)");

    // 구글 시트 선택 (config.js의 defaultSheetName 사용)
    const sheets = config.sheets;
    let selectedSheet = sheets.find(sheet => sheet.name === config.defaultSheetName);
    
    if (!selectedSheet) {
      // 기본 시트가 없으면 첫 번째 시트 선택
      if (sheets.length === 0) {
        console.log("사용 가능한 시트가 없습니다. 프로그램을 종료합니다.");
        rl.close();
        return;
      }
      selectedSheet = sheets[0];
      console.log(`⚠️ 기본 시트 '${config.defaultSheetName}'를 찾을 수 없어 첫 번째 시트를 선택했습니다.`);
    }
    
    const sheetConfig = {
      sheetUrl: selectedSheet.url,
      sheetName: selectedSheet.sheetName,
      columnMapping: selectedSheet.columnMapping || {
        productName: 4,
        customerName: 5,
        accountInfo: 8,
        amount: 10
      }
    };
    
    console.log(`\n선택된 시트: ${selectedSheet.name} (자동 선택)`);
    
    const { sheetUrl, sheetName, columnMapping } = sheetConfig;
    console.log(`시트 URL: ${sheetUrl}`);
    console.log(`시트명: ${sheetName}\n`);

    // 이체 프로세스 실행
    await executeTransferProcess(
      newPage,
      { sheetUrl, sheetName, columnMapping },
      autoTransfer,
    );

    // 브라우저 종료 감지 (dev 모드가 아니면 프로세스 종료)
    browser.on('disconnected', () => {
      console.log('브라우저가 닫혔습니다.');
      if (!isDevMode) process.exit(0);
    });

    // 무한 대기 (프로세스가 계속 실행되어 크롬 창 유지)
    await new Promise(() => {});

  } catch (error) {
    console.error('오류:', error.message);
    rl.close();
    process.exit(1);
  } finally {
    rl.close();
  }
}

// Ctrl+C 종료 처리
process.on('SIGINT', async () => {
  console.log('\n종료 중...');
  rl.close();
  // dev 모드: 크롬만 연결 해제하고 프로세스 종료 → 크롬 창은 유지
  if (isDevMode && browserInstance) {
    try {
      browserInstance.disconnect();
      console.log('크롬 창은 유지됩니다.');
    } catch (_) {}
  }
  process.exit(0);
});

openCoupang();
