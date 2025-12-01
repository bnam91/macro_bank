const puppeteer = require('puppeteer');
const fs = require('fs');
const fsPromises = require('fs').promises;
const path = require('path');
const readline = require('readline');
const transferModule = require('./transfer-module');
const { setReadlineInterface } = require('./user-input-module');

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

// config.txt 파일에서 경로 읽기
function readPathFromFile() {
  const configPath = path.join(__dirname, 'config.txt');
  
  try {
    if (!fs.existsSync(configPath)) {
      console.error(`\n❌ config.txt 파일을 찾을 수 없습니다.`);
      console.error(`프로젝트 루트에 config.txt 파일을 생성하고 경로를 입력해주세요.`);
      console.error(`예시 (Windows): C:\\Users\\신현빈\\Desktop\\github\\user_data`);
      console.error(`예시 (Mac): /Users/a1/Documents/github/user_data\n`);
      process.exit(1);
    }
    
    const content = fs.readFileSync(configPath, 'utf-8');
    const pathValue = content.trim();
    
    if (!pathValue) {
      console.error(`\n❌ config.txt 파일이 비어있습니다.`);
      console.error(`경로를 입력해주세요.\n`);
      process.exit(1);
    }
    
    return pathValue;
  } catch (error) {
    console.error(`\n❌ config.txt 파일 읽기 중 오류: ${error.message}\n`);
    process.exit(1);
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

async function openCoupang() {
  let browser;
  
  try {
    // 사용자 프로필 경로 설정 (config.txt에서 읽기)
    const userDataParent = readPathFromFile();
    
    // 프로필 선택
    const selectedProfile = await selectProfile(userDataParent);
    if (!selectedProfile) {
      console.log("프로필을 선택할 수 없습니다. 프로그램을 종료합니다.");
      rl.close();
      return;
    }
    
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
        '--start-maximized',
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
    console.log('✅ 크롬이 열렸습니다. 종료하려면 Ctrl+C를 누르세요.\n');

    // 첫 번째 페이지 사용
    const pages = await browser.pages();
    const page = pages[0];

    // 구글로 이동
    await page.goto('https://www.google.com');

    // 새 탭 열기 - 한은 로그인 페이지
    const newPage = await browser.newPage();
    await newPage.goto('https://www.kebhana.com/common/login.do');
    console.log('한은 로그인 페이지로 이동했습니다.\n');

    // 다계좌이체진행 자동 처리 (개발 중이므로 n으로 설정)
    const autoTransfer = false;
    console.log("🟠다계좌이체진행(자동): 자동으로 n으로 처리합니다. (개발 중)");

    // 엑셀 파일 경로 설정
    const excelPath = path.join(__dirname, "이체정보.xlsx");
    
    // 이체 프로세스 실행
    if (fs.existsSync(excelPath)) {
      await transferModule.executeTransferProcess(newPage, excelPath, autoTransfer);
    } else {
      console.log(`엑셀 파일을 찾을 수 없습니다: ${excelPath}`);
      console.log("수동으로 이체를 진행해주세요.");
    }

    // 브라우저 종료 감지
    browser.on('disconnected', () => {
      console.log('브라우저가 닫혔습니다.');
      process.exit(0);
    });

    // 무한 대기
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
  process.exit(0);
});

openCoupang();

