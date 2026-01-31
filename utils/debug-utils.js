import fs from 'fs';

/**
 * 페이지의 HTML 소스와 프레임 정보를 파일로 저장하는 디버깅 유틸리티
 * @param {Page} page - Puppeteer 페이지 객체
 * @param {string} filename - 저장할 파일명
 */
async function debugPage(page, filename = 'page-debug.html') {
  try {
    // 페이지의 HTML 소스 가져오기
    const content = await page.content();
    
    // 파일로 저장
    fs.writeFileSync(filename, content, 'utf-8');
    
    // 프레임 정보 출력 (frames() 메서드가 있는 경우에만)
    try {
      if (typeof page.frames === 'function') {
        const frames = page.frames();
        console.log(`디버깅 정보 저장 완료: ${filename}`);
        console.log(`발견된 프레임 수: ${frames.length}`);
        
        frames.forEach((frame, index) => {
          console.log(`  프레임 ${index}: ${frame.name() || '(이름 없음)'} - ${frame.url()}`);
        });
      } else {
        console.log(`디버깅 정보 저장 완료: ${filename} (iframe이므로 프레임 정보 없음)`);
      }
    } catch (frameError) {
      // frames() 호출 실패해도 HTML 저장은 성공
      console.log(`디버깅 정보 저장 완료: ${filename} (프레임 정보 조회 실패)`);
    }
    
    return true;
  } catch (error) {
    console.error(`디버깅 정보 저장 중 오류: ${error.message}`);
    return false;
  }
}

export {
  debugPage
};

