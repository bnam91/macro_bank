import { loadCertStoreConfig, getCertStoreKeywords } from './config.js';
import {
  findPopupFrame,
  waitForPopupContainer,
  clickLocalDisk,
  selectCertStore,
  selectCertificate,
  handlePasswordInput,
  clickOkButton
} from './steps.js';

async function handleCertPopup(page) {
  try {
    console.log("공인인증서 팝업 대기 중...");

    const { targetPage } = await findPopupFrame(page);
    await waitForPopupContainer(targetPage);

    console.log("\n[1/5] 로컬디스크 선택");
    await clickLocalDisk(targetPage, page);

    console.log("\n[2/5] 확장매체(저장매체) 선택");
    const certStoreConfig = await loadCertStoreConfig();
    const certStoreKeywords = getCertStoreKeywords(certStoreConfig);
    await selectCertStore(targetPage, page, certStoreKeywords);

    console.log("\n[3/5] 인증서 선택");
    await selectCertificate(targetPage, page, "신현빈");

    console.log("\n[4/5] 비밀번호 입력");
    const password = certStoreConfig?.certPassword || "@gusqls120";
    const { success, manualPopupHandled } = await handlePasswordInput(targetPage, page, password);
    if (!success) {
      console.log("❌ 비밀번호 입력이 실패하여 프로세스를 중단합니다.");
      return false;
    }
    if (manualPopupHandled) {
      return true;
    }

    console.log("\n[5/5] 확인 버튼 클릭");
    return await clickOkButton(targetPage, page);
  } catch (error) {
    console.log(`공인인증서 팝업 처리 중 오류: ${error.message}`);
    return false;
  }
}

export {
  handleCertPopup
};
