import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

async function loadCertStoreConfig() {
  try {
    // ES 모듈에서는 동적 import 사용 (상대 경로)
    const configModule = await import('../../config/cert-store-config.js');
    const config = configModule.default || configModule;
    if (!config || (typeof config === 'object' && Object.keys(config).length === 0)) {
      console.log("⚠️ cert-store-config.js 파일이 비어있어 기본 선택자를 사용합니다.");
      return null;
    }
    return config;
  } catch (error) {
    console.log(`⚠️ cert-store-config.js 로드 실패: ${error.message}`);
    return null;
  }
}

function getCertStoreKeywords(config) {
  if (!config) return [];
  const keywords = [];
  const add = (value) => {
    if (!value) return;
    if (Array.isArray(value)) {
      value.forEach((item) => {
        if (typeof item === 'string' && item.trim()) {
          keywords.push(item.trim());
        }
      });
      return;
    }
    if (typeof value === 'string' && value.trim()) {
      keywords.push(value.trim());
    }
  };

  add(config.certStoreKeywords);
  add(config.certStoreLabels);
  add(config.certStoreLabel);

  return keywords;
}

export {
  loadCertStoreConfig,
  getCertStoreKeywords
};
