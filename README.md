# 크롬 자동 실행 스크립트

Node.js와 Puppeteer를 사용하여 크롬 창을 여는 스크립트입니다.

## 설치 방법

```bash
npm install
```

## 실행 방법

### 프로필 모드 (권장)
```bash
npm run name
```

또는

```bash
node index-profile.js
```

프로필 모드에서는:
- 여러 프로필을 생성하고 선택할 수 있습니다
- 각 프로필의 로그인 상태가 독립적으로 유지됩니다
- 프로필은 `/Users/a1/Documents/github/user_data` 폴더에 저장됩니다
- 프로필 이름은 자동으로 `google_` 접두사가 추가됩니다

### 일반 모드
```bash
npm start
```

또는

```bash
node index.js
```

### 시크릿 모드
```bash
npm run start:secret
```

또는

```bash
node index-secret.js
```

## 기능

- 구글 사이트를 크롬 창으로 자동으로 엽니다
- **프로필 모드** (`index-profile.js`): 여러 프로필을 관리하고 선택할 수 있으며, 각 프로필의 로그인 상태가 유지됩니다
- **일반 모드** (`index.js`): 캐시와 쿠키를 유지합니다
- **시크릿 모드** (`index-secret.js`): 실행할 때마다 모든 쿠키와 캐시를 자동으로 삭제하며, 완전히 분리된 브라우징 환경에서 실행됩니다
- 크롬 창이 최대화된 상태로 열립니다
- 브라우저를 닫으려면 Ctrl+C를 누르세요

## 프로필 사용법

1. `npm run name` 실행
2. 프로필 목록에서 번호를 선택하거나 새 프로필을 생성합니다
3. 새 프로필 생성 시 이메일 주소만 입력하면 자동으로 `google_` 접두사가 추가됩니다
   - 예: `coq3820@gmail.com` 입력 → `google_coq3820@gmail.com`으로 저장
4. 선택한 프로필로 크롬이 실행되며, 해당 프로필의 로그인 상태가 유지됩니다

