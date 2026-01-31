// 은행 선택을 위한 맵핑
const bankOptions = {
  "하나은행": "081",
  "경남은행": "039",
  "광주은행": "034",
  "국민은행": "004",
  "기업은행": "003",
  "농협": "011",
  "iM뱅크(대구)": "031",
  "도이치뱅크": "055",
  "부산은행": "032",
  "산업은행": "002",
  "저축은행": "050",
  "새마을금고": "045",
  "수협은행": "007",
  "신협": "048",
  "신한은행": "088",
  "우리은행": "020",
  "우체국": "071",
  "전북은행": "037",
  "제주은행": "035",
  "카카오뱅크": "090",
  "케이뱅크": "089",
  "한국씨티은행": "027",
  "BOA": "060",
  "HSBC": "054",
  "JP모간": "057",
  "SC제일은행": "023",
  "하나증권": "270",
  "교보증권": "261",
  "대신증권": "267",
  "미래에셋증권": "238",
  "DB금융투자": "279",
  "유안타증권": "209",
  "메리츠증권": "287",
  "부국증권": "290",
  "삼성증권": "240",
  "신영증권": "291",
  "신한투자증권": "278",
  "NH투자증권": "247",
  "유진증권": "280",
  "키움증권": "264",
  "하이투자증권": "262",
  "한국투자": "243",
  "한화투자증권": "269",
  "KB증권": "218",
  "LS증권": "265",
  "현대차증권": "263",
  "케이프증권": "292",
  "SK증권": "266",
  "산림조합": "064",
  "중국공상은행": "062",
  "중국은행": "063",
  "중국건설은행": "067",
  "BNP파리바은행": "061",
  "한국포스증권": "294",
  "다올투자증권": "227",
  "BNK투자증권": "224",
  "카카오페이증권": "288",
  "IBK투자증권": "225",
  "토스증권": "271",
  "토스뱅크": "092",
  "상상인증권": "221"
};

// 은행명 표준화 함수
function standardizeBankName(bankName) {
  const name = bankName.toLowerCase().trim();
  if (name.includes("sc제일") || name.includes("제일")) {
    return "SC제일은행";
  } else if (name.includes("제일은행")) {
    return "SC제일은행";
  } else if (name.includes("하나")) {
    return "하나은행";
  } else if (name.includes("경남")) {
    return "경남은행";
  } else if (name.includes("광주")) {
    return "광주은행";
  } else if (name.includes("국민")) {
    return "국민은행";
  } else if (name.includes("기업")) {
    return "기업은행";
  } else if (name.includes("농협은행") || name.includes("nh농협") || name.includes("농협/")) {
    return "농협";
  } else if (name.includes("대구") || name.includes("im뱅크") || name.includes("대구은행")) {
    return "iM뱅크(대구)";
  } else if (name.includes("부산")) {
    return "부산은행";
  } else if (name.includes("새마을")) {
    return "새마을금고";
  } else if (name.includes("수협")) {
    return "수협은행";
  } else if (name.includes("신한")) {
    return "신한은행";
  } else if (name.includes("우리")) {
    return "우리은행";
  } else if (name.includes("전북")) {
    return "전북은행";
  } else if (name.includes("제주")) {
    return "제주은행";
  } else if (name.includes("카카오") || name.includes("카뱅")) {
    return "카카오뱅크";
  } else if (name.includes("씨티")) {
    return "한국씨티은행";
  } else if (name.includes("토스")) {
    return "토스뱅크";
  } else if (name.includes("카카오페이증권")) {
    return "카카오페이증권";
  } else if (name.includes("미래에셋대우") || name.includes("미래에셋")) {
    return "미래에셋증권";
  }
  return bankName;
}

// 계좌정보 전처리 함수
function preprocessAccountInfo(accountInfo) {
  const match = accountInfo.match(/(\D+)\s*([\d\s\-]+)/);
  if (match) {
    const bankName = standardizeBankName(match[1].trim());
    const accountNumber = match[2].replace(/[\s\-]/g, '');
    return { bankName, accountNumber };
  }
  return { bankName: "", accountNumber: "" };
}

export {
  bankOptions,
  standardizeBankName,
  preprocessAccountInfo
};

