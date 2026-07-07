// ===== 설정 =====
// 임의의 비밀키 문자열로 바꿔주세요 (영문/숫자 조합 추천). HTML 쪽 설정값과 똑같이 맞춰야 합니다.
const SECRET_KEY = "여기에-나만의-비밀키-입력";

// 독립형(Standalone) Apps Script 프로젝트라 시트에 자동 연결되어 있지 않으므로, 시트 ID로 직접 열어서 사용
const SHEET_ID = "1AMuQ5FtI8fypOkId7WkoWrX98b9o6i1EUIRAEE9jD0A";
const SS = SpreadsheetApp.openById(SHEET_ID);
const STOCKS_SHEET_NAME = "Stocks";
const HISTORY_SHEET_NAME = "History";
const SETTINGS_SHEET_NAME = "Settings";
const PORTFOLIO_HISTORY_SHEET_NAME = "PortfolioHistory";
// V2 (10년 대시보드)용 시트
const SNAPSHOTS_SHEET_NAME = "Snapshots";   // 월별 총자산 기록
const JOURNAL_SHEET_NAME = "Journal";       // 매매 일지

function getOrCreateSheet(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function doGet(e) {
  if (e.parameter.key !== SECRET_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ error: "잘못된 비밀키입니다." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const stocksSheet = getOrCreateSheet(STOCKS_SHEET_NAME, ["id", "ticker", "name", "quantity", "buyPrice"]);
  const historySheet = getOrCreateSheet(HISTORY_SHEET_NAME, ["stockId", "date", "currentPrice", "profitLoss", "profitLossPercent", "promptType", "modelUsed", "aiText"]);
  const settingsSheet = getOrCreateSheet(SETTINGS_SHEET_NAME, ["key", "value"]);

  const stocksRows = stocksSheet.getDataRange().getValues().slice(1);
  const historyRows = historySheet.getDataRange().getValues().slice(1);
  const settingsRows = settingsSheet.getDataRange().getValues().slice(1);

  const stocks = stocksRows
    .filter(row => row[0] !== "")
    .map(row => ({
      id: String(row[0]),
      ticker: String(row[1]),
      name: String(row[2] || ""),
      quantity: Number(row[3]),
      buyPrice: Number(row[4]),
      history: []
    }));

  historyRows.forEach(row => {
    if (row[0] === "") return;
    const stock = stocks.find(s => s.id === String(row[0]));
    if (!stock) return;
    stock.history.push({
      date: String(row[1]),
      currentPrice: Number(row[2]),
      profitLoss: Number(row[3]),
      profitLossPercent: Number(row[4]),
      promptType: String(row[5]),
      modelUsed: String(row[6]),
      aiText: String(row[7])
    });
  });

  const settings = {};
  settingsRows.forEach(row => {
    if (row[0] === "") return;
    settings[String(row[0])] = String(row[1]);
  });

  // 자동분석이 쌓아둔 종합분석 히스토리도 함께 반환 (앱에서 차트/기록으로 표시)
  const phSheet = getOrCreateSheet(PORTFOLIO_HISTORY_SHEET_NAME, ["date", "totalProfitLossPercent", "modelUsed", "aiText"]);
  const phRows = phSheet.getDataRange().getValues().slice(1);
  const portfolioHistory = phRows
    .filter(row => row[0] !== "")
    .map(row => ({
      date: String(row[0]),
      totalProfitLossPercent: Number(row[1]),
      modelUsed: String(row[2]),
      aiText: String(row[3])
    }));

  // V2: 월별 자산 스냅샷
  const snapSheet = getOrCreateSheet(SNAPSHOTS_SHEET_NAME, ["month", "stocksValue", "otherAssets", "totalAssets", "contribution", "note"]);
  const snapshots = snapSheet.getDataRange().getValues().slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      month: String(row[0]),
      stocksValue: Number(row[1]),
      otherAssets: Number(row[2]),
      totalAssets: Number(row[3]),
      contribution: Number(row[4]),
      note: String(row[5] || "")
    }));

  // V2: 매매 일지
  const jSheet = getOrCreateSheet(JOURNAL_SHEET_NAME, ["id", "date", "ticker", "action", "price", "quantity", "reason", "stopLoss", "targetPrice", "emotion"]);
  const journal = jSheet.getDataRange().getValues().slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      id: String(row[0]),
      date: String(row[1]),
      ticker: String(row[2]),
      action: String(row[3]),
      price: Number(row[4]),
      quantity: Number(row[5]),
      reason: String(row[6] || ""),
      stopLoss: String(row[7] || ""),
      targetPrice: String(row[8] || ""),
      emotion: String(row[9] || "")
    }));

  return ContentService.createTextOutput(JSON.stringify({ stocks, settings, portfolioHistory, snapshots, journal }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.key !== SECRET_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ error: "잘못된 비밀키입니다." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "saveStocks") {
    saveStocks(body.stocks || []);
  } else if (body.action === "saveSettings") {
    saveSettings(body.settings || {});
  } else if (body.action === "saveSnapshots") {
    saveSnapshots(body.snapshots || []);
  } else if (body.action === "saveJournal") {
    saveJournal(body.journal || []);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveStocks(stocks) {
  const stocksSheet = getOrCreateSheet(STOCKS_SHEET_NAME, ["id", "ticker", "name", "quantity", "buyPrice"]);
  const historySheet = getOrCreateSheet(HISTORY_SHEET_NAME, ["stockId", "date", "currentPrice", "profitLoss", "profitLossPercent", "promptType", "modelUsed", "aiText"]);

  stocksSheet.clearContents();
  stocksSheet.appendRow(["id", "ticker", "name", "quantity", "buyPrice"]);
  historySheet.clearContents();
  historySheet.appendRow(["stockId", "date", "currentPrice", "profitLoss", "profitLossPercent", "promptType", "modelUsed", "aiText"]);

  // 종목코드가 005930처럼 숫자로만 구성되면 구글시트가 자동으로 숫자로 변환해 앞자리 0을 지워버리므로,
  // 앞에 ' (텍스트 표시 기호)를 붙여 항상 텍스트로 저장되게 한다
  stocks.forEach(stock => {
    stocksSheet.appendRow(["'" + stock.id, "'" + stock.ticker, stock.name || "", stock.quantity, stock.buyPrice]);
    (stock.history || []).forEach(entry => {
      historySheet.appendRow([
        "'" + stock.id, entry.date, entry.currentPrice, entry.profitLoss,
        entry.profitLossPercent, entry.promptType, entry.modelUsed, entry.aiText
      ]);
    });
  });
}

function saveSettings(settings) {
  const settingsSheet = getOrCreateSheet(SETTINGS_SHEET_NAME, ["key", "value"]);
  // 기존 값과 병합해서 저장 — 앱이 일부 키만 보내도 telegram_bot_token 등
  // 시트에 직접 넣어둔 다른 설정이 지워지지 않게 한다 (통째 교체 방식의 버그 수정)
  const existing = {};
  settingsSheet.getDataRange().getValues().slice(1).forEach(row => {
    if (row[0] !== "") existing[String(row[0])] = String(row[1]);
  });
  Object.keys(settings).forEach(key => {
    existing[key] = settings[key];
  });
  settingsSheet.clearContents();
  settingsSheet.appendRow(["key", "value"]);
  Object.keys(existing).forEach(key => {
    settingsSheet.appendRow([key, existing[key]]);
  });
}

// V2: 월별 자산 스냅샷 저장 (전체 교체 방식 — 클라이언트가 병합해서 보냄)
function saveSnapshots(snapshots) {
  const sheet = getOrCreateSheet(SNAPSHOTS_SHEET_NAME, ["month", "stocksValue", "otherAssets", "totalAssets", "contribution", "note"]);
  sheet.clearContents();
  sheet.appendRow(["month", "stocksValue", "otherAssets", "totalAssets", "contribution", "note"]);
  snapshots.forEach(s => {
    sheet.appendRow(["'" + s.month, s.stocksValue, s.otherAssets, s.totalAssets, s.contribution, s.note || ""]);
  });
}

// V2: 매매 일지 저장 (전체 교체 방식)
function saveJournal(journal) {
  const sheet = getOrCreateSheet(JOURNAL_SHEET_NAME, ["id", "date", "ticker", "action", "price", "quantity", "reason", "stopLoss", "targetPrice", "emotion"]);
  sheet.clearContents();
  sheet.appendRow(["id", "date", "ticker", "action", "price", "quantity", "reason", "stopLoss", "targetPrice", "emotion"]);
  journal.forEach(j => {
    sheet.appendRow(["'" + j.id, j.date, "'" + j.ticker, j.action, j.price, j.quantity, j.reason || "", j.stopLoss || "", j.targetPrice || "", j.emotion || ""]);
  });
}

// =====================================================================
// ===== 매일 자동 종합분석 (시간 기반 트리거로 실행) =====
// =====================================================================
// setup: Apps Script 편집기 상단에서 함수 'setupDailyTrigger'를 한 번 실행하면
//        매일 오전 7~8시 사이에 runDailyAnalysis가 자동 실행되도록 트리거가 걸린다.
// (Apps Script 시간 트리거는 "정확히 7:30"이 아니라 "7~8시 구간 중"으로만 지정 가능)

const MAX_PORTFOLIO_HISTORY = 30;

// OpenRouter 무료 모델 (하나가 혼잡하면 다음으로 자동 전환)
const FREE_MODELS = [
  "meta-llama/llama-3.3-70b-instruct:free",
  "openai/gpt-oss-120b:free",
  "qwen/qwen3-next-80b-a3b-instruct:free",
  "nvidia/nemotron-3-super-120b-a12b:free"
];

// 한 번만 실행하면 자동분석 트리거가 설치된다.
// 미국장은 한국시간 새벽에 마감되고 주말엔 시세가 안 변하므로,
// "월~금 미장 마감 결과"를 반영하는 화~토 오전에만 분석한다. (일/월 아침은 스킵 — 금요일 종가와 동일)
function setupDailyTrigger() {
  // 기존 자동분석 트리거는 모두 제거하고 새로 건다 (중복 방지)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runDailyAnalysis") {
      ScriptApp.deleteTrigger(t);
    }
  });
  const weekdays = [
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
    ScriptApp.WeekDay.SATURDAY
  ];
  weekdays.forEach(day => {
    ScriptApp.newTrigger("runDailyAnalysis")
      .timeBased()
      .onWeekDay(day)
      .atHour(7)          // 오전 7시대 실행 (7:00~7:59 사이 구글이 알아서)
      .inTimezone("Asia/Seoul")
      .create();
  });
  Logger.log("화~토 오전 7시대 자동분석 트리거가 설치되었습니다. (일/월 아침은 시세 동일로 스킵)");
}

// 매일 아침 트리거가 호출하는 함수 — 종합분석을 돌리고 시트에 저장한다.
function runDailyAnalysis() {
  const analysis = runPortfolioAnalysis_();
  if (!analysis) return; // 실패 사유는 runPortfolioAnalysis_ 안에서 로그로 남김
  Logger.log("자동 종합분석 완료 및 저장. 수익률 " + analysis.totalProfitLossPercent + "%");
}

// 종합분석 공통 로직 — 시세 조회 → 프롬프트 → OpenRouter → 시트 저장까지 하고 결과를 반환.
// 실패 시 null 반환(사유는 로그). 일간 자동분석과 주간 텔레그램 리포트가 공유한다.
function runPortfolioAnalysis_() {
  const settings = readSettings_();
  const apiKey = settings.openrouter_key;
  const priceApiKey = settings.twelvedata_key;
  if (!apiKey) {
    Logger.log("OpenRouter 키가 Settings 시트에 없습니다. 분석 중단.");
    return null;
  }

  const stocksSheet = getOrCreateSheet(STOCKS_SHEET_NAME, ["id", "ticker", "name", "quantity", "buyPrice"]);
  const stockRows = stocksSheet.getDataRange().getValues().slice(1).filter(r => r[0] !== "");
  if (stockRows.length === 0) {
    Logger.log("등록된 종목이 없습니다. 분석 중단.");
    return null;
  }

  const holdings = [];
  stockRows.forEach(row => {
    const ticker = String(row[1]);
    const name = String(row[2] || "");
    const quantity = Number(row[3]);
    const buyPrice = Number(row[4]);
    const isKorean = isKoreanTicker_(ticker);
    try {
      const currentPrice = fetchPrice_(ticker, isKorean, priceApiKey);
      const totalValue = quantity * currentPrice;
      const totalBuyAmount = quantity * buyPrice;
      const profitLossPercent = (((totalValue - totalBuyAmount) / totalBuyAmount) * 100).toFixed(2);
      holdings.push({
        ticker: ticker, name: name, quantity: quantity,
        market: isKorean ? "국내" : "미국",
        currencySymbol: isKorean ? "₩" : "$",
        buyPrice: buyPrice, currentPrice: currentPrice,
        totalValue: totalValue, profitLossPercent: profitLossPercent
      });
    } catch (e) {
      Logger.log("시세 조회 실패(" + ticker + "): " + e.message);
    }
  });

  if (holdings.length === 0) {
    Logger.log("모든 종목 시세 조회에 실패했습니다. 분석 중단.");
    return null;
  }

  const totalPortfolioValue = holdings.reduce((s, h) => s + h.totalValue, 0);
  holdings.forEach(h => {
    h.weightPercent = ((h.totalValue / totalPortfolioValue) * 100).toFixed(1);
  });

  const prompt = buildPortfolioRebalancePrompt_(holdings);
  const result = callOpenRouterWithFallback_(apiKey, prompt);
  if (!result) {
    Logger.log("OpenRouter 모든 모델 응답 실패. 분석 중단.");
    return null;
  }

  const totalBuyAmt = holdings.reduce((s, h) => s + h.buyPrice * h.quantity, 0);
  const totalProfitLossPercent = totalBuyAmt > 0
    ? parseFloat(((totalPortfolioValue - totalBuyAmt) / totalBuyAmt * 100).toFixed(2))
    : 0;

  savePortfolioHistoryEntry_({
    date: new Date().toISOString(),
    totalProfitLossPercent: totalProfitLossPercent,
    modelUsed: result.modelUsed,
    aiText: result.content
  });

  return {
    totalProfitLossPercent: totalProfitLossPercent,
    aiText: result.content,
    modelUsed: result.modelUsed,
    holdings: holdings,
    apiKey: apiKey
  };
}

function readSettings_() {
  const settingsSheet = getOrCreateSheet(SETTINGS_SHEET_NAME, ["key", "value"]);
  const rows = settingsSheet.getDataRange().getValues().slice(1);
  const settings = {};
  rows.forEach(row => {
    if (row[0] !== "") settings[String(row[0])] = String(row[1]);
  });
  return settings;
}

function isKoreanTicker_(ticker) {
  return /^[0-9A-Z]{6}$/.test(ticker) && /\d/.test(ticker);
}

// 시세 조회: 한국=네이버 직접(서버라 프록시 불필요), 미국=Twelve Data
function fetchPrice_(ticker, isKorean, priceApiKey) {
  if (isKorean) {
    const url = "https://polling.finance.naver.com/api/realtime/domestic/stock/" + encodeURIComponent(ticker);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    const entry = data && data.datas && data.datas[0];
    if (!entry || !entry.closePrice) throw new Error("국내 시세 없음");
    return parseFloat(String(entry.closePrice).replace(/,/g, ""));
  }
  const url = "https://api.twelvedata.com/price?symbol=" + encodeURIComponent(ticker) + "&apikey=" + encodeURIComponent(priceApiKey);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (!data.price || data.status === "error") throw new Error(data.message || "미국 시세 없음");
  return parseFloat(data.price);
}

function callOpenRouterWithFallback_(apiKey, promptMessage) {
  for (let i = 0; i < FREE_MODELS.length; i++) {
    const model = FREE_MODELS[i];
    try {
      const res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "post",
        contentType: "application/json",
        headers: { "Authorization": "Bearer " + apiKey },
        payload: JSON.stringify({ model: model, messages: [{ role: "user", content: promptMessage }] }),
        muteHttpExceptions: true
      });
      const data = JSON.parse(res.getContentText());
      if (data.choices && data.choices[0]) {
        return { content: data.choices[0].message.content, modelUsed: model };
      }
      Logger.log("모델 응답 이상(" + model + "): " + res.getContentText().slice(0, 200));
    } catch (e) {
      Logger.log("모델 호출 실패(" + model + "): " + e.message);
    }
  }
  return null;
}

function savePortfolioHistoryEntry_(entry) {
  const sheet = getOrCreateSheet(PORTFOLIO_HISTORY_SHEET_NAME, ["date", "totalProfitLossPercent", "modelUsed", "aiText"]);
  sheet.appendRow([entry.date, entry.totalProfitLossPercent, entry.modelUsed, entry.aiText]);
  // 상한 초과 시 가장 오래된 행(헤더 다음 행) 삭제
  const lastRow = sheet.getLastRow();
  if (lastRow - 1 > MAX_PORTFOLIO_HISTORY) {
    sheet.deleteRow(2);
  }
}

// 투자자 프로필 (index.html의 INVESTOR_PROFILE과 동일하게 유지)
const INVESTOR_PROFILE = "\n## 투자자 프로필 (이 정보를 반드시 모든 판단의 기준으로 삼으세요)\n"
  + "- **목표**: 경제적 자유. 10년 이상 장기 복리로 시드머니를 공격적으로 불리는 단계.\n"
  + "- **전략**: 코어-새틀라이트. 나스닥100 지수추종 ETF(QQQM 등)를 포트폴리오의 50~70% 코어로 유지하고, 나머지 30~50%는 테마+실적이 동시에 검증된 주도주에 압축 투자해 알파 수익 추구.\n"
  + "- **월 투자 가능 금액**: 약 100~150만원 (매수형 리밸런싱 재원).\n"
  + "- **손절 기준**: 아직 없음 → 이 종목 특성에 맞는 손절/익절 기준을 AI가 명확히 제안해줘야 함.\n"
  + "- **성향**: 공격적이지만 코어 비중이 흔들리면 불안함. 단기 등락보다 장기 복리 우상향에 집중.\n"
  + "\n## 운용 원칙 (모든 제안은 아래 원칙에 어긋나면 안 됩니다)\n"
  + "1. **복리의 수학**: -50% 손실은 +100% 수익으로만 복구된다. 큰 손실을 피하는 것이 수익을 쫓는 것보다 우선한다.\n"
  + "2. **물타기 금지, 불타기 원칙**: 새틀라이트 추가매수는 상승 추세(200일선 위)에서 실적으로 검증될 때만. 하락 추세 종목에 \"평단 낮추기\" 목적의 물타기는 절대 제안하지 말 것. 이기는 종목에 추가하는 피라미딩이 원칙.\n"
  + "3. **기회비용 벤치마크**: 모든 새틀라이트는 QQQM 대비 초과수익이 기대될 때만 보유 가치가 있다. 6개월 이상 지수에 뒤처지면 교체를 검토하라.\n"
  + "4. **집중 한도**: 새틀라이트 단일 종목은 포트폴리오의 15%를 넘지 않게. 초과 시 수익 실현으로 비중 조절을 제안하라.\n"
  + "5. **월 투입 우선순위**: ① 코어 50~70% 유지가 최우선 → ② 원칙 2를 충족한 새틀라이트 → ③ 조건 미달이면 현금 대기도 유효한 선택이다. 무리하게 살 곳을 찾아주지 말 것.";

// index.html의 buildPortfolioRebalancePrompt와 동일하게 유지
function buildPortfolioRebalancePrompt_(holdings) {
  const holdingsText = holdings.map(function(h) {
    return "- " + (h.name || h.ticker) + " (" + h.ticker + ", " + h.market + "): 평균매입가 " + h.currencySymbol + h.buyPrice
      + ", 현재가 " + h.currencySymbol + h.currentPrice.toFixed(2)
      + ", 평가금액 " + h.currencySymbol + Math.round(h.totalValue).toLocaleString()
      + ", 수익률 " + h.profitLossPercent + "%, 비중 " + h.weightPercent + "%";
  }).join("\n");

  return "당신은 경제적 자유를 목표로 하는 직장인 투자자의 전담 포트폴리오 매니저입니다.\n"
    + "아래 투자자 프로필과 전체 포트폴리오를 종합해 구조적 진단과 실행 계획을 제시하세요.\n"
    + INVESTOR_PROFILE + "\n\n"
    + "## 현재 보유 포트폴리오 전체\n" + holdingsText + "\n\n"
    + "## 요청 사항\nJSON 없이, 아래 4개 섹션을 순서대로 한국어로 작성하세요.\n\n"
    + "### 1. 🏗️ 코어-새틀라이트 구조 진단\n"
    + "- 현재 포트폴리오에서 코어(나스닥100 등 지수 ETF)와 새틀라이트(주도주)의 실제 비중을 각각 합산하세요.\n"
    + "- 목표 비율(코어 50~70% : 새틀라이트 30~50%)과 비교해 현재 어디가 과소/과대 배분인지 명시하세요.\n"
    + "- 국내/미국 시장 편중 리스크가 있는지 짚어주세요.\n\n"
    + "### 2. ⚖️ 매수형 리밸런싱 제안 (월 100~150만원 기준)\n"
    + "매도 없이 월급으로 비중을 맞추는 방향을 우선하되, 운용 원칙 5의 우선순위(코어 유지 → 검증된 새틀라이트 → 현금 대기)를 따르세요.\n"
    + "- 이번 달 월급 중 코어에 얼마, 어떤 새틀라이트에 얼마를 배분해야 하는지 구체적 금액으로 제시하세요.\n"
    + "- 하락 추세 새틀라이트에 물타기 배분은 금지. 살 만한 곳이 없으면 \"이번 달은 코어 + 현금 대기\"라고 정직하게 답하세요.\n"
    + "- 만약 즉시 비중 조정이 필요한 종목이 있다면 분할매도 계획도 함께 제시하세요.\n\n"
    + "### 3. 🔥 새틀라이트 종목 옥석 가리기\n"
    + "보유 새틀라이트 종목 각각에 대해:\n"
    + "- **유지**: 테마+실적 모두 살아있어 계속 보유할 이유가 있는 종목\n"
    + "- **축소/매도 검토**: 실적 없이 기대감만 남았거나 주도주 역할을 잃은 종목\n"
    + "매도 검토 종목이 있다면, 그 자금을 어디에 재배치할지도 제시하세요.\n\n"
    + "### 4. 📅 다음 월급날 행동 지침\n"
    + "10년 복리 목표를 위해 이번 달 당장 실행할 것을 3줄 이내로 명확하게 요약하세요.\n"
    + "모호한 표현 없이, 종목명과 금액이 포함된 실행 가능한 지침으로 작성하세요.";
}

// =====================================================================
// ===== 일요일 오전 텔레그램 주간 리포트 =====
// =====================================================================
// 사전 준비 (Settings 시트에 아래 2개 key/value를 추가):
//   telegram_bot_token  = 봇파더(@BotFather)에서 받은 토큰
//   telegram_chat_id    = 내 채팅 ID (봇에게 아무 메시지나 보낸 뒤,
//                         https://api.telegram.org/bot<토큰>/getUpdates 열면 "chat":{"id":숫자} 확인 가능)
// 설치: 함수 'setupWeeklyTelegramTrigger'를 한 번 실행하면 매주 일요일 오전에 자동 발송된다.

function setupWeeklyTelegramTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runWeeklyTelegramReport") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runWeeklyTelegramReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)          // 일요일 오전 9시대 발송
    .inTimezone("Asia/Seoul")
    .create();
  Logger.log("매주 일요일 오전 9시대 텔레그램 리포트 트리거가 설치되었습니다.");
}

// 일요일 트리거가 호출하는 함수.
// 주말엔 시세가 금요일 종가 그대로이므로, 새로 분석하지 않고
// 토요일 아침에 이미 나온 마지막 종합분석 기록을 가져다 요약해서 보낸다. (중복 분석/429 위험 제거)
function runWeeklyTelegramReport() {
  const settings = readSettings_();
  const botToken = settings.telegram_bot_token;
  const chatId = settings.telegram_chat_id;
  const apiKey = settings.openrouter_key;
  if (!botToken || !chatId) {
    Logger.log("텔레그램 봇 토큰/chat_id가 Settings 시트에 없습니다. 리포트 중단.");
    return;
  }

  const last = getLastPortfolioHistoryEntry_();
  if (!last) {
    sendTelegramMessage_(botToken, chatId,
      "⏰ 일요일 아침입니다.\n아직 저장된 종합분석 기록이 없어요.\n앱에서 직접 종합분석을 눌러 확인해보세요.");
    return;
  }

  // 긴 분석 전문을 텔레그램용 주간 리포트로 요약 (일요일은 새 분석 없이 요약만 하므로 AI 호출은 이 1회뿐)
  const summary = summarizeForTelegram_(apiKey, last.aiText, last.totalProfitLossPercent);

  // 앱 주소는 Settings 시트의 app_url 값 우선, 없으면 기본 GitHub Pages 주소 사용
  const appUrl = settings.app_url || "https://whiteinsun1.github.io/portfolio-ai/";

  const sign = last.totalProfitLossPercent >= 0 ? "+" : "";
  const analyzedStr = Utilities.formatDate(new Date(last.date), "Asia/Seoul", "M/d(E)");

  // V2 목표 진척 라인 (스냅샷·목표 설정이 있을 때만 붙음)
  const goalLine = buildGoalProgressLine_(settings);

  const message =
    "📊 *주간 포트폴리오 요약*\n" +
    "(" + analyzedStr + " 종가 기준 · 전체 수익률 *" + sign + last.totalProfitLossPercent + "%*)\n" +
    (goalLine ? goalLine + "\n" : "") + "\n" +
    summary + "\n\n" +
    "━━━━━━━━━━━━━\n" +
    "🔗 [앱에서 자세한 분석 보기](" + appUrl + ")";

  sendTelegramMessage_(botToken, chatId, message);
  Logger.log("일요일 텔레그램 리포트 발송 완료. (분석시점: " + analyzedStr + ")");
}

// V2: 목표(10년 대시보드) 진척 라인 생성 — 목표 설정과 스냅샷이 있으면 "🎯 목표 진행 ..." 한 줄 반환, 없으면 빈 문자열
function buildGoalProgressLine_(settings) {
  try {
    const target = Number(settings.goal_target || 0);
    if (!target) return "";

    const snapSheet = getOrCreateSheet(SNAPSHOTS_SHEET_NAME, ["month", "stocksValue", "otherAssets", "totalAssets", "contribution", "note"]);
    const lastRow = snapSheet.getLastRow();
    if (lastRow < 2) return "";
    const snap = snapSheet.getRange(lastRow, 1, 1, 6).getValues()[0];
    const totalAssets = Number(snap[3]);
    if (!totalAssets) return "";

    const progressPct = (totalAssets / target * 100).toFixed(1);
    const eokCur = (totalAssets / 100000000).toFixed(2);
    const eokTarget = (target / 100000000).toFixed(0);

    // 계획 곡선 대비 위/아래 판정
    let planNote = "";
    const startAmount = Number(settings.goal_start_amount || 0);
    const startDate = settings.goal_start_date ? new Date(settings.goal_start_date) : null;
    const cagr = Number(settings.goal_cagr || 0);
    const monthly = Number(settings.goal_monthly || 0);
    if (startDate && cagr > 0) {
      const monthsElapsed = (new Date() - startDate) / (1000 * 60 * 60 * 24 * 30.44);
      const i = Math.pow(1 + cagr / 100, 1 / 12) - 1;
      const n = Math.max(monthsElapsed, 0);
      const planValue = startAmount * Math.pow(1 + i, n) + monthly * ((Math.pow(1 + i, n) - 1) / i);
      const diffPct = ((totalAssets - planValue) / planValue * 100).toFixed(1);
      planNote = totalAssets >= planValue ? " · 계획보다 +" + diffPct + "% 앞섬 ✅" : " · 계획 대비 " + diffPct + "% ⚠️";
    }

    return "🎯 목표 진행: " + eokCur + "억/" + eokTarget + "억 (" + progressPct + "%)" + planNote;
  } catch (e) {
    return "";
  }
}

// V2: 매월 1일 오전, 자산 스냅샷 기록 리마인드 발송 트리거 설치 (한 번만 실행)
function setupMonthlySnapshotReminder() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runMonthlySnapshotReminder") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("runMonthlySnapshotReminder")
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .inTimezone("Asia/Seoul")
    .create();
  Logger.log("매월 1일 오전 10시대 스냅샷 리마인드 트리거가 설치되었습니다.");
}

function runMonthlySnapshotReminder() {
  const settings = readSettings_();
  const botToken = settings.telegram_bot_token;
  const chatId = settings.telegram_chat_id;
  if (!botToken || !chatId) return;
  const appUrl = settings.app_url_v2 || settings.app_url || "https://whiteinsun1.github.io/portfolio-ai/v2.html";
  sendTelegramMessage_(botToken, chatId,
    "📅 새 달이 시작됐어요!\n" +
    "이번 달 자산 스냅샷을 기록하고 적립을 실행하세요.\n" +
    "(1분이면 됩니다 — 10년 그래프에 점 하나가 찍혀요)\n\n" +
    "🔗 [대시보드 열기](" + appUrl + ")");
}

// PortfolioHistory 시트의 가장 최근 기록 한 건을 읽어온다 (없으면 null)
function getLastPortfolioHistoryEntry_() {
  const sheet = getOrCreateSheet(PORTFOLIO_HISTORY_SHEET_NAME, ["date", "totalProfitLossPercent", "modelUsed", "aiText"]);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null; // 헤더만 있음
  const row = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
  return {
    date: String(row[0]),
    totalProfitLossPercent: Number(row[1]),
    modelUsed: String(row[2]),
    aiText: String(row[3])
  };
}

// 긴 종합분석 텍스트를 텔레그램용 주간 리포트 요약으로 압축
function summarizeForTelegram_(apiKey, fullAnalysisText, profitPercent) {
  const prompt =
    "아래는 한 투자자(경제적 자유 목표, 코어-새틀라이트 전략, 10년 장기 복리)의 포트폴리오 종합분석 전문입니다.\n" +
    "이걸 '일요일 아침에 받아보는 주간 리포트'로 요약하세요. 커피 마시며 30초 안에 읽고 이번 주 할 일을 파악할 수 있게.\n\n" +
    "아래 형식을 그대로 지키세요 (각 섹션 제목과 이모지 포함, 마크다운 굵게(*)는 쓰지 말 것):\n\n" +
    "🧭 한 줄 총평\n" +
    "(포트폴리오가 지금 건강한지, 코어-새틀라이트 균형이 맞는지 한 문장으로)\n\n" +
    "✅ 이번 주 할 일\n" +
    "(실제 행동 2~3줄. 각 줄 앞에 · 를 붙이고 '무엇을 · 얼마에 · 왜'를 담아 구체적으로. 살 곳이 없으면 '이번 주는 코어 적립 + 현금 대기'라고 정직하게.)\n\n" +
    "⚠️ 주의할 종목\n" +
    "(손절선 근접·비중 과다·추세 이탈 등 지켜봐야 할 종목 1~2개. 없으면 '특이사항 없음'.)\n\n" +
    "규칙:\n" +
    "- 종목명과 구체적 숫자(가격·금액·%)를 반드시 포함.\n" +
    "- '상황 봐서', '적절히' 같은 모호한 표현 금지.\n" +
    "- 물타기(하락 종목 평단 낮추기)는 절대 권하지 말 것.\n" +
    "- 전체 12줄 이내로 간결하게.\n\n" +
    "=== 종합분석 전문 ===\n" + fullAnalysisText;

  const result = callOpenRouterWithFallback_(apiKey, prompt);
  const raw = (result && result.content) ? result.content.trim() : (fullAnalysisText.slice(0, 400) + "...");
  // AI가 실수로 남긴 마크다운 특수문자(* _ ` [ ])는 텔레그램 Markdown 파싱을 깨뜨려 발송 실패를 유발하므로 제거
  return raw.replace(/[*_`\[\]]/g, "");
}

// 텔레그램 메시지 발송
function sendTelegramMessage_(botToken, chatId, text) {
  const url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: text,
        parse_mode: "Markdown",
        disable_web_page_preview: true  // 하단 링크의 큰 미리보기 카드가 안 뜨게
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (!data.ok) {
      Logger.log("텔레그램 응답 오류: " + res.getContentText());
    }
  } catch (e) {
    Logger.log("텔레그램 발송 실패: " + e.message);
  }
}
