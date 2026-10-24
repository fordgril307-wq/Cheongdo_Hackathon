// 사투리 예약 문장 → 예약 정보(JSON)
// 1순위: Azure OpenAI (환경변수가 있을 때)
// 2순위: 규칙 기반 해석 (키가 없거나 AI 호출이 실패해도 데모가 멈추지 않도록)

const { villages, destinations, findVillage, branchFor } = require('./places');
const { todayKST, fmt, addDays } = require('./dates');

const DAY_WORDS = { 오늘: 0, 내일: 1, 낼: 1, 모레: 2, 모래: 2, 글피: 3 };
const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];
const KOR_NUM = { 한: 1, 두: 2, 세: 3, 네: 4, 다섯: 5, 여섯: 6, 일곱: 7, 여덟: 8, 아홉: 9, 열: 10, 열한: 11, 열두: 12 };

function parseDate(text, base = todayKST()) {
  for (const [w, n] of Object.entries(DAY_WORDS)) {
    if (text.includes(w)) return fmt(addDays(base, n));
  }
  const m = text.match(/([월화수목금토일])요일/);
  if (m) {
    const target = WEEKDAYS.indexOf(m[1]);
    let diff = (target - base.getUTCDay() + 7) % 7;
    if (diff === 0) diff = 7;
    if (text.includes('다음 주') || text.includes('담주')) diff += diff < 7 ? 7 : 0;
    return fmt(addDays(base, diff));
  }
  return null;
}

function parseTime(text) {
  let hour = null;
  let minute = 0;
  const digit = text.match(/(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분|\s*반)?/);
  if (digit) {
    hour = parseInt(digit[1], 10);
    if (digit[2]) minute = parseInt(digit[2], 10);
    else if (/시\s*반/.test(digit[0])) minute = 30;
  } else {
    const word = text.match(/(열두|열한|열|한|두|세|네|다섯|여섯|일곱|여덟|아홉)\s*시(\s*반)?/);
    if (word) {
      hour = KOR_NUM[word[1]];
      if (word[2]) minute = 30;
    }
  }
  const pm = /오후|저녁|낮/.test(text);
  if (hour !== null) {
    if (pm && hour < 12) hour += 12;
    else if (!pm && hour <= 6) hour += 12; // "세 시"는 보통 오후
  } else if (/새벽|첫차/.test(text)) hour = 7;
  else if (/아침/.test(text)) hour = 9;
  else if (/오전/.test(text)) hour = 10;
  else if (/점심/.test(text)) hour = 12;
  else if (/오후|낮/.test(text)) hour = 14;
  else if (/저녁/.test(text)) hour = 17;
  if (hour === null) return null;
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function parseVillage(text, fallbackId) {
  const byRi = villages.find((v) => text.includes(v.name.split(' ')[1]));
  if (byRi) return byRi;
  const byMyeon = villages.find((v) => text.includes(v.myeon) || text.includes(v.myeon.replace('면', '')));
  if (byMyeon && /(에서|사는|살아|우리)/.test(text)) return byMyeon;
  return findVillage(fallbackId) || null;
}

function parseDestination(text, village) {
  if (/보건지소|지소/.test(text) && village) return branchFor(village.myeon);
  for (const d of destinations) {
    if (text.includes(d.name) || d.keywords.some((k) => text.includes(k))) return d;
  }
  if (/병원|진료|아파|쑤시|약/.test(text)) return destinations.find((d) => d.id === 'daenam');
  return null;
}

function parsePurpose(text) {
  if (/투석/.test(text)) return '투석';
  if (/물리치료|침 맞|침맞|한방|한의/.test(text)) return '물리치료·한방';
  if (/약|처방/.test(text)) return '약 처방';
  if (/검사|건강검진|피검사/.test(text)) return '검사';
  if (/시장|장보|장에/.test(text)) return '장보기';
  if (/병원|진료|아파|쑤시|결리|어지럽/.test(text)) return '진료';
  return '기타';
}

function ruleParse(text, profile = {}) {
  const t = (text || '').replace(/\s+/g, ' ').trim();
  const village = parseVillage(t, profile.villageId);
  const destination = parseDestination(t, village);
  const date = parseDate(t);
  const time = parseTime(t);
  const passengers = /(둘이|두 명|할배랑|할매랑|영감이랑|집사람이랑|같이 가)/.test(t) ? 2 : 1;
  const wheelchair = /휠체어|못 걸|걷기 힘들|다리가 불편/.test(t);
  const missing = [];
  if (!village) missing.push('어느 마을에서 타실지');
  if (!destination) missing.push('어디로 가실지');
  if (!date) missing.push('며칠에 가실지');
  if (!time) missing.push('몇 시까지 가셔야 하는지');
  return {
    villageId: village ? village.id : null,
    destinationId: destination ? destination.id : null,
    date,
    time,
    purpose: parsePurpose(t),
    passengers,
    wheelchair,
    clarify: missing.length ? `${missing.join(', ')} 말씀해 주이소.` : null,
    summary: [
      village && village.name,
      destination && `→ ${destination.name}`,
      date,
      time && `${time}까지`
    ].filter(Boolean).join(' '),
    engine: 'rules'
  };
}

function buildPrompt(text, profile) {
  const today = fmt(todayKST());
  const weekday = WEEKDAYS[todayKST().getUTCDay()];
  const placeList = villages.map((v) => `${v.id}: ${v.name}`).join('\n');
  const destList = destinations.map((d) => `${d.id}: ${d.name}(${d.kind})`).join('\n');
  return [
    {
      role: 'system',
      content:
        `너는 경북 청도군 어르신의 택시 예약을 받는 상담원이다. 어르신은 경상도 사투리로 말한다.\n` +
        `오늘은 ${today}(${weekday}요일)이다. 말한 내용을 아래 JSON 하나로만 답하라.\n` +
        `{"villageId": 마을 id 또는 null, "destinationId": 목적지 id 또는 null, "date": "YYYY-MM-DD" 또는 null, ` +
        `"time": "HH:MM"(도착해야 하는 시각, 24시간) 또는 null, "purpose": "진료|물리치료·한방|약 처방|검사|투석|장보기|기타", ` +
        `"passengers": 숫자, "wheelchair": true/false, "clarify": 빠진 정보를 묻는 짧은 사투리 질문 또는 null, ` +
        `"summary": 표준어 한 줄 요약}\n` +
        `- "보건지소"는 그 마을이 속한 면의 보건지소로 고른다. 면에 보건지소가 없으면 health-center.\n` +
        `- 시간이 "아침"이면 09:00, "점심"이면 12:00, 숫자 1~6시는 오후로 본다.\n` +
        `- 마을을 말하지 않았으면 등록된 마을(${profile.villageId || '없음'})을 쓴다.\n\n` +
        `[마을]\n${placeList}\n\n[목적지]\n${destList}`
    },
    { role: 'user', content: text }
  ];
}

async function aiParse(text, profile = {}) {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const key = process.env.AZURE_OPENAI_API_KEY;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-10-21';
  if (!endpoint || !key || !deployment) return null;

  const url = `${endpoint.replace(/\/$/, '')}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 12000);
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'api-key': key },
      body: JSON.stringify({
        messages: buildPrompt(text, profile),
        temperature: 0,
        response_format: { type: 'json_object' }
      }),
      signal: controller.signal
    });
    if (!res.ok) throw new Error(`Azure OpenAI ${res.status}`);
    const data = await res.json();
    const parsed = JSON.parse(data.choices[0].message.content);
    return { ...parsed, engine: 'azure-openai' };
  } finally {
    clearTimeout(timer);
  }
}

async function parseRequest(text, profile = {}) {
  try {
    const ai = await aiParse(text, profile);
    if (ai) {
      // AI가 빠뜨린 값은 규칙 기반 결과로 보완
      const rules = ruleParse(text, profile);
      for (const k of ['villageId', 'destinationId', 'date', 'time']) {
        if (!ai[k] && rules[k]) ai[k] = rules[k];
      }
      return ai;
    }
  } catch (err) {
    const fallback = ruleParse(text, profile);
    fallback.engine = 'rules (AI 실패: ' + err.message + ')';
    return fallback;
  }
  return ruleParse(text, profile);
}

module.exports = { parseRequest, ruleParse, parseDate, parseTime };
