// 청도효도택시 화면 로직
// 어르신 화면: 한 화면에 한 질문 (어디로 → 언제 → 확인 → 완료)
// 신청 방법 두 가지: 어르신 직접 / 경로당(마을회관)에서 여러 분을 대신 신청
// 참고한 패턴: 호출 앱의 '자주 가는 곳'·배차 상태 타임라인·하단 고정 버튼,
// 시니어 모드의 큰 글씨·적은 선택지·가족 공유, 전화 예약 대안
const COLORS = ['#F26B21', '#4F9A4A', '#2F74D0', '#8A5CC2', '#D99A00', '#D9456B', '#3E8C8C', '#7A6A5C'];
const EXAMPLES = [
  '내일 아침 열 시까지 대구 효성병원 가야 되는데예',
  '삭신이 쑤시가 모레 아홉 시에 대남병원 함 가볼라꼬',
  '낼 오후 두 시 보건지소 물리치료 받으러 갈라카는데'
];
const TIMES = [
  ['09:00', '오전 9시'], ['10:00', '오전 10시'], ['11:00', '오전 11시'],
  ['13:00', '오후 1시'], ['14:00', '오후 2시'], ['15:00', '오후 3시']
];
const ICONS = {
  hospital: '<svg viewBox="0 0 24 24"><path fill="currentColor" d="M10 3h4v7h7v4h-7v7h-4v-7H3v-4h7z"/></svg>',
  branch: '<svg viewBox="0 0 24 24"><path fill="currentColor" d="M12 3 2 11h3v10h5v-6h4v6h5V11h3L12 3Z"/></svg>',
  city: '<svg viewBox="0 0 24 24"><path fill="currentColor" d="M3 21V9l6-4v4l6-4v16H3Zm14 0V11h4v10h-4Z"/></svg>',
  market: '<svg viewBox="0 0 24 24"><path fill="currentColor" d="M7 18a2 2 0 1 0 0 4 2 2 0 0 0 0-4Zm10 0a2 2 0 1 0 0 4 2 2 0 0 0 0-4ZM5.2 4H2V2h4.6l1 2H21l-3.4 8H8.5l-1 2H20v2H4.2l2.7-5.3L5.2 4Z"/></svg>'
};
const WEEK = '일월화수목금토';

const state = {
  places: null, profiles: [], requests: [], map: null, layer: null,
  mode: 'self', hall: null, screen: 'where', booking: {}, lastPlan: null, lastIds: []
};
const $ = (id) => document.getElementById(id);
const config = window.HYODO_CONFIG || {};

async function api(path, body) {
  const res = await fetch(`/api/${path}`, body ? {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  } : undefined);
  if (!res.ok) throw new Error(`${path} 실패 (${res.status})`);
  return res.json();
}

const village = (id) => state.places.villages.find((v) => v.id === id);
const dest = (id) => state.places.destinations.find((d) => d.id === id);
const riName = (id) => (village(id)?.name || '').split(' ')[1] || '';
const hallName = (id) => `${riName(id)} 경로당`;
const pickupPoint = (id) => {
  const stop = village(id)?.stop || riName(id);
  return /정류장|터미널|역\(/.test(stop) ? stop : `${stop} 정류장`;
};
const placeName = (id) => village(id)?.name || '–';
const destName = (id) => dest(id)?.name || '–';
const profile = () => state.profiles[$('profile').value] || {};
const homeVillage = () => (state.mode === 'hall' ? state.hall : profile().villageId);
const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

// 날짜는 항상 요일까지: "10월 25일 일요일"
function dateLabel(s) {
  if (!s) return '–';
  const dt = new Date(s + 'T00:00:00');
  return `${dt.getMonth() + 1}월 ${dt.getDate()}일 ${WEEK[dt.getDay()]}요일`;
}
function relWord(s) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const diff = Math.round((new Date(s + 'T00:00:00') - today) / 86400000);
  return ['오늘', '내일', '모레'][diff] || '';
}
// "내일, 10월 25일 일요일"
function dayFull(s) {
  const w = relWord(s);
  return w ? `${w}, ${dateLabel(s)}` : dateLabel(s);
}
function timeLabel(t) {
  if (!t) return '–';
  const [h, m] = t.split(':').map(Number);
  const ampm = h < 12 ? '오전' : '오후';
  return `${ampm} ${h > 12 ? h - 12 : h}시${m ? ` ${m}분` : ''}`;
}
function minutesBetween(a, b) {
  const R = 6371, rad = (x) => (x * Math.PI) / 180;
  const h = Math.sin(rad(b.lat - a.lat) / 2) ** 2 + Math.cos(rad(a.lat)) * Math.cos(rad(b.lat)) * Math.sin(rad(b.lng - a.lng) / 2) ** 2;
  const km = 2 * R * Math.asin(Math.sqrt(h)) * 1.35;
  return Math.max(5, Math.round((km / 40) * 60 / 5) * 5);
}

// ---------- 음성 안내 ----------
function speak(text) {
  if (!('speechSynthesis' in window)) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text);
  u.lang = 'ko-KR';
  u.rate = 0.9;
  const ko = speechSynthesis.getVoices().find((v) => v.lang && v.lang.startsWith('ko'));
  if (ko) u.voice = ko;
  speechSynthesis.speak(u);
}

// ---------- 글자 크게 ----------
function setBigText(on) {
  document.documentElement.classList.toggle('big', on);
  $('bigtext').setAttribute('aria-pressed', on);
  try { localStorage.setItem('hyodo-big', on ? '1' : '0'); } catch (e) { /* 저장 불가 환경 */ }
}
$('bigtext').addEventListener('click', () => setBigText(!document.documentElement.classList.contains('big')));
try { if (localStorage.getItem('hyodo-big') === '1') setBigText(true); } catch (e) { /* 무시 */ }

// 어둡게 (빛 번짐이 심한 분)
function setDark(on) {
  document.documentElement.classList.toggle('dark', on);
  $('darkmode').setAttribute('aria-pressed', on);
  try { localStorage.setItem('hyodo-dark', on ? '1' : '0'); } catch (e) { /* 저장 불가 환경 */ }
  if (state.places) renderHero();
  if (state.lastPlan) drawMap(state.lastPlan.vehicles);
}
$('darkmode').addEventListener('click', () => setDark(!document.documentElement.classList.contains('dark')));
try { if (localStorage.getItem('hyodo-dark') === '1') setDark(true); } catch (e) { /* 무시 */ }

// ---------- 탭 ----------
document.querySelectorAll('.tab').forEach((btn) => btn.addEventListener('click', () => showTab(btn.dataset.tab)));
$('staff-link').addEventListener('click', () => showTab('dash'));
function showTab(name) {
  document.querySelectorAll('.tab').forEach((b) => {
    const on = b.dataset.tab === name;
    b.classList.toggle('active', on);
    b.setAttribute('aria-selected', on);
  });
  document.querySelectorAll('.panel').forEach((p) => p.classList.toggle('active', p.id === `tab-${name}`));
  document.body.dataset.tab = name;
  window.scrollTo({ top: 0 });
  $('cta-bar').classList.toggle('hidden', name !== 'book' || !['when', 'confirm'].includes(state.screen));
  if (name === 'dash') {
    setTimeout(() => {
      if (!state.map) return;
      state.map.invalidateSize();
      if (state.lastPlan) drawMap(state.lastPlan.vehicles);
    }, 60);
  }
}

// ---------- 단계 이동 ----------
function go(screen) {
  state.screen = screen;
  document.body.dataset.screen = screen;
  document.querySelectorAll('.screen').forEach((s) => s.classList.toggle('active', s.id === `s-${screen}`));
  const order = ['where', 'when', 'confirm', 'done'];
  document.querySelectorAll('#steps li').forEach((li) => {
    const i = order.indexOf(li.dataset.step);
    li.classList.toggle('on', i === order.indexOf(screen));
    li.classList.toggle('past', i < order.indexOf(screen));
  });
  $('cta-bar').classList.toggle('hidden', !['when', 'confirm'].includes(screen));
  if (screen === 'when') { renderWhen(); $('cta').textContent = '다음'; }
  if (screen === 'confirm') { renderConfirm(); $('cta').textContent = state.mode === 'hall' ? `${riders().length}분 예약할게요` : '예약할게요'; }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
document.querySelectorAll('[data-go]').forEach((b) => b.addEventListener('click', () => go(b.dataset.go)));

$('cta').addEventListener('click', () => {
  if (state.screen === 'when') {
    const b = state.booking;
    if (!b.date || !b.time) {
      $('clarify').textContent = !b.date ? '날짜를 골라 주세요' : '도착 시각을 골라 주세요';
      $('clarify').classList.remove('hidden');
      return;
    }
    go('confirm');
  } else if (state.screen === 'confirm') {
    book();
  }
});

// ---------- 1. 누가 · 어디로 ----------
// 이미 예약이 들어간 분은 경로당 명단에서 뺌
function bookedNames() {
  return new Set(state.requests.map((r) => r.name));
}
function riders() {
  if (state.mode === 'self') return [profile()];
  return [...document.querySelectorAll('#roster input:checked')].map((i) => state.profiles[Number(i.value)]);
}

function renderProfiles() {
  $('profile').innerHTML = state.profiles
    .map((p, i) => `<option value="${i}">${p.name} 어르신 · ${placeName(p.villageId)}</option>`).join('');
  const halls = [...new Set(state.profiles.map((p) => p.villageId))];
  $('hall').innerHTML = halls.map((id) => `<option value="${id}">${hallName(id)} · ${village(id).myeon}</option>`).join('');
  state.hall = halls[0];
  refreshWhere();
}

function setMode(mode) {
  state.mode = mode;
  document.querySelectorAll('.mode-btn').forEach((b) => {
    const on = b.dataset.mode === mode;
    b.classList.toggle('active', on);
    b.setAttribute('aria-pressed', on);
  });
  $('self-box').classList.toggle('hidden', mode !== 'self');
  $('hall-box').classList.toggle('hidden', mode !== 'hall');
  refreshWhere();
}
document.querySelectorAll('.mode-btn').forEach((b) => b.addEventListener('click', () => setMode(b.dataset.mode)));
$('profile').addEventListener('change', refreshWhere);
$('hall').addEventListener('change', () => { state.hall = $('hall').value; refreshWhere(); });

function renderRoster() {
  const booked = bookedNames();
  const list = state.profiles.map((p, i) => ({ p, i })).filter(({ p }) => p.villageId === state.hall && !booked.has(p.name));
  $('roster').innerHTML = list.length
    ? list.map(({ p, i }) => `
        <label><input type="checkbox" value="${i}" />
          <span>${p.name}<small>${p.age}세 · ${p.gender}</small></span></label>`).join('')
    : '<p class="empty">이 경로당 어르신은 모두 예약돼 있어요.</p>';
}

function refreshWhere() {
  state.booking = { villageId: homeVillage() };
  if (state.mode === 'self') $('greet').textContent = '어디 가세요?';
  else { $('greet').textContent = '어르신들 어디 가세요?'; renderRoster(); }
  renderPlaces();
  renderTogether();
  renderHero();
}

// ---------- 지도 그림 (실제 정류장·병원 좌표를 그대로 투영) ----------
function tomorrow() { const d = new Date(); d.setDate(d.getDate() + 1); return ymd(d); }

function togetherGroups() {
  const home = village(homeVillage());
  const t = tomorrow();
  const groups = new Map();
  for (const r of state.requests) {
    if (r.date !== t) continue;
    const g = groups.get(r.destinationId) || { id: r.destinationId, n: 0, near: 0, times: {} };
    g.n += 1;
    if (home && village(r.villageId)?.myeon === home.myeon) g.near += 1;
    g.times[r.time] = (g.times[r.time] || 0) + 1;
    groups.set(r.destinationId, g);
  }
  return [...groups.values()].sort((a, b) => b.near - a.near || b.n - a.n).slice(0, 3);
}

function renderTogether() {
  const gs = togetherGroups();
  $('together').classList.toggle('hidden', !gs.length);
  $('together-chips').innerHTML = gs.map((g) => `
    <button type="button" class="t-chip" data-id="${g.id}">${destName(g.id)}<small>이웃 ${g.n}분</small></button>`).join('');
}
$('together-chips').addEventListener('click', (e) => {
  const btn = e.target.closest('.t-chip');
  if (!btn || needRiders()) return;
  const g = togetherGroups().find((x) => x.id === btn.dataset.id);
  const time = Object.entries(g.times).sort((a, b) => b[1] - a[1])[0][0];
  state.booking = { villageId: homeVillage(), destinationId: g.id, date: tomorrow(), time, purpose: '진료', engine: 'tap' };
  go('when');
});

function renderHero() {
  const cs = getComputedStyle(document.documentElement);
  const cv = (n) => cs.getPropertyValue(n).trim();
  const home = village(homeVillage());
  if (!home) return;
  const inCounty = togetherGroups().map((g) => dest(g.id)).filter((d) => d && d.zone !== 'outside');
  const focus = inCounty[0] || dest('daenam');
  const branch = myBranch();
  const far = (d) => Math.hypot(d.lat - home.lat, (d.lng - home.lng) * 0.81) > 0.012;
  const targets = [focus, branch].filter((d, i, a) => d && far(d) && a.findIndex((x) => x.id === d.id) === i);
  if (!targets.length) targets.push(dest('daenam'));
  const pts = [home, ...targets];
  const W = 390, H = 300, top = 92, bottom = 96, side = 80;
  const kx = Math.cos((35.7 * Math.PI) / 180);
  const minLat = Math.min(...pts.map((p) => p.lat)), maxLat = Math.max(...pts.map((p) => p.lat));
  const minLng = Math.min(...pts.map((p) => p.lng)), maxLng = Math.max(...pts.map((p) => p.lng));
  const scale = Math.min((W - side * 2) / Math.max((maxLng - minLng) * kx, 0.01), (H - top - bottom) / Math.max(maxLat - minLat, 0.01));
  const cx = (minLng + maxLng) / 2, cy = (minLat + maxLat) / 2;
  const proj = (p) => ({ x: W / 2 + (p.lng - cx) * kx * scale, y: top + (H - top - bottom) / 2 - (p.lat - cy) * scale });
  const h = proj(home);
  const pins = targets.map((d) => ({ d, ...proj(d) }));
  const dots = state.places.villages.filter((v) => v.id !== home.id).map(proj)
    .filter((p) => p.x > 10 && p.x < W - 10 && p.y > 10 && p.y < H - 40);
  const f = pins[0];
  const mx = (h.x + f.x) / 2, my = (h.y + f.y) / 2 - 26;
  const label = (text, x, y, dark) => {
    const w = text.length * 14 + 22;
    const lx = Math.min(Math.max(x - w / 2, 8), W - w - 8);
    return `<g><rect x="${lx}" y="${y}" width="${w}" height="30" rx="15" fill="${dark ? cv('--ink') : cv('--paper')}"/><text x="${lx + w / 2}" y="${y + 20}" font-size="14" font-weight="800" text-anchor="middle" fill="${dark ? cv('--paper') : cv('--ink')}">${text}</text></g>`;
  };
  $('hero').innerHTML = `
  <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid slice">
    <rect width="${W}" height="${H}" fill="${cv('--map-bg')}"/>
    <ellipse cx="60" cy="60" rx="130" ry="70" fill="${cv('--map-hill')}"/>
    <ellipse cx="360" cy="240" rx="150" ry="80" fill="${cv('--map-hill')}"/>
    <path d="M-20 ${H - 64} C 90 ${H - 110}, 170 ${H - 34}, 260 ${H - 84} S 380 ${H - 130}, ${W + 20} ${H - 100}" stroke="${cv('--map-river')}" stroke-width="16" fill="none" stroke-linecap="round"/>
    ${dots.map((p) => `<circle cx="${p.x}" cy="${p.y}" r="4" fill="${cv('--map-dot')}"/>`).join('')}
    ${pins.slice(1).map((p) => `<path d="M${h.x} ${h.y} L${p.x} ${p.y}" stroke="${cv('--paper')}" stroke-width="7" stroke-linecap="round"/>`).join('')}
    <path d="M${h.x} ${h.y} Q ${mx} ${my} ${f.x} ${f.y}" stroke="${cv('--paper')}" stroke-width="9" fill="none" stroke-linecap="round"/>
    <path d="M${h.x} ${h.y} Q ${mx} ${my} ${f.x} ${f.y}" stroke="${cv('--orange')}" stroke-width="4" fill="none" stroke-dasharray="1 9" stroke-linecap="round"/>
    <g transform="translate(${mx - 22} ${my - 10})">
      <rect x="0" y="10" width="44" height="16" rx="7" fill="${cv('--sun')}"/><path d="M8 10 L14 3 H30 L36 10Z" fill="${cv('--sun')}"/>
      <rect x="15" y="-3" width="14" height="6" rx="2" fill="${cv('--orange')}"/><rect x="12" y="5" width="8" height="5" rx="1" fill="${cv('--paper')}"/><rect x="24" y="5" width="8" height="5" rx="1" fill="${cv('--paper')}"/>
      <circle cx="11" cy="27" r="4.5" fill="${cv('--ink')}"/><circle cx="33" cy="27" r="4.5" fill="${cv('--ink')}"/>
    </g>
    ${pins.map((p, i) => `<g transform="translate(${p.x} ${p.y})"><rect x="-15" y="-15" width="30" height="30" rx="10" fill="${i === 0 ? cv('--ink') : cv('--leaf')}" stroke="${cv('--paper')}" stroke-width="3"/><text x="0" y="5" fill="${cv('--paper')}" font-size="13" font-weight="800" text-anchor="middle">${i === 0 ? 'H' : '+'}</text></g>`).join('')}
    ${pins.map((p, i) => label(p.d.name, p.x, p.y + 20, i === 0)).join('')}
    <g transform="translate(${h.x - 34} ${h.y - 30})"><circle cx="0" cy="12" r="12" fill="#5E9F57"/><rect x="-2" y="20" width="4" height="12" fill="#8A5A3A"/><circle cx="-4" cy="9" r="3" fill="#E8590C"/><circle cx="5" cy="15" r="3" fill="#E8590C"/></g>
    <circle cx="${h.x}" cy="${h.y}" r="13" fill="${cv('--orange')}" stroke="${cv('--paper')}" stroke-width="4"/>
    ${label('우리 마을', h.x, h.y + 20, false)}
  </svg>
  <div class="hero-pill">📍 ${pickupPoint(home.id)}</div>`;
}

function myBranch() {
  const v = village(homeVillage());
  return state.places.destinations.find((d) => d.kind === '보건지소' && d.myeon === v?.myeon) || dest('health-center');
}

function renderPlaces() {
  const from = village(homeVillage());
  const tiles = [
    [dest('daenam'), ICONS.hospital, '응급실 있는 병원', '병원'],
    [myBranch(), ICONS.branch, '우리 동네 진료', '보건'],
    [dest('daegu-hyosung'), ICONS.city, '대구 협약병원', '대구'],
    [dest('cheongdo-market'), ICONS.market, '장보기', '시장']
  ];
  $('places').innerHTML = tiles.map(([d, icon, sub, kind]) => `
    <button class="place" type="button" data-id="${d.id}" data-kind="${kind}">
      <span class="place-icon" aria-hidden="true">${icon}</span>
      <span><b>${d.name}</b><small>${sub}</small></span>
      <span class="eta">${from ? `약 ${minutesBetween(from, d)}분` : ''}<i aria-hidden="true">›</i></span>
    </button>`).join('');
}

function needRiders() {
  if (state.mode === 'hall' && riders().length === 0) {
    $('heard').textContent = '같이 가실 어르신을 먼저 골라 주세요.';
    $('heard').classList.remove('hidden');
    $('roster').scrollIntoView({ behavior: 'smooth', block: 'center' });
    return true;
  }
  return false;
}

$('places').addEventListener('click', (e) => {
  const btn = e.target.closest('.place');
  if (!btn || needRiders()) return;
  const d = dest(btn.dataset.id);
  state.booking = { villageId: homeVillage(), destinationId: d.id, purpose: d.kind === '전통시장' ? '장보기' : '진료', engine: 'tap' };
  go('when');
});

$('examples').innerHTML = EXAMPLES.map((e) => `<button class="chip" type="button">${e}</button>`).join('');
$('examples').addEventListener('click', (e) => {
  if (e.target.classList.contains('chip')) $('utter').value = e.target.textContent;
});
$('parse').addEventListener('click', () => parseText($('utter').value.trim()));

// 음성 인식: 말이 끝나면 자동으로 AI가 정리
const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
if (!Recognition) {
  $('mic-label').textContent = '말로 예약하기';
  $('mic-sub').textContent = '이 기기에선 음성이 안 돼요. 아래 목록을 눌러 주세요';
  $('mic').disabled = true;
} else {
  const rec = new Recognition();
  rec.lang = 'ko-KR';
  rec.interimResults = true;
  let listening = false;
  let heard = '';
  rec.onresult = (e) => {
    heard = Array.from(e.results).map((r) => r[0].transcript).join('');
    $('heard').textContent = `“${heard}”`;
    $('heard').classList.remove('hidden');
  };
  rec.onend = () => {
    listening = false;
    $('mic').classList.remove('listening');
    $('mic-label').textContent = '말로 예약하기';
    $('mic-sub').textContent = '사투리 그대로 말씀하셔도 돼요';
    if (heard) parseText(heard);
  };
  $('mic').addEventListener('click', () => {
    if (listening) return rec.stop();
    if (needRiders()) return;
    listening = true;
    heard = '';
    $('mic').classList.add('listening');
    $('mic-label').textContent = '듣고 있어요';
    $('mic-sub').textContent = '다 말씀하시면 잠깐 기다리세요';
    rec.start();
  });
}

async function parseText(text) {
  if (!text || needRiders()) return;
  $('heard').textContent = `“${text}” 정리하고 있어요`;
  $('heard').classList.remove('hidden');
  try {
    const r = await api('parse', { text, profile: { villageId: homeVillage() } });
    state.booking = { ...r, villageId: state.mode === 'hall' ? state.hall : (r.villageId || homeVillage()), text };
    $('heard').textContent = `“${text}”`;
    if (r.destinationId && r.date && r.time) go('confirm');
    else if (r.destinationId) {
      go('when');
      if (r.clarify) { $('clarify').textContent = r.clarify; $('clarify').classList.remove('hidden'); }
    } else {
      const ask = r.clarify || '어디로 가실지 다시 말씀해 주이소';
      $('heard').textContent = `“${text}” — ${ask}`;
      speak(ask);
    }
  } catch (err) {
    $('heard').textContent = '말씀을 정리하지 못했어요. 아래 목록에서 가실 곳을 눌러 주세요.';
  }
}

// ---------- 2. 언제 ----------
function renderWhen() {
  const b = state.booking;
  const who = state.mode === 'hall' ? `${riders().length}분이 ` : '';
  const kind = { daenam: ['병원', ICONS.hospital], 'daegu-hyosung': ['대구', ICONS.city], 'cheongdo-market': ['시장', ICONS.market] }[b.destinationId] || ['보건', ICONS.branch];
  $('picked-dest').dataset.kind = kind[0];
  $('picked-dest').innerHTML = `<span class="place-icon" aria-hidden="true">${kind[1]}</span>
    <span><small>${who}가시는 곳</small><b>${destName(b.destinationId)}</b></span>
    <button type="button" class="change" data-go="where">바꾸기</button>`;
  $('picked-dest').querySelector('.change').addEventListener('click', () => go('where'));
  $('clarify').classList.add('hidden');
  const days = [0, 1, 2].map((n) => { const d = new Date(); d.setDate(d.getDate() + n); return ymd(d); });
  if (b.date && !days.includes(b.date)) days.push(b.date);
  $('days').innerHTML = days.map((d) => {
    const dt = new Date(d + 'T00:00:00');
    return `<button type="button" class="opt ${b.date === d ? 'sel' : ''}" data-date="${d}" aria-pressed="${b.date === d}">
      <b>${relWord(d) || `${dt.getMonth() + 1}/${dt.getDate()}`}</b>
      <small>${dt.getMonth() + 1}월 ${dt.getDate()}일 ${WEEK[dt.getDay()]}요일</small></button>`;
  }).join('');
  const times = [...TIMES];
  if (b.time && !times.some(([t]) => t === b.time)) times.push([b.time, timeLabel(b.time)]);
  $('times').innerHTML = times.map(([t, label]) => `
    <button type="button" class="opt ${b.time === t ? 'sel' : ''}" data-time="${t}" aria-pressed="${b.time === t}">${label}</button>`).join('');
}
$('days').addEventListener('click', (e) => {
  const btn = e.target.closest('[data-date]');
  if (!btn) return;
  state.booking.date = btn.dataset.date;
  renderWhen();
});
$('times').addEventListener('click', (e) => {
  const btn = e.target.closest('[data-time]');
  if (!btn) return;
  state.booking.time = btn.dataset.time;
  renderWhen();
});

// ---------- 3. 확인 ----------
function summaryText(b) {
  const who = state.mode === 'hall' ? `${riders().map((p) => p.name.replace('○', ' ')).join(', ')} 어르신이 ` : '';
  return `${who}${dayFull(b.date)}, ${timeLabel(b.time)}까지 ${destName(b.destinationId)}에 가십니다. ${pickupPoint(b.villageId)}에서 타십니다.`;
}
function renderConfirm() {
  const b = state.booking;
  const rows = [];
  if (state.mode === 'hall') rows.push(['누가', riders().map((p) => p.name).join(', ')]);
  rows.push(['타는 곳', pickupPoint(b.villageId)], ['가는 곳', destName(b.destinationId)],
    ['날짜', dayFull(b.date)], ['도착', `${timeLabel(b.time)}까지`]);
  const d = dest(b.destinationId);
  const outside = d && d.zone === 'outside';
  rows.push(['타는 차', outside
    ? '행복택시 구간 확대(제안)<br><small>같은 방향 어르신이 모이면 대구 협약병원까지</small>'
    : `행복택시<br><small>본인 부담 ${(config.happyTaxiFare ?? 100).toLocaleString('ko-KR')}원 · 나머지는 군이 정산</small>`]);
  $('ticket').innerHTML = rows.map(([k, v]) => `<div class="t-row"><span>${k}</span><b>${v}</b></div>`).join('');
  const unregistered = riders().filter((p) => p.registered === false);
  $('reg-notice').classList.toggle('hidden', !unregistered.length);
  if (unregistered.length) {
    const myeon = village(b.villageId)?.myeon || '';
    $('reg-notice').innerHTML = `<b>${unregistered.map((p) => p.name).join(', ')} 어르신은 아직 행복택시 등록 전이에요.</b><br>
      예약하시면 ${myeon} 행정복지센터로 등록 요청도 같이 보내 드려요. 지금처럼 면사무소에 직접 가실 필요가 없어요.`;
  }
  $('consent-title').textContent = state.mode === 'hall'
    ? '어르신들께 여쭤보고 함께 타는 데 동의받았어요'
    : '같은 방향 어르신과 함께 탈게요';
  const ai = String(b.engine || '').startsWith('azure');
  $('engine').textContent = ai ? 'Azure OpenAI가 정리했어요' : '규칙 기반으로 정리했어요(AI 키 없음)';
  $('engine').classList.toggle('hidden', b.engine === 'tap' || !b.engine);
  $('engine').classList.toggle('rules', !ai);
  speak(summaryText(b) + ' 맞으시면 아래 버튼을 누르세요.');
}
$('speak').addEventListener('click', () => speak(summaryText(state.booking)));

// ---------- 4. 완료 ----------
async function book() {
  const b = state.booking;
  const people = riders();
  const groupId = state.mode === 'hall' && people.length > 1 ? `g${Date.now()}` : undefined;
  state.lastIds = people.map((p, i) => {
    const id = `u${Date.now()}${i}`;
    state.requests.push({
      id, groupId, name: p.name, age: p.age, gender: p.gender, villageId: b.villageId,
      destinationId: b.destinationId, date: b.date, time: b.time, purpose: b.purpose || '진료',
      wheelchair: Boolean(b.wheelchair), consent: $('consent').checked, registered: p.registered !== false,
      text: b.text || (state.mode === 'hall' ? `(${hallName(b.villageId)} 대리 신청)` : '(화면에서 선택)'),
      via: state.mode === 'hall' ? hallName(b.villageId) : '본인'
    });
    return id;
  });
  $('cta').disabled = true;
  $('cta').textContent = '자리 맞추는 중…';
  try { await runMatch(); } finally { $('cta').disabled = false; }
  renderDone();
  go('done');
}

function findSeat(id) {
  if (!state.lastPlan) return {};
  for (const [i, v] of state.lastPlan.vehicles.entries()) {
    const s = v.stops.find((x) => x.requestId === id);
    if (s) return { v, s, no: i + 1 };
  }
  return {};
}

// 좌석 지도: 차를 위에서 본 모습. 나=반시, 함께=잎 연두, 빈 자리=회색
function seatMapSVG(v, myId) {
  const isVan = v.vehicle === 'van' || v.capacity === 6;
  const taken = new Map(v.stops.map((s) => [s.seat, s.requestId === myId ? 'me' : 'co']));
  const fill = (state) => (state === 'me' ? 'var(--red)' : state === 'co' ? 'var(--green)' : 'var(--seat)');
  const label = (state) => (state === 'me' ? '나' : '');
  let seats;
  let h;
  if (!isVan) {
    h = 200;
    seats = [
      { x: 24, y: 64, driver: true },
      { x: 62, y: 64, name: '조수석' },
      { x: 24, y: 124, name: '뒷좌석 왼쪽' },
      { x: 62, y: 124, name: '뒷좌석 오른쪽(문 쪽)' }
    ];
  } else {
    h = 250;
    seats = [{ x: 24, y: 52, driver: true }];
    for (let i = 0; i < 6; i++) seats.push({ x: i % 2 ? 62 : 24, y: 104 + Math.floor(i / 2) * 46, name: `${i + 1}번 좌석` });
  }
  const parts = seats.map((s, idx) => {
    if (s.driver) {
      return `<rect x="${s.x}" y="${s.y}" width="34" height="38" rx="9" fill="var(--driver)"/>
        <text x="${s.x + 17}" y="${s.y + 24}" text-anchor="middle" font-size="11" font-weight="800" fill="var(--on-driver)">기사</text>`;
    }
    const st = taken.get(s.name) || '';
    const pop = st ? ` class="s-pop" style="animation-delay:${0.15 * idx}s"` : '';
    return `<g${pop}><rect x="${s.x}" y="${s.y}" width="34" height="38" rx="9" fill="${fill(st)}"/>
      <text x="${s.x + 17}" y="${s.y + 24}" text-anchor="middle" font-size="13" font-weight="800" fill="${st === 'me' ? 'var(--on-primary)' : st === 'co' ? 'var(--on-co)' : 'var(--ink)'}">${label(st)}</text></g>`;
  }).join('');
  return `<svg viewBox="0 0 120 ${h}" role="img" aria-label="좌석 지도: 내 자리 ${taken.size ? '' : ''}">
    <rect x="8" y="8" width="104" height="${h - 16}" rx="34" fill="var(--paper)" stroke="var(--rule)" stroke-width="3"/>
    <path d="M24 42 Q60 26 96 42" fill="none" stroke="var(--rule)" stroke-width="3" stroke-linecap="round"/>
    ${parts}</svg>`;
}

function renderDone() {
  const hall = state.mode === 'hall';
  $('hall-list').classList.toggle('hidden', !hall);
  $('pickup').classList.toggle('hidden', hall);
  const timeline = (v, s, no, others) => [
    ['예약 접수', hall ? `${hallName(state.booking.villageId)}에서 신청` : '방금', true],
    [others ? `합승 배차 완료 · ${no}호차` : `배차 완료 · ${no}호차`, others ? `같은 방향 ${others}분과 함께 가요` : '', true],
    ['기사님 확인 전화', `타시기 1시간 전에 ${v.service.driver} 기사님이 전화드려요`, false],
    ['탑승', `${timeLabel(s.pickup)} · ${pickupPoint(state.booking.villageId)}`, false],
    ['도착', `${v.destination.name} ${timeLabel(v.arriveBy)}까지`, false]
  ];

  if (!hall) {
    const { v, s, no } = findSeat(state.lastIds[0]);
    if (!v) { $('pickup').innerHTML = '<p>자리를 맞추고 있어요.</p>'; return; }
    const others = s.coRiders.length;
    const empty = v.capacity - v.seatsUsed;
    $('pickup').innerHTML = `
      <p class="p-label">${dayFull(v.date)}</p>
      <p class="p-time">${timeLabel(s.pickup)}</p>
      <p class="p-where">${pickupPoint(state.booking.villageId)}에서 타세요</p>
      <div class="pass-cut" aria-hidden="true"></div>
      <div class="seatmap">
        ${seatMapSVG(v, state.lastIds[0])}
        <ul class="seat-legend">
          <li class="me"><i></i>내 자리 · ${s.seat}</li>
          <li class="co"><i></i>함께 타는 분 ${others}명</li>
          <li class="empty"><i></i>빈 자리 ${empty}개</li>
        </ul>
      </div>
      <div class="pass-cut" aria-hidden="true"></div>
      <div class="p-grid">
        <div><span>차</span><b>${no}호차 ${v.type}</b></div>
        <div><span>도착</span><b>${v.destination.name}<br>${timeLabel(v.arriveBy)}까지</b></div>
        <div class="wide"><span>요금</span><b>${v.service.fareText}</b></div>
      </div>`;
    $('timeline').innerHTML = timeline(v, s, no, others).map(([t, d, done]) =>
      `<li class="${done ? 'done' : ''}"><b>${t}</b>${d ? `<small>${d}</small>` : ''}</li>`).join('');
    $('share').textContent = '👪 가족에게 알리기';
    speak(`예약됐습니다. ${dayFull(v.date)}, ${timeLabel(s.pickup)}에 ${pickupPoint(state.booking.villageId)}에서 타세요. ${s.seat}에 앉으시면 됩니다.`);
  } else {
    const rows = state.lastIds.map((id) => ({ id, ...findSeat(id) })).filter((r) => r.v);
    const req = (id) => state.requests.find((r) => r.id === id);
    $('hall-list').innerHTML = `<p class="p-label" style="padding-top:14px">${dayFull(state.booking.date)} · ${pickupPoint(state.booking.villageId)}</p>` +
      rows.map(({ id, v, s, no }) => `
        <div class="hall-row"><b>${req(id).name} 어르신</b><time>${timeLabel(s.pickup)}</time>
          <small>${no}호차 · ${s.seat} · ${v.destination.name} ${timeLabel(v.arriveBy)}까지${s.coRiders.length ? ` · 함께 ${s.coRiders.length}명` : ''}</small></div>`).join('');
    const first = rows[0];
    const cars = [...new Set(rows.map((r) => r.no))];
    const steps = first ? timeline(first.v, first.s, first.no, 0) : [];
    if (first) {
      steps[1] = [`배차 완료 · ${cars.map((n) => `${n}호차`).join(', ')}`,
        cars.length === 1 ? `${rows.length}분 모두 같은 차로 가세요` : `${cars.length}대로 나눠 가세요`, true];
    }
    $('timeline').innerHTML = steps.map(([t, d, done]) =>
      `<li class="${done ? 'done' : ''}"><b>${t}</b>${d ? `<small>${d}</small>` : ''}</li>`).join('');
    $('share').textContent = '📋 명단 보내기 (가족·담당자)';
    speak(`${rows.length}분 예약됐습니다. ${dayFull(state.booking.date)}, ${pickupPoint(state.booking.villageId)}에서 타십니다.`);
  }
}

$('share').addEventListener('click', async () => {
  const lines = state.lastIds.map((id) => {
    const { v, s, no } = findSeat(id);
    const r = state.requests.find((x) => x.id === id);
    return v ? `${r.name} 어르신 ${timeLabel(s.pickup)} 탑승 (${no}호차, ${s.seat}) → ${v.destination.name} ${timeLabel(v.arriveBy)}까지` : '';
  }).filter(Boolean);
  const text = `[청도효도택시] ${dayFull(state.booking.date)} · ${pickupPoint(state.booking.villageId)}\n${lines.join('\n')}`;
  try {
    if (navigator.share) await navigator.share({ title: '청도효도택시 예약', text });
    else location.href = `sms:?&body=${encodeURIComponent(text)}`;
  } catch (e) { /* 공유 취소 */ }
});
$('again').addEventListener('click', () => { refreshWhere(); $('heard').classList.add('hidden'); go('where'); });

if (config.callCenter) {
  $('callcenter').href = `tel:${config.callCenter}`;
  $('callcenter').classList.remove('hidden');
}

// ---------- 배차 담당 화면 ----------
function renderRequests() {
  $('req-count').textContent = state.requests.length;
  $('m-req').textContent = state.requests.length;
  $('requests').innerHTML = state.requests.map((r, i) => `
    <li>
      <div><b>${r.name}</b> · ${placeName(r.villageId)} → ${destName(r.destinationId)}
        <small>${dateLabel(r.date)} ${r.time}까지 · ${r.gender} · ${r.consent ? '합승 동의' : '혼자 타기'}${r.registered === false ? ' · 등록 요청 함께 접수' : ''}${r.via ? ` · 신청: ${r.via}` : ''}${r.text ? ` · "${r.text}"` : ''}</small></div>
      <button type="button" data-i="${i}" aria-label="예약 삭제">✕</button>
    </li>`).join('');
}
$('requests').addEventListener('click', (e) => {
  if (e.target.dataset.i === undefined) return;
  state.requests.splice(Number(e.target.dataset.i), 1);
  runMatch();
});
$('load-demo').addEventListener('click', async () => {
  const { requests } = await api('demo');
  state.requests = requests;
  runMatch();
});
$('clear').addEventListener('click', () => { state.requests = []; runMatch(); });
$('match').addEventListener('click', runMatch);
$('vans').addEventListener('change', runMatch);

async function runMatch() {
  renderRequests();
  if (!state.requests.length) {
    state.lastPlan = null;
    $('vehicles').innerHTML = '';
    $('compare').innerHTML = '';
    ['m-trips', 'm-km', 'm-fill'].forEach((id) => ($(id).textContent = '–'));
    if (state.layer) state.layer.clearLayers();
    return;
  }
  const out = await api('match', { requests: state.requests, fleet: { van: Number($('vans').value) } });
  state.lastPlan = out;
  const s = out.summary;
  $('m-trips').textContent = `${s.tripsBefore}→${s.tripsAfter}`;
  $('m-km').textContent = `${s.kmSavedPct}%`;
  $('m-fill').textContent = `${s.seatFillPct}%`;
  $('compare').innerHTML = `
    <div><span>지금 방식</span><b>기사님께 한 분씩 전화 · ${s.settleBefore}번 운행 · 군 정산 ${s.settleBefore}건</b></div>
    <div class="now"><span>청도효도택시</span><b>AI 통합 배차 · ${s.settleAfter}번 운행 · 군 정산 ${s.settleAfter}건</b></div>
    ${s.outsideRiders ? `<p>대구 병원 가시는 ${s.outsideRiders}분은 지금은 행복택시 구간 밖이에요. 함께 모이면 구간을 넓히자는 제안으로 묶었어요.</p>` : ''}`;
  renderVehicles(out.vehicles);
  drawMap(out.vehicles);
}

function renderVehicles(vehicles) {
  $('vehicles').innerHTML = vehicles.map((v, i) => `
    <article class="card vehicle" style="--c:${COLORS[i % COLORS.length]}">
      <h3>${i + 1}호차 · ${v.type}</h3>
      <p class="meta">${dateLabel(v.date)} ${v.arriveBy}까지 ${v.destination.name}<br>${v.km}km · ${v.seatsUsed}/${v.capacity}석</p>
      <span class="rule">${v.rule}</span>
      <p class="meta">${v.service.driver} · ${v.service.fareText}</p>
      <ol class="stops">
        ${v.stops.map((s) => `
          <li class="${state.lastIds.includes(s.requestId) ? 'mine' : ''}"><time>${s.pickup}</time>
            <div>${s.name} · ${s.village}
              <small>${s.seat}${s.coRiders.length ? ` · 함께 타는 분 ${s.coRiders.length}명` : ''}</small></div></li>`).join('')}
      </ol>
      <a class="sos" href="tel:112">긴급신고 112</a>
    </article>`).join('');
}

function initMap() {
  if (!window.L) {
    $('map').innerHTML = '<p class="hint" style="padding:16px">지도를 불러오지 못했어요. 배차 결과는 아래 카드에서 볼 수 있어요.</p>';
    return;
  }
  state.map = L.map('map', { scrollWheelZoom: false }).setView([35.67, 128.76], 11);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 18, attribution: '&copy; OpenStreetMap'
  }).addTo(state.map);
  state.layer = L.layerGroup().addTo(state.map);
}

function pin(label, color, cls = '') {
  return L.divIcon({ className: '', html: `<div class="pin ${cls}" style="background:${color}">${label}</div>`, iconSize: [30, 30], iconAnchor: [15, 15] });
}

function drawMap(vehicles) {
  if (!state.layer) return;
  state.layer.clearLayers();
  const bounds = [];
  const seenDest = new Set();
  vehicles.forEach((v, i) => {
    const color = COLORS[i % COLORS.length];
    const path = v.stops.map((s) => [s.lat, s.lng]).concat([[v.destination.lat, v.destination.lng]]);
    L.polyline(path, { color, weight: 5, opacity: 0.85 }).addTo(state.layer);
    v.stops.forEach((s, j) => {
      L.marker([s.lat, s.lng], { icon: pin(`${i + 1}`, color) })
        .bindPopup(`<b>${i + 1}호차 ${j + 1}번째 탑승</b><br>${s.name} · ${s.village}<br>${s.pickup} 탑승 · ${s.seat}` +
          (s.coRiders.length ? `<br><small>함께: ${s.coRiders.join(', ')}</small>` : ''))
        .addTo(state.layer);
      bounds.push([s.lat, s.lng]);
    });
    if (!seenDest.has(v.destination.id)) {
      seenDest.add(v.destination.id);
      L.marker([v.destination.lat, v.destination.lng], { icon: pin('H', getComputedStyle(document.documentElement).getPropertyValue('--ink').trim(), 'dest') })
        .bindPopup(`<b>${v.destination.name}</b>`).addTo(state.layer);
      bounds.push([v.destination.lat, v.destination.lng]);
    }
  });
  if (bounds.length) state.map.fitBounds(bounds, { padding: [30, 30] });
}

// ---------- 시작 ----------
(async function init() {
  initMap();
  try {
    const [places, demo] = await Promise.all([api('places'), api('demo')]);
    state.places = places;
    state.profiles = demo.profiles || demo.requests.map(({ name, age, gender, villageId }) => ({ name, age, gender, villageId }));
    // 이웃 어르신들의 예약이 이미 들어와 있는 상태로 시작 → 새 예약이 바로 합승으로 묶임
    state.requests = demo.requests;
    renderProfiles();
    go('where');
    runMatch();
  } catch (err) {
    document.querySelector('main').insertAdjacentHTML('afterbegin',
      `<div class="card clarify">서버에 연결하지 못했어요 (${err.message}). 배포 상태를 확인해 주세요.</div>`);
  }
})();
