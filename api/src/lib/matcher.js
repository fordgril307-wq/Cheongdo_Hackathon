// 합승 배차: 같은 날·같은 목적지·비슷한 시간의 예약을 한 차로 묶는다.
// 합법 합승 조건(택시발전법 시행규칙, 2022.6.15 시행)을 배차 규칙으로 반영:
//  - 사전 동의한 승객만 합승
//  - 세단(일반 택시)은 같은 성별끼리만, 승합(대형) 차량은 성별 제한 없음
//  - 탑승 위치·시각과 좌석을 미리 안내 (결과에 포함)
// 청도 행복택시 개선안 (현행 운영 방식은 HS 제공 정보, 군 확인 필요):
//  - 현행: 마을 전담 기사에게 개별 전화, 한 분씩 운행, 구간은 읍·면 소재지·군청 소재지까지
//  - 개선: AI 통합 배차로 면 경계를 넘어 같은 방향끼리 합승 → 운행·정산 건수 감소
//  - 제안: 군 밖(대구 협약병원)은 합승으로 모일 때 행복택시 구간 확대
// ※ 정책 해석은 데모용이며, 실제 운영 전 관할관청 확인이 필요합니다.

const { findVillage, findDestination, zoneOf } = require('./places');

const HAPPY_TAXI_FARE = Number(process.env.HAPPY_TAXI_FARE || 100); // 본인 부담금(원), HS 제공 값 — 군 확인 필요

const ROAD_FACTOR = 1.35; // 직선거리 → 도로거리 보정
const SPEED_KMH = 40; // 농촌 평균 주행속도 가정
const BUFFER_MIN = 15; // 병원 도착 여유
const TIME_WINDOW_MIN = 60; // 도착 시각 차이 허용 범위
const MAX_DETOUR_KM = 8; // 한 명 더 태우려고 돌아가는 최대 거리
const SEDAN_CAP = 3; // 택시 정원은 4명(관광택시 1대 최대 4인)이지만 어르신 편의상 뒷좌석 가운데는 비움
const VAN_CAP = 6;
const SEDAN_SEATS = ['뒷좌석 오른쪽(문 쪽)', '뒷좌석 왼쪽', '조수석'];

function km(a, b) {
  const R = 6371;
  const toRad = (x) => (x * Math.PI) / 180;
  const dLat = toRad(b.lat - a.lat);
  const dLng = toRad(b.lng - a.lng);
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h)) * ROAD_FACTOR;
}

function toMin(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function toHHMM(min) {
  const m = ((Math.round(min) % 1440) + 1440) % 1440;
  return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
}

function enrich(r) {
  const v = findVillage(r.villageId);
  const d = findDestination(r.destinationId);
  if (!v || !d || !r.date || !r.time) return null;
  return { ...r, village: v, dest: d, zone: zoneOf(d), pos: { lat: v.lat, lng: v.lng }, arriveMin: toMin(r.time) };
}

function routeLength(stops, dest) {
  let total = 0;
  for (let i = 0; i < stops.length - 1; i++) total += km(stops[i].pos, stops[i + 1].pos);
  return total + km(stops[stops.length - 1].pos, dest);
}

// 목적지에서 먼 사람부터 태우고, 이후 가장 가까운 다음 사람을 태우는 순서
function orderStops(riders, dest) {
  const rest = [...riders].sort((a, b) => km(b.pos, dest) - km(a.pos, dest));
  const route = [rest.shift()];
  while (rest.length) {
    const last = route[route.length - 1];
    rest.sort((a, b) => km(last.pos, a.pos) - km(last.pos, b.pos));
    route.push(rest.shift());
  }
  return route;
}

function ruleText(type, riders) {
  if (riders.length === 1) return '단독 운행';
  const groups = new Set(riders.map((r) => r.groupId || r.id));
  if (groups.size === 1) return '일행 함께 탑승';
  if (type !== 'sedan') return '승합 차량 · 성별 제한 없음';
  const genders = new Set(riders.map((r) => r.gender));
  return genders.size === 1 ? `같은 성별(${riders[0].gender}) 합승` : '일행 + 같은 성별 합승';
}

function serviceOf(type, riders) {
  const zone = riders[0].zone;
  const myeons = [...new Set(riders.map((r) => r.village.myeon))];
  const fareText = `본인 부담 ${HAPPY_TAXI_FARE.toLocaleString('ko-KR')}원 · 나머지는 군이 정산`;
  if (zone === 'county') {
    return {
      kind: 'happy',
      label: type === 'van' ? '승합 행복택시(제안)' : '행복택시',
      driver: myeons.length === 1 ? `${myeons[0]} 행복택시` : `통합 배차 행복택시(${myeons.join('·')})`,
      fare: HAPPY_TAXI_FARE,
      fareText
    };
  }
  return {
    kind: 'extended',
    label: '행복택시 구간 확대(제안)',
    driver: myeons.length === 1 ? `${myeons[0]} 행복택시` : `통합 배차 행복택시(${myeons.join('·')})`,
    fare: HAPPY_TAXI_FARE,
    fareText: riders.length > 1 ? `${riders.length}분이 차 1대로 함께 가요 (구간 확대 제안)` : '혼자 가시면 지금은 구간 밖이에요 (구간 확대 제안)'
  };
}

function buildVehicle(type, riders, dest, index) {
  const stops = orderStops(riders, dest);
  const totalKm = routeLength(stops, dest);
  const arrive = Math.min(...riders.map((r) => r.arriveMin)) - BUFFER_MIN;
  // 역산: 도착 시각에서 각 구간 주행시간을 빼서 픽업 시각 계산
  let t = arrive - (km(stops[stops.length - 1].pos, dest) / SPEED_KMH) * 60;
  const times = new Array(stops.length);
  times[stops.length - 1] = t;
  for (let i = stops.length - 2; i >= 0; i--) {
    const leg = km(stops[i].pos, stops[i + 1].pos);
    t -= leg > 0.05 ? (leg / SPEED_KMH) * 60 + 3 : 0; // 다른 곳이면 승차 3분, 같은 곳이면 같은 시각
    times[i] = t;
  }
  const bySeatPriority = [...stops].sort((a, b) => (b.wheelchair ? 1 : 0) - (a.wheelchair ? 1 : 0) || (b.age || 0) - (a.age || 0));
  const seatOf = new Map();
  if (type === 'sedan') bySeatPriority.forEach((r, i) => seatOf.set(r.id, SEDAN_SEATS[i]));
  else stops.forEach((r, i) => seatOf.set(r.id, r.wheelchair ? '휠체어 자리' : `${i + 1}번 좌석`));
  const service = serviceOf(type, riders);
  return {
    id: `car-${index}`,
    type: service.label,
    vehicle: type,
    service,
    destination: { id: dest.id, name: dest.name, lat: dest.lat, lng: dest.lng },
    date: riders[0].date,
    arriveBy: toHHMM(arrive + BUFFER_MIN),
    km: Math.round(totalKm * 10) / 10,
    seatsUsed: riders.length,
    capacity: type === 'sedan' ? SEDAN_CAP : VAN_CAP,
    rule: ruleText(type, riders),
    stops: stops.map((r, i) => ({
      requestId: r.id,
      name: r.name,
      gender: r.gender,
      village: r.village.name,
      stop: r.village.stop,
      lat: r.pos.lat,
      lng: r.pos.lng,
      pickup: toHHMM(times[i]),
      seat: seatOf.get(r.id),
      needsRegistration: r.zone === 'county' && r.registered === false,
      // 함께 타는 사람 정보 공개(합승 조건): 이름은 가리고 마을·시각만
      coRiders: stops.filter((o) => o.id !== r.id).map((o) => `${o.village.stop} 정류장 ${toHHMM(times[stops.indexOf(o)])} 탑승`)
    }))
  };
}

// 같은 일행(경로당에서 함께 신청 등)은 한 단위로 묶음.
// 일행끼리는 아는 사이라 성별 규칙은 '모르는 사람과의 합승'에만 적용.
function toUnits(riders) {
  const map = new Map();
  for (const r of riders) {
    const key = r.groupId || r.id;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(r);
  }
  return [...map.values()].map((members) => {
    const genders = new Set(members.map((m) => m.gender));
    return {
      members,
      size: members.length,
      pos: members[0].pos,
      gender: genders.size === 1 ? members[0].gender : 'mixed',
      wheelchair: members.some((m) => m.wheelchair),
      consent: members.every((m) => m.consent)
    };
  });
}

const sizeOf = (units) => units.reduce((s, u) => s + u.size, 0);
const flat = (units) => units.flatMap((u) => u.members);

// 한 그룹(같은 날·목적지·시간대) 안에서 차량 묶기
function packGroup(group, dest, fleet, out) {
  const units = toUnits(group);
  const solo = units.filter((u) => !u.consent);
  let pool = units.filter((u) => u.consent);
  const far = (a, b) => km(b.pos, dest) - km(a.pos, dest);

  // 1) 승합차가 있으면 4명 이상 모인 묶음을 먼저 승합에 태움(성별 무관)
  while (fleet.van > 0 && sizeOf(pool) >= 4) {
    const seed = [...pool].sort(far)[0];
    const near = pool.filter((u) => u !== seed && km(seed.pos, u.pos) <= MAX_DETOUR_KM * 2)
      .sort((a, b) => km(seed.pos, a.pos) - km(seed.pos, b.pos));
    const picked = [seed];
    for (const u of near) if (sizeOf(picked) + u.size <= VAN_CAP) picked.push(u);
    if (sizeOf(picked) < 4) break;
    out.push(['van', flat(picked), dest]);
    fleet.van -= 1;
    pool = pool.filter((u) => !picked.includes(u));
  }

  // 2) 나머지는 세단: 모르는 사람끼리는 같은 성별만, 최대 3명
  while (pool.length) {
    const seed = [...pool].sort(far)[0];
    const picked = [seed];
    if (seed.gender !== 'mixed' && !seed.wheelchair) {
      const base = routeLength([seed.members[0]], dest);
      const candidates = pool
        .filter((u) => u !== seed && u.gender === seed.gender && !u.wheelchair)
        .map((u) => ({ u, detour: routeLength(orderStops([seed.members[0], u.members[0]], dest), dest) - base }))
        .filter((c) => c.detour <= MAX_DETOUR_KM)
        .sort((a, b) => a.detour - b.detour);
      for (const { u } of candidates) if (sizeOf(picked) + u.size <= SEDAN_CAP) picked.push(u);
    }
    // 일행이 세단 정원을 넘으면 나눠 태움
    const riders = flat(picked);
    for (let i = 0; i < riders.length; i += SEDAN_CAP) out.push(['sedan', riders.slice(i, i + SEDAN_CAP), dest]);
    pool = pool.filter((u) => !picked.includes(u));
  }
  solo.forEach((u) => {
    for (let i = 0; i < u.members.length; i += SEDAN_CAP) out.push(['sedan', u.members.slice(i, i + SEDAN_CAP), dest]);
  });
}

function matchRides(requests, fleet = { van: 0 }) {
  const valid = requests.map(enrich).filter(Boolean);
  const skipped = requests.length - valid.length;
  const fleetLeft = { van: fleet.van || 0 };

  // 날짜+목적지별로 나누고(면 경계 없이 통합 배차), 그 안에서 도착 시각 기준 시간창으로 다시 나눔
  const byKey = new Map();
  for (const r of valid) {
    const k = `${r.date}|${r.dest.id}`;
    if (!byKey.has(k)) byKey.set(k, []);
    byKey.get(k).push(r);
  }
  const plans = [];
  for (const group of byKey.values()) {
    group.sort((a, b) => a.arriveMin - b.arriveMin);
    let window = [];
    for (const r of group) {
      if (window.length && r.arriveMin - window[0].arriveMin > TIME_WINDOW_MIN) {
        packGroup(window, window[0].dest, fleetLeft, plans);
        window = [];
      }
      window.push(r);
    }
    if (window.length) packGroup(window, window[0].dest, fleetLeft, plans);
  }

  const vehicles = plans.map(([type, riders, dest], i) => buildVehicle(type, riders, dest, i + 1));
  const soloKm = valid.reduce((s, r) => s + km(r.pos, r.dest), 0);
  const pooledKm = vehicles.reduce((s, v) => s + v.km, 0);
  const seatsOffered = vehicles.reduce((s, v) => s + v.capacity, 0);
  return {
    vehicles,
    summary: {
      requests: valid.length,
      skipped,
      tripsBefore: valid.length,
      tripsAfter: vehicles.length,
      tripsSaved: valid.length - vehicles.length,
      kmBefore: Math.round(soloKm * 10) / 10,
      kmAfter: Math.round(pooledKm * 10) / 10,
      kmSavedPct: soloKm ? Math.round((1 - pooledKm / soloKm) * 100) : 0,
      seatFillPct: seatsOffered ? Math.round((valid.length / seatsOffered) * 100) : 0,
      // 현행: 한 분씩 따로 운행 → 운행마다 군 정산 1건
      settleBefore: valid.length,
      settleAfter: vehicles.length,
      outsideRiders: valid.filter((r) => r.zone === 'outside').length
    }
  };
}

module.exports = { matchRides, km };
