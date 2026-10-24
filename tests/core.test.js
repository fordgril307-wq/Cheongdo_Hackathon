const test = require('node:test');
const assert = require('node:assert');
const { ruleParse, parseTime } = require('../api/src/lib/parser');
const { matchRides } = require('../api/src/lib/matcher');
const { demoRequests } = require('../api/src/lib/demo');

test('사투리 문장에서 목적지·날짜·시간을 뽑는다', () => {
  const r = ruleParse('내일 아침 열 시까지 대구 효성병원 가야 되는데예', { villageId: 'unmun-sinwon' });
  assert.equal(r.destinationId, 'daegu-hyosung');
  assert.equal(r.villageId, 'unmun-sinwon');
  assert.equal(r.time, '10:00');
  assert.ok(r.date);
  assert.equal(r.clarify, null);
});

test('보건지소는 그 마을 면의 보건지소로 간다', () => {
  const r = ruleParse('모레 보건지소 물리치료 받으러 갈라카는데', { villageId: 'geumcheon-sinji' });
  assert.equal(r.destinationId, 'geumcheon-branch');
  assert.equal(r.purpose, '물리치료·한방');
});

test('"아침"을 한방 침으로 오해하지 않는다', () => {
  const r = ruleParse('내일 아침 열 시까지 대구 효성병원 가야 되는데예', { villageId: 'unmun-sinwon' });
  assert.equal(r.purpose, '진료');
});

test('정보가 빠지면 되묻는다', () => {
  const r = ruleParse('병원 좀 가야 되는데', {});
  assert.ok(r.clarify.includes('며칠'));
});

test('시간 표현', () => {
  assert.equal(parseTime('세 시'), '15:00');
  assert.equal(parseTime('오후 2시 반'), '14:30');
  assert.equal(parseTime('아침에'), '09:00');
});

test('합승 배차: 차량 수가 줄고, 세단은 같은 성별만, 미동의자는 단독', () => {
  const reqs = demoRequests();
  const out = matchRides(reqs, { van: 1 });
  assert.equal(out.summary.requests, 12);
  assert.ok(out.summary.tripsAfter < out.summary.tripsBefore, '차량 수 감소');
  for (const v of out.vehicles) {
    if (v.capacity === 3 && v.stops.length > 1) {
      const genders = new Set(v.stops.map((s) => s.gender));
      assert.equal(genders.size, 1, '모르는 사람끼리 세단은 같은 성별');
    }
    assert.ok(v.stops.length <= v.capacity);
  }
  const noConsent = out.vehicles.find((v) => v.stops.some((s) => s.requestId === 'r8'));
  assert.equal(noConsent.stops.length, 1, '미동의자는 단독');
  console.log(JSON.stringify(out.summary));
});

test('픽업 시각은 도착 시각보다 앞선다', () => {
  const out = matchRides(demoRequests(), { van: 1 });
  for (const v of out.vehicles) {
    for (const s of v.stops) assert.ok(s.pickup < v.arriveBy, `${s.pickup} < ${v.arriveBy}`);
  }
});

test('경로당 일행은 성별이 달라도 한 차에, 모르는 사람과는 성별 규칙 유지', () => {
  const reqs = demoRequests();
  const d = reqs[0].date;
  reqs.push(
    { id: 'g1', groupId: 'hall-1', name: '문○자', age: 79, gender: '여', villageId: 'unmun-daecheon', destinationId: 'daenam', date: d, time: '09:00', consent: true },
    { id: 'g2', groupId: 'hall-1', name: '하○수', age: 81, gender: '남', villageId: 'unmun-daecheon', destinationId: 'daenam', date: d, time: '09:00', consent: true }
  );
  const out = matchRides(reqs, { van: 1 });
  const carOf = (id) => out.vehicles.find((v) => v.stops.some((s) => s.requestId === id));
  assert.equal(carOf('g1').id, carOf('g2').id, '일행은 같은 차');
  const party = carOf('g1');
  assert.ok(party.stops.every((s) => ['g1', 'g2'].includes(s.requestId)) || party.capacity === 6, '혼성 일행 세단엔 모르는 사람 안 태움');
});

test('통합 배차가 현행(한 분씩 운행)보다 운행·정산 건수를 줄이고, 군 안은 행복택시로 표시', () => {
  const out = matchRides(demoRequests(), { van: 0 });
  assert.ok(out.summary.settleAfter < out.summary.settleBefore);
  const county = out.vehicles.filter((v) => v.destination.id !== 'daegu-hyosung');
  assert.ok(county.every((v) => v.service.kind === 'happy'));
  const daegu = out.vehicles.filter((v) => v.destination.id === 'daegu-hyosung');
  assert.ok(daegu.every((v) => v.service.kind === 'extended'), '대구는 구간 확대 제안');
});
