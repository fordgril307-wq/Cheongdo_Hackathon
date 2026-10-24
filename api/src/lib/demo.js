// 대시보드 시연용 예약 샘플 (가상 인물)
const { todayKST, fmt, addDays } = require('./dates');

function demoRequests() {
  const d = fmt(addDays(todayKST(), 1));
  return [
    { id: 'r1', name: '김○순', age: 84, gender: '여', villageId: 'unmun-sinwon', destinationId: 'daegu-hyosung', date: d, time: '10:00', purpose: '진료', consent: true, text: '내일 아침 열 시까지 대구 효성병원 가야 되는데예' },
    { id: 'r2', name: '박○자', age: 79, gender: '여', villageId: 'unmun-bangeum', destinationId: 'daegu-hyosung', date: d, time: '10:30', purpose: '검사', consent: true, text: '낼 대구 병원 피검사 받으러 갈라카는데' },
    { id: 'r3', name: '이○희', age: 81, gender: '여', villageId: 'geumcheon-donggok', destinationId: 'daegu-hyosung', date: d, time: '10:00', purpose: '진료', consent: true, text: '내일 효성병원 열 시 예약이라 카던데' },
    { id: 'r4', name: '최○식', age: 86, gender: '남', villageId: 'unmun-daecheon', destinationId: 'daegu-hyosung', date: d, time: '11:00', purpose: '진료', consent: true, text: '대구 병원 내일 열한 시까지 가야 된다' },
    { id: 'r5', name: '정○수', age: 77, gender: '남', villageId: 'geumcheon-bangji', destinationId: 'daegu-hyosung', date: d, time: '10:40', purpose: '약 처방', consent: true, text: '내일 대구 가서 약 타야 되는데' },
    { id: 'r6', name: '강○분', age: 88, gender: '여', villageId: 'maejeon-onmak', destinationId: 'daenam', date: d, time: '09:30', purpose: '진료', consent: true, text: '삭신이 쑤시가 내일 아침 대남병원 함 가볼라꼬' },
    { id: 'r7', name: '조○례', age: 82, gender: '여', villageId: 'maejeon-hapyeong', destinationId: 'daenam', date: d, time: '09:00', purpose: '진료', consent: true, text: '내일 아홉 시까지 청도병원' },
    { id: 'r8', name: '윤○덕', age: 90, gender: '여', villageId: 'geumcheon-sinji', destinationId: 'daenam', date: d, time: '10:00', purpose: '물리치료·한방', consent: false, text: '내일 열 시 대남병원 물리치료' },
    { id: 'r9', name: '장○호', age: 80, gender: '남', villageId: 'gaknam-hwa', destinationId: 'health-center', date: d, time: '14:00', purpose: '약 처방', consent: true, text: '내일 오후 두 시 보건소' },
    { id: 'r10', name: '임○옥', age: 76, gender: '여', villageId: 'gaknam-chilseong', destinationId: 'health-center', date: d, time: '14:30', purpose: '검사', consent: true, text: '낼 오후에 보건소 검사' },
    { id: 'r11', name: '한○자', age: 85, gender: '여', villageId: 'punggak-songseo', destinationId: 'health-center', date: d, time: '14:00', purpose: '진료', consent: true, text: '내일 두 시 보건소 갈라카는데' },
    { id: 'r12', name: '오○남', age: 83, gender: '남', villageId: 'gakbuk-osan', destinationId: 'health-center', date: d, time: '15:00', purpose: '약 처방', consent: true, text: '내일 세 시 보건소 약' }
  ];
}

// 예약 화면에서 고를 수 있는 어르신 (첫 번째는 아직 예약이 없는 분)
function demoProfiles() {
  const fromRequests = demoRequests().map(({ name, age, gender, villageId }) => ({ name, age, gender, villageId, registered: true }));
  const newcomers = [
    { name: '서○분', age: 82, gender: '여', villageId: 'unmun-daecheon' },
    { name: '문○자', age: 79, gender: '여', villageId: 'unmun-daecheon' },
    { name: '하○수', age: 81, gender: '남', villageId: 'unmun-daecheon' },
    { name: '권○순', age: 85, gender: '여', villageId: 'unmun-sinwon' },
    { name: '배○길', age: 78, gender: '남', villageId: 'unmun-sinwon', registered: false },
    { name: '노○례', age: 83, gender: '여', villageId: 'geumcheon-donggok' },
    { name: '남○희', age: 80, gender: '여', villageId: 'maejeon-onmak' }
  ];
  return [...newcomers.map((p) => ({ registered: true, ...p })), ...fromRequests];
}

module.exports = { demoRequests, demoProfiles };
