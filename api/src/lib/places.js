// 청도군 장소 데이터
// 마을: 공공데이터포털 '경상북도 청도군_버스정류장'(2025-08-28 기준)에서 마을마다 대표 정류장 1곳을 골라
//       그 정류장 좌표·이름을 씀 (타는 곳 = 실제 정류장)
// 병원: '경상북도 청도군_집단급식소운영병원'(2026-06-08 기준) 주소. 좌표는 주소 근처 정류장 기준 근사값
// 보건소·보건지소: 위치 자료가 없어 면 소재지 정류장 좌표로 근사 (군 확인 필요)

const villages = [
  {id: "unmun-sinwon", name: "운문면 신원리", myeon: "운문면", ri: "신원리", stop: "삼계리", stopId: "TSB366000525", lat: 35.667465, lng: 129.00254},
  {id: "unmun-bongha", name: "운문면 봉하리", myeon: "운문면", ri: "봉하리", stop: "봉하", stopId: "TSB366000414", lat: 35.802273, lng: 128.968908},
  {id: "unmun-mail", name: "운문면 마일리", myeon: "운문면", ri: "마일리", stop: "마일", stopId: "TSB366000232", lat: 35.812145, lng: 128.988043},
  {id: "unmun-bangeum", name: "운문면 방음리", myeon: "운문면", ri: "방음리", stop: "방음리", stopId: "TSB366000335", lat: 35.708265, lng: 128.952994},
  {id: "unmun-daecheon", name: "운문면 대천리", myeon: "운문면", ri: "대천리", stop: "대천공용버스터미널", stopId: "TSB366000464", lat: 35.716407, lng: 128.914301},
  {id: "geumcheon-donggok", name: "금천면 동곡리", myeon: "금천면", ri: "동곡리", stop: "동곡버스터미널", stopId: "TSB366000534", lat: 35.693389, lng: 128.890681},
  {id: "geumcheon-bangji", name: "금천면 방지리", myeon: "금천면", ri: "방지리", stop: "방지리", stopId: "TSB366000516", lat: 35.713089, lng: 128.911199},
  {id: "geumcheon-sinji", name: "금천면 신지리", myeon: "금천면", ri: "신지리", stop: "박곡리", stopId: "TSB366000405", lat: 35.668371, lng: 128.915465},
  {id: "maejeon-hapyeong", name: "매전면 하평리", myeon: "매전면", ri: "하평리", stop: "매전", stopId: "TSB366000394", lat: 35.668315, lng: 128.857998},
  {id: "maejeon-onmak", name: "매전면 온막리", myeon: "매전면", ri: "온막리", stop: "온막리", stopId: "TSB366000144", lat: 35.634737, lng: 128.822625},
  {id: "gaknam-hwa", name: "각남면 화리", myeon: "각남면", ri: "화리", stop: "화리", stopId: "TSB366000388", lat: 35.64159, lng: 128.681592},
  {id: "gaknam-chilseong", name: "각남면 칠성리", myeon: "각남면", ri: "칠성리", stop: "칠성1리", stopId: "TSB366000020", lat: 35.650949, lng: 128.683663},
  {id: "gakbuk-osan", name: "각북면 오산리", myeon: "각북면", ri: "오산리", stop: "오산1리", stopId: "TSB366000382", lat: 35.709243, lng: 128.57613},
  {id: "gakbuk-namsan", name: "각북면 남산리", myeon: "각북면", ri: "남산리", stop: "남산3리", stopId: "TSB366000187", lat: 35.695559, lng: 128.569309},
  {id: "punggak-songseo", name: "풍각면 송서리", myeon: "풍각면", ri: "송서리", stop: "풍각버스정류장", stopId: "TSB366000484", lat: 35.640564, lng: 128.615518},
  {id: "iseo-haksan", name: "이서면 학산리", myeon: "이서면", ri: "학산리", stop: "이서", stopId: "TSB366000476", lat: 35.662658, lng: 128.666855},
  {id: "cheongdo-gosu", name: "청도읍 고수리", myeon: "청도읍", ri: "고수리", stop: "청도역(군청방향)", stopId: "TSB366000544", lat: 35.640525, lng: 128.746638},
  {id: "hwayang-seosang", name: "화양읍 서상리", myeon: "화양읍", ri: "서상리", stop: "합천2리", stopId: "TSB366000532", lat: 35.659431, lng: 128.706901}
];

// zone: 'county'(군 안) / 'outside'(군 밖)
const destinations = [
  { id: 'daenam', name: '청도대남병원', kind: '응급의료기관', detail: '응급실 10병상', address: '화양읍 청화로 79-7', lat: 35.6486, lng: 128.7366, keywords: ['대남', '청도병원', '응급실'] },
  { id: 'health-center', name: '청도군보건소', kind: '보건소', detail: '위치 근사', lat: 35.6478, lng: 128.7357, keywords: ['보건소', '소아과'] },
  { id: 'unmun-branch', name: '운문보건지소', kind: '보건지소', myeon: '운문면', detail: '위치 근사(대천)', lat: 35.716407, lng: 128.914301, keywords: [] },
  { id: 'geumcheon-branch', name: '금천보건지소', kind: '보건지소', myeon: '금천면', detail: '위치 근사(동곡)', lat: 35.686899, lng: 128.893332, keywords: [] },
  { id: 'maejeon-branch', name: '매전보건지소', kind: '보건지소', myeon: '매전면', detail: '위치 근사(하평)', lat: 35.668315, lng: 128.857998, keywords: [] },
  { id: 'daegu-hyosung', name: '대구 효성병원', kind: '대구 협약병원', zone: 'outside', lat: 35.857, lng: 128.627, keywords: ['대구', '효성'] },
  { id: 'cheongdo-market', name: '청도시장', kind: '전통시장', detail: '청도시장 정류장', lat: 35.643196, lng: 128.74397, keywords: ['시장', '장보', '장에'] }
];

const zoneOf = (d) => (d && d.zone) || 'county';

function findVillage(id) {
  return villages.find((v) => v.id === id);
}

function findDestination(id) {
  return destinations.find((d) => d.id === id);
}

// 같은 면의 보건지소가 있으면 그곳, 없으면 보건소
function branchFor(myeon) {
  return destinations.find((d) => d.kind === '보건지소' && d.myeon === myeon) || findDestination('health-center');
}

module.exports = { villages, destinations, findVillage, findDestination, branchFor, zoneOf };
