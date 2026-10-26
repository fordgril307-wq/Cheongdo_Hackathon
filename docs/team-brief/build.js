// 청도 해커톤 팀 브리프 — HS 공모전 PPT 양식(리포트형)
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/554c16fb-4080-4068-9fc1-43c46e0b6dee_c628754b-c24e-43b5-8b45-2f50b98300c1/pptx/scripts/apply_theme.js");
const OUT = process.argv[2] || "brief.pptx";
const FONT = process.argv[3] || "Pretendard";
const K = { o: "FF8200", d: "00358F", s: "307FE2", l: "A9C0E4", ink: "222222", sub: "555555", gray: "8C8C8C", bar: "EEF1F6", rule: "C8CDD6", g: "B5BAC2", nv: "0B1F4D", white: "FFFFFF", pale: "F5F7FB", oPale: "FFF3E6" };
const THEME = { name: "Team Brief", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: K.ink, lt1: K.white, dk2: K.d, lt2: K.pale, accent1: K.o, accent2: K.d, accent3: K.s, accent4: K.l, accent5: K.gray, accent6: "595959", hlink: K.s, folHlink: K.d } };
async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}
(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "청도 해커톤 팀 브리프"; pres.author = "김혜수";
  const C = pres.SchemeColor;
  const T = (o) => Object.assign({ isTextBox: true, margin: 0, fontFace: FONT }, o);
  pres.defineSlideMaster({ title: "COVER", background: { color: K.d }, objects: [] });
  pres.defineSlideMaster({
    title: "REPORT", background: { color: K.white },
    objects: [
      { line: { x: 0.5, y: 0.98, w: 12.33, h: 0, line: { color: K.rule, width: 1 } } },
      { line: { x: 0.5, y: 7.02, w: 12.33, h: 0, line: { color: K.rule, width: 0.75 } } },
      { text: { text: "청도 해커톤 팀 브리프", options: { x: 1.0, y: 7.08, w: 6, h: 0.28, fontSize: 8, color: "A0A6AF", fontFace: FONT, margin: 0 } } },
      { text: { text: "2026 청도 바이브코딩 해커톤 with 해커그라운드", options: { x: 7.3, y: 7.08, w: 5.53, h: 0.28, fontSize: 8, color: "A0A6AF", align: "right", fontFace: FONT, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.33, h: 0.62, fontSize: 24, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 0.5, y: 7.08, w: 0.5, h: 0.28, fontSize: 8, color: C.accent5, align: "left" },
  });
  const head = (s, sec, title) => s.addText([{ text: "[" + sec + "] ", options: { color: K.ink, bold: true } }, { text: title, options: { color: K.ink, bold: true } }], { placeholder: "title" });
  const msg = (s, key, bullets, opt = {}) => {
    const y = opt.y || 1.25, h = opt.h || 1.45;
    s.addText(key, T({ x: 0.5, y, w: 2.75, h, fontSize: 16, bold: true, color: K.d, valign: "top", lineSpacingMultiple: 1.15 }));
    s.addText(bullets.map((b, i) => ({ text: b, options: { bullet: true, breakLine: i < bullets.length - 1, bold: i === bullets.length - 1 && !opt.noBoldLast } })),
      T({ x: 3.45, y, w: 9.38, h, fontSize: 12.5, color: K.ink, valign: "top", paraSpaceAfter: 4, lineSpacingMultiple: 1.1 }));
  };
  const panel = (s, x, y, w, h, title, source) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.34, fill: { color: K.bar }, line: { color: K.bar }, objectName: "panel bar " + title });
    s.addText(title, T({ x: x + 0.12, y, w: w - 0.24, h: 0.34, fontSize: 11.5, bold: true, color: K.nv, valign: "middle" }));
    if (source) s.addText("자료: " + source, T({ x, y: y + h + 0.05, w, h: 0.24, fontSize: 7.5, color: "A0A6AF" }));
  };
  const ax = { catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, dataLabelFontFace: FONT, legendFontFace: FONT, titleFontFace: FONT,
    catAxisLabelColor: K.sub, valAxisLabelColor: K.sub, catAxisLabelFontSize: 10, valAxisLabelFontSize: 9.5, dataLabelColor: K.ink, dataLabelFontSize: 10,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLineColor: "7F7F7F", valAxisLineColor: "7F7F7F", showTitle: false };
  const tbl = (s, rows, opt) => s.addTable(rows, Object.assign({ fontFace: FONT, fontSize: 11, color: K.ink, border: { type: "solid", color: "D9D9D9", pt: 0.75 }, valign: "middle", margin: [0.04, 0.08, 0.04, 0.08] }, opt));
  const th = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: K.white, fill: { color: K.d }, align: "center", fontSize: 11 }, o) });
  const td = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 11 }, o) });
  const tdk = (t, o = {}) => td(t, Object.assign({ bold: true, color: K.d, fill: { color: K.pale } }, o));
  const P = { sy: { color: K.d, fill: { color: "E3ECFA" } }, cw: { color: "1F7A6B", fill: { color: "E2F3EF" } }, hs: { color: "C25E00", fill: { color: K.oPale } } };
  const who = (t, k) => td(t, Object.assign({ bold: true }, P[k]));
  const hot = (t, o = {}) => td(t, Object.assign({ bold: true, color: K.o, fill: { color: K.oPale } }, o));
  const I = { users: await icon(fa.FaUsers, K.d), flag: await icon(fa.FaFlag, K.d), clock: await icon(fa.FaClock, K.d), cloud: await icon(fa.FaCloudUploadAlt, K.d), check: await icon(fa.FaCheckCircle, K.d) };

  // ===== 표지
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  s.addText([{ text: "Team Brief", options: { bold: true } }, { text: "  |  2026.10" }], T({ x: 0.8, y: 0.6, w: 6, h: 0.35, fontSize: 13, color: "C9D6EE" }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.95, w: 1.35, h: 0.42, rectRadius: 0.21, fill: { color: K.o }, line: { color: K.o }, objectName: "cover tag" });
  s.addText("해커톤", T({ x: 0.8, y: 1.95, w: 1.35, h: 0.42, fontSize: 14, bold: true, color: K.white, align: "center", valign: "middle" }));
  s.addText("청도 해커톤 팀 브리프", T({ x: 0.8, y: 2.5, w: 9, h: 1.2, fontSize: 54, bold: true, color: K.white }));
  s.addText("팀 운영 · 아이디어 준비 · 부록\n11/5 하루 안에 기획·코드·배포까지", T({ x: 0.8, y: 3.8, w: 7.6, h: 1.0, fontSize: 22, color: "DCE6F7", valign: "top", lineSpacingMultiple: 1.2 }));
  const cv = [await icon(fa.FaLaptopCode, K.d), await icon(fa.FaLongArrowAltRight, K.white), await icon(fa.FaCloudUploadAlt, K.white)];
  s.addShape(pres.shapes.OVAL, { x: 8.75, y: 2.55, w: 1.5, h: 1.5, fill: { color: K.white }, line: { color: K.white }, objectName: "cover from" });
  s.addImage({ data: cv[0], x: 9.13, y: 2.93, w: 0.74, h: 0.74, objectName: "cover from icon" });
  s.addImage({ data: cv[1], x: 10.36, y: 3.05, w: 0.55, h: 0.5, objectName: "cover arrow" });
  s.addShape(pres.shapes.OVAL, { x: 11.03, y: 2.55, w: 1.5, h: 1.5, fill: { color: K.o }, line: { color: K.o }, objectName: "cover to" });
  s.addImage({ data: cv[2], x: 11.41, y: 2.93, w: 0.74, h: 0.74, objectName: "cover to icon" });
  s.addText("만들고", T({ x: 8.5, y: 4.15, w: 2.0, h: 0.35, fontSize: 12, color: "DCE6F7", align: "center" }));
  s.addText("배포까지", T({ x: 10.78, y: 4.15, w: 2.0, h: 0.35, fontSize: 12, color: "DCE6F7", align: "center" }));
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 6.15, w: 11.73, h: 0, line: { color: "5470A8", width: 0.75 }, objectName: "cover foot rule" });
  s.addText("2026 청도 바이브코딩 해커톤 with 해커그라운드", T({ x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 13, color: "C9D6EE" }));
  s.addText([{ text: "팀  ", options: { color: "C9D6EE" } }, { text: "백승유 · 김채원 · 김혜수", options: { bold: true, color: K.white } }], T({ x: 7.3, y: 6.3, w: 5.23, h: 0.4, fontSize: 13, align: "right" }));
  s.addNotes("사전 회의용 팀 브리프입니다. 본문 5장은 역할·일정·준비, 뒤쪽 부록은 데이터와 참고 자료입니다.");

  // ===== 요약
  pres.addSection({ title: "팀 운영" });
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "요약", "로컬에서 다 확인하고, 배포는 정해진 두 번만");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 12.33, h: 1.38, fill: { color: K.pale }, line: { color: K.pale }, objectName: "metrics strip" });
  const km = [["3인", "우리 팀 (최대 4인)"], ["4시간 10분", "실제 개발 가능 시간"], ["2번", "배포 기회 (확인 필요)"], ["14:50", "1차 제출 목표 시각"]];
  km.forEach(([n, l], i) => {
    const x = 0.8 + i * 3.08;
    if (i) s.addShape(pres.shapes.LINE, { x: x - 0.25, y: 1.42, w: 0, h: 0.84, line: { color: "D3DAE6", width: 0.75 }, objectName: "metric sep " + i });
    s.addText(n, T({ x, y: 1.32, w: 2.8, h: 0.62, fontSize: 30, bold: true, color: i === 2 ? K.o : K.d, valign: "middle" }));
    s.addText(l, T({ x, y: 1.97, w: 2.8, h: 0.35, fontSize: 11.5, color: K.gray, valign: "top" }));
  });
  const B = (k, t) => [{ text: k, options: { bold: true, color: K.d } }, { text: "  " + t }];
  const rows = [
    ["대회", I.flag, [B("11/5(목) 하루", "10:00 주제 공개 → 16:30 기획서·코드·배포된 서비스 제출, 심사 기준은 기획·코드 구현·사용자 경험"), B("발표 없음", "기획서·README가 발표 역할, 15:00 1차 자동 검수 · 16:30 2차 AI 사전 검수")]],
    ["원칙", I.cloud, [B("배포는 두 번만", "14:30(1차 제출용) · 16:00(2차 제출용), 그 전에는 로컬에서 확인"), B("14:50 1차 제출, 16:10 코드 동결", "마지막 20분은 기획서·README만 수정")]],
    ["역할", I.users, [B("승유 백엔드·배포 / 채원 화면 / 혜수 기획·QA·제출", "폴더로 담당을 나눠 같은 파일 동시 수정 금지")]],
    ["준비", I.check, [B("11/3까지", "팀원 각자 참가 신청, GitHub·Copilot·Azure 계정, 개인 계정으로 연습 배포 1회")]],
  ];
  const hts = [0.98, 0.98, 0.98, 0.98]; // 요약 행 높이·아이콘 크기 통일 (HS 양식 규칙)
  let yy = 2.68;
  rows.forEach(([lab, ic, bl], i) => {
    const h = hts[i];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yy, w: 12.33, h, rectRadius: 0.06, fill: { color: K.white }, line: { color: "D3D8E0", width: 0.75 }, objectName: "sum row " + lab });
    const bs = 0.58;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.68, y: yy + (h - bs) / 2, w: bs, h: bs, rectRadius: 0.08, fill: { color: "E3ECFA" }, line: { color: "E3ECFA" }, objectName: "sum icon box " + lab });
    s.addImage({ data: ic, x: 0.68 + bs * 0.2, y: yy + (h - bs) / 2 + bs * 0.2, w: bs * 0.6, h: bs * 0.6, objectName: "sum icon " + lab });
    s.addText(lab, T({ x: 1.42, y: yy, w: 0.9, h, fontSize: 16, bold: true, color: K.d, valign: "middle" }));
    s.addShape(pres.shapes.LINE, { x: 2.3, y: yy + 0.12, w: 0, h: h - 0.24, line: { color: "DDE1E7", width: 0.75 }, objectName: "sum sep " + lab });
    const runs = [];
    bl.forEach((parts, k) => parts.forEach((r, j) => runs.push({ text: r.text, options: Object.assign({}, r.options, j === 0 ? { bullet: true } : {}, (j === parts.length - 1 && k < bl.length - 1) ? { breakLine: true } : {}) })));
    s.addText(runs, T({ x: 2.5, y: yy, w: 10.2, h, fontSize: 12, color: K.ink, valign: "middle", paraSpaceAfter: 3 }));
    yy += h + 0.1;
  });
  s.addNotes("한 장 요약입니다. 우리 팀은 3명, 실제 개발 시간은 4시간 남짓이고, 배포 기회가 2번뿐이라는 정보가 있어 로컬에서 다 확인한 뒤 14:30과 16:00 두 번만 배포합니다. 1차 제출은 14:50, 코드 동결은 16:10입니다.");

  // ===== 역할
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "역할", "개발 2명은 구현, 기획 1명은 기획서·검수·제출 전담");
  msg(s, "기획서는 필수 제출물이자\n심사 기준", [
    "개발자가 기획서까지 쓰면 구현 시간이 줄어듦 → 기획서·검수·제출은 분리해 한 명이 전담",
    "폴더로 담당 분리: 화면 app/ = 채원, API api/ = 승유, README·기획서 = 혜수 (같은 파일 동시 수정 금지)",
    "배포는 팀장만 실행, 30분마다 한 줄 공유(끝난 것 / 막힌 것), 15분 넘게 막히면 현장 멘토에게 바로 질문",
    "제출 양식의 팀 이름·저장소 주소·GitHub ID는 팀 등록 때와 글자 하나까지 동일하게",
  ]);
  panel(s, 0.5, 2.95, 12.33, 3.6, "시간대별 역할 분담 (안, 최종 판단은 팀장)", null);
  tbl(s, [
    [th("담당"), th("10:00~11:00"), th("11:00~13:00"), th("14:00~16:10"), th("마감")],
    [who("백승유\n팀장·백엔드", "sy"), td("아이디어 최종 결정, 기술 구조"), td("저장소·배포 파이프라인, API·AI 연결"), td("배포 ①(14:30)·②(16:00) 실행\n검수 지적 대응"), td("코드 동결 판단")],
    [who("김채원\n프론트엔드", "cw"), td("화면 흐름 함께 설계"), td("화면 뼈대·핵심 화면 구현"), td("API 연결, 디자인 다듬기"), td("화면 최종 확인")],
    [who("김혜수\n기획·PM", "hs"), td("문제 정의·근거 숫자\n화면 흐름·문구"), td("기획서 작성, 데이터 준비"), td("배포 전 체크리스트, 휴대폰 QA, README"), td("1·2차 제출 체크리스트")],
  ], { x: 0.5, y: 3.42, w: 12.33, colW: [1.75, 2.55, 2.75, 3.13, 2.15], rowH: [0.42, 0.85, 0.85, 0.85], fontSize: 10.5 });
  s.addNotes("역할입니다. 기획서가 필수 제출물이자 심사 기준이라, 개발 두 명은 구현에 집중하고 기획서·검수·제출은 제가 전담하겠습니다. 폴더로 담당을 나눠 충돌을 막고, 배포는 팀장님만 실행합니다.");

  // ===== 시간표
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "당일 시간표", "점심 전 핵심 기능, 점심 뒤 화면 — 배포는 14:30·16:00");
  msg(s, "개발 가능 시간\n약 4시간 10분", [
    "10:30~16:10 중 점심·제출 시간을 빼면 약 4시간 10분",
    "오전은 핵심 기능(API·AI)과 화면 뼈대를 로컬에서 완성, 오후는 화면 붙이기와 README·기획서",
    "시간 부족 시 버리는 순서: 부가 기능 → 지도·애니메이션 → AI 연결(규칙 기반으로 대체)",
    "배포 2번이 아니라고 확인되면 11:30에 빈 화면부터 배포하는 방식으로 변경",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "오전 (08:00~13:00)", "해커그라운드 공식 일정 + 팀 계획");
  tbl(s, [
    [th("시각"), th("공식"), th("우리 팀")],
    [td("08:00", { align: "center" }), td("동대구역 셔틀"), td("07:50까지 탑승 장소 도착")],
    [td("09:00", { align: "center" }), td("등록·체크인"), td("와이파이·계정 로그인 확인")],
    [td("10:00", { align: "center" }), td("주제 공개"), td("주제 기록, 아이디어 결정(15분)")],
    [td("10:15", { align: "center" }), td(""), td("기획서 뼈대·화면 흐름 확정")],
    [td("10:30", { align: "center" }), td("팀빌딩"), td("팀 등록, 저장소·배포 설정")],
    [td("11:00", { align: "center" }), td("개발 Ⅰ"), td("핵심 기능·화면 뼈대(로컬)")],
  ], { x: 0.5, y: 3.38, w: 6.0, colW: [0.9, 1.6, 3.5], rowH: 0.44, fontSize: 10.5 });
  panel(s, 6.83, 2.95, 6.0, 3.6, "오후 (13:00~17:30)", "해커그라운드 공식 일정 + 팀 계획");
  tbl(s, [
    [th("시각"), th("공식"), th("우리 팀")],
    [td("13:00", { align: "center" }), td("중식"), td("로컬 화면 점검, 오후 할 일 정리")],
    [td("14:00", { align: "center" }), td("개발 Ⅱ"), td("화면 붙이기, README·기획서")],
    [hot("14:30", { align: "center" }), hot(""), hot("배포 ① → 배포 주소 확인")],
    [td("14:50", { align: "center", bold: true }), td("15:00 1차 제출"), td("1차 제출(10분 여유)", { bold: true })],
    [hot("16:00", { align: "center" }), hot(""), hot("배포 ②(마지막) → 전 기능 확인")],
    [td("16:20", { align: "center", bold: true }), td("16:30 2차 제출"), td("16:10 동결 → 2차 제출", { bold: true })],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [0.9, 1.6, 3.5], rowH: 0.44, fontSize: 10.5 });
  s.addNotes("당일 시간표입니다. 실제 개발 시간은 4시간 10분 정도라, 오전에 핵심 기능을 로컬에서 끝내고 오후에 화면을 붙입니다. 주황색 두 줄이 배포 시각입니다.");

  // ===== 사전 회의
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "사전 회의", "역할·도구·운영 방식을 미리 맞춰, 당일엔 개발에만 집중");
  msg(s, "온라인 1시간,\n안건 6개", [
    "가장 먼저 '배포 2번'의 정확한 의미 확인 — 배포가 2번인지, 제출(1차·2차)이 2번인지에 따라 당일 방식이 달라짐",
    "역할·기술 스택을 정하고, 예상 주제는 방향만 이야기(코드는 당일 작성)",
    "셔틀·귀가 방법과 연락 채널까지 정리",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "회의 안건 (안)", null);
  tbl(s, [
    [th("안건"), th("시간"), th("정할 것")],
    [tdk("1. 배포 횟수"), td("10분", { align: "center" }), td("'2번'의 출처와 의미")],
    [tdk("2. 역할"), td("10분", { align: "center" }), td("분담 확정, 디자이너 추가 여부")],
    [tdk("3. 기술 스택"), td("15분", { align: "center" }), td("화면·API·AI·배포 도구")],
    [tdk("4. 예상 주제"), td("15분", { align: "center" }), td("분야별 방향 1개씩")],
    [tdk("5. 운영 규칙"), td("5분", { align: "center" }), td("저장소·폴더, 30분 공유")],
    [tdk("6. 이동·연락"), td("5분", { align: "center" }), td("셔틀, 귀가 합승, 채널")],
  ], { x: 0.5, y: 3.38, w: 6.0, colW: [1.8, 0.8, 3.4], rowH: 0.44, fontSize: 10.5 });
  panel(s, 6.83, 2.95, 6.0, 3.6, "배포 방식 — '2번 제한' 여부에 따라", "배포 2번: 팀 내 공유 정보(공식 출처 확인 필요)");
  tbl(s, [
    [th("구분"), th("배포가 2번이면"), th("제한이 없으면")],
    [tdk("배포 시점"), hot("14:30 · 16:00 두 번만"), td("11:30 빈 화면부터, 수시로")],
    [tdk("확인 방법"), td("로컬(SWA CLI)에서 끝내고 배포"), td("배포 주소에서 바로 확인")],
    [tdk("자동 배포"), td("끔(수동 실행만)"), td("push마다 자동")],
    [tdk("배포 담당"), td("팀장만"), td("팀장 중심")],
    [tdk("대회 전"), td("각자 개인 Azure로 연습 배포 1회", { colspan: 2 })],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [1.3, 2.35, 2.35], rowH: 0.5, fontSize: 10.5 });
  s.addNotes("사전 회의 안건입니다. 가장 먼저 배포가 2번뿐이라는 정보가 정확히 무슨 뜻인지 확인해야 합니다. 그에 따라 오른쪽 표처럼 당일 배포 방식이 달라집니다.");

  // ===== 준비
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "대회 전 준비", "계정·설치는 11/3까지 — 당일 아침엔 늦음");
  msg(s, "11/3(화)까지\n각자 완료", [
    "참가 신청은 팀원 각자 hackersground.kr에서(마감 11/3 23:59), 팀 등록은 현장",
    "행사장 와이파이가 느릴 수 있어 계정·설치는 미리 끝내기",
    "귀가 셔틀은 없음 — 행사장(운문면)에서 동대구역까지 택시 약 1시간, 팀 합승 가능",
  ]);
  panel(s, 0.5, 2.95, 12.33, 3.6, "개인 준비 체크리스트", null);
  tbl(s, [
    [th("구분"), th("항목"), th("메모"), th("확인")],
    [tdk("신청"), td("해커톤 참가 신청"), td("팀원 각자, 11/3 23:59 마감"), td("☐", { align: "center" })],
    [tdk("계정"), td("GitHub + GitHub Copilot"), td("무료 플랜 가능, 학생은 GitHub Education 확인"), td("☐", { align: "center" })],
    [tdk("계정"), td("Azure"), td("무료 계정(카드 인증), 학생은 Azure for Students 확인"), td("☐", { align: "center" })],
    [tdk("설치"), td("VS Code · Git · Node.js 20 · Azure CLI"), td("Copilot 확장 로그인까지"), td("☐", { align: "center" })],
    [tdk("연습"), td("개인 Azure로 연습 배포 1회"), td("연습 저장소는 당일 제출에 쓰지 않음"), td("☐", { align: "center" })],
    [tdk("당일"), td("노트북·충전기·보조배터리·이어폰·신분증"), td("핫스팟 백업, 음성 기능 테스트용 이어폰"), td("☐", { align: "center" })],
  ], { x: 0.5, y: 3.38, w: 12.33, colW: [1.3, 4.2, 5.83, 1.0], rowH: 0.44, fontSize: 10.5 });
  s.addNotes("대회 전 준비입니다. 참가 신청은 각자 해야 하고, 계정과 설치는 11/3까지 끝내 주세요. 개인 Azure로 연습 배포를 한 번 해 보면 당일 막히는 지점을 미리 알 수 있습니다.");

  // ===== 부록1-3 아이디어 정하는 법
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "팀 운영" });
  head(s, "아이디어 정하는 법", "주제 공개 후 15분 안에 아이디어 정하는 법");
  msg(s, "PSST 순서로\n한 줄씩 먼저", [
    "주제 문장에서 '누가, 무엇이 불편한가'를 뽑고, 청도에 이미 있는데 덜 쓰이는 자원과 연결",
    "기획서는 문제 → 해결 → 강점 → 기대효과(PSST) 순서로 한 줄씩 먼저 작성",
    "핵심 기능은 3개로 제한, 나머지는 '확장 계획'에 한 줄",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "15분 진행 순서", null);
  tbl(s, [
    [th("단계"), th("할 일"), th("시간")],
    [tdk("1"), td("주제 문장 그대로 기록"), td("1분", { align: "center" })],
    [tdk("2"), td("누가 불편한지 + 근거 숫자 1개"), td("4분", { align: "center" })],
    [tdk("3"), td("덜 쓰이는 자원 찾기"), td("5분", { align: "center" })],
    [tdk("4"), td("핵심 기능 3개로 제한"), td("3분", { align: "center" })],
    [tdk("5"), td("사용자 정의(어르신이면 큰 글씨·말로 입력)"), td("2분", { align: "center" })],
  ], { x: 0.5, y: 3.38, w: 6.0, colW: [0.8, 4.2, 1.0], rowH: 0.5, fontSize: 10.5 });
  panel(s, 6.83, 2.95, 6.0, 3.6, "기획서 첫 줄 틀 (PSST)", "2023 대상 DEPS 팀 발표 구성(PSST) 참고");
  tbl(s, [
    [th("순서"), th("쓸 것"), th("예시 질문")],
    [tdk("P 문제"), td("누가·무엇이 불편한가 + 숫자"), td("몇 명이, 얼마나 자주?")],
    [tdk("S 해결"), td("무엇으로 어떻게 푸나"), td("이미 있는 자원은?")],
    [tdk("S 강점"), td("기존 방식과 다른 점"), td("왜 지금 안 되고 있나?")],
    [tdk("T 효과"), td("숫자로 본 변화"), td("전후 비교가 되나?")],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [1.2, 2.5, 2.3], rowH: 0.55, fontSize: 10.5 });


  // ===== 아이디어 준비 표지
  pres.addSection({ title: "아이디어 준비" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "아이디어 준비" });
  s.addText("아이디어 준비", T({ x: 0.8, y: 2.3, w: 11, h: 1.2, fontSize: 60, bold: true, color: K.white }));
  s.addText("1. 지난 대상팀은 어떻게 했나   2. 청도 데이터: 있는데 덜 쓰이는 자원   3. 가안: 청도효도택시 — 기능·화면", T({ x: 0.8, y: 3.6, w: 11.7, h: 0.6, fontSize: 18, color: "DCE6F7" }));
  s.addText("가안은 대회 전 참고용 — 제출 코드는 당일 주제 공개 뒤 새 저장소에서 작성", T({ x: 0.8, y: 4.4, w: 11.7, h: 0.4, fontSize: 13, color: "C9D6EE" }));

  // ===== 부록3 지난 대상
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "지난 대상팀", "지난 대상팀은 '놀고 있는 자원 × 부족한 것'을 잇고 배포까지 확실히 했음");
  msg(s, "아이디어와\n배포 신뢰성이 우위", [
    "2023·2024 대상 모두 지역의 놀고 있는 자원을 부족한 것과 연결, 기능이 화려한 팀보다 우위",
    "2024년엔 검증단이 README를 그대로 따라 해 배포되는지 확인, 대부분 팀이 Azure 배포에서 막힘",
    "2026년 15:00 자동 검수·16:30 AI 사전 검수는 이 검증의 하루짜리 자동화일 가능성(추정)",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "역대 대상작", "2023 DEPS 팀 회고(zedy_dev.log), 2024 나누리 팀 회고(김건우 velog)");
  tbl(s, [
    [th("연도"), th("대상작"), th("놀고 있던 자원"), th("부족한 것")],
    [tdk("2024 의성"), td("나누리", { bold: true }), td("지역 잉여 농산물"), td("나눔·유통")],
    [tdk("2023 대구"), td("DEPS", { bold: true }), td("카센터의 고압 전력"), td("지방 전기차 충전·정비")],
  ], { x: 0.5, y: 3.38, w: 6.0, colW: [1.2, 1.1, 1.9, 1.8], rowH: 0.62, fontSize: 10.5 });
  s.addText([{ text: "공통점  ", options: { bold: true, color: K.o } }, { text: "새 제도보다 이미 있는 자원을 잇는 아이디어", options: { breakLine: true } }, { text: "숫자로 필요성 증명(예: 정비 가능 카센터 2% 미만)" }], T({ x: 0.5, y: 5.45, w: 6.0, h: 0.8, fontSize: 11, color: K.ink, valign: "middle" }));
  panel(s, 6.83, 2.95, 6.0, 3.6, "2024년 심사 과정", "2024 나누리 팀 회고(김건우 velog)");
  tbl(s, [
    [th("단계"), th("확인 내용"), th("탈락 사례")],
    [tdk("배포 검증"), td("README대로 따라 해서 배포되는지"), td("대부분 Azure 배포에서 막힘")],
    [tdk("제출 대조"), td("팀명·저장소·GitHub ID 봇 대조"), td("표기 불일치")],
    [tdk("기술 심사"), td("실제 배포·CI/CD 동작"), td("로컬로만 띄운 팀")],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [1.2, 2.6, 2.2], rowH: 0.62, fontSize: 10.5 });

  // ===== 부록1-1 인구 차트
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "청도 데이터 1", "청도는 주민 10명 중 6명이 60세 이상, 행사장 운문면이 가장 나이 든 면");
  msg(s, "고령·의료·이동이\n가장 유력한 문제 영역", [
    "총인구 39,954명 중 60세 이상 23,736명(59.4%), 가장 큰 연령대는 60대 → 10년 안에 돌봄 수요 한 번 더 증가",
    "80대 이상은 여성이 남성의 약 2배(3,464명 대 1,691명)",
    "읍 두 곳(청도읍·화양읍)만 군 평균(59.1세)보다 젊고, 면 지역은 모두 59세 이상",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "연령대별 인구 (명)", "행정안전부 주민등록 인구통계(2026.9)");
  const ag = ["0~9세", "10대", "20대", "30대", "40대", "50대", "60대", "70대", "80대", "90세+"];
  const pop = [990, 1939, 1624, 2078, 3440, 6147, 10343, 8238, 4156, 999];
  s.addChart(pres.charts.BAR, [
    { name: "60세 미만", labels: ag, values: pop.map((v, i) => i < 6 ? v : 0) },
    { name: "60세 이상", labels: ag, values: pop.map((v, i) => i >= 6 ? v : 0) },
  ], Object.assign({}, ax, { x: 0.55, y: 3.35, w: 5.9, h: 3.15, barDir: "col", barGrouping: "stacked", chartColors: [K.g, K.o], showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "#,##0;;", dataLabelColor: K.white, dataLabelFontSize: 9,
    showLegend: false, valAxisHidden: true, barGapWidthPct: 40, catAxisLabelFontSize: 9.5 }));
  panel(s, 6.83, 2.95, 6.0, 3.6, "읍면별 평균 연령 (세, 군 평균 59.1세)", "행정안전부 주민등록 인구통계(2026.9)");
  const my = ["운문면", "금천면", "매전면", "각남면", "각북면", "풍각면", "이서면", "청도읍", "화양읍"];
  const av = [64.1, 64.0, 63.9, 63.0, 62.6, 61.0, 59.7, 56.6, 53.9];
  s.addChart(pres.charts.BAR, [
    { name: "행사장", labels: my, values: av.map((v, i) => i === 0 ? v : 0) },
    { name: "기타", labels: my, values: av.map((v, i) => i === 0 ? 0 : v) },
  ], Object.assign({}, ax, { x: 6.88, y: 3.35, w: 5.9, h: 3.15, barDir: "bar", barGrouping: "stacked", chartColors: [K.o, K.g], catAxisOrientation: "maxMin", showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0.0;;", dataLabelColor: K.white,
    showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 70, barGapWidthPct: 35, catAxisLabelFontSize: 10 }));
  s.addNotes("부록입니다. 청도는 60세 이상이 59%이고, 행사장이 있는 운문면이 평균 64.1세로 가장 나이 든 면입니다.");


  // ===== 청도 데이터 2 의료
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "청도 데이터 2", "응급실은 군 전체에 1곳, 운문면 마을에선 차로 최대 약 60분");
  msg(s, "병원은 읍에 몰려 있고\n어르신은 면에 삼", [
    "청도군은 국립중앙의료원 지정 의료취약지(경북 13곳·전국 76곳 중 하나)",
    "응급실은 청도대남병원 1곳(10병상), 이용 연 8,000명 이상 — 각남·각북·운문·매전엔 동네 의원 없음",
    "운문면 마을에서 대남병원까지 차로 35~60분(추정), 가장 나이 든 면일수록 가장 멂",
  ]);
  panel(s, 0.5, 2.95, 7.3, 3.6, "마을별 청도대남병원까지 차량 예상 시간 (분, 추정)", "공공데이터포털 「청도군_버스정류장 현황」(2025.8.28)·「집단급식소 운영 병원」(2026.6.8) 좌표, 직선거리×1.35·시속 40km 가정");
  const vl = ["운문 마일리", "운문 봉하리", "운문 신원리", "운문 방음리", "운문 대천리", "금천 방지리", "금천 신지리", "금천 동곡리", "각북 오산리", "각북 남산리", "매전 하평리", "풍각 송서리"];
  const vm = [60, 55, 50, 40, 35, 35, 35, 30, 30, 30, 25, 20];
  s.addChart(pres.charts.BAR, [
    { name: "운문면", labels: vl, values: vm.map((v, i) => i < 5 ? v : 0) },
    { name: "기타", labels: vl, values: vm.map((v, i) => i < 5 ? 0 : v) },
  ], Object.assign({}, ax, { x: 0.55, y: 3.35, w: 7.2, h: 3.15, barDir: "bar", barGrouping: "stacked", chartColors: [K.o, K.g], catAxisOrientation: "maxMin", showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0;;", dataLabelColor: K.white,
    showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 65, barGapWidthPct: 30, catAxisLabelFontSize: 9.5 }));
  panel(s, 8.13, 2.95, 4.7, 3.6, "청도군 병원 현황", "공공데이터포털 「청도군_집단급식소 운영 병원」(2026.6.8)");
  tbl(s, [
    [th("병원"), th("위치"), th("응급실")],
    [td("청도대남병원", { bold: true }), td("화양읍"), hot("10병상", { align: "center" })],
    [td("하나병원"), td("청도읍"), td("없음", { align: "center" })],
    [td("청도메타병원"), td("청도읍"), td("없음", { align: "center" })],
    [td("윤성병원"), td("각북면"), td("없음", { align: "center" })],
    [td("청도자연요양병원"), td("각북면"), td("없음", { align: "center" })],
    [td("감꽃요양병원"), td("각북면"), td("없음", { align: "center" })],
  ], { x: 8.13, y: 3.38, w: 4.7, colW: [2.1, 1.3, 1.3], rowH: 0.44, fontSize: 10.5 });
  s.addNotes("의료 접근입니다. 응급실은 군 전체에 대남병원 한 곳이고, 운문면 마을에서는 차로 35분에서 60분이 걸립니다. 가장 나이 든 면이 병원에서 가장 멉니다. 시간은 정류장 좌표로 계산한 추정치입니다.");

  // ===== 청도 데이터 3 이동
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "청도 데이터 3", "정류장은 508곳, 그래도 어르신에게 이동은 '부를 수 있느냐'의 문제");
  msg(s, "수단은 있는데\n부르기가 어려움", [
    "버스정류장 508곳 중 운문면은 34곳으로 9개 읍면 중 8위 — 면적 대비 띄엄띄엄",
    "청도군은 대중교통 취약지역에 행복택시·행복버스를 운영(세부 기준은 군 확인 중)",
    "70세 이상은 택시 호출 앱을 거의 못 씀(조사 42명 중 8명만 가능) → 전화·말로 부르는 방식이 필요",
  ]);
  panel(s, 0.5, 2.95, 6.6, 3.6, "읍면별 버스정류장 수 (곳)", "공공데이터포털 「경상북도 청도군_버스정류장 현황」(2025.8.28), 주소 불명확 11곳 제외");
  const sl = ["매전면", "청도읍", "이서면", "화양읍", "풍각면", "금천면", "각남면", "운문면", "각북면"];
  const sv = [85, 77, 65, 56, 55, 53, 41, 34, 31];
  s.addChart(pres.charts.BAR, [
    { name: "기타", labels: sl, values: sv.map((v, i) => i === 7 ? 0 : v) },
    { name: "운문면", labels: sl, values: sv.map((v, i) => i === 7 ? v : 0) },
  ], Object.assign({}, ax, { x: 0.55, y: 3.35, w: 6.5, h: 3.15, barDir: "col", barGrouping: "stacked", chartColors: [K.g, K.o], showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0;;", dataLabelColor: K.white,
    showLegend: false, valAxisHidden: true, barGapWidthPct: 40, catAxisLabelFontSize: 10 }));
  panel(s, 7.43, 2.95, 5.4, 3.6, "이동 관련 현황", "헤럴드경제 2026.8.2, 경북매일 2021, 공공데이터 좌표 기반 추정");
  tbl(s, [
    [th("항목"), th("현황")],
    [tdk("버스정류장"), td("군 전체 508곳 (2025.8.28 기준)")],
    [tdk("행복택시·버스"), td("대중교통 취약지역 운영, 세부 기준 확인 중")],
    [tdk("호출 앱"), hot("70세 이상 42명 중 8명만 사용 가능")],
    [tdk("행사장→응급실"), td("운문면 대천리 기준 차로 약 35분(추정)")],
  ], { x: 7.43, y: 3.38, w: 5.4, colW: [1.5, 3.9], rowH: 0.62, fontSize: 10.5 });

  // ===== 청도 데이터 4 덜 쓰이는 자원
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "청도 데이터 4", "있는데 덜 쓰이는 자원: 관광주민증 1.2%, 관광택시 3.5%만 이용");
  msg(s, "놀고 있는 자원 ×\n부족한 것", [
    "관광주민증은 약 9.9만 건 발급, 혜택 이용은 1,223건 / 군 관광택시는 정원 999명에 신청 35명",
    "경로당(전 경로당 비상벨)·노인일자리 1,830명·보건지소 8곳처럼 사람과 거점도 이미 있음",
    "지난 대상팀처럼 '있는데 덜 쓰이는 자원'을 '부족한 것'과 이으면 새 제도 없이 해결 가능",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "준비된 만큼 쓰이고 있나 (이용률, %)", "경북매일 2026.5.26(관광주민증), 청도군 통합예약시스템 관광택시 화면(2026.10.24)");
  s.addChart(pres.charts.BAR, [{ name: "이용률", labels: ["관광주민증 혜택 이용 (1,223 / 9.9만)", "관광택시 신청 (35 / 999명)"], values: [1.2, 3.5] }],
    Object.assign({}, ax, { x: 0.55, y: 3.35, w: 5.9, h: 3.15, barDir: "bar", chartColors: [K.o], catAxisOrientation: "maxMin", showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0\"%\"", dataLabelColor: K.o, dataLabelFontSize: 14,
      showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 25, valAxisLabelFormatCode: "0\"%\"", barGapWidthPct: 60, catAxisLabelFontSize: 10 }));
  panel(s, 6.83, 2.95, 6.0, 3.6, "놀고 있는 자원 × 부족한 것 (브레인스토밍 출발점)", "부록 3 출처 참고");
  tbl(s, [
    [th("놀고 있는 자원"), th("이용 현황"), th("부족한 것")],
    [td("관광주민증 보유자", { bold: true }), td("발급 대비 이용 1.2%"), td("농번기 일손, 체류 시간")],
    [td("군 관광택시", { bold: true }), hot("정원 대비 3.5%"), td("외곽 면 어르신 이동")],
    [td("경로당·노인일자리", { bold: true }), td("전 경로당 비상벨, 1,830명"), td("돌봄 인력·기록")],
    [td("보건지소 8곳", { bold: true }), td("전문의 주 1회 이상"), td("동네 의원 없는 면 진료")],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [1.9, 2.0, 2.1], rowH: 0.55, fontSize: 10.5 });

  // ===== 준비한 아이디어 1
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "가안", "청도효도택시 — 어르신 병원길, 같은 방향 이웃과 한 차로");
  msg(s, "그래서: 지금 행복택시를\nAI 합승으로 개선", [
    "어르신이 사투리로 말하면 AI가 예약을 정리하고, 같은 날·같은 방향 이웃을 한 차로 묶어 배차",
    "새 제도가 아니라 청도군이 운영 중인 행복택시의 개선안 — 전화 예약·한 분씩 운행 → AI 접수·합승",
    "합승 규칙은 택시발전법 시행규칙(2022.6.15)의 플랫폼 중개 합승 기준 반영: 사전 동의, 모르는 분은 같은 성별끼리",
    "주제가 이동·고령·의료면 당일 이 방향으로 새 저장소에서 구현, 다른 주제면 '접수 + 묶기' 구조만 각색",
  ]);
  panel(s, 0.5, 2.95, 7.3, 3.6, "지금 행복택시 → 청도효도택시", "청도군 행복택시 운영 보도(경북매일 2021), 세부 기준은 군 확인 중 · 관광택시: 청도군 통합예약시스템(2026.10.24)");
  tbl(s, [
    [th("지금"), th("청도효도택시")],
    [td("읍·면사무소나 이장님을 거쳐 이용자 등록"), td("이장님·경로당이 앱에서 대리 신청, 어르신 직접도 가능")],
    [td("전담 기사님께 개인 전화로 예약, 기록 없음"), td("사투리 그대로 AI가 접수, 모든 예약 기록")],
    [td("같은 곳에 가도 한 분씩 따로 운행"), hot("같은 방향끼리 합승 → 운행·정산 건수 절반")],
    [td("정해진 구간까지만 운행(확인 중)"), td("(제안) 모이면 병원·대구 협약병원까지 구간 확대")],
    [td("관광택시는 정원 999명에 신청 35명"), td("(제안) 평일 관광택시를 어르신 이동 보조로")],
  ], { x: 0.5, y: 3.38, w: 7.3, colW: [3.4, 3.9], rowH: 0.5, fontSize: 10.5 });
  panel(s, 8.13, 2.95, 4.7, 3.6, "샘플 예약 12건 시뮬레이션", "공공데이터포털 「청도군_버스정류장 현황」(2025.8.28) 좌표 기반 자체 계산, 샘플 기준");
  s.addChart(pres.charts.BAR, [
    { name: "지금", labels: ["지금 방식", "청도효도택시"], values: [12, 0] },
    { name: "효도", labels: ["지금 방식", "청도효도택시"], values: [0, 6] },
  ], Object.assign({}, ax, { x: 8.2, y: 3.45, w: 2.7, h: 3.0, barDir: "col", barGrouping: "stacked", chartColors: [K.g, K.o], showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0\"번\";;", dataLabelColor: K.white, dataLabelFontSize: 12,
    showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 14, barGapWidthPct: 45, catAxisLabelFontSize: 10 }));
  s.addText([{ text: "운행 횟수", options: { fontSize: 11.5, color: K.gray, breakLine: true } }, { text: "12 → 6번", options: { fontSize: 24, bold: true, color: K.d, breakLine: true } }, { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "주행거리", options: { fontSize: 11.5, color: K.gray, breakLine: true } }, { text: "45% 절감", options: { fontSize: 24, bold: true, color: K.o } }], T({ x: 11.05, y: 3.6, w: 1.75, h: 2.6, valign: "middle" }));
  s.addNotes("제가 준비한 아이디어는 청도효도택시입니다. 새 제도가 아니라 청도군이 이미 운영하는 행복택시를 고치는 안입니다. 지금은 기사님께 전화로 한 분씩 예약하고 따로 운행하는데, 어르신이 말로 예약하면 AI가 정리해서 같은 방향 이웃을 한 차로 묶습니다. 실제 정류장 좌표로 샘플 12건을 돌려 보니 운행은 6번, 주행거리는 45% 줄었습니다. 주제가 맞으면 당일 이 방향으로 새로 구현하고, 다른 주제면 구조만 가져다 쓰면 됩니다.");


  // ===== 가안 기능
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "가안", "필요 기능 — 필수 4개 먼저, 나머지는 시간이 남으면");
  msg(s, "필수 4개로\n14:30 배포 ①", [
    "필수 4개(말로 예약·합승 배차·어르신 화면·배포)는 14:30 배포 ① 전에 끝내기",
    "있으면 좋은 기능은 15:20 이후, 시간이 없으면 기획서 '확장 계획'에 한 줄로",
    "AI 연결이 막히면 규칙 기반 해석으로 대체해 데모는 멈추지 않게",
  ]);
  panel(s, 0.5, 2.95, 7.6, 3.6, "기능 목록과 담당 (안)", null);
  tbl(s, [
    [th("구분"), th("기능"), th("내용"), th("담당")],
    [hot("필수", { align: "center" }), td("말로 예약", { bold: true }), td("사투리 → AI가 날짜·시간·목적지 정리, 실패 시 규칙 기반"), who("승유", "sy")],
    [hot("필수", { align: "center" }), td("합승 배차", { bold: true }), td("같은 날·같은 방향·60분 이내 묶기, 모르는 분은 동성"), who("승유", "sy")],
    [hot("필수", { align: "center" }), td("어르신 화면", { bold: true }), td("어디로 → 언제 → 확인 → 완료, 시니어 컬러"), who("채원", "cw")],
    [hot("필수", { align: "center" }), td("배포·README", { bold: true }), td("README대로 따라 하면 배포되게"), who("승유·혜수", "hs")],
    [td("선택", { align: "center" }), td("탑승권·좌석 지도"), td("타는 시각·정류장·내 자리 안내"), who("채원", "cw")],
    [td("선택", { align: "center" }), td("경로당 대리 신청"), td("이장님·경로당이 여러 분을 한 번에"), who("채원", "cw")],
    [td("선택", { align: "center" }), td("배차 담당 화면"), td("지도, '지금 방식 vs 효도택시' 비교"), who("승유", "sy")],
  ], { x: 0.5, y: 3.38, w: 7.6, colW: [0.8, 1.75, 4.05, 1.0], rowH: 0.41, fontSize: 10 });
  panel(s, 8.43, 2.95, 4.4, 3.6, "구조 (Azure + GitHub)", null);
  const box = (y, t1, t2, col) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.6, y, w: 4.06, h: 0.52, rectRadius: 0.06, fill: { color: col === K.o ? K.oPale : "E3ECFA" }, line: { color: col, width: 1 }, objectName: "arch " + t1 });
    s.addText([{ text: t1, options: { bold: true, color: col, breakLine: true } }, { text: t2, options: { color: K.sub } }], T({ x: 8.72, y, w: 3.82, h: 0.52, fontSize: 10, valign: "middle", lineSpacingMultiple: 1.0 }));
  };
  const down = (y) => s.addShape(pres.shapes.DOWN_ARROW, { x: 10.5, y, w: 0.26, h: 0.18, fill: { color: K.g }, line: { color: K.g }, objectName: "arch arrow " + y });
  box(3.42, "어르신 휴대폰 (웹)", "말하기·버튼 — 앱 설치 없이 주소로", K.d); down(3.96);
  box(4.16, "Azure Static Web Apps", "화면(HTML·CSS·JS) 호스팅", K.d); down(4.70);
  box(4.90, "Azure Functions (API)", "사투리 해석 · 합승 배차", K.d); down(5.44);
  box(5.64, "Azure OpenAI", "사투리 → 예약 정보(실패 시 규칙 기반)", K.o);
  s.addText("GitHub + Copilot으로 작성 → GitHub Actions가 테스트 후 배포", T({ x: 8.6, y: 6.2, w: 4.06, h: 0.3, fontSize: 9.5, color: K.sub }));
  s.addNotes("필요 기능입니다. 필수 네 가지, 말로 예약, 합승 배차, 어르신 화면, 배포까지를 14시 30분 첫 배포 전에 끝내고, 나머지는 시간이 남으면 붙입니다. 구조는 Azure Static Web Apps에 화면, Functions에 API, Azure OpenAI로 사투리를 해석하는 단순한 형태입니다.");

  // ===== 준비한 아이디어 2 (화면)
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "가안", "화면은 70세 이상 어르신 기준 — 한 화면에 한 질문");
  msg(s, "시니어 컬러 기준\n대비 4.5:1 이상", [
    "어디로 → 언제 → 확인 → 완료 4단계, 한 화면에 질문 하나, 주요 버튼 높이 64px 이상",
    "흰 바탕 + 짙은 검정 글자(대비 18.9:1), 누를 곳만 딥블루(6.8:1), '어둡게'·'가+'(글자 크게) 버튼",
    "화면 시안은 대회 전 참고용 — 당일엔 주제에 맞춰 Copilot으로 새로 구현",
  ]);
  panel(s, 0.5, 2.95, 6.6, 3.6, "주요 화면 (휴대폰)", "청도효도택시 화면 시안(2026.10), 화면 속 인물·예약은 가상");
  const shots = [["s1.png", "1. 어디로"], ["s2.png", "2. 언제"], ["s3.png", "3. 확인"], ["s4.png", "4. 완료"]];
  shots.forEach(([f, cap], i) => {
    const x = 0.72 + i * 1.6;
    s.addImage({ path: f, x, y: 3.42, w: 1.3, h: 2.81, objectName: "screen " + cap, altText: "청도효도택시 " + cap + " 화면" });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.42, w: 1.3, h: 2.81, fill: { type: "none" }, line: { color: "D3D8E0", width: 0.75 }, objectName: "screen frame " + cap });
    s.addText(cap, T({ x: x - 0.1, y: 6.27, w: 1.5, h: 0.24, fontSize: 10, bold: true, color: K.d, align: "center" }));
  });
  panel(s, 7.43, 2.95, 5.4, 3.6, "사용자 경험 설계", "시니어(70세 이상) UX 컬러 가이드 기준, 대비는 WCAG 계산");
  tbl(s, [
    [th("원칙"), th("적용")],
    [tdk("말로 예약"), td("사투리 그대로 → AI가 정리, 확인 내용 소리로 다시 읽기")],
    [tdk("아는 이름"), td("타는 곳은 실제 버스정류장 이름, 날짜는 요일까지")],
    [tdk("좌석 지도"), td("내 자리·함께 타는 분을 미리 안내")],
    [tdk("색"), td("파랑(내 자리)·딥오렌지(이웃) 구분, 파스텔+흰 글씨 금지")],
    [tdk("대리 신청"), td("이장님·경로당이 여러 분을 한 번에, 가족에게 알리기")],
  ], { x: 7.43, y: 3.38, w: 5.4, colW: [1.25, 4.15], rowH: 0.5, fontSize: 10.5 });
  s.addNotes("화면은 70세 이상 어르신 기준으로 설계했습니다. 한 화면에 질문 하나, 글자와 바탕 대비는 18.9대 1, 누를 곳만 딥블루입니다. 사투리로 말하면 AI가 정리하고, 타는 곳은 실제 버스정류장 이름으로 안내합니다. 이 시안은 참고용이고, 당일에는 주제에 맞춰 새로 만듭니다.");

  // ===== 예비 아이디어 6건
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "예비 아이디어", "주제 '청도에서 이게 되네?' 맞춤 6건 — 도시에서나 되는 게 청도에서");
  msg(s, "효도택시 1순위 유지\n+ 예비 3건", [
    "효도택시는 문장만 바꿔 유지: '청도에서 카카오T 합승이 되네?'",
    "예비 1 마을방송 번역기 — 근거 가장 탄탄, 효도택시 음성 엔진 재사용으로 개발 위험 최소",
    "예비 2 경로당 픽업 — 양양군이 조례·예산으로 이미 시행, 실현 가능성 근거 강함",
    "예비 3 터널 창고 — 청도에만 있는 자원, 단 터널 온도는 미확인이라 확인 전엔 발표에서 제외",
  ]);
  panel(s, 0.5, 2.95, 12.33, 3.6, "놀고 있는 자원 × 부족한 것 (✅ 확인 / ⚠️ 확인 필요)", "부록 3 출처 참고 (행안부 외국인주민 현황, 경북매일·매일신문·아주경제·뉴스서울 보도)");
  tbl(s, [
    [th("순위"), th("아이디어 · 청도에서 ○○가 되네?"), th("놀고 있는 자원 → 부족한 것"), th("확인된 근거"), th("확인 필요")],
    [hot("예비 1", { align: "center" }), td("마을방송 번역기 · 실시간 알림", { bold: true }), td("이장 마을방송 → 못 들은 어르신·외국인 근로자"), td("✅ 외국인 주민 1,680명(4.2%), 근로자 867명 · 계절근로자 84→450명"), td("⚠️ 군 스마트 마을방송 구축 범위")],
    [td("예비 2", { align: "center", bold: true }), td("경로당 픽업 · 문 앞 배송", { bold: true }), td("낮에 빈 경로당 → 외곽 면 택배·반품"), td("✅ 경로당 319곳 · 양양군 산간 택배 조례·시범"), td("⚠️ 청도 면 단위 미배송 실태")],
    [td("예비 3", { align: "center", bold: true }), td("터널 창고 · 냉장 보관", { bold: true }), td("와인터널(옛 경부선) → 농가 저온 저장"), td("✅ 길이 1.01km, 와인 15만 병 저장"), td("⚠️ 내부 온도·남는 공간")],
    [td("4", { align: "center" }), td("농기계 나눔 · 공유경제"), td("며칠만 쓰는 농기계 → 소농·귀농인"), td("✅ 군 임대 785대, 연 11,736건"), td("⚠️ 군 사업과 겹쳐 신선도 약함")],
    [td("5", { align: "center" }), td("빈 교실 워케이션 · 원격근무"), td("빈 교실·폐교 → 체류 인구"), td("✅ 금천고 2027.3 폐교 · 생활인구 34.4만"), td("⚠️ 폐교 활용 현황·개방 규정")],
    [td("6", { align: "center" }), td("할매 손맛 클래스 · 원데이클래스"), td("어르신 손기술 → 체류형 관광"), td("✅ 관광주민증 이용 1.2% · 65세 이상 45.1%"), td("⚠️ 기존 농촌체험 프로그램")],
  ], { x: 0.5, y: 3.38, w: 12.33, colW: [0.8, 2.75, 3.1, 3.5, 2.18], rowH: 0.45, fontSize: 9.5 });
  s.addNotes("주제가 '청도에서 이게 되네?'라서, 도시에서나 되는 서비스가 청도에서 된다는 방향으로 예비 아이디어 여섯 개를 공공데이터로 확인했습니다. 효도택시는 그대로 1순위로 두고, 예비 1은 마을방송 번역기, 예비 2는 경로당 픽업, 예비 3은 터널 창고입니다.");

  // ===== 예비 아이디어 근거
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "아이디어 준비" });
  head(s, "예비 아이디어 근거", "외국인 계절근로자 4년 새 5배, 경로당은 319곳");
  msg(s, "예비 1·2는\n숫자로 설명 가능", [
    "청도 외국인 주민 1,680명 중 근로자 867명(51.6%) — 마을방송·안내문을 못 알아듣는 주민이 늘어남",
    "계절근로자 2023년 19농가 84명 → 2026년 127농가 450명, 누적 951명(필리핀 카빈티시 출신 62%)",
    "경로당 319곳(운영 지원 47억 원)이 낮엔 비어 있음 → 택배 거점으로 쓰면 새 건물 없이 해결",
  ]);
  panel(s, 0.5, 2.95, 6.0, 3.6, "청도 외국인 계절근로자 (명)", "매일신문 2026.10.8(청도군 발표), 2024·2025년 수치는 미공개");
  s.addChart(pres.charts.BAR, [
    { name: "2023", labels: ["2023년 (19농가)", "2026년 (127농가)"], values: [84, 0] },
    { name: "2026", labels: ["2023년 (19농가)", "2026년 (127농가)"], values: [0, 450] },
  ], Object.assign({}, ax, { x: 0.55, y: 3.4, w: 5.9, h: 3.1, barDir: "col", barGrouping: "stacked", chartColors: [K.g, K.o], showValue: true, dataLabelPosition: "inEnd", dataLabelFormatCode: "0\"명\";;", dataLabelColor: K.white, dataLabelFontSize: 13,
    showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 520, barGapWidthPct: 70, catAxisLabelFontSize: 10.5 }));
  panel(s, 6.83, 2.95, 6.0, 3.6, "경로당 픽업 — 다른 군은 이미 시행", "경북매일 2025.5.6, 아주경제 2026.8.4, 강원일보 2025.9.22");
  tbl(s, [
    [th("지역"), th("현황")],
    [tdk("청도"), td("경로당 319곳, 운영 지원 47억 원 — 낮 시간 거점으로 활용 가능")],
    [tdk("양양"), hot("2025.4 조례 → 2026.8 시범 2개 마을 69가구, 마을회관 앞 택배보관함")],
    [tdk("양양 계획"), td("5개 읍·면 13개 마을 129가구로 확대")],
    [tdk("홍천"), td("풍천1·2리 150여 가구 집에서 택배 못 받음, '마을회관으로 찾으러'")],
    [tdk("청도"), td("⚠️ 면 단위 미배송 실태는 군청·택배사 확인 필요")],
  ], { x: 6.83, y: 3.38, w: 6.0, colW: [1.2, 4.8], rowH: 0.5, fontSize: 10.5 });
  s.addNotes("예비 아이디어 근거입니다. 청도 외국인 계절근로자는 2023년 84명에서 2026년 450명으로 늘었습니다. 마을방송을 못 알아듣는 주민이 그만큼 늘었다는 뜻입니다. 경로당은 319곳이 있고, 양양군은 이미 조례를 만들어 마을회관 앞 택배보관함을 시범 운영 중입니다.");

  // ===== 부록 표지
  pres.addSection({ title: "부록" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "부록" });
  s.addText("부록", T({ x: 0.8, y: 2.3, w: 11, h: 1.2, fontSize: 60, bold: true, color: K.white }));
  s.addText("1. 대회 개요와 규칙   2. 팀원 제안 검토(길찾기 앱)·버스 노선   3. 출처", T({ x: 0.8, y: 3.6, w: 11.7, h: 0.6, fontSize: 18, color: "DCE6F7" }));

  // ===== 부록2 대회 개요
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "부록" });
  head(s, "부록 1", "대회 개요와 규칙");
  msg(s, "청도 지역 문제가\n주제일 가능성 높음", [
    "주최 해커그라운드 · 주관 대구한의대 · 후원 Microsoft·청도군 → 청도 지역 문제가 주제로 나올 가능성 높음(추정)",
    "대회 전 작성 코드 사용 가능 여부는 공지에 없음 → 대회 전엔 방향·역할·도구 연습까지, 제출 코드는 당일 새 저장소에서",
    "운영진의 팀별 저장소·Azure 리소스 제공 여부는 미정(2024년엔 제공) → 각자 개인 계정도 준비",
  ], { noBoldLast: true });
  panel(s, 0.5, 2.95, 12.33, 3.6, "대회 개요", "해커그라운드 공식 페이지, 참가 신청 접수 메일(2026.10.25)");
  tbl(s, [
    [th("항목"), th("내용"), th("항목"), th("내용")],
    [tdk("일시"), td("11/5(목) 09:00 등록 → 17:30 폐회·석식"), tdk("장소"), td("청도 신화랑풍류마을(운문면 신화랑길 1)")],
    [tdk("주제"), td("10:00 현장 공개"), tdk("팀"), td("1~4인, 공식 등록은 현장")],
    [tdk("제출물"), td("기획서 · 코드(GitHub) · 배포된 서비스"), tdk("심사 기준"), td("기획 · 코드 구현 · 사용자 경험")],
    [tdk("필수 도구"), td("GitHub Copilot, Azure"), tdk("발표"), td("없음(기획서·README가 대신)")],
    [tdk("셔틀"), td("11/5 08:00 동대구역 출발"), tdk("귀가"), td("셔틀 없음, 택시·합승")],
  ], { x: 0.5, y: 3.38, w: 12.33, colW: [1.3, 4.86, 1.3, 4.87], rowH: 0.5, fontSize: 10.5 });

  // ===== 부록 버스 현황
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "부록" });
  head(s, "부록 2", "팀원 제안 '어르신 길찾기 앱' 검토 → 효도택시 기능으로 흡수");
  msg(s, "제안 내용\n(10/26 팀원 제안)", [
    "제안: 목적지만 넣으면 탈 버스·내릴 곳·남은 시간을 큰 화면과 음성으로 알려 주는 '어르신용 네이버 길찾기'",
    "제안 이유: 네이버 길찾기는 어르신에게 복잡하고, 버스에서 '어디서 내려요?', '몇 번 타요?' 묻는 어르신이 많음",
    "검토 결과: 청도는 노선·운행 횟수가 적어 '어떤 버스'보다 '버스가 안 온다'가 더 큰 문제 → 단독 아이디어로는 약함",
    "반영: 효도택시 안에 '곧 오는 버스 있으면 버스 안내, 없으면 합승택시 예약'으로 흡수",
  ]);
  panel(s, 0.5, 2.95, 12.33, 3.6, "검토 근거 — 청도 버스 현황", "나무위키 청도 버스 1·7번(2016.5 시간표, 현재 시간표 확인 필요), 주간경향 2026.4.27(국토부·통계청 인용), 경북매일 2021.7.4");
  tbl(s, [
    [th("포인트"), th("근거"), th("출처")],
    [tdk("노선마다 운행 횟수가 적음"), td("청도 7번: 덕암 방면 하루 7회, 남성현·숯골 방면 하루 5회, 버스 1대가 전부 운행"), td("나무위키(2016년 시간표)")],
    [tdk("주요 간선도 듬성듬성"), td("1번(청도터미널–풍각): 하루 약 21회, 배차 15~60분"), td("나무위키(2016년)")],
    [tdk("경북 자체가 배차가 김"), hot("경북 평균 배차간격 70.7분, 서울 10.1분"), td("국토부 대중교통 현황조사(2022)")],
    [tdk("대중교통 없는 마을 증가"), td("걸어서 15분 안에 대중교통이 없는 마을 전국 2,224곳, 2015년의 2배 이상"), td("통계청 농림어업총조사(2020)")],
    [tdk("군도 버스가 안 닿는다고 인정"), td("청도군이 대중교통 취약지역에 행복버스·행복택시 운영 중"), td("경북매일 2021.7.4")],
    [tdk("사용자 대부분이 어르신"), td("65세 이상 45.1%, 외곽 면 평균 64세"), td("행안부·군 발표")],
  ], { x: 0.5, y: 3.38, w: 12.33, colW: [2.9, 6.4, 3.03], rowH: 0.45, fontSize: 10.5 });

  // ===== 부록 버스 노선 구성
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "부록" });
  head(s, "부록 2-2", "청도 농어촌버스 노선 구성 — 운수사 1곳, 주요 노선 6개");
  msg(s, "노선이 적어\n'고를 경로'가 적음", [
    "운수사는 청도버스 1곳, 모든 노선이 청도공용버스터미널에서 출발",
    "주요 노선 6개(0·1·2·3·5·7번)가 방면별로 나뉨 → 목적지가 정해지면 탈 버스도 거의 정해짐",
    "일부 노선은 2~4시간 간격(팀 수집 자료, 공식 시간표로 확인 필요)",
    "네이버·카카오 지도에서 청도 버스 실시간 조회 가능하다는 정보 → 사실이면 길찾기 앱 차별점은 더 약해짐(확인 필요)",
  ], { noBoldLast: true });
  panel(s, 0.5, 2.95, 12.33, 3.6, "주요 노선 (청도공용버스터미널 기점)", "HS 수집 자료(2026.10.26, 원 출처 미기재), 나무위키 청도 버스 1·7번, 노선 정보 사이트(3·5번) — 청도군 공식 버스 시간표로 재확인 필요");
  tbl(s, [
    [th("노선"), th("주요 경유지"), th("특징"), th("확인 상태")],
    [tdk("0번"), td("청도터미널 ↔ 팔조령 ↔ 가창 ↔ 대구(가창중학교)"), td("대구 시내 연계"), td("⚠️ 노선 존재 확인, 경유지 미확인")],
    [tdk("1번"), td("청도터미널 ↔ 화양 ↔ 각남 ↔ 풍각버스터미널"), td("서부 권역 핵심"), td("✅ 하루 약 21회, 15~60분(2016)")],
    [tdk("2번"), td("청도터미널 ↔ 유등 ↔ 대곡 ↔ 신촌 ↔ 수야"), td("이서면 마을 연계"), td("⚠️ 노선 존재 확인, 경유지 미확인")],
    [tdk("3번"), td("청도터미널 ↔ 금천동 ↔ 운산 ↔ 운문사"), td("동부 권역, 운문사"), td("✅ 3번 노선 정류소 12개 구간 확인")],
    [tdk("5번"), td("청도터미널 ↔ 유천 ↔ 온막 ↔ 상동역(밀양)"), td("밀양 연계"), td("✅ 동곡행 정류소 40개, 5-2번 상동행")],
    [tdk("7번"), td("청도터미널 ↔ 남성현 ↔ 와인터널 ↔ 덕암"), td("관광지 연계"), hot("✅ 방면별 하루 5~7회, 버스 1대(2016)")],
  ], { x: 0.5, y: 3.38, w: 12.33, colW: [0.9, 5.0, 2.2, 4.23], rowH: 0.45, fontSize: 10 });

  // ===== 부록4 출처
  s = pres.addSlide({ masterName: "REPORT", sectionTitle: "부록" });
  head(s, "부록 3", "출처");
  tbl(s, [
    [th("구분"), th("자료"), th("쓰인 곳")],
    [tdk("대회"), td("해커그라운드 공식 페이지(hackersground.kr), 참가 신청 접수 메일(2026.10.25)"), td("일정·제출물·심사 기준·셔틀")],
    [tdk("대회"), td("2023 대상 DEPS 팀 회고(zedy_dev.log), 2024 대상 나누리 팀 회고(김건우 velog)"), td("지난 대상작·심사 과정")],
    [tdk("대회"), td("배포 기회 2번 — 팀 내 공유 정보"), hot("공식 출처 확인 필요")],
    [tdk("인구"), td("행정안전부 주민등록 인구통계(2026.9)"), td("연령대별 인구, 읍면별 평균 연령")],
    [tdk("인구·돌봄"), td("매일신문 2025.5.8 「행복하고 건강한 노년, 인생 이모작」"), td("65세 이상 45.1%, 노인일자리, 비상벨")],
    [tdk("의료"), td("뉴스1 2026.8.18(의료취약지 76곳), 경북매일 2026.7.20"), td("의료취약지, 병원 없는 면, 응급실")],
    [tdk("의료"), td("공공데이터포털 「청도군_집단급식소 운영 병원」(2026.6.8)"), td("병원 현황·응급실 병상")],
    [tdk("관광"), td("경북매일 2026.5.26, 청도군 통합예약시스템 관광택시 화면(2026.10.24)"), td("관광주민증, 관광택시")],
    [tdk("관광"), td("경북매일 2026.10.7(생활인구) — 원문 링크 재확인 필요"), td("생활인구 34.4만")],
    [tdk("교통"), td("공공데이터포털 「청도군_버스정류장 현황」(2025.8.28), 헤럴드경제 2026.8.2, 한국관광공사 관광정보"), td("정류장 수, 병원까지 시간(추정), 호출 앱, 행사장")],
    [tdk("예비 1"), td("행정안전부 외국인주민 현황(청도군, 2026.9 집계), 매일신문 2026.10.8, 경북매일 2026.10.1"), td("외국인 주민·근로자, 계절근로자")],
    [tdk("예비 2"), td("경북매일 2025.5.6(경로당), 아주경제 2026.8.4(양양 택배), 강원일보 2025.9.22(홍천)"), td("경로당 319곳, 산간 택배 사례")],
    [tdk("예비 3~6"), td("한국관광공사 여행코스(와인터널), 뉴스서울 2024.7.17(농기계), 경북매일 2026.10.20(금천고)"), td("터널 규모, 임대 실적, 폐교")],
    [tdk("버스"), td("나무위키 청도 버스 1·7번(2016.5), 노선 정보 사이트, HS 수집 자료(10.26), 주간경향 2026.4.27"), td("노선별 운행 횟수, 경북 배차간격")],
  ], { x: 0.5, y: 1.2, w: 12.33, colW: [1.3, 7.3, 3.73], rowH: 0.38, fontSize: 9.5 });

  // ===== 마무리
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "마무리" });
  s.addText("회의에서 정할 것", T({ x: 0.8, y: 1.9, w: 11, h: 1.0, fontSize: 48, bold: true, color: K.white }));
  s.addText([
    { text: "1. '배포 2번'의 정확한 의미와 출처", options: { breakLine: true } },
    { text: "2. 역할 분담 확정과 디자이너 추가 여부", options: { breakLine: true } },
    { text: "3. 기술 스택(화면·API·AI·배포)과 대회 전 연습 배포 일정" },
  ], T({ x: 0.8, y: 3.1, w: 11.5, h: 1.8, fontSize: 22, color: "DCE6F7", paraSpaceAfter: 8 }));
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 6.15, w: 11.73, h: 0, line: { color: "5470A8", width: 0.75 }, objectName: "end foot rule" });
  s.addText([{ text: "작성  ", options: { color: "C9D6EE" } }, { text: "김혜수", options: { bold: true, color: K.white } }], T({ x: 8.8, y: 6.3, w: 3.73, h: 0.4, fontSize: 13, align: "right" }));

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
