import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

// =========================================================
// 스트레스 취약성 검사 고도화 버전
// - 48문항 전체 적용
// - 영역별 점수 / 취약·보통·건강 판정
// - 원문 문항 유지 / 직접 점수 합산
// - 개인 결과 리포트
// - 조직/부서 분석 대시보드
// - 위험군 탐지
// - 기간별 변화 분석
// - AI형 자동 해석/개선 제안
// - Excel 다중 시트 다운로드
// - 인쇄/PDF 저장
// - Firebase 연동을 위한 데이터 모델/어댑터 구조 포함
// =========================================================

// 인쇄 시 결과지만 나오도록 스타일 주입
const printStyle = `
button, input, select {
  -webkit-tap-highlight-color: transparent;
}

button,
button:focus,
button:active,
button:focus:not(:focus-visible),
input:focus:not(:focus-visible),
select:focus:not(:focus-visible) {
  outline: none;
  box-shadow: none;
}

button:focus-visible,
input:focus-visible,
select:focus-visible {
  outline: 2px solid #1D9E75;
  outline-offset: 2px;
}

@media print {
  body * { visibility: hidden; }
  #print-area, #print-area *, #admin-print-area, #admin-print-area * { visibility: visible; }
  #print-area, #admin-print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 24px; }
  .no-print { display: none !important; }
}
`;

// ─── Types ───────────────────────────────────────────────
type AnswerValue = 0 | 1 | 3;
type CatKey = "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H";
type Level = "취약" | "보통" | "건강";
type AppMode = "test" | "dashboard" | "report" | "settings";
type PeriodPreset = "all" | "30d" | "90d" | "180d" | "365d";

interface Question {
  id: number;
  text: string;
  cat: CatKey;
}

interface SavedResult {
  id: string;
  name: string;
  dept: string;
  total: number;
  status: string;
  scores: Record<CatKey, number>;
  levels: Record<CatKey, Level>;
  answers: Record<number, AnswerValue>;
  weakCats: CatKey[];
  riskLevel: "낮음" | "주의" | "위험";
  date: string;
  createdAt: string;
  consent: boolean;
}

interface DeptStat {
  dept: string;
  count: number;
  avgTotal: number;
  riskRate: number;
  weakAvgCount: number;
  catAvg: Record<CatKey, number>;
}

// ─── Data ────────────────────────────────────────────────
const CAT_KEYS: CatKey[] = ["A", "B", "C", "D", "E", "F", "G", "H"];

const CAT_NAMES: Record<CatKey, string> = {
  A: "식습관",
  B: "음주·흡연",
  C: "운동",
  D: "사회적관계",
  E: "성생활",
  F: "여가·휴식",
  G: "일·삶균형",
  H: "자기이해",
};

const CAT_ADVICE: Record<CatKey, string> = {
  A: "규칙적인 식사와 충분한 수분 섭취, 균형 잡힌 영양 관리가 필요합니다.",
  B: "음주·흡연 빈도를 줄이고 건강한 스트레스 해소 방법을 찾아보세요.",
  C: "주 3회 이상 가벼운 유산소 운동부터 시작해보세요.",
  D: "정서적 지지를 줄 수 있는 관계를 회복하고 소통 빈도를 늘려보세요.",
  E: "파트너와의 정서적 대화 시간을 늘리고 친밀감을 회복해보세요.",
  F: "쉼과 취미 활동을 일정에 포함하고 충분한 수면을 취하세요.",
  G: "업무와 삶의 경계를 명확히 하고 야근을 줄여보세요.",
  H: "명상·기도·상담 등을 통해 자기 자신을 돌보는 시간을 가지세요.",
};

const CAT_PROGRAMS: Record<CatKey, string[]> = {
  A: ["식사 리듬 점검", "카페인·수분 섭취 체크", "건강 간식 가이드"],
  B: ["음주 빈도 줄이기", "흡연 대체 행동 계획", "회식 문화 개선"],
  C: ["점심 산책 챌린지", "주 3회 20분 운동", "계단 이용 캠페인"],
  D: ["동료 지지 모임", "감정표현 훈련", "상담 접근성 안내"],
  E: ["관계 대화 가이드", "친밀감 회복 교육", "개인 상담 연계"],
  F: ["수면 루틴 만들기", "무자극 휴식 시간", "취미 회복 워크숍"],
  G: ["퇴근 경계 세우기", "야근 원인 분석", "업무량 재조정"],
  H: ["마음챙김 명상", "자기이해 저널링", "회복탄력성 코칭"],
};

const QUESTIONS: Question[] = [
  { id: 1, text: "적절한 음식을 적당량 먹는다.", cat: "A" },
  { id: 2, text: "점심식사 때 술을 마시지 않는다.", cat: "B" },
  { id: 3, text: "일주일에 적어도 3번 정도는 땀이 날 정도로 운동을 한다.", cat: "C" },
  { id: 4, text: "친구나 친척과 정기적 계 모임이나 동창모임을 갖는다.", cat: "D" },
  { id: 5, text: "나의 성생활에 만족한다.", cat: "E" },
  { id: 6, text: "적어도 한 가지 이상의 여가활동이나 규칙적으로 즐기는 흥미거리가 있다.", cat: "F" },
  { id: 7, text: "주말에는 일하지 않고 쉬거나 여가활동을 한다.", cat: "G" },
  { id: 8, text: "규칙적으로 기도를 하거나 명상을 한다.", cat: "H" },
  { id: 9, text: "하루에 커피, 또는 콜라 등을 5잔 이상 마시지 않는다.", cat: "A" },
  { id: 10, text: "음료수를 술 마시기보다 더 좋아한다.", cat: "B" },
  { id: 11, text: "일상생활 가운데 적당한 신체에너지를 사용하는 일을 한다.", cat: "C" },
  { id: 12, text: "감정을 솔직하게 표현한다.", cat: "D" },
  { id: 13, text: "규칙적으로 애정을 주고받는다.", cat: "E" },
  { id: 14, text: "매일 휴식시간을 갖는다.", cat: "F" },
  { id: 15, text: "할 수 있는 일(현실적인 일)만 하며 과도하게 일하지 않는다.", cat: "G" },
  { id: 16, text: "나의 문제를 내 스스로 해결한다.", cat: "H" },
  { id: 17, text: "나의 키에 알맞은 적절한 체중을 유지한다.", cat: "A" },
  { id: 18, text: "나 혼자서는 술을 마시지 않는다.", cat: "B" },
  { id: 19, text: "엘리베이터나 승강기를 이용하기보다는 계단을 걸어 올라간다.", cat: "C" },
  { id: 20, text: "부정적 감정을 쌓아두기보다는 잘 표현한다.", cat: "D" },
  { id: 21, text: "성적으로 무력하거나 냉담하지 않는다.", cat: "E" },
  { id: 22, text: "적어도 일주일에 나흘 정도는 7~8시간 잠을 잔다.", cat: "F" },
  { id: 23, text: "결코 일(사업)이 나의 생활을 지배하도록 두지 않는다.", cat: "G" },
  { id: 24, text: "나 자신을 믿는다.", cat: "H" },
  { id: 25, text: "음식에 소금을 넣지 않는다.", cat: "A" },
  { id: 26, text: "퇴근 후 집에서 술을 잘 마시지 않는다.", cat: "B" },
  { id: 27, text: "규칙적인 운동 프로그램을 실천한다.", cat: "C" },
  { id: 28, text: "매우 절친한 친구가 있어 그들과 은밀한 문제를 논의할 수 있다.", cat: "D" },
  { id: 29, text: "사랑이 충만한 성관계를 갖는다.", cat: "E" },
  { id: 30, text: "일주일에 적어도 한번은 재미있는 일을 한다.", cat: "F" },
  { id: 31, text: "사회생활에서 나 자신의 사업에 대해 말하는 것을 피한다.", cat: "G" },
  { id: 32, text: "내 마음이 평안하다는 느낌을 갖는다.", cat: "H" },
  { id: 33, text: "매일 규칙적으로 식사를 하고 군것질은 하지 않는다.", cat: "A" },
  { id: 34, text: "적당하게 술을 마시거나(하루 2잔 미만), 전혀 마시지 않는다.", cat: "B" },
  { id: 35, text: "매주 한가지 이상의 운동을 한다.", cat: "C" },
  { id: 36, text: "직장에서 나를 정서적으로 지지해 주는 동료가 있다.", cat: "D" },
  { id: 37, text: "안정된 성관계를 유지하길 좋아한다.", cat: "E" },
  { id: 38, text: "나 스스로 즐거운 일거리를 마련할 수 있다.", cat: "F" },
  { id: 39, text: "야간에는 일하지 않는다.", cat: "G" },
  { id: 40, text: "어떤 일이나 어떤 단체에 대해 높은 소속감을 갖는다.", cat: "H" },
  { id: 41, text: "규칙적으로 물을 마신다.", cat: "A" },
  { id: 42, text: "하루 3개피 이하의 담배를 피우거나 아예 피우지 않는다.", cat: "B" },
  { id: 43, text: "매일 산책하거나 길을 걷는다.", cat: "C" },
  { id: 44, text: "친구로부터 도움을 구하거나 필요하면 전문적 충고를 받는다.", cat: "D" },
  { id: 45, text: "거의 성적 좌절감을 느끼지 않는다.", cat: "E" },
  { id: 46, text: "아무것도 하지 않고도 시간을 보낼 수 있다.", cat: "F" },
  { id: 47, text: "가정생활과 직장생활이 나에게는 똑같이 중요하다.", cat: "G" },
  { id: 48, text: "스트레스 상황을 극복하는 것을 배운 적이 있다.", cat: "H" },
];

// ─── Helpers ─────────────────────────────────────────────
const storageKey = "stressResults.v2";
const MIN_DEPT_SAMPLE = 5;

function uid() {
  return `sr_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function formatDate(date = new Date()) {
  return date.toLocaleDateString("ko-KR");
}

function normalizeDateKey(iso: string) {
  const d = new Date(iso);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function getLevel(score: number): Level {
  if (score <= 9) return "취약";
  if (score <= 13) return "보통";
  return "건강";
}

function getLevelStyle(level: Level) {
  if (level === "취약") return { bg: "#FCEBEB", border: "#F09595", text: "#791F1F" };
  if (level === "보통") return { bg: "#FAEEDA", border: "#EF9F27", text: "#633806" };
  return { bg: "#E1F5EE", border: "#5DCAA5", text: "#085041" };
}

function getTotalStatus(total: number) {
  if (total >= 112) return { label: "매우 건강", bg: "#E1F5EE", color: "#085041" };
  if (total >= 96) return { label: "양호", bg: "#E1F5EE", color: "#085041" };
  if (total >= 64) return { label: "보통", bg: "#FAEEDA", color: "#633806" };
  return { label: "취약", bg: "#FCEBEB", color: "#791F1F" };
}

function hasRepeatedWeak(scores?: Record<CatKey, number>, previous?: SavedResult | null) {
  if (!scores || !previous) return false;
  return CAT_KEYS.some((cat) => getLevel(scores[cat]) === "취약" && previous.levels[cat] === "취약");
}

function getRiskLevel(
  weakCount: number,
  total: number,
  scores?: Record<CatKey, number>,
  previous?: SavedResult | null
): SavedResult["riskLevel"] {
  const workLifeWeak = scores ? getLevel(scores.G) === "취약" : false;
  const restWeak = scores ? getLevel(scores.F) === "취약" : false;
  const previousDrop = previous ? previous.total - total : 0;

  if (
    weakCount >= 3 ||
    total < 64 ||
    ((workLifeWeak || restWeak) && total < 96) ||
    previousDrop >= 15
  ) return "위험";

  if (
    weakCount >= 1 ||
    total < 96 ||
    previousDrop >= 8 ||
    hasRepeatedWeak(scores, previous)
  ) return "주의";

  return "낮음";
}

function scoreAnswer(val: AnswerValue): number {
  return val;
}

function emptyScores(): Record<CatKey, number> {
  return CAT_KEYS.reduce((acc, cat) => {
    acc[cat] = 0;
    return acc;
  }, {} as Record<CatKey, number>);
}

function calcScores(answers: Record<number, AnswerValue>) {
  const scores = emptyScores();
  QUESTIONS.forEach((q) => {
    const val = answers[q.id];
    if (val !== undefined) scores[q.cat] += scoreAnswer(val);
  });
  return scores;
}

function calcLevels(scores: Record<CatKey, number>) {
  return CAT_KEYS.reduce((acc, cat) => {
    acc[cat] = getLevel(scores[cat]);
    return acc;
  }, {} as Record<CatKey, Level>);
}

function average(nums: number[]) {
  if (!nums.length) return 0;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function round1(n: number) {
  return Math.round(n * 10) / 10;
}

function loadResults(): SavedResult[] {
  try {
    const raw = localStorage.getItem(storageKey) || "[]";
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveResults(results: SavedResult[]) {
  localStorage.setItem(storageKey, JSON.stringify(results));
}

function filterByPeriod(results: SavedResult[], preset: PeriodPreset) {
  if (preset === "all") return results;
  const days = Number(preset.replace("d", ""));
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  return results.filter((r) => new Date(r.createdAt) >= cutoff);
}

function makeResult(params: {
  name: string;
  dept: string;
  answers: Record<number, AnswerValue>;
  consent: boolean;
  previousResult?: SavedResult | null;
}): SavedResult {
  const scores = calcScores(params.answers);
  const levels = calcLevels(scores);
  const weakCats = CAT_KEYS.filter((cat) => levels[cat] === "취약");
  const total = Object.values(scores).reduce((a, b) => a + b, 0);
  const totalStatus = getTotalStatus(total);
  return {
    id: uid(),
    name: params.name.trim() || "이름 미입력",
    dept: params.dept.trim() || "부서 미입력",
    total,
    status: totalStatus.label,
    scores,
    levels,
    answers: params.answers,
    weakCats,
    riskLevel: getRiskLevel(weakCats.length, total, scores, params.previousResult || null),
    date: formatDate(),
    createdAt: new Date().toISOString(),
    consent: params.consent,
  };
}

function findPreviousResult(results: SavedResult[], name: string, dept: string) {
  const cleanName = name.trim();
  const cleanDept = dept.trim();
  if (!cleanName && !cleanDept) return null;
  return [...results]
    .filter((r) => (!cleanName || r.name === cleanName) && (!cleanDept || r.dept === cleanDept))
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0] || null;
}

function findPreviousComparableResult(results: SavedResult[], current: SavedResult) {
  return [...results]
    .filter((r) => r.id !== current.id && r.name === current.name && r.dept === current.dept && new Date(r.createdAt) < new Date(current.createdAt))
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0] || null;
}

function displayNameForAdmin(result: SavedResult) {
  if (!result.consent) return "익명";
  if (!result.name || result.name === "이름 미입력") return "이름 미입력";
  if (result.name.length <= 2) return `${result.name[0]}*`;
  return `${result.name[0]}${"*".repeat(result.name.length - 2)}${result.name[result.name.length - 1]}`;
}

function getPatternInsights(result: SavedResult) {
  const isWeak = (cat: CatKey) => result.levels[cat] === "취약";
  const patterns: { title: string; body: string; recommendation: string }[] = [];

  if (isWeak("G") && isWeak("F")) {
    patterns.push({
      title: "과로·회복 부족 패턴",
      body: "일·삶균형과 여가·휴식이 함께 낮아 업무 경계와 회복 시간이 동시에 약화되어 있을 가능성이 있습니다.",
      recommendation: "퇴근 후 업무 연결을 줄이고, 주 1회 이상 회복 시간을 일정에 먼저 배치하세요.",
    });
  }
  if (isWeak("D") && isWeak("H")) {
    patterns.push({
      title: "정서적 지지 자원 부족 패턴",
      body: "사회적관계와 자기이해가 함께 낮아 스트레스 상황에서 혼자 버티는 양상이 나타날 수 있습니다.",
      recommendation: "신뢰할 수 있는 동료·친구와의 대화 루틴을 만들고 필요 시 상담 자원을 연결하세요.",
    });
  }
  if (isWeak("A") && isWeak("B") && isWeak("C")) {
    patterns.push({
      title: "신체 건강 루틴 취약 패턴",
      body: "식습관, 음주·흡연, 운동 영역이 함께 낮아 스트레스가 신체 컨디션 관리로 이어지지 못할 수 있습니다.",
      recommendation: "식사·수분·걷기처럼 부담이 작은 행동부터 2주간 고정 루틴으로 만드세요.",
    });
  }
  if (isWeak("G") && isWeak("D")) {
    patterns.push({
      title: "업무 부담 고립 패턴",
      body: "업무와 삶의 경계가 약하고 지지 관계도 낮아 업무 스트레스를 혼자 떠안을 가능성이 있습니다.",
      recommendation: "업무량 조정 대화, 동료 지원 체계, 관리자 면담을 병행하세요.",
    });
  }
  if (isWeak("F") && isWeak("H")) {
    patterns.push({
      title: "내적 회복감 저하 패턴",
      body: "휴식과 자기이해가 함께 낮아 쉬어도 회복감을 느끼기 어려울 수 있습니다.",
      recommendation: "짧은 명상, 저널링, 수면 루틴처럼 자극을 낮추는 회복 전략을 우선 적용하세요.",
    });
  }

  return patterns;
}

function getTwoWeekActionPlan(result: SavedResult) {
  const lowest = CAT_KEYS.map((cat) => ({ cat, score: result.scores[cat] })).sort((a, b) => a.score - b.score)[0];
  const programs = CAT_PROGRAMS[lowest.cat];
  return [
    {
      label: "1주차",
      body: `${CAT_NAMES[lowest.cat]} 영역을 중심으로 '${programs[0]}'를 작게 시작합니다. 매일 5~10분 안에 끝나는 행동으로 설정하세요.`,
    },
    {
      label: "2주차",
      body: `'${programs[1] || programs[0]}'를 추가하고, 실천 여부를 주 3회 이상 체크합니다. 실패한 날은 원인을 기록하고 강도를 낮추세요.`,
    },
    {
      label: "체크 항목",
      body: `실천 횟수, 피로감, 수면·휴식감, 업무 경계 유지 여부를 간단히 기록한 뒤 2~4주 후 재검사를 권장합니다.`,
    },
  ];
}

function buildDeptStats(results: SavedResult[]): DeptStat[] {
  const map = new Map<string, SavedResult[]>();
  results.forEach((r) => {
    const key = r.dept || "부서 미입력";
    map.set(key, [...(map.get(key) || []), r]);
  });
  return [...map.entries()]
    .map(([dept, rows]) => {
      const catAvg = emptyScores();
      CAT_KEYS.forEach((cat) => {
        catAvg[cat] = round1(average(rows.map((r) => r.scores[cat])));
      });
      const riskCount = rows.filter((r) => r.riskLevel === "위험" || r.riskLevel === "주의").length;
      return {
        dept,
        count: rows.length,
        avgTotal: round1(average(rows.map((r) => r.total))),
        riskRate: round1((riskCount / rows.length) * 100),
        weakAvgCount: round1(average(rows.map((r) => r.weakCats.length))),
        catAvg,
      };
    })
    .sort((a, b) => b.riskRate - a.riskRate || a.avgTotal - b.avgTotal);
}

function buildTrend(results: SavedResult[]) {
  const map = new Map<string, SavedResult[]>();
  results.forEach((r) => {
    const key = normalizeDateKey(r.createdAt);
    map.set(key, [...(map.get(key) || []), r]);
  });
  return [...map.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, rows]) => ({
      date,
      count: rows.length,
      avgTotal: round1(average(rows.map((r) => r.total))),
      riskRate: round1((rows.filter((r) => r.riskLevel !== "낮음").length / rows.length) * 100),
    }));
}

function getWeakCategoryRanking(results: SavedResult[]) {
  const counts = CAT_KEYS.reduce((acc, cat) => {
    acc[cat] = 0;
    return acc;
  }, {} as Record<CatKey, number>);
  results.forEach((r) => r.weakCats.forEach((cat) => counts[cat]++));
  return CAT_KEYS.map((cat) => ({ cat, count: counts[cat] })).sort((a, b) => b.count - a.count);
}

function getOrgAiInsight(results: SavedResult[], deptStats: DeptStat[]) {
  if (!results.length) {
    return {
      summary: "아직 저장된 검사 결과가 없습니다. 검사를 완료하고 저장하면 조직 분석이 활성화됩니다.",
      bullets: ["검사 완료 후 저장 버튼을 눌러 조직 대시보드에 반영하세요."],
    };
  }

  const avgTotal = round1(average(results.map((r) => r.total)));
  const riskRate = round1((results.filter((r) => r.riskLevel !== "낮음").length / results.length) * 100);
  const weakRank = getWeakCategoryRanking(results).filter((x) => x.count > 0);
  const topWeak = weakRank[0];
  const eligibleDeptStats = deptStats.filter((d) => d.count >= MIN_DEPT_SAMPLE);
  const riskiestDept = eligibleDeptStats[0];

  const bullets: string[] = [];
  if (topWeak) {
    bullets.push(`${CAT_NAMES[topWeak.cat]} 영역의 취약 빈도가 가장 높습니다. ${CAT_PROGRAMS[topWeak.cat].slice(0, 2).join(" · ")} 프로그램을 우선 배치하세요.`);
  }
  if (riskiestDept && riskiestDept.riskRate >= 40) {
    bullets.push(`${riskiestDept.dept}의 주의/위험 비율이 ${riskiestDept.riskRate}%로 높습니다. 부서장 인터뷰와 업무량·근무시간 점검이 필요합니다.`);
  }
  if (riskRate >= 30) {
    bullets.push("조직 차원의 회복탄력성 개입이 필요합니다. 개인 상담보다 팀 단위 예방 교육과 제도 개선을 병행하는 것이 좋습니다.");
  } else {
    bullets.push("전반적 위험 비율은 관리 가능한 수준입니다. 취약 영역을 조기 발견하는 정기 모니터링 체계를 유지하세요.");
  }
  if (avgTotal < 96) {
    bullets.push("평균 총점이 양호 기준보다 낮습니다. 생활습관·휴식·업무경계 영역의 공동 개선 캠페인을 권장합니다.");
  }

  return {
    summary: `총 ${results.length}건의 결과 기준, 조직 평균은 ${avgTotal}점이며 주의/위험군 비율은 ${riskRate}%입니다.`,
    bullets,
  };
}

function getPersonalAiInsight(result: SavedResult | null) {
  if (!result) return [];
  const lines: string[] = [];
  const patterns = getPatternInsights(result);

  if (result.weakCats.length === 0) {
    lines.push("현재 취약 영역은 발견되지 않았습니다. 지금의 생활 리듬과 회복 루틴을 유지하는 것이 핵심입니다.");
  } else if (result.weakCats.length === 1) {
    lines.push(`취약 영역은 ${CAT_NAMES[result.weakCats[0]]} 1개입니다. 생활 습관 전체를 바꾸기보다 해당 영역의 작은 행동 1개부터 시작하는 것이 좋습니다.`);
  } else if (result.weakCats.length === 2) {
    lines.push(`취약 영역은 ${result.weakCats.map((c) => CAT_NAMES[c]).join(", ")}입니다. 두 영역이 서로 영향을 주고 있을 수 있으므로 2주 단위의 집중 개선이 필요합니다.`);
  } else {
    lines.push(`취약 영역이 ${result.weakCats.length}개입니다. 개인 노력만으로 해결하기보다 EAP 상담, 업무 조정, 회복 프로그램을 함께 검토하는 것이 좋습니다.`);
  }

  const lowest = CAT_KEYS.map((cat) => ({ cat, score: result.scores[cat] })).sort((a, b) => a.score - b.score)[0];
  lines.push(`가장 우선적으로 볼 영역은 ${CAT_NAMES[lowest.cat]}입니다. 권장 시작 행동: ${CAT_PROGRAMS[lowest.cat][0]}.`);

  if (patterns.length) {
    lines.push(`주요 패턴: ${patterns.map((p) => p.title).join(" · ")}. 단일 영역보다 연결된 생활·업무 패턴을 함께 조정하는 것이 효과적입니다.`);
  }

  if (result.riskLevel === "위험") {
    lines.push("현재 결과는 위험군 기준에 해당합니다. 취약 영역 수, 총점, 회복 부족 또는 이전 대비 하락 가능성을 함께 고려해 전문가 상담 또는 관리자와의 업무 조정 논의를 권장합니다.");
  } else if (result.riskLevel === "주의") {
    lines.push("주의 신호가 있으므로 1개월 후 재검사를 통해 변화 추이를 확인하는 것이 좋습니다.");
  } else {
    lines.push("위험도는 낮습니다. 회복 자원을 유지하고 과로 시그널을 조기에 점검하세요.");
  }
  return lines;
}

// ─── UI Components ───────────────────────────────────────
function Pill({ children, level }: { children: React.ReactNode; level: Level }) {
  const s = getLevelStyle(level);
  return (
    <span style={{ fontSize: 11, padding: "3px 10px", borderRadius: 999, background: s.bg, color: s.text, whiteSpace: "nowrap" }}>
      {children}
    </span>
  );
}

function SectionCard({ children, title, right }: { children: React.ReactNode; title?: string; right?: React.ReactNode }) {
  return (
    <section style={{ background: "#fff", border: "0.5px solid #e0e0e0", borderRadius: 16, padding: 18, boxShadow: "0 8px 24px rgba(0,0,0,0.03)" }}>
      {(title || right) && (
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", marginBottom: 14 }}>
          {title && <h3 style={{ fontSize: 15, fontWeight: 650, margin: 0 }}>{title}</h3>}
          {right}
        </div>
      )}
      {children}
    </section>
  );
}

function StatCard({ label, value, hint }: { label: string; value: React.ReactNode; hint?: string }) {
  return (
    <div style={{ background: "#f8f8f8", borderRadius: 14, padding: "16px 18px" }}>
      <p style={{ fontSize: 12, color: "#888", margin: "0 0 8px" }}>{label}</p>
      <div style={{ fontSize: 26, fontWeight: 700, lineHeight: 1 }}>{value}</div>
      {hint && <p style={{ fontSize: 11, color: "#999", margin: "8px 0 0", lineHeight: 1.5 }}>{hint}</p>}
    </div>
  );
}

function ProgressBar({ answered }: { answered: number }) {
  const pct = Math.round((answered / QUESTIONS.length) * 100);
  return (
    <div style={{ marginBottom: 28 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
        <span style={{ fontSize: 12, color: "#888" }}>진행률</span>
        <span style={{ fontSize: 12, fontWeight: 600 }}>{answered} / {QUESTIONS.length} · {pct}%</span>
      </div>
      <div style={{ height: 6, background: "#f0f0f0", borderRadius: 99, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${pct}%`, background: "#1D9E75", borderRadius: 99, transition: "width 0.4s ease" }} />
      </div>
    </div>
  );
}

function QuestionCard({ question, answer, onAnswer }: { question: Question; answer: AnswerValue | undefined; onAnswer: (id: number, val: AnswerValue) => void }) {
  const isDone = answer !== undefined;
  const options: { label: string; value: AnswerValue; activeStyle: React.CSSProperties }[] = [
    { label: "언제나 그렇다", value: 3, activeStyle: { background: "#E1F5EE", borderColor: "#5DCAA5", color: "#085041" } },
    { label: "가끔 그렇다", value: 1, activeStyle: { background: "#FAEEDA", borderColor: "#EF9F27", color: "#633806" } },
    { label: "전혀 아니다", value: 0, activeStyle: { background: "#FCEBEB", borderColor: "#F09595", color: "#791F1F" } },
  ];

  return (
    <div style={{ background: "#fff", border: `0.5px solid ${isDone ? "#5DCAA5" : "#e0e0e0"}`, borderRadius: 14, padding: "14px 16px", marginBottom: 8, transition: "border-color 0.2s" }}>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 12 }}>
        <div style={{ minWidth: 30, height: 30, borderRadius: 10, background: isDone ? "#E1F5EE" : "#f5f5f5", color: isDone ? "#085041" : "#888", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, flexShrink: 0 }}>
          {question.id}
        </div>
        <div>
          <p style={{ fontSize: 14, lineHeight: 1.55, margin: 0, paddingTop: 2 }}>{question.text}</p>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 7 }}>
        {options.map((opt) => (
          <button
            key={opt.value}
            onClick={(e) => {
              e.currentTarget.blur();
              onAnswer(question.id, opt.value);
            }}
            style={{ fontSize: 12, padding: "10px 12px", borderRadius: 999, border: "0.5px solid #ccc", background: "transparent", color: "#666", cursor: "pointer", transition: "all 0.15s", touchAction: "manipulation", ...(answer === opt.value ? opt.activeStyle : {}) }}
          >
            {opt.label} <b>({opt.value})</b>
          </button>
        ))}
      </div>
    </div>
  );
}

function CategoryCard({ cat, score }: { cat: CatKey; score: number }) {
  const lv = getLevel(score);
  const lc = getLevelStyle(lv);
  const pct = Math.min(100, Math.round((score / 18) * 100));
  return (
    <div style={{ background: "#fff", border: "0.5px solid #e0e0e0", borderRadius: 14, padding: "14px 16px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, gap: 8 }}>
        <span style={{ fontSize: 13, fontWeight: 650 }}>{cat}. {CAT_NAMES[cat]}</span>
        <Pill level={lv}>{lv}</Pill>
      </div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 4, marginBottom: 8 }}>
        <span style={{ fontSize: 24, fontWeight: 750 }}>{score}</span>
        <span style={{ fontSize: 12, color: "#888" }}>/ 18점</span>
      </div>
      <div style={{ height: 5, background: "#f0f0f0", borderRadius: 99, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${pct}%`, background: lc.border, borderRadius: 99 }} />
      </div>
      {lv !== "건강" && <p style={{ fontSize: 11, color: "#777", marginTop: 9, lineHeight: 1.5 }}>{CAT_ADVICE[cat]}</p>}
    </div>
  );
}

function AlertBox({ result }: { result: SavedResult | null }) {
  if (!result) return null;
  const weakNames = result.weakCats.map((c) => CAT_NAMES[c]);
  if (result.riskLevel === "위험") {
    return (
      <div style={{ background: "#FCEBEB", borderRadius: 14, padding: "16px 18px", display: "flex", gap: 12, alignItems: "flex-start", marginBottom: 20 }}>
        <span style={{ fontSize: 20, color: "#791F1F", flexShrink: 0 }}>⚠</span>
        <div>
          <p style={{ fontWeight: 700, color: "#791F1F", margin: "0 0 4px" }}>위험군 경고</p>
          <p style={{ fontSize: 13, color: "#791F1F", lineHeight: 1.55, margin: 0 }}>취약 영역 {result.weakCats.length}개 발견: {weakNames.join(", ") || "총점 취약"}. 즉각적인 EAP 전문가 상담 또는 관리자 면담을 권장드립니다.</p>
        </div>
      </div>
    );
  }
  if (result.riskLevel === "주의") {
    return (
      <div style={{ background: "#FAEEDA", borderRadius: 14, padding: "16px 18px", display: "flex", gap: 12, alignItems: "flex-start", marginBottom: 20 }}>
        <span style={{ fontSize: 20, color: "#633806", flexShrink: 0 }}>!</span>
        <div>
          <p style={{ fontWeight: 700, color: "#633806", margin: "0 0 4px" }}>주의 필요</p>
          <p style={{ fontSize: 13, color: "#633806", lineHeight: 1.55, margin: 0 }}>취약 영역: {weakNames.join(", ") || "총점 보통"}. 해당 영역의 개선 프로그램 참여와 1개월 후 재검사를 추천합니다.</p>
        </div>
      </div>
    );
  }
  return (
    <div style={{ background: "#E1F5EE", borderRadius: 14, padding: "16px 18px", display: "flex", gap: 12, alignItems: "flex-start", marginBottom: 20 }}>
      <span style={{ fontSize: 20, color: "#085041", flexShrink: 0 }}>✓</span>
      <div>
        <p style={{ fontWeight: 700, color: "#085041", margin: "0 0 4px" }}>건강한 상태</p>
        <p style={{ fontSize: 13, color: "#085041", lineHeight: 1.55, margin: 0 }}>모든 영역이 양호합니다. 현재의 생활 습관과 회복 루틴을 꾸준히 유지하세요.</p>
      </div>
    </div>
  );
}

function SimpleBarChart({ data, max = 18 }: { data: { label: string; value: number }[]; max?: number }) {
  return (
    <div style={{ display: "grid", gap: 9 }}>
      {data.map((d) => {
        const pct = Math.min(100, Math.round((d.value / max) * 100));
        return (
          <div key={d.label} style={{ display: "grid", gridTemplateColumns: "96px 1fr 44px", gap: 10, alignItems: "center" }}>
            <span style={{ fontSize: 12, color: "#555" }}>{d.label}</span>
            <div style={{ height: 9, borderRadius: 99, background: "#f0f0f0", overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${pct}%`, borderRadius: 99, background: "#1D9E75" }} />
            </div>
            <span style={{ fontSize: 12, textAlign: "right", fontWeight: 650 }}>{d.value}</span>
          </div>
        );
      })}
    </div>
  );
}

function TrendChart({ trend }: { trend: ReturnType<typeof buildTrend> }) {
  if (!trend.length) return <p style={{ fontSize: 13, color: "#999", margin: 0 }}>기간별 데이터가 없습니다.</p>;
  const max = Math.max(144, ...trend.map((t) => t.avgTotal));
  return (
    <div style={{ display: "flex", alignItems: "end", gap: 8, minHeight: 180, paddingTop: 12, overflowX: "auto" }}>
      {trend.slice(-18).map((t) => {
        const h = Math.max(16, Math.round((t.avgTotal / max) * 150));
        return (
          <div key={t.date} style={{ width: 44, flex: "0 0 44px", textAlign: "center" }}>
            <div title={`${t.date} · 평균 ${t.avgTotal}점`} style={{ height: h, background: "#E1F5EE", border: "0.5px solid #5DCAA5", borderRadius: "8px 8px 2px 2px" }} />
            <p style={{ fontSize: 10, color: "#999", margin: "6px 0 0" }}>{t.date.slice(5)}</p>
            <p style={{ fontSize: 10, color: "#555", margin: "2px 0 0", fontWeight: 650 }}>{t.avgTotal}</p>
          </div>
        );
      })}
    </div>
  );
}

function NavButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: React.ReactNode }) {
  return (
    <button onClick={onClick} style={{ border: "0.5px solid #ddd", background: active ? "#1D9E75" : "#fff", color: active ? "#fff" : "#555", borderRadius: 999, padding: "9px 14px", fontSize: 13, fontWeight: 650, cursor: "pointer" }}>
      {children}
    </button>
  );
}

// ─── Main App ─────────────────────────────────────────────
export default function App() {
  const [mode, setMode] = useState<AppMode>("test");
  const [name, setName] = useState("");
  const [dept, setDept] = useState("");
  const [consent, setConsent] = useState(true);
  const [answers, setAnswers] = useState<Record<number, AnswerValue>>({});
  const [results, setResults] = useState<SavedResult[]>([]);
  const [selectedId, setSelectedId] = useState<string>("");
  const [period, setPeriod] = useState<PeriodPreset>("all");
  const [deptFilter, setDeptFilter] = useState("전체");
  const [search, setSearch] = useState("");
  const [resetVersion, setResetVersion] = useState(0);

  useEffect(() => {
    const tag = document.createElement("style");
    tag.innerHTML = printStyle;
    document.head.appendChild(tag);
    setResults(loadResults());
    return () => { document.head.removeChild(tag); };
  }, []);

  useEffect(() => {
    saveResults(results);
  }, [results]);

  const answeredCount = Object.keys(answers).length;
  const catScores = useMemo(() => calcScores(answers), [answers]);
  const catLevels = useMemo(() => calcLevels(catScores), [catScores]);
  const totalScore = Object.values(catScores).reduce((a, b) => a + b, 0);
  const totalStatus = getTotalStatus(totalScore);
  const unansweredIds = useMemo(() => QUESTIONS.filter((q) => answers[q.id] === undefined).map((q) => q.id), [answers]);
  const firstUnansweredId = unansweredIds[0];
  const previewResult = useMemo(() => {
    if (!answeredCount) return null;
    return makeResult({ name, dept, answers, consent, previousResult: findPreviousResult(results, name, dept) });
  }, [name, dept, answers, consent, answeredCount, results]);

  const selectedResult = useMemo(() => results.find((r) => r.id === selectedId) || previewResult || results[0] || null, [results, selectedId, previewResult]);
  const selectedPreviousResult = useMemo(() => selectedResult ? findPreviousComparableResult(results, selectedResult) : null, [results, selectedResult]);

  const filteredResults = useMemo(() => {
    let rows = filterByPeriod(results, period);
    if (deptFilter !== "전체") rows = rows.filter((r) => r.dept === deptFilter);
    const q = search.trim().toLowerCase();
    if (q) rows = rows.filter((r) => `${r.name} ${r.dept} ${r.status} ${r.riskLevel}`.toLowerCase().includes(q));
    return rows;
  }, [results, period, deptFilter, search]);

  const departments = useMemo(() => ["전체", ...Array.from(new Set(results.map((r) => r.dept))).sort()], [results]);
  const deptStats = useMemo(() => buildDeptStats(filteredResults), [filteredResults]);
  const trend = useMemo(() => buildTrend(filteredResults), [filteredResults]);
  const weakRank = useMemo(() => getWeakCategoryRanking(filteredResults), [filteredResults]);
  const orgInsight = useMemo(() => getOrgAiInsight(filteredResults, deptStats), [filteredResults, deptStats]);

  const handleAnswer = (id: number, val: AnswerValue) => {
    setAnswers((prev) => ({ ...prev, [id]: val }));
    const nextId = id + 1;
    setTimeout(() => document.getElementById(`q-${nextId}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 80);
  };

  const btnBase: React.CSSProperties = { fontSize: 14, padding: "10px 18px", borderRadius: 10, cursor: "pointer", border: "none", fontWeight: 700, transition: "opacity 0.15s", touchAction: "manipulation" };

  const saveCurrent = () => {
    if (answeredCount < QUESTIONS.length) {
      alert(`미응답 문항이 ${QUESTIONS.length - answeredCount}개 있습니다. 첫 미응답 문항으로 이동합니다.`);
      if (firstUnansweredId) {
        setTimeout(() => document.getElementById(`q-${firstUnansweredId}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 50);
      }
      return;
    }
    const result = makeResult({ name, dept, answers, consent, previousResult: findPreviousResult(results, name, dept) });
    setResults((prev) => [result, ...prev]);
    setSelectedId(result.id);
    setMode("report");
    alert("검사 결과가 저장되었습니다.");
  };

  const resetAll = () => {
    if (!confirm("현재 입력한 답변을 초기화할까요?")) return;
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    document.querySelectorAll("button").forEach((button) => button.blur());
    setAnswers({});
    setName("");
    setDept("");
    setConsent(true);
    setResetVersion((v) => v + 1);
  };

  const deleteResult = (id: string) => {
    if (!confirm("선택한 결과를 삭제할까요?")) return;
    setResults((prev) => prev.filter((r) => r.id !== id));
    if (selectedId === id) setSelectedId("");
  };

  const exportExcel = () => {
    const rows = filteredResults.map((r) => {
      const row: Record<string, string | number> = {
        ID: r.id,
        이름: r.consent ? r.name : "익명",
        부서: r.dept,
        날짜: r.date,
        생성시각: r.createdAt,
        총점: r.total,
        총점상태: r.status,
        위험도: r.riskLevel,
        취약영역수: r.weakCats.length,
        취약영역: r.weakCats.map((c) => CAT_NAMES[c]).join(", "),
      };
      CAT_KEYS.forEach((c) => {
        row[`${c}.${CAT_NAMES[c]} 점수`] = r.scores[c];
        row[`${c}.${CAT_NAMES[c]} 판정`] = r.levels[c];
      });
      return row;
    });

    const deptRows = deptStats.map((d) => {
      const row: Record<string, string | number> = {
        부서: d.dept,
        인원수: d.count,
        평균총점: d.avgTotal,
        주의위험비율: `${d.riskRate}%`,
        평균취약영역수: d.weakAvgCount,
      };
      CAT_KEYS.forEach((c) => { row[`${c}.${CAT_NAMES[c]} 평균`] = d.catAvg[c]; });
      return row;
    });

    const trendRows = trend.map((t) => ({ 날짜: t.date, 건수: t.count, 평균총점: t.avgTotal, 주의위험비율: `${t.riskRate}%` }));
    const weakRows = weakRank.map((w) => ({ 영역: `${w.cat}.${CAT_NAMES[w.cat]}`, 취약건수: w.count }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "개인결과");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deptRows), "부서분석");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trendRows), "기간추이");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(weakRows), "취약영역랭킹");
    XLSX.writeFile(wb, `스트레스취약성_조직리포트_${normalizeDateKey(new Date().toISOString())}.xlsx`);
  };

  const exportSelectedExcel = () => {
    if (!selectedResult) return;
    const r = selectedResult;
    const row: Record<string, string | number> = {
      이름: r.consent ? r.name : "익명",
      부서: r.dept,
      날짜: r.date,
      총점: r.total,
      총점상태: r.status,
      위험도: r.riskLevel,
      취약영역수: r.weakCats.length,
      취약영역: r.weakCats.map((c) => CAT_NAMES[c]).join(", "),
    };
    CAT_KEYS.forEach((c) => {
      row[`${c}.${CAT_NAMES[c]} 점수`] = r.scores[c];
      row[`${c}.${CAT_NAMES[c]} 판정`] = r.levels[c];
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([row]), "개인결과");
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(CAT_KEYS.map((c) => ({ 영역: `${c}.${CAT_NAMES[c]}`, 점수: r.scores[c], 판정: r.levels[c], 개선방향: CAT_ADVICE[c] }))),
      "영역해석"
    );
    XLSX.writeFile(wb, `스트레스검사_${r.name || "결과"}_${normalizeDateKey(r.createdAt)}.xlsx`);
  };

  const importExcel = async (file: File | null) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json<Record<string, any>>(sheet);
    const imported: SavedResult[] = json.map((row) => {
      const scores = emptyScores();
      CAT_KEYS.forEach((c) => { scores[c] = Number(row[`${c}.${CAT_NAMES[c]} 점수`] ?? row[`${c}.${CAT_NAMES[c]}`] ?? 0); });
      const levels = calcLevels(scores);
      const weakCats = CAT_KEYS.filter((c) => levels[c] === "취약");
      const total = Number(row["총점"] ?? Object.values(scores).reduce((a, b) => a + b, 0));
      return {
        id: String(row.ID || uid()),
        name: String(row["이름"] || "이름 미입력"),
        dept: String(row["부서"] || "부서 미입력"),
        total,
        status: String(row["총점상태"] || getTotalStatus(total).label),
        scores,
        levels,
        answers: {},
        weakCats,
        riskLevel: getRiskLevel(weakCats.length, total, scores),
        date: String(row["날짜"] || formatDate()),
        createdAt: String(row["생성시각"] || new Date().toISOString()),
        consent: true,
      };
    });
    setResults((prev) => [...imported, ...prev]);
    alert(`${imported.length}건을 가져왔습니다.`);
  };

  const seedDemoData = () => {
    const depts = ["인사팀", "영업팀", "개발팀", "고객지원팀", "마케팅팀"];
    const demo = Array.from({ length: 36 }, (_, i) => {
      const answersObj: Record<number, AnswerValue> = {};
      QUESTIONS.forEach((q) => {
        const bias = q.cat === "G" && i % 3 === 0 ? 0.35 : q.cat === "F" && i % 4 === 0 ? 0.45 : 0.72;
        const r = Math.random();
        answersObj[q.id] = r < bias ? 3 : r < bias + 0.2 ? 1 : 0;
      });
      const result = makeResult({ name: `데모${i + 1}`, dept: depts[i % depts.length], answers: answersObj, consent: false });
      const d = new Date();
      d.setDate(d.getDate() - Math.floor(Math.random() * 150));
      result.createdAt = d.toISOString();
      result.date = formatDate(d);
      return result;
    });
    setResults((prev) => [...demo, ...prev]);
  };

  const clearAllStored = () => {
    if (!confirm("저장된 모든 결과를 삭제할까요? 이 작업은 되돌릴 수 없습니다.")) return;
    setResults([]);
    setSelectedId("");
  };

  return (
    <div style={{ minHeight: "100vh", background: "#fbfbfa", fontFamily: "-apple-system, BlinkMacSystemFont, 'Pretendard', 'Noto Sans KR', sans-serif", color: "#1a1a1a" }}>
      <div style={{ maxWidth: 1120, margin: "0 auto", padding: "32px 20px 56px" }}>
        <header style={{ display: "flex", justifyContent: "space-between", gap: 18, alignItems: "flex-start", marginBottom: 24, paddingBottom: 20, borderBottom: "0.5px solid #e8e8e8" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 8 }}>
              <div style={{ width: 44, height: 44, borderRadius: 14, background: "#E1F5EE", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>🧠</div>
              <div>
                <h1 style={{ fontSize: 22, fontWeight: 800, margin: 0 }}>스트레스 취약성 검사 시스템</h1>
                <p style={{ fontSize: 13, color: "#888", margin: "5px 0 0", lineHeight: 1.5 }}>48문항 · 8개 영역 · 개인 리포트 · 조직 분석 · AI 개선 제안</p>
              </div>
            </div>
          </div>
          <nav className="no-print" style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
            <NavButton active={mode === "test"} onClick={() => setMode("test")}>검사</NavButton>
            <NavButton active={mode === "report"} onClick={() => setMode("report")}>개인 리포트</NavButton>
            <NavButton active={mode === "dashboard"} onClick={() => setMode("dashboard")}>관리자 대시보드</NavButton>
            <NavButton active={mode === "settings"} onClick={() => setMode("settings")}>설정/클라우드</NavButton>
          </nav>
        </header>

        {mode === "test" && (
          <main style={{ maxWidth: 760, margin: "0 auto" }}>
            <SectionCard title="응답자 정보" right={<span style={{ fontSize: 12, color: "#999" }}>태블릿 입력 최적화</span>}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 12 }}>
                <label style={{ display: "grid", gap: 6 }}>
                  <span style={{ fontSize: 12, color: "#888", fontWeight: 650 }}>이름</span>
                  <input value={name} onChange={(e) => setName(e.target.value)} placeholder="홍길동" style={{ padding: "11px 13px", fontSize: 15, border: "0.5px solid #ddd", borderRadius: 10, outline: "none", background: "#fff" }} />
                </label>
                <label style={{ display: "grid", gap: 6 }}>
                  <span style={{ fontSize: 12, color: "#888", fontWeight: 650 }}>부서</span>
                  <input value={dept} onChange={(e) => setDept(e.target.value)} placeholder="인사팀" style={{ padding: "11px 13px", fontSize: 15, border: "0.5px solid #ddd", borderRadius: 10, outline: "none", background: "#fff" }} />
                </label>
              </div>
              <label style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 12, color: "#777", lineHeight: 1.5 }}>
                <input type="checkbox" checked={consent} onChange={(e) => setConsent(e.target.checked)} style={{ marginTop: 2 }} />
                개인정보를 결과 저장 및 조직 통계에 활용하는 것에 동의합니다. 미동의 시 리포트에는 익명 처리할 수 있습니다.
              </label>
            </SectionCard>

            <div style={{ marginTop: 24 }}>
              <ProgressBar answered={answeredCount} />
              {unansweredIds.length > 0 && answeredCount > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", background: "#FAEEDA", color: "#633806", borderRadius: 14, padding: "12px 14px", marginBottom: 14 }}>
                  <span style={{ fontSize: 13, fontWeight: 700 }}>미응답 문항 {unansweredIds.length}개가 남았습니다.</span>
                  <button onClick={() => document.getElementById(`q-${firstUnansweredId}`)?.scrollIntoView({ behavior: "smooth", block: "center" })} style={{ border: "0.5px solid #EF9F27", background: "#fff", color: "#633806", borderRadius: 999, padding: "7px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>첫 미응답으로 이동</button>
                </div>
              )}
              <div key={resetVersion}>
                {QUESTIONS.map((q) => (
                  <div id={`q-${q.id}`} key={q.id}>
                    <QuestionCard question={q} answer={answers[q.id]} onAnswer={handleAnswer} />
                  </div>
                ))}
              </div>
            </div>

            <div className="no-print" style={{ position: "sticky", bottom: 16, display: "flex", gap: 10, marginTop: 28, padding: 12, background: "rgba(255,255,255,0.92)", border: "0.5px solid #e0e0e0", borderRadius: 16, backdropFilter: "blur(8px)", boxShadow: "0 10px 30px rgba(0,0,0,0.08)" }}>
              <button onClick={saveCurrent} style={{ ...btnBase, background: "#1D9E75", color: "#fff", flex: 1 }}>결과 저장 및 리포트 보기</button>
              <button onClick={() => setMode("report")} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>미리보기</button>
              <button onClick={resetAll} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>초기화</button>
            </div>
          </main>
        )}

        {mode === "report" && (
          <main id="print-area" style={{ maxWidth: 980, margin: "0 auto" }}>
            <div className="no-print" style={{ display: "flex", gap: 10, justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap" }}>
              <select value={selectedId} onChange={(e) => setSelectedId(e.target.value)} style={{ padding: "10px 12px", borderRadius: 10, border: "0.5px solid #ddd", minWidth: 240 }}>
                <option value="">현재 입력값 또는 최신 결과</option>
                {results.map((r) => <option key={r.id} value={r.id}>{r.date} · {r.name} · {r.dept} · {r.total}점</option>)}
              </select>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => window.print()} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>🖨️ PDF 저장/인쇄</button>
                <button onClick={exportSelectedExcel} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>📊 개인 Excel</button>
              </div>
            </div>

            {!selectedResult ? (
              <SectionCard><p style={{ color: "#999", fontSize: 14, margin: 0 }}>아직 표시할 결과가 없습니다.</p></SectionCard>
            ) : (
              <div style={{ display: "grid", gap: 18 }}>
                <SectionCard>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 20, flexWrap: "wrap" }}>
                    <div>
                      <p style={{ fontSize: 11, color: "#999", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", margin: "0 0 5px" }}>개인 검사 결과</p>
                      <h2 style={{ fontSize: 22, fontWeight: 800, margin: 0 }}>스트레스 취약성 프로파일</h2>
                      <p style={{ fontSize: 13, color: "#888", margin: "8px 0 0" }}>{selectedResult.consent ? selectedResult.name : "익명"} · {selectedResult.dept} · {selectedResult.date}</p>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <p style={{ fontSize: 12, color: "#888", margin: "0 0 4px" }}>총점</p>
                      <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                        <span style={{ fontSize: 46, fontWeight: 850 }}>{selectedResult.total}</span>
                        <span style={{ fontSize: 14, color: "#888" }}>/ 144점</span>
                        <span style={{ fontSize: 12, padding: "4px 10px", borderRadius: 999, background: getTotalStatus(selectedResult.total).bg, color: getTotalStatus(selectedResult.total).color, fontWeight: 700 }}>{selectedResult.status}</span>
                      </div>
                    </div>
                  </div>
                </SectionCard>

                <AlertBox result={selectedResult} />

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 10 }}>
                  {CAT_KEYS.map((cat) => <CategoryCard key={cat} cat={cat} score={selectedResult.scores[cat]} />)}
                </div>

                <SectionCard title="AI 자동 해석">
                  <div style={{ display: "grid", gap: 10 }}>
                    {getPersonalAiInsight(selectedResult).map((line, idx) => (
                      <div key={idx} style={{ display: "flex", gap: 10, alignItems: "flex-start", background: "#f8f8f8", borderRadius: 12, padding: "12px 14px" }}>
                        <span style={{ color: "#1D9E75", fontWeight: 800 }}>AI</span>
                        <p style={{ fontSize: 13, lineHeight: 1.6, color: "#555", margin: 0 }}>{line}</p>
                      </div>
                    ))}
                  </div>
                </SectionCard>

                <SectionCard title="재검사 비교">
                  {selectedPreviousResult ? (() => {
                    const diff = selectedResult.total - selectedPreviousResult.total;
                    const catDiffs = CAT_KEYS.map((cat) => ({ cat, diff: selectedResult.scores[cat] - selectedPreviousResult.scores[cat] }))
                      .sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
                    return (
                      <div style={{ display: "grid", gap: 12 }}>
                        <div style={{ background: "#f8f8f8", borderRadius: 12, padding: "13px 14px", fontSize: 13, lineHeight: 1.6, color: "#555" }}>
                          이전 검사({selectedPreviousResult.date}) 대비 총점이 <b>{Math.abs(diff)}점 {diff >= 0 ? "상승" : "하락"}</b>했습니다.
                          {diff <= -15 ? " 하락 폭이 커서 위험 신호로 함께 반영됩니다." : diff <= -8 ? " 주의 깊은 추적이 필요합니다." : diff >= 8 ? " 긍정적인 개선 추세입니다." : " 큰 변화는 없습니다."}
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 8 }}>
                          {catDiffs.slice(0, 4).map((x) => (
                            <div key={x.cat} style={{ border: "0.5px solid #e0e0e0", borderRadius: 12, padding: "11px 12px" }}>
                              <p style={{ margin: "0 0 5px", fontSize: 12, color: "#888" }}>{x.cat}. {CAT_NAMES[x.cat]}</p>
                              <p style={{ margin: 0, fontSize: 18, fontWeight: 800 }}>{x.diff > 0 ? "+" : ""}{x.diff}점</p>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })() : (
                    <p style={{ fontSize: 13, color: "#999", margin: 0 }}>동일 이름·부서의 이전 검사 결과가 있으면 변화 추이를 자동 비교합니다.</p>
                  )}
                </SectionCard>

                {getPatternInsights(selectedResult).length > 0 && (
                  <SectionCard title="스트레스 원인 패턴 분석">
                    <div style={{ display: "grid", gap: 10 }}>
                      {getPatternInsights(selectedResult).map((p) => (
                        <div key={p.title} style={{ background: "#f8f8f8", borderRadius: 12, padding: "13px 14px" }}>
                          <p style={{ fontSize: 13, fontWeight: 800, margin: "0 0 5px", color: "#333" }}>{p.title}</p>
                          <p style={{ fontSize: 12, lineHeight: 1.6, color: "#666", margin: "0 0 5px" }}>{p.body}</p>
                          <p style={{ fontSize: 12, lineHeight: 1.6, color: "#1D9E75", margin: 0 }}>권장: {p.recommendation}</p>
                        </div>
                      ))}
                    </div>
                  </SectionCard>
                )}

                <SectionCard title="2주 실천 계획">
                  <div style={{ display: "grid", gap: 8 }}>
                    {getTwoWeekActionPlan(selectedResult).map((item) => (
                      <div key={item.label} style={{ display: "grid", gridTemplateColumns: "70px 1fr", gap: 10, background: "#f8f8f8", borderRadius: 12, padding: "12px 14px" }}>
                        <b style={{ fontSize: 13, color: "#1D9E75" }}>{item.label}</b>
                        <p style={{ fontSize: 13, lineHeight: 1.6, color: "#555", margin: 0 }}>{item.body}</p>
                      </div>
                    ))}
                  </div>
                </SectionCard>

                <SectionCard title="점수 기준">
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                    <span style={{ fontSize: 13 }}><Pill level="건강">건강</Pill> 14~18점</span>
                    <span style={{ fontSize: 13 }}><Pill level="보통">보통</Pill> 10~13점</span>
                    <span style={{ fontSize: 13 }}><Pill level="취약">취약</Pill> 0~9점</span>
                  </div>
                </SectionCard>
              </div>
            )}
          </main>
        )}

        {mode === "dashboard" && (
          <main id="admin-print-area" style={{ display: "grid", gap: 18 }}>
            <div className="no-print" style={{ display: "flex", gap: 10, justifyContent: "space-between", flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <select value={period} onChange={(e) => setPeriod(e.target.value as PeriodPreset)} style={{ padding: "10px 12px", borderRadius: 10, border: "0.5px solid #ddd" }}>
                  <option value="all">전체 기간</option>
                  <option value="30d">최근 30일</option>
                  <option value="90d">최근 90일</option>
                  <option value="180d">최근 180일</option>
                  <option value="365d">최근 1년</option>
                </select>
                <select value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)} style={{ padding: "10px 12px", borderRadius: 10, border: "0.5px solid #ddd" }}>
                  {departments.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
                <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="이름/부서/상태 검색" style={{ padding: "10px 12px", borderRadius: 10, border: "0.5px solid #ddd", minWidth: 180 }} />
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={exportExcel} style={{ ...btnBase, background: "#1D9E75", color: "#fff" }}>관리자 Excel</button>
                <button onClick={() => window.print()} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>관리자 PDF</button>
                <button onClick={seedDemoData} style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555" }}>데모 데이터</button>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12 }}>
              <StatCard label="분석 대상" value={`${filteredResults.length}건`} hint="현재 필터 기준" />
              <StatCard label="조직 평균" value={`${round1(average(filteredResults.map((r) => r.total)))}점`} hint="144점 만점" />
              <StatCard label="주의/위험군" value={`${filteredResults.length ? round1((filteredResults.filter((r) => r.riskLevel !== "낮음").length / filteredResults.length) * 100) : 0}%`} hint="취약 영역 또는 총점 기준" />
              <StatCard label="위험군" value={`${filteredResults.filter((r) => r.riskLevel === "위험").length}명`} hint="즉각 개입 권장" />
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.2fr) minmax(320px, 0.8fr)", gap: 18 }}>
              <SectionCard title="기간별 평균 총점 변화">
                <TrendChart trend={trend} />
              </SectionCard>
              <SectionCard title="취약 영역 랭킹">
                <SimpleBarChart data={weakRank.map((w) => ({ label: `${w.cat}. ${CAT_NAMES[w.cat]}`, value: w.count }))} max={Math.max(1, filteredResults.length)} />
              </SectionCard>
            </div>

            <SectionCard title="AI 조직개선 제안">
              <p style={{ fontSize: 14, lineHeight: 1.65, color: "#444", margin: "0 0 12px" }}>{orgInsight.summary}</p>
              <div style={{ display: "grid", gap: 8 }}>
                {orgInsight.bullets.map((b, idx) => (
                  <div key={idx} style={{ background: "#f8f8f8", borderRadius: 12, padding: "12px 14px", fontSize: 13, lineHeight: 1.6, color: "#555" }}>• {b}</div>
                ))}
              </div>
            </SectionCard>

            <SectionCard title="부서별 비교">
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#777", borderBottom: "0.5px solid #ddd" }}>
                      <th style={{ padding: 10 }}>부서</th>
                      <th style={{ padding: 10 }}>인원</th>
                      <th style={{ padding: 10 }}>평균총점</th>
                      <th style={{ padding: 10 }}>주의/위험</th>
                      <th style={{ padding: 10 }}>평균 취약수</th>
                      <th style={{ padding: 10 }}>가장 낮은 영역</th>
                    </tr>
                  </thead>
                  <tbody>
                    {deptStats.map((d) => {
                      const lowest = CAT_KEYS.map((cat) => ({ cat, score: d.catAvg[cat] })).sort((a, b) => a.score - b.score)[0];
                      const protectedDept = d.count < MIN_DEPT_SAMPLE;
                      return (
                        <tr key={d.dept} style={{ borderBottom: "0.5px solid #eee" }}>
                          <td style={{ padding: 10, fontWeight: 700 }}>{d.dept}</td>
                          <td style={{ padding: 10 }}>{d.count}</td>
                          <td style={{ padding: 10 }}>{protectedDept ? "표본 보호" : d.avgTotal}</td>
                          <td style={{ padding: 10 }}>{protectedDept ? `최소 ${MIN_DEPT_SAMPLE}명 필요` : `${d.riskRate}%`}</td>
                          <td style={{ padding: 10 }}>{protectedDept ? "표본 보호" : d.weakAvgCount}</td>
                          <td style={{ padding: 10 }}>{protectedDept ? "표본 보호" : `${lowest.cat}. ${CAT_NAMES[lowest.cat]} (${lowest.score})`}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </SectionCard>

            <SectionCard title="위험군 탐지 목록">
              <div style={{ display: "grid", gap: 8 }}>
                {filteredResults.filter((r) => r.riskLevel !== "낮음").slice(0, 20).map((r) => (
                  <div key={r.id} style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 10, alignItems: "center", background: "#f8f8f8", borderRadius: 12, padding: "10px 12px" }}>
                    <div>
                      <p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{displayNameForAdmin(r)} · {r.dept}</p>
                      <p style={{ margin: "4px 0 0", color: "#888", fontSize: 11 }}>{r.date} · 취약: {r.weakCats.map((c) => CAT_NAMES[c]).join(", ") || "없음"}</p>
                    </div>
                    <span style={{ fontSize: 13, fontWeight: 700 }}>{r.total}점</span>
                    <span style={{ fontSize: 12, padding: "3px 10px", borderRadius: 999, background: r.riskLevel === "위험" ? "#FCEBEB" : "#FAEEDA", color: r.riskLevel === "위험" ? "#791F1F" : "#633806", fontWeight: 700 }}>{r.riskLevel}</span>
                    <button onClick={() => { setSelectedId(r.id); setMode("report"); }} style={{ border: "0.5px solid #ccc", background: "#fff", borderRadius: 8, padding: "7px 10px", cursor: "pointer" }}>보기</button>
                  </div>
                ))}
                {!filteredResults.some((r) => r.riskLevel !== "낮음") && <p style={{ margin: 0, fontSize: 13, color: "#999" }}>주의/위험군이 없습니다.</p>}
              </div>
            </SectionCard>
          </main>
        )}

        {mode === "settings" && (
          <main style={{ display: "grid", gap: 18 }}>
            <SectionCard title="데이터 관리">
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <button onClick={exportExcel} style={{ ...btnBase, background: "#1D9E75", color: "#fff" }}>전체 Excel 다운로드</button>
                <label style={{ ...btnBase, background: "#fff", border: "0.5px solid #ccc", color: "#555", display: "inline-flex", alignItems: "center", gap: 8 }}>
                  Excel 가져오기
                  <input type="file" accept=".xlsx,.xls" onChange={(e) => importExcel(e.target.files?.[0] || null)} style={{ display: "none" }} />
                </label>
                <button onClick={clearAllStored} style={{ ...btnBase, background: "#fff", border: "0.5px solid #F09595", color: "#791F1F" }}>저장 데이터 전체 삭제</button>
              </div>
              <p style={{ fontSize: 12, color: "#999", lineHeight: 1.6, margin: "12px 0 0" }}>현재 버전은 브라우저 localStorage에 저장합니다. 운영 배포 시에는 Firebase Firestore 또는 사내 DB로 저장소를 교체하세요.</p>
            </SectionCard>

            <SectionCard title="Firebase 클라우드 연동 구조">
              <div style={{ background: "#f8f8f8", borderRadius: 14, padding: 16 }}>
                <p style={{ fontSize: 13, lineHeight: 1.7, color: "#555", margin: 0 }}>
                  이 앱의 저장 단위는 <b>SavedResult</b>입니다. Firestore 컬렉션 예시는 <code>stressResults</code>이며, 관리자 대시보드는 해당 컬렉션을 실시간 구독하면 여러 태블릿에서 동기화됩니다.
                </p>
                <pre style={{ overflowX: "auto", margin: "14px 0 0", fontSize: 12, lineHeight: 1.6, color: "#555" }}>{`// Firebase 연결 예시
// 1) npm i firebase
// 2) firebaseConfig 입력
// 3) saveCurrent()에서 setResults 대신 addDoc(collection(db, "stressResults"), result)
// 4) useEffect에서 onSnapshot(collection(db, "stressResults"), snapshot => setResults(...))

interface SavedResult {
  id: string;
  name: string;
  dept: string;
  total: number;
  scores: Record<CatKey, number>;
  levels: Record<CatKey, Level>;
  weakCats: CatKey[];
  riskLevel: "낮음" | "주의" | "위험";
  createdAt: string;
  consent: boolean;
}`}</pre>
              </div>
            </SectionCard>

            <SectionCard title="운영 시 권장 보안 정책">
              <div style={{ display: "grid", gap: 8 }}>
                {["관리자 계정과 일반 응답자 계정을 분리하세요.", "이름 대신 사번/익명 ID 사용을 기본값으로 두세요.", "성생활 문항이 포함되어 있으므로 민감정보 수집 동의 문구를 별도 고지하세요.", "부서별 통계는 최소 인원 기준, 예: 5명 이상일 때만 표시하세요.", "개인 리포트와 관리자 리포트의 접근 권한을 분리하세요."].map((x) => <div key={x} style={{ background: "#f8f8f8", borderRadius: 12, padding: "11px 13px", fontSize: 13, color: "#555" }}>• {x}</div>)}
              </div>
            </SectionCard>

            <SectionCard title="저장된 결과 목록">
              <div style={{ display: "grid", gap: 8 }}>
                {results.slice(0, 30).map((r) => (
                  <div key={r.id} style={{ display: "grid", gridTemplateColumns: "1fr auto auto", gap: 10, alignItems: "center", background: "#f8f8f8", borderRadius: 12, padding: "10px 12px" }}>
                    <div>
                      <p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{r.name} · {r.dept}</p>
                      <p style={{ margin: "4px 0 0", color: "#888", fontSize: 11 }}>{r.date} · {r.total}점 · {r.riskLevel}</p>
                    </div>
                    <button onClick={() => { setSelectedId(r.id); setMode("report"); }} style={{ border: "0.5px solid #ccc", background: "#fff", borderRadius: 8, padding: "7px 10px", cursor: "pointer" }}>보기</button>
                    <button onClick={() => deleteResult(r.id)} style={{ border: "0.5px solid #F09595", color: "#791F1F", background: "#fff", borderRadius: 8, padding: "7px 10px", cursor: "pointer" }}>삭제</button>
                  </div>
                ))}
                {!results.length && <p style={{ margin: 0, fontSize: 13, color: "#999" }}>저장된 결과가 없습니다.</p>}
              </div>
            </SectionCard>
          </main>
        )}
      </div>
    </div>
  );
}
