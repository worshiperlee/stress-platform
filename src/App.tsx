import { useMemo, useState, useEffect } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
} from "chart.js";
import { Bar } from "react-chartjs-2";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

type AnswerValue = 0 | 1 | 3;

type Question = {
  id: number;
  text: string;
  category: string;
};

const questions: Question[] = [
  { id: 1, text: "적절한 음식을 적당량 먹는다.", category: "식습관" },
  { id: 2, text: "점심시간 음주를 하지 않는다.", category: "금연/금주" },
  { id: 3, text: "일주일 3회 이상 운동한다.", category: "운동" },
  { id: 4, text: "친구나 동료와 정기적으로 교류한다.", category: "사회적관계" },
  { id: 5, text: "부부관계에 만족한다.", category: "부부관계" },
  { id: 6, text: "여가활동이나 취미가 있다.", category: "여가(취미)" },
  { id: 7, text: "일과 삶의 균형을 유지한다.", category: "가정과 일의 균형감" },
  { id: 8, text: "기도·명상·자기성찰 시간을 갖는다.", category: "자기이해수용" },

  { id: 9, text: "규칙적으로 식사한다.", category: "식습관" },
  { id: 10, text: "과도한 음주를 하지 않는다.", category: "금연/금주" },
  { id: 11, text: "걷기나 활동량이 충분하다.", category: "운동" },
  { id: 12, text: "감정을 솔직하게 표현한다.", category: "사회적관계" },
  { id: 13, text: "정서적 친밀감을 느낀다.", category: "부부관계" },
  { id: 14, text: "충분한 휴식을 취한다.", category: "여가(취미)" },
  { id: 15, text: "업무 스트레스를 조절한다.", category: "가정과 일의 균형감" },
  { id: 16, text: "나 자신을 긍정적으로 바라본다.", category: "자기이해수용" },
];

const categoryAdvice: Record<string, string> = {
  식습관: "규칙적인 식사와 충분한 수분 섭취가 필요합니다.",
  "금연/금주": "음주 빈도를 줄이고 건강한 스트레스 해소법을 찾아보세요.",
  운동: "주 3회 이상 가벼운 운동부터 시작해보세요.",
  사회적관계: "정서적 지지를 줄 수 있는 관계를 회복해보세요.",
  부부관계: "배우자와의 정서적 대화 시간을 늘려보세요.",
  "여가(취미)": "쉼과 취미를 일정 안에 포함해보세요.",
  "가정과 일의 균형감": "업무와 삶의 경계를 조정할 필요가 있습니다.",
  자기이해수용: "기도·명상·상담 등을 통해 자신을 돌보세요.",
};

function getLevel(score: number) {
  if (score <= 9) return "취약";
  if (score <= 13) return "보통";
  return "건강";
}

export default function App() {
  const [employeeName, setEmployeeName] = useState("");
  const [department, setDepartment] = useState("");
  const [answers, setAnswers] = useState<Record<number, AnswerValue>>({});
  const [savedResults, setSavedResults] = useState<any[]>([]);
  const [adminMode, setAdminMode] = useState(false);

  useEffect(() => {
    const data = JSON.parse(
      localStorage.getItem("stressResults") || "[]"
    );

    setSavedResults(data);
  }, []);

  const handleAnswer = (id: number, value: AnswerValue) => {
    setAnswers((prev) => ({
      ...prev,
      [id]: value,
    }));
  };

  const categoryResults = useMemo(() => {
    const result: Record<string, number> = {};

    questions.forEach((q) => {
      result[q.category] =
        (result[q.category] || 0) + (answers[q.id] || 0);
    });

    return result;
  }, [answers]);

  const totalScore = Object.values(categoryResults).reduce(
    (a, b) => a + b,
    0
  );

  const vulnerableCategories = Object.entries(categoryResults)
    .filter(([_, score]) => getLevel(score) === "취약")
    .map(([category]) => category);

  const saveResult = () => {
    const existing = JSON.parse(
      localStorage.getItem("stressResults") || "[]"
    );

    const newData = {
      name: employeeName,
      department,
      totalScore,
      results: categoryResults,
      createdAt: new Date().toISOString(),
    };

    existing.push(newData);

    localStorage.setItem(
      "stressResults",
      JSON.stringify(existing)
    );

    setSavedResults(existing);

    alert("결과 저장 완료");
  };

  const exportExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(savedResults);

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Stress Results"
    );

    XLSX.writeFile(workbook, "stress-results.xlsx");
  };

  const exportPDF = () => {
    const doc = new jsPDF();

    doc.setFontSize(20);
    doc.text("스트레스 취약성 검사 결과", 20, 20);

    doc.setFontSize(12);
    doc.text(`이름: ${employeeName}`, 20, 40);
    doc.text(`부서: ${department}`, 20, 50);
    doc.text(`총점: ${totalScore}`, 20, 60);

    let y = 80;

    Object.entries(categoryResults).forEach(
      ([category, score]) => {
        doc.text(
          `${category}: ${score}점 (${getLevel(score)})`,
          20,
          y
        );

        y += 10;
      }
    );

    doc.save("stress-report.pdf");
  };

  const averageData = useMemo(() => {
    const totals: Record<string, number> = {};
    const counts: Record<string, number> = {};

    savedResults.forEach((item) => {
      Object.entries(item.results).forEach(
        ([category, score]: any) => {
          totals[category] =
            (totals[category] || 0) + score;

          counts[category] =
            (counts[category] || 0) + 1;
        }
      );
    });

    const averages: Record<string, number> = {};

    Object.keys(totals).forEach((category) => {
      averages[category] =
        totals[category] / counts[category];
    });

    return averages;
  }, [savedResults]);

  return (
    <div
      style={{
        padding: 30,
        maxWidth: 1400,
        margin: "0 auto",
        fontFamily: "sans-serif",
      }}
    >
      <h1>스트레스 취약성 조직 분석 플랫폼</h1>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "1fr 1fr",
          gap: 20,
          marginBottom: 30,
        }}
      >
        <input
          placeholder="이름"
          value={employeeName}
          onChange={(e) =>
            setEmployeeName(e.target.value)
          }
          style={{
            padding: 15,
            fontSize: 16,
          }}
        />

        <input
          placeholder="부서"
          value={department}
          onChange={(e) =>
            setDepartment(e.target.value)
          }
          style={{
            padding: 15,
            fontSize: 16,
          }}
        />
      </div>

      <button
        onClick={() => setAdminMode(!adminMode)}
        style={{
          padding: 12,
          marginBottom: 30,
        }}
      >
        관리자 모드
      </button>

      {questions.map((q) => (
        <div
          key={q.id}
          style={{
            border: "1px solid #ddd",
            padding: 20,
            borderRadius: 15,
            marginBottom: 15,
          }}
        >
          <div
            style={{
              marginBottom: 15,
              fontWeight: "bold",
            }}
          >
            {q.id}. {q.text}
          </div>

          <div
            style={{
              display: "flex",
              gap: 10,
              flexWrap: "wrap",
            }}
          >
            <button
              onClick={() =>
                handleAnswer(q.id, 3)
              }
            >
              언제나 그렇다 (3)
            </button>

            <button
              onClick={() =>
                handleAnswer(q.id, 1)
              }
            >
              가끔 그렇다 (1)
            </button>

            <button
              onClick={() =>
                handleAnswer(q.id, 0)
              }
            >
              전혀 아니다 (0)
            </button>
          </div>
        </div>
      ))}

      <div
        style={{
          marginTop: 40,
          background: "#f5f5f5",
          padding: 30,
          borderRadius: 20,
        }}
      >
        <h2>검사 결과</h2>

        <h3>총점: {totalScore}</h3>

        {Object.entries(categoryResults).map(
          ([category, score]) => (
            <div
              key={category}
              style={{
                marginBottom: 20,
                paddingBottom: 10,
                borderBottom:
                  "1px solid #ddd",
              }}
            >
              <h3>{category}</h3>

              <p>점수: {score}</p>

              <p>
                상태:{" "}
                <strong>
                  {getLevel(score)}
                </strong>
              </p>

              <p>
                개선방향:{" "}
                {categoryAdvice[category]}
              </p>
            </div>
          )
        )}

        <div
          style={{
            marginTop: 30,
            background:
              vulnerableCategories.length >= 3
                ? "#ffe5e5"
                : "#e8ffe8",
            padding: 20,
            borderRadius: 15,
          }}
        >
          <h2>AI 조직 개선 제안</h2>

          {vulnerableCategories.length > 0 ? (
            <ul>
              {vulnerableCategories.map(
                (item) => (
                  <li key={item}>
                    {item} 영역 강화 프로그램
                    필요
                  </li>
                )
              )}
            </ul>
          ) : (
            <p>
              현재 조직 상태는 비교적
              안정적입니다.
            </p>
          )}
        </div>

        {vulnerableCategories.length >= 3 && (
          <div
            style={{
              marginTop: 20,
              background: "#ffdddd",
              padding: 20,
              borderRadius: 15,
            }}
          >
            <h2>⚠ 위험군 경고</h2>

            <p>
              다수 영역에서 취약 상태가
              발견되었습니다.
            </p>
          </div>
        )}

        <div
          style={{
            display: "flex",
            gap: 15,
            marginTop: 30,
            flexWrap: "wrap",
          }}
        >
          <button onClick={saveResult}>
            결과 저장
          </button>

          <button onClick={exportExcel}>
            Excel 다운로드
          </button>

          <button onClick={exportPDF}>
            PDF 다운로드
          </button>
        </div>
      </div>

      {adminMode && (
        <div
          style={{
            marginTop: 50,
            background: "white",
            padding: 30,
            borderRadius: 20,
            border: "1px solid #ddd",
          }}
        >
          <h2>관리자 대시보드</h2>

          <p>
            전체 응답 수: {savedResults.length}
          </p>

          <div
            style={{
              marginTop: 40,
            }}
          >
            <Bar
              data={{
                labels:
                  Object.keys(averageData),
                datasets: [
                  {
                    label:
                      "영역별 평균 점수",
                    data: Object.values(
                      averageData
                    ),
                  },
                ],
              }}
            />
          </div>
        </div>
      )}
    </div>
  );
}