// 로그인 설정 — 이 파일 한 곳만 바꾸면 된다(생성: data10-00/scripts/apply-auth-kit.mjs, 대문 문구 출처 docs/plan.json).
// 내 Supabase 로 옮길 때: url · anonKey 를 내 프로젝트 값으로, 카카오를 설정하지 않았으면 kakao: false. 안내 = docs/내_깃허브로_옮기기.md
// prefix 는 supabase/auth.sql 의 표 이름 앞부분과 같아야 한다.
window.AUTH_CONFIG = {
  "url": "https://hcmgdztsgjvzcyxyayaj.supabase.co",
  "anonKey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhjbWdkenRzZ2p2emN5eHlheWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MzU4ODcsImV4cCI6MjA4NzAxMTg4N30.gznaPzY1l8qDAPsEyYNR9KS7f7VqS3xaw-_2HTSwSZw",
  "prefix": "data1011",
  "kakao": true,
  "idleMinutes": 30,
  "localBypass": true,
  "site": {
    "title": "Deep drawing press 불량 최소화",
    "summary": "Deep drawing press 작업일보 엑셀을 여러 달 쌓아 일별 · 월별 · 금형 · 소재별 작업량 대비 불량률과 4M · 계절별 「불량이 몰린 조건」, 금형 보수 · 설비 수선 전후, 압력 작업표준 대비를 보여 주는 브라우저 도구입니다.",
    "course": "현장 데이터 디지털화 전문가과정 · 2차수",
    "features": [
      {
        "name": "작업일보 여러 달 가져오기",
        "desc": "엑셀 여러 개 · 폴더째 넣으면 머리행과 열(두 줄 머리 · 별칭)을 찾고 소재 칸을 값 모양으로 확인합니다."
      },
      {
        "name": "불량률 추이 · 항목별 표",
        "desc": "일별 · 월별, 금형 · 소재 메이커 · lot · 규격 · 모델 · 작업자별 작업량 대비 불량률을 표본과 함께 봅니다."
      },
      {
        "name": "4M · 계절 「불량이 몰린 조건」",
        "desc": "한 조건 · 두 조건 조합을 표본 · 배수 · 95% 하한으로 걸러 근거 숫자와 겹침 표시를 붙여 보여 줍니다."
      },
      {
        "name": "압력 기준 · 보수 · 수선",
        "desc": "모델별 압력 작업표준 대비 실제와 전파 목록, 금형 보수 · 설비 수선 전후 불량률을 비교합니다."
      }
    ],
    "steps": [
      "로그인 · 승인",
      "작업일보 가져오기",
      "열 · 소재 칸 확인",
      "불량률 집계",
      "4M · 계절 분석",
      "압력 · 설비 점검"
    ]
  }
};
