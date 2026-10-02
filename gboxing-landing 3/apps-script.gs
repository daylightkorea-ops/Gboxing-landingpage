/**
 * G복싱 랜딩페이지 → Google Sheets 연동 (Apps Script)
 *
 * 설치 방법
 * 1. Google Sheets 새 파일 만들기 (예: "G복싱 랜딩 리드")
 * 2. 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 3. 기존 코드를 지우고 이 파일 내용을 전부 붙여넣기 > 저장
 * 4. [배포] > [새 배포] > 유형: 웹 앱
 *      - 다음 사용자로 실행: 나
 *      - 액세스 권한: 모든 사용자
 * 5. 배포 후 나오는 "웹 앱 URL"을 복사 → index.html 의 CONFIG.SHEET_WEBHOOK 에 붙여넣기
 *
 * 결과
 *  - "리드" 시트: 무료체험 신청자 (이름, 연락처, 지점, 목적, 시간대, 유입경로)
 *  - "이벤트" 시트: 네이버 예약 클릭 / 카톡 상담 클릭 / 전화 클릭 (지점, 유입경로)
 *  - 새 신청이 들어오면 NOTIFY_EMAIL 로 메일 알림 (비워두면 알림 없음)
 */

const NOTIFY_EMAIL = ""; // 예: "gboxing@gmail.com"

const LEAD_HEADERS = ["신청일시", "이름", "연락처", "지점", "관심클래스", "희망시간", "utm_source", "utm_medium", "utm_campaign", "utm_content", "세션ID", "상태(직원 입력)", "메모"];
const EVENT_HEADERS = ["일시", "이벤트", "지점", "위치", "신청 후 클릭", "utm_source", "utm_medium", "utm_campaign", "utm_content", "세션ID"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    if (d.type === "lead") {
      const sh = getSheet_(ss, "리드", LEAD_HEADERS);
      sh.appendRow([now, d.name, "'" + d.phone, d.branch, d.goal, d.time,
        d.utm_source || "", d.utm_medium || "", d.utm_campaign || "", d.utm_content || "", d.session_id || "", "신규", ""]);
      if (NOTIFY_EMAIL) {
        MailApp.sendEmail(NOTIFY_EMAIL, `[G복싱] 무료체험 신청 - ${d.branch} ${d.name}`,
          `이름: ${d.name}\n연락처: ${d.phone}\n지점: ${d.branch}\n관심 클래스: ${d.goal}\n희망시간: ${d.time}\n유입: ${d.utm_source || "-"} / ${d.utm_campaign || "-"}`);
      }
    } else {
      const sh = getSheet_(ss, "이벤트", EVENT_HEADERS);
      const label = { naver_booking_click: "네이버 예약 클릭", kakao_click: "카톡 상담 클릭", phone_click: "전화 클릭" }[d.type] || d.type;
      sh.appendRow([now, label, d.branch || "", d.location || "", d.from_lead || "",
        d.utm_source || "", d.utm_medium || "", d.utm_campaign || "", d.utm_content || "", d.session_id || ""]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#2B3482").setFontColor("#FFFFFF");
  }
  return sh;
}
