/**
 * 「위성신, 묻다」 2027 캘린더 신청 폼 — 자동 생성 스크립트
 *
 * 사용법
 *   1. script.google.com → 새 프로젝트 → 이 코드를 전부 붙여넣기
 *   2. 위쪽 함수 선택에서 setup → 실행 → 권한 허용
 *   3. 실행 로그에 나온 "신청 링크"를 게시물에 사용
 *
 * setup을 한 번 실행하면 아래가 모두 갖춰집니다.
 *   - 신청 폼 + 응답 시트(발송관리 탭 포함)
 *   - 신청이 들어올 때마다 작가에게 메일 알림 (순번·배정 형태 포함)
 *   - 가로형/세로형 각 10부가 소진되면 해당 선택지 자동 삭제
 *   - 20부 전부 소진되거나 마감일이 되면 폼 자동 마감
 */

const CONFIG = {
  title: '「위성신, 묻다」 2027 캘린더 신청',
  stock: { 가로형: 10, 세로형: 10 },
  deadline: '2026-10-18T23:59:00+09:00',
  shipDate: '10월 20일 전후',
  anyChoice: '둘 중 아무거나 좋습니다',
};

function setup() {
  const form = FormApp.create(CONFIG.title);

  form
    .setDescription(
      [
        '열두 달에 열두 개의 질문을 담았습니다.',
        '그림 앞에 멈춰 스스로에게 묻는 한 해가 되기를 바랍니다.',
        '',
        '· 가로형 10부 · 세로형 10부, 선착순으로 보내드립니다.',
        '· 신청 마감: 10월 18일(일) 또는 20부 소진 시',
        `· 발송: ${CONFIG.shipDate}`,
        '',
        '— 위성신 Sungshin Wee',
      ].join('\n')
    )
    .setConfirmationMessage(
      [
        '신청해 주셔서 감사합니다.',
        `캘린더는 ${CONFIG.shipDate} 발송할 예정입니다.`,
        '',
        '열두 개의 질문이 당신의 한 해에 조용히 머물기를.',
      ].join('\n')
    )
    .setProgressBar(false)
    .setShowLinkToRespondAgain(false)
    .setAllowResponseEdits(false);

  form.addTextItem().setTitle('성함').setRequired(true);

  form
    .addTextItem()
    .setTitle('연락처')
    .setHelpText('배송 중 연락이 필요할 때만 사용합니다. (예: 010-1234-5678)')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .setHelpText('휴대폰 번호 형식으로 적어주세요. (예: 010-1234-5678)')
        .requireTextMatchesPattern('^01[016789]-?[0-9]{3,4}-?[0-9]{4}$')
        .build()
    );

  const typeItem = form
    .addMultipleChoiceItem()
    .setTitle('원하시는 형태')
    .setHelpText('가로형 〈네가 모르는 크기의 꿈〉 · 세로형 〈찬란한 고요〉')
    .setChoiceValues(['가로형', '세로형', CONFIG.anyChoice])
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle('받으실 주소')
    .setHelpText('우편번호를 포함해 적어주세요.\n(예: 06234 서울시 강남구 테헤란로 00길 00, 101동 1001호)')
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle('남기고 싶은 말 (선택)')
    .setHelpText('열두 개의 질문 중 마음에 남는 것이 있다면 들려주세요.');

  form
    .addCheckboxItem()
    .setTitle('개인정보 수집·이용 동의')
    .setHelpText(
      [
        '· 수집 항목: 성함, 연락처, 주소',
        '· 이용 목적: 캘린더 발송',
        '· 보유 기간: 발송 완료 후 즉시 파기',
        '동의하지 않으실 수 있으나, 이 경우 캘린더를 보내드릴 수 없습니다.',
      ].join('\n')
    )
    .setChoiceValues(['동의합니다'])
    .setRequired(true);

  // 응답 시트 + 발송관리 탭
  const ss = SpreadsheetApp.create(`${CONFIG.title} — 응답`);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  const ship = ss.insertSheet('발송관리', 0);
  ship
    .getRange(1, 1, 1, 9)
    .setValues([['순번', '접수시각', '성함', '연락처', '신청', '배정', '주소', '남기는 말', '발송완료']])
    .setFontWeight('bold')
    .setBackground('#262c3e')
    .setFontColor('#f4f1ec');
  ship.setFrozenRows(1);
  ship.setColumnWidth(7, 360);
  ship.setColumnWidth(8, 280);
  ship.getRange('I2:I21').insertCheckboxes();

  // 기존 트리거 정리 후 새로 등록
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSubmit').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('closeForm').timeBased().at(new Date(CONFIG.deadline)).create();

  PropertiesService.getScriptProperties().setProperties({
    formId: form.getId(),
    typeItemId: String(typeItem.getId()),
    sheetId: ss.getId(),
    count_가로형: '0',
    count_세로형: '0',
    seq: '0',
  });

  Logger.log('신청 링크(게시물용): ' + form.shortenFormUrl(form.getPublishedUrl()));
  Logger.log('폼 편집: ' + form.getEditUrl());
  Logger.log('발송관리 시트: ' + ss.getUrl());
}

function onSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const answers = {};
    e.response.getItemResponses().forEach((r) => {
      answers[r.getItem().getTitle()] = r.getResponse();
    });

    const wanted = answers['원하시는 형태'];
    const left = (type) => CONFIG.stock[type] - Number(props.getProperty('count_' + type));

    // 원하는 형태가 남아 있으면 그대로, 아니면 남은 쪽으로, 둘 다 없으면 대기
    let assigned = '대기';
    const order = wanted === '세로형' ? ['세로형', '가로형'] : ['가로형', '세로형'];
    if (wanted === CONFIG.anyChoice) order.sort((a, b) => left(b) - left(a));
    for (const type of order) {
      if (left(type) > 0) {
        assigned = type;
        break;
      }
    }

    const seq = Number(props.getProperty('seq')) + 1;
    props.setProperty('seq', String(seq));
    if (assigned !== '대기') {
      props.setProperty('count_' + assigned, String(Number(props.getProperty('count_' + assigned)) + 1));
    }

    const row = [
      seq,
      e.response.getTimestamp(),
      answers['성함'],
      answers['연락처'],
      wanted,
      assigned,
      answers['받으실 주소'],
      answers['남기고 싶은 말 (선택)'] || '',
      false,
    ];
    SpreadsheetApp.openById(props.getProperty('sheetId')).getSheetByName('발송관리').appendRow(row);

    notifyOwner(seq, assigned, answers, left('가로형'), left('세로형'));
    refreshChoices(left('가로형'), left('세로형'));
  } finally {
    lock.releaseLock();
  }
}

function notifyOwner(seq, assigned, a, leftH, leftV) {
  const esc = (s) => String(s || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: `[캘린더 신청 ${seq}번] ${a['성함']} · ${assigned}`,
    htmlBody: `
      <div style="font-family:Georgia,serif;color:#262c3e;max-width:520px">
        <p style="letter-spacing:2px;font-size:12px;color:#787680">SUNGSHIN WEE · 2027 CALENDAR</p>
        <h2 style="font-weight:normal;margin:4px 0 16px">${seq}번째 신청 — ${esc(assigned)}</h2>
        <table style="border-collapse:collapse;font-size:14px;line-height:1.7">
          <tr><td style="color:#787680;padding-right:16px">성함</td><td>${esc(a['성함'])}</td></tr>
          <tr><td style="color:#787680">연락처</td><td>${esc(a['연락처'])}</td></tr>
          <tr><td style="color:#787680">신청</td><td>${esc(a['원하시는 형태'])}</td></tr>
          <tr><td style="color:#787680;vertical-align:top">주소</td><td>${esc(a['받으실 주소'])}</td></tr>
          <tr><td style="color:#787680;vertical-align:top">남긴 말</td><td>${esc(a['남기고 싶은 말 (선택)'])}</td></tr>
        </table>
        <p style="margin-top:20px;font-size:13px;color:#787680">남은 수량 — 가로형 ${leftH}부 · 세로형 ${leftV}부</p>
      </div>`,
  });
}

function refreshChoices(leftH, leftV) {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('formId'));
  if (leftH <= 0 && leftV <= 0) {
    closeForm();
    return;
  }
  const item = form.getItemById(Number(props.getProperty('typeItemId'))).asMultipleChoiceItem();
  const choices = [];
  if (leftH > 0) choices.push('가로형');
  if (leftV > 0) choices.push('세로형');
  if (choices.length === 2) choices.push(CONFIG.anyChoice);
  item.setChoiceValues(choices);
}

function closeForm() {
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('formId'));
  form
    .setAcceptingResponses(false)
    .setCustomClosedFormMessage(
      '준비한 캘린더 20부가 모두 신청되어 마감되었습니다.\n관심 가져 주셔서 진심으로 감사합니다.\n\n— 위성신'
    );
}
