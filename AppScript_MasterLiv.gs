// ============================================================
// Master Liv - Google Apps Script v3
// 用途：接收肝龄测试问卷数据 + 双状态追踪 + APITemplate.io PDF 生成
//      form_completed (填表完成) → report_claimed (领取报告)
// 最后更新：2026年5月
// ============================================================
//
// 数据流：
//   1) 用户填完表 → 前端送 formStatus="form_completed"
//      → append 一笔新 row，Lead状态 = "form_completed"
//
//   2) 用户点【📲 领取免费报告】→ 前端送 formStatus="report_claimed" + pdfData JSON
//      → 根据 WhatsApp 号码找到原 row，UPDATE Lead状态 + 领报告时间
//      → 调 APITemplate.io REST API → 生成客制化 PDF
//      → PDF download URL 写到 Sheet 「PDF链接」column
//      → 整 row 染浅绿色
//
// ⚠️ 部署前必做的 3 件事：
//   1. 跑 setupApiKeys() — 设定 APITemplate.io API key + template ID
//      (key 存在 Apps Script Properties，比 hardcode 安全)
//   2. 跑 setupHeaders() — Sheet 加新 column 「PDF链接」
//   3. Deploy → New version (URL 不变)
// ============================================================

// 26 个 column = 旧 22 + 新 4 (Lead状态/填表时间/领报告时间/PDF链接)
var HEADERS = [
  '提交时间',          // 1
  '姓名',              // 2
  '邮箱',              // 3
  'WhatsApp',          // 4   ← 用这个 column 来 match update
  '年龄',              // 5
  '肝龄评估(岁)',      // 6
  '总分',              // 7
  'Q1-睡眠',           // 8
  'Q2-饮酒',           // 9
  'Q3-饮食',           // 10
  'Q4-疲劳',           // 11
  'Q5-运动',           // 12
  'Q6-体型',           // 13
  'Q7-压力',           // 14
  'Q8-消化',           // 15
  'Q9-服药',           // 16
  'Q10-肝指数',        // 17
  '当前情况',          // 18
  '健康目标',          // 19
  '最大障碍',          // 20
  '方案意向',          // 21
  '备注',              // 22
  'Lead状态',          // 23
  '填表时间',          // 24
  '领报告时间',        // 25
  'PDF链接'            // 26  ← NEW v3
];

var COL = {
  PHONE:        4,
  LEAD_STATUS:  23,
  FORM_TIME:    24,
  CLAIM_TIME:   25,
  PDF_URL:      26   // NEW v3
};

// ============================================================
// formatPhone：电话号码标准化成 +60 格式
// ============================================================
function formatPhone(phone) {
  if (!phone) return '';
  phone = phone.toString().replace(/\s+/g, '').replace(/-/g, '');
  if (phone.startsWith('+60')) return phone;
  if (phone.startsWith('60')) return '+' + phone;
  if (phone.startsWith('0')) return '+6' + phone;
  return '+60' + phone;
}

// ============================================================
// _phoneMatchKey：宽松比对 key (去所有非数字 + 取后 9 位)
// 解决 +60123456789 / 60123456789 / 0123456789 都能 match
// ============================================================
function _phoneMatchKey(phone) {
  if (!phone) return '';
  var digits = phone.toString().replace(/\D/g, '');
  return digits.slice(-9);
}

// ============================================================
// doPost：前端 fetch 进来的入口
// ============================================================
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var p = e.parameter || {};
    var formStatus = (p.formStatus || 'form_completed').toString();

    if (formStatus === 'report_claimed') {
      return handleReportClaimed(sheet, p);
    } else {
      return handleFormCompleted(sheet, p);
    }

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// handleFormCompleted：用户填完表 → 新建一行
// ============================================================
function handleFormCompleted(sheet, p) {
  var now = new Date();
  var formCompletedAt = p.formCompletedTimestamp ? new Date(p.formCompletedTimestamp) : now;

  sheet.appendRow([
    now,                            // 1  提交时间
    p.userName  || '',              // 2  姓名
    p.userEmail || '',              // 3  邮箱
    formatPhone(p.userPhone),       // 4  WhatsApp
    p.userAge   || '',              // 5
    p.liverAge  || '',              // 6
    p.totalScore|| '',              // 7
    p.q1  || '', p.q2  || '', p.q3  || '', p.q4  || '', p.q5  || '',
    p.q6  || '', p.q7  || '', p.q8  || '', p.q9  || '', p.q10 || '',
    p.big1 || '', p.big2 || '', p.big3 || '', p.big4 || '', p.big5 || '',
    'form_completed',               // 23
    formCompletedAt,                // 24
    '',                             // 25  领报告时间（暂空）
    ''                              // 26  PDF链接（暂空）
  ]);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL.PHONE).setNumberFormat('@');

  return jsonResponse({
    status: 'success',
    action: 'form_completed_appended',
    row: lastRow
  });
}

// ============================================================
// handleReportClaimed：用户点按钮领报告 → UPDATE row + 生成 PDF
// ============================================================
function handleReportClaimed(sheet, p) {
  var now = new Date();
  var claimedAt = p.claimTimestamp ? new Date(p.claimTimestamp) : now;
  var targetKey = _phoneMatchKey(p.userPhone);

  if (!targetKey) {
    return jsonResponse({ status: 'error', message: 'No userPhone provided' });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return appendClaimRowFallback(sheet, p, claimedAt, 'sheet_empty');
  }

  // 找原 row (宽松 phone 比对)
  var phoneRange = sheet.getRange(2, COL.PHONE, lastRow - 1, 1);
  var phones = phoneRange.getValues();
  var foundRow = -1;
  for (var i = phones.length - 1; i >= 0; i--) {
    if (_phoneMatchKey(phones[i][0]) === targetKey) {
      foundRow = i + 2;
      break;
    }
  }

  // ── 调 APITemplate.io 生成 PDF（即使没找到 row 也试着生）──
  var pdfUrl = '';
  if (p.pdfData) {
    pdfUrl = generatePdfViaApi(p.pdfData);
  }

  if (foundRow > 0) {
    // ✅ 找到 → update 状态 + 时间 + PDF URL
    sheet.getRange(foundRow, COL.LEAD_STATUS).setValue('report_claimed');
    sheet.getRange(foundRow, COL.CLAIM_TIME).setValue(claimedAt);
    if (pdfUrl) {
      sheet.getRange(foundRow, COL.PDF_URL).setValue(pdfUrl);
    }
    sheet.getRange(foundRow, 1, 1, HEADERS.length).setBackground('#e8f8f5');

    return jsonResponse({
      status: 'success',
      action: 'report_claimed_updated',
      row: foundRow,
      pdfUrl: pdfUrl || null
    });
  } else {
    // race condition fallback
    return appendClaimRowFallback(sheet, p, claimedAt, 'no_matching_row', pdfUrl);
  }
}

// ============================================================
// appendClaimRowFallback：找不到原 row 时直接 append 一笔
// ============================================================
function appendClaimRowFallback(sheet, p, claimedAt, reason, pdfUrl) {
  var now = new Date();
  sheet.appendRow([
    now,
    p.userName  || '',
    p.userEmail || '',
    formatPhone(p.userPhone),
    p.userAge   || '',
    p.liverAge  || '',
    p.totalScore|| '',
    p.q1 || '', p.q2 || '', p.q3 || '', p.q4 || '', p.q5 || '',
    p.q6 || '', p.q7 || '', p.q8 || '', p.q9 || '', p.q10 || '',
    p.big1 || '', p.big2 || '', p.big3 || '', p.big4 || '', p.big5 || '',
    'report_claimed',
    '',
    claimedAt,
    pdfUrl || ''                 // 26  PDF链接
  ]);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL.PHONE).setNumberFormat('@');
  sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#e8f8f5');

  return jsonResponse({
    status: 'success',
    action: 'report_claimed_appended_fallback',
    reason: reason,
    row: lastRow,
    pdfUrl: pdfUrl || null
  });
}

// ============================================================
// 🔥 generatePdfViaApi：调 APITemplate.io REST API 生成 PDF
// 拿 download_url 回来。失败回 ''（不阻挡 row update）
// ============================================================
function generatePdfViaApi(pdfDataJson) {
  try {
    var props = PropertiesService.getScriptProperties();
    var apiKey = props.getProperty('APITEMPLATE_API_KEY');
    var templateId = props.getProperty('APITEMPLATE_TEMPLATE_ID');

    if (!apiKey || !templateId) {
      Logger.log('⚠️ APITemplate keys 没设定，跳过 PDF 生成。请跑 setupApiKeys()');
      return '';
    }

    // 解析 frontend 送来的 JSON 字符串
    var pdfData;
    if (typeof pdfDataJson === 'string') {
      pdfData = JSON.parse(pdfDataJson);
    } else {
      pdfData = pdfDataJson;
    }

    var url = 'https://rest.apitemplate.io/v2/create-pdf?template_id=' + templateId;
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-API-KEY': apiKey },
      payload: JSON.stringify(pdfData),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code === 200) {
      var json = JSON.parse(body);
      if (json.status === 'success' && json.download_url) {
        Logger.log('✅ PDF 生成成功: ' + json.download_url);
        return json.download_url;
      }
      Logger.log('⚠️ APITemplate response 异常: ' + body);
      return '';
    }
    Logger.log('❌ APITemplate HTTP ' + code + ': ' + body);
    return '';

  } catch (e) {
    Logger.log('❌ generatePdfViaApi error: ' + e.toString());
    return '';
  }
}

// ============================================================
// jsonResponse helper
// ============================================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet：浏览器测试可用
// ============================================================
function doGet(e) {
  return jsonResponse({
    status: 'ok',
    message: 'Master Liv Script 正常运行中 ✅',
    version: 'v3 (form_completed + report_claimed + APITemplate.io PDF)',
    headers: HEADERS.length + ' columns'
  });
}

// ============================================================
// formatHeaders：表头美化
// ============================================================
function formatHeaders(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground('#0b4a43');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

// ============================================================
// ⚠️ setupApiKeys：必跑！设定 API key + Template ID 到 PropertiesService
// 跑完后这两个 key 就安全存在 Apps Script 里，不在代码里
// ============================================================
function setupApiKeys() {
  var props = PropertiesService.getScriptProperties();

  // ⚠️ 把下面 2 个值换成你自己的（从 APITemplate.io 抓）
  var API_KEY     = '786cNTMzMDI6NTA1NTI6YjNGRlZONEdRVHZsNEs1Rg=';   // ← 你的 API Key
  var TEMPLATE_ID = 'fcd77b231e60ae88';                              // ← 你的 Template ID (肝龄)

  props.setProperty('APITEMPLATE_API_KEY', API_KEY);
  props.setProperty('APITEMPLATE_TEMPLATE_ID', TEMPLATE_ID);

  Logger.log('✅ API key + Template ID 已存到 PropertiesService');
  Logger.log('   API Key 前 6 位: ' + API_KEY.substring(0, 6) + '...');
  Logger.log('   Template ID: ' + TEMPLATE_ID);
  Logger.log('⚠️ 跑完这个 function 后，建议把上面 2 个值删掉（或改成空字串）');
  Logger.log('   因为 PropertiesService 已经存了，代码里不需要再写');
}

// ============================================================
// setupHeaders：⚠️ 部署后必跑一次 — Sheet 表头同步到 26 col
// ============================================================
function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  formatHeaders(sheet);
  sheet.autoResizeColumns(1, HEADERS.length);

  Logger.log('✅ 表头已更新完成！共 ' + HEADERS.length + ' 列');
  Logger.log('新增 column: PDF链接 (col 26)');
}

// ============================================================
// migrateOldRows：旧 row 标 legacy
// ============================================================
function migrateOldRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, COL.LEAD_STATUS, lastRow - 1, 1);
  var values = range.getValues();
  var migrated = 0;

  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      values[i][0] = 'legacy';
      migrated++;
    }
  }
  range.setValues(values);
  Logger.log('✅ 已把 ' + migrated + ' 笔旧 row 标为 legacy');
}

// ============================================================
// 🧪 测试函数（手动 Run 用）
// ============================================================
function testFormCompleted() {
  var fakeEvent = {
    parameter: {
      formStatus: 'form_completed',
      userName: '测试用户',
      userEmail: 'test@example.com',
      userPhone: '0123456789',
      userAge: '40-49岁',
      liverAge: '55',
      totalScore: '45',
      q1: '常熬夜', q2: '应酬多', q3: '高油', q4: '常累', q5: '没运动',
      q6: '有大肚腩', q7: '压力大', q8: '消化差', q9: '长期吃药', q10: '指数偏高',
      big1: '体检发现肝指数偏高',
      big2: '降低并稳定肝指数',
      big3: '工作忙应酬多',
      big4: '准备好完整疗程',
      big5: '无',
      formCompletedTimestamp: new Date().toISOString()
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}

function testReportClaimedWithPdf() {
  var IMG = 'https://cdn.jsdelivr.net/gh/ahmoimotivation/master-liv-pdf-assets@main';
  var pdfData = {
    user_name: '测试用户',
    actual_age: '45',
    actual_age_num: 45,
    liver_age: 60,
    age_diff_num: 15,
    age_diff_display: '+15 岁',
    risk_class: 'red',
    risk_label: '高风险 · 发炎期',
    risk_emoji: '🔴',
    is_red: true, is_yellow: false, is_green: false,
    report_id: 'ML-TEST-001',
    test_date: new Date().toISOString().slice(0,10),
    situation_display: '刚体检发现肝指数偏高',
    outcome_display: '降低并稳定肝指数',
    obstacle_display: '工作忙、应酬多',
    difficulty_line: '你工作忙、应酬多，要完全改作息和戒酒，确实难。',
    risk_signals: ['长期熬夜','经常饮酒应酬','饮食高油','缺乏运动','肝指数偏高'],
    diet_items: [{tag:'针对你「高油煎炸饮食」', action:'戒掉油炸食品，每天加深绿色蔬菜。'}],
    sleep_items: [{tag:'针对你「经常熬夜」', action:'23:00 前关灯入睡。'}],
    exercise_items: [{tag:'针对你「缺乏运动」', action:'每天饭后快走 15 分钟。'}],
    reasoning_bullets: ['你的现状：<strong>肝指数偏高</strong> → 已经在发炎','你的目标：<strong>降低肝指数</strong> → 需要时间'],
    recommended_boxes: 7, months: '3.5', total_days: 105,
    liver_image_url:    IMG + '/risk-red.png',
    master_liv_box_url: IMG + '/master_liv_box.png',
    ginger_image_url:   IMG + '/ginger.jpg',
    silymarin_image_url:IMG + '/silymarin.jpg',
    package_image_url:  IMG + '/package_7boxes.png',
    trust_image_url:    IMG + '/trust-30k.jpg',
    testimonial_1_url:  IMG + '/testimonial_P1.jpg',
    testimonial_2_url:  IMG + '/testimonial_P8.jpg',
    testimonial_3_url:  IMG + '/testimonial_P9.jpg',
    testimonial_4_url:  IMG + '/testimonial_P12.jpg',
    testimonial_5_url:  IMG + '/testimonial_P14.jpg'
  };
  var fakeEvent = {
    parameter: {
      formStatus: 'report_claimed',
      userName: '测试用户',
      userPhone: '0123456789',
      pdfData: JSON.stringify(pdfData),
      claimTimestamp: new Date().toISOString()
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}
