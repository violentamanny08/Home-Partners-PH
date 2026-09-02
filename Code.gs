const CONFIG = {
  SPREADSHEET_ID: '1kyic8J-NQgELsIAa-6FkkvKDSozItBTPsLJ1fr_yhqM',
  AGENT_EMAIL: 'violentamanny08@gmail.com',
  BUSINESS_NAME: 'Home Partners PH',
  PHONE: '0995 358 1519'
};

/**
 * I-run ito NANG ISANG BESES lang (Run menu sa Apps Script editor) para
 * ma-set up ang daily trigger ng 7-stage follow-up automation. Ligtas
 * itong i-run ulit — aalisin muna ang dating trigger bago gumawa ng bago.
 */
function createDailyFollowUpTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyFollowUps') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('sendDailyFollowUps')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Home Partners PH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Receives a consultation/inquiry from Index.html, saves it to the first
 * sheet in the supplied spreadsheet, then emails both the agent and client.
 */
function submitLead(data) {
  data = data || {};

  const name = clean_(data.name);
  const email = clean_(data.email);
  const phone = clean_(data.phone);
  const buyerType = clean_(data.buyerType);
  const budget = clean_(data.budget);
  const consultationDate = clean_(data.consultationDate);
  const property = clean_(data.property);
  const location = clean_(data.location);
  const needs = Array.isArray(data.needs) ? data.needs.map(clean_).filter(Boolean) : [];
  const message = clean_(data.message);
  const wantPdfGuide = !!data.wantPdfGuide;

  Logger.log('DEBUG submitLead - raw data.wantPdfGuide: ' + data.wantPdfGuide + ' | computed wantPdfGuide: ' + wantPdfGuide);

  if (!name) throw new Error('Please enter your full name.');
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) throw new Error('Please enter a valid email address.');

  const timestamp = new Date();
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheets()[0];
  ensureLeadSheetHeaders_(sheet);

  const row = [
    timestamp,
    name,
    email,
    phone,
    property,
    location,
    budget,
    consultationDate,
    buildMessage_(message, buyerType, needs) + (wantPdfGuide ? ' | Gusto ng PDF guide' : ' | Ayaw ng PDF guide'),
    'New Lead',
    0,
    'Progressing'
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Manila';
  const submittedAt = Utilities.formatDate(timestamp, tz, 'MMM d, yyyy h:mm a');
  const subject = 'New Property Inquiry   ' + name;

  const leadData = {
    name: name,
    email: email,
    phone: phone,
    buyerType: buyerType,
    budget: budget,
    consultationDate: consultationDate,
    property: property,
    location: location,
    needs: needs,
    message: message,
    wantPdfGuide: wantPdfGuide
  };

  const agentHtml = buildAgentEmail_(submittedAt, leadData);

  MailApp.sendEmail({
    to: CONFIG.AGENT_EMAIL,
    subject: subject,
    htmlBody: agentHtml,
    body: stripHtml_(agentHtml),
    replyTo: email,
    name: CONFIG.BUSINESS_NAME
  });

  const clientHtml = buildClientEmail_(leadData);

  const clientEmailOptions = {
    to: email,
    subject: 'We received your consultation request   ' + CONFIG.BUSINESS_NAME,
    htmlBody: clientHtml,
    body: stripHtml_(clientHtml),
    replyTo: CONFIG.AGENT_EMAIL,
    name: CONFIG.BUSINESS_NAME
  };

  // Kung na-check ng client ang "Isama rin ang libreng PDF Gabay" bago mag-submit,
  // ilakip ang PDF sa kanilang confirmation email. Kung hindi, walang PDF na ipapadala.
  if (wantPdfGuide) {
    Logger.log('DEBUG submitLead - attaching PDF now');
    const pdfBlob = Utilities.newBlob(
      Utilities.base64Decode(LEADMAGNET_PDF_BASE64),
      'application/pdf',
      'Gabay-sa-Pagbili-ng-Bahay.pdf'
    );
    clientEmailOptions.attachments = [pdfBlob];
  }

  MailApp.sendEmail(clientEmailOptions);

  return {
    success: true,
    message: wantPdfGuide
      ? 'Thank you! Your inquiry has been received. Check your email for your free PDF guide.'
      : 'Thank you! Your inquiry has been received. We will contact you shortly.'
  };
}

function buildMessage_(message, buyerType, needs) {
  const parts = [];
  if (buyerType) parts.push('Buyer type: ' + buyerType);
  if (needs.length) parts.push('Looking for: ' + needs.join(', '));
  if (message) parts.push(message);
  return parts.join(' | ');
}

function formatConsultationDate_(value) {
  if (!value) return 'Not specified';
  const date = new Date(value + 'T00:00:00');
  if (isNaN(date.getTime())) return clean_(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Manila', 'MMMM d, yyyy');
}

function emailValue_(value, fallback) {
  const text = clean_(value);
  return text || fallback || 'Not specified';
}

function icon_(emoji) {
  return '<span style="font-size:20px;line-height:1">' + emoji + '</span>';
}

function infoCard_(title, rows) {
  const rowHtml = rows.map(function(r) {
    return '<tr>' +
      '<td style="padding:9px 0;border-bottom:1px solid #e7edf5;color:#0b2148;font-weight:700;font-size:13px;width:42%;vertical-align:top">' +
      escapeHtml_(r[0]) + '</td>' +
      '<td style="padding:9px 0;border-bottom:1px solid #e7edf5;color:#26364f;font-size:13px;vertical-align:top">' +
      r[1] + '</td></tr>';
  }).join('');

  return '<div style="background:#ffffff;border:1px solid #e4eaf2;border-radius:12px;padding:18px;margin-bottom:14px">' +
    '<div style="font-size:14px;font-weight:800;color:#0b2b63;text-transform:uppercase;letter-spacing:.3px;margin-bottom:8px">' +
    escapeHtml_(title) + '</div>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse">' + rowHtml + '</table>' +
    '</div>';
}

function buildAgentEmail_(submittedAt, d) {
  const needsText = d.needs.length ? d.needs.join(', ') : 'Not specified';
  const consultation = formatConsultationDate_(d.consultationDate);
  const replyUrl = 'mailto:' + encodeURIComponent(d.email) + '?subject=' + encodeURIComponent('Re: Your Home Partners PH consultation request');

  const clientRows = [
    ['Name', escapeHtml_(d.name)],
    ['Email', '<a href="mailto:' + escapeHtml_(d.email) + '" style="color:#0b5bd3;text-decoration:none">' + escapeHtml_(d.email) + '</a>'],
    ['Phone', escapeHtml_(emailValue_(d.phone, 'Not provided'))],
    ['Buyer Type', escapeHtml_(emailValue_(d.buyerType, 'Not provided'))]
  ];

  const propertyRows = [
    ['Property Type', escapeHtml_(emailValue_(d.property))],
    ['Location', escapeHtml_(emailValue_(d.location))],
    ['Budget / Income', escapeHtml_(emailValue_(d.budget))],
    ['Looking For', escapeHtml_(needsText)],
    ['Message', escapeHtml_(emailValue_(d.message, ' '))]
  ];

  return '<!doctype html><html><body style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#17233a">' +
    '<div style="width:100%;background:#f4f7fb;padding:24px 0">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:720px;margin:0 auto;border-collapse:collapse;background:#ffffff">' +
    '<tr><td style="background:#061b3d;padding:22px 28px;color:#ffffff">' +
      '<table role="presentation" width="100%"><tr>' +
      '<td style="font-size:25px;font-weight:900;letter-spacing:-.5px">\uD83C\uDFE0 HOME<br><span style="font-size:12px;letter-spacing:2px">PARTNERS PH</span></td>' +
      '<td align="right" style="font-size:15px;font-weight:700">New Property Inquiry &nbsp; \uD83D\uDCE9</td>' +
      '</tr></table></td></tr>' +
    '<tr><td style="padding:22px 28px;background:linear-gradient(90deg,#f3fbf6,#ffffff);border-bottom:1px solid #e3eee7">' +
      '<div style="font-size:19px;font-weight:800;color:#13753a;margin-bottom:7px">\u2705 &nbsp; NEW LEAD RECEIVED</div>' +
      '<div style="font-size:14px;line-height:1.7;color:#42526a">A new consultation request was submitted through the Home Partners PH landing page.</div>' +
    '</td></tr>' +
    '<tr><td style="padding:22px 28px 8px">' +
      infoCard_('Client Information', clientRows) +
      infoCard_('Property Preferences', propertyRows) +
      '<div style="background:#f5fbf7;border:1px solid #d6eadc;border-radius:12px;padding:16px 18px;margin-bottom:16px">' +
        '<table role="presentation" width="100%"><tr><td><div style="font-size:13px;color:#8a5a00;font-weight:800;text-transform:uppercase">Lead Status</div><div style="display:inline-block;margin-top:7px;background:#138a45;color:#fff;border-radius:20px;padding:6px 12px;font-size:12px;font-weight:800">NEW LEAD</div></td>' +
        '<td align="right" style="font-size:13px;color:#59677d"><strong>Submitted</strong><br>' + escapeHtml_(submittedAt) + '</td></tr></table>' +
      '</div>' +
      '<div style="background:#eef4ff;border:1px solid #d6e3fb;border-radius:12px;padding:20px;text-align:center;margin-bottom:20px">' +
        '<div style="font-size:14px;font-weight:800;color:#0b2b63">\uD83D\uDCAC &nbsp; Reply directly to this email to respond to the client.</div>' +
        '<div style="font-size:12px;color:#66758b;margin:6px 0 14px">Your reply will go straight to ' + escapeHtml_(d.name) + '&#39;s inbox.</div>' +
        '<a href="' + replyUrl + '" style="display:inline-block;background:#061b3d;color:#ffffff;text-decoration:none;padding:13px 28px;border-radius:7px;font-size:13px;font-weight:800">\u2709\uFE0F &nbsp; REPLY TO CLIENT</a>' +
      '</div>' +
    '</td></tr>' +
    '<tr><td style="background:#061b3d;padding:22px 28px;color:#ffffff">' +
      '<div style="font-size:20px;font-weight:900">\uD83C\uDFE0 HOME PARTNERS PH</div>' +
      '<div style="font-size:13px;color:#dbe5f4;margin-top:5px">We help you find a home that fits your plans.</div>' +
      '<div style="border-top:1px solid #294365;margin-top:15px;padding-top:13px;font-size:12px;color:#e8eef7">\uD83D\uDCDE ' + escapeHtml_(CONFIG.PHONE) + ' &nbsp; | &nbsp; \uD83D\uDCE7 ' + escapeHtml_(CONFIG.AGENT_EMAIL) + '</div>' +
    '</td></tr>' +
    '</table></div></body></html>';
}

function buildClientEmail_(d) {
  const consultation = formatConsultationDate_(d.consultationDate);
  const needsText = d.needs.length ? d.needs.join(', ') : 'Not specified';

  const summaryRows = [
    ['\uD83C\uDFE0', 'Property Type', emailValue_(d.property)],
    ['\uD83D\uDC64', 'Buyer Type', emailValue_(d.buyerType)],
    ['\uD83D\uDCC5', 'Preferred Date', consultation],
    ['\uD83D\uDCCD', 'Location', emailValue_(d.location)],
    ['\uD83D\uDCB0', 'Budget / Income', emailValue_(d.budget)],
    ['\uD83D\uDD0D', 'Looking For', needsText]
  ];

  const cells = summaryRows.map(function(r) {
    return '<td style="width:33.33%;padding:14px 12px;border-bottom:1px solid #e8edf3;vertical-align:top">' +
      '<div style="font-size:20px">' + r[0] + '</div>' +
      '<div style="font-size:12px;font-weight:800;color:#142443;margin-top:5px">' + escapeHtml_(r[1]) + '</div>' +
      '<div style="font-size:12px;color:#56647a;margin-top:4px">' + escapeHtml_(r[2]) + '</div>' +
      '</td>';
  });

  const summaryTable = '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse"><tr>' + cells.slice(0,3).join('') + '</tr><tr>' + cells.slice(3,6).join('') + '</tr></table>';

  return '<!doctype html><html><body style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#17233a">' +
    '<div style="width:100%;background:#f4f7fb;padding:24px 0">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:720px;margin:0 auto;border-collapse:collapse;background:#ffffff">' +
    '<tr><td style="background:#061b3d;padding:24px 28px;color:#ffffff">' +
      '<div style="font-size:24px;font-weight:900;letter-spacing:-.5px">\uD83C\uDFE0 HOME</div><div style="font-size:12px;letter-spacing:2px">PARTNERS PH</div>' +
    '</td></tr>' +
    '<tr><td style="padding:32px 28px;background:#111d2c;color:#ffffff">' +
      '<div style="font-size:30px;font-weight:900;line-height:1.15">Thank you, ' + escapeHtml_(d.name) + '! \uD83C\uDF89</div>' +
      '<div style="font-size:18px;font-weight:800;color:#f7c843;margin-top:10px">We received your consultation request.</div>' +
      '<div style="width:55px;height:3px;background:#f7c843;margin-top:15px"></div>' +
    '</td></tr>' +
    '<tr><td style="padding:24px 28px">' +
      '<div style="font-size:14px;line-height:1.7;color:#34445b;margin-bottom:14px">Here&#39;s a summary of your request:</div>' +
      '<div style="border:1px solid #e0e7ef;border-radius:11px;padding:3px 12px;margin-bottom:22px">' + summaryTable + '</div>' +
      '<div style="font-size:14px;font-weight:900;color:#0b2b63;text-transform:uppercase;letter-spacing:.4px;margin:8px 0 12px">WHAT HAPPENS NEXT?</div>' +
      '<div style="background:#f7fafc;border:1px solid #e3eaf1;border-radius:12px;padding:16px">' +
        '<table role="presentation" width="100%" style="border-collapse:collapse"><tr>' +
          '<td align="center" style="width:33.33%;padding:8px"><div style="margin:auto;width:30px;height:30px;border-radius:50%;background:#f7c843;line-height:30px;font-weight:900">1</div><div style="font-size:12px;line-height:1.5;margin-top:9px">Our property specialist will review your preferences.</div></td>' +
          '<td align="center" style="width:33.33%;padding:8px"><div style="margin:auto;width:30px;height:30px;border-radius:50%;background:#f7c843;line-height:30px;font-weight:900">2</div><div style="font-size:12px;line-height:1.5;margin-top:9px">We will contact you shortly to schedule your consultation.</div></td>' +
          '<td align="center" style="width:33.33%;padding:8px"><div style="margin:auto;width:30px;height:30px;border-radius:50%;background:#f7c843;line-height:30px;font-weight:900">3</div><div style="font-size:12px;line-height:1.5;margin-top:9px">We&#39;ll help you find the right property for your needs.</div></td>' +
        '</tr></table>' +
      '</div>' +
      '<div style="margin-top:18px;background:#f2faf5;border:1px solid #cfe8d8;border-radius:11px;padding:16px 18px">' +
        '<div style="font-size:14px;font-weight:900;color:#18733c">Need to talk to us sooner?</div>' +
        '<div style="font-size:12px;color:#4d5c70;margin:5px 0 12px">You may also contact us directly.</div>' +
        '<a href="tel:' + escapeHtml_(CONFIG.PHONE.replace(/\s/g,'')) + '" style="display:inline-block;border:1px solid #2b8a57;border-radius:7px;color:#18733c;text-decoration:none;padding:10px 18px;font-size:13px;font-weight:800">\uD83D\uDCDE &nbsp; ' + escapeHtml_(CONFIG.PHONE) + '</a>' +
      '</div>' +
      '<div style="margin-top:24px;font-size:13px;line-height:1.7">Talk soon,<br><strong style="font-size:15px;color:#0b2148">The Home Partners PH Team</strong><br><span style="color:#e89b16;font-size:18px;font-style:italic">Property made simple.</span></div>' +
    '</td></tr>' +
    '<tr><td style="background:#061b3d;padding:15px 24px;color:#ffffff;font-size:11px;text-align:center">\uD83D\uDCDE ' + escapeHtml_(CONFIG.PHONE) + ' &nbsp; | &nbsp; \uD83D\uDCE7 ' + escapeHtml_(CONFIG.AGENT_EMAIL) + '</td></tr>' +
    '<tr><td style="padding:16px 24px;text-align:center;color:#7a8798;font-size:11px;line-height:1.6">This is an automated message from Home Partners PH.<br>If you need assistance, contact us at ' + escapeHtml_(CONFIG.PHONE) + '.</td></tr>' +
    '</table></div></body></html>';
}

/**
 * Nilalagyan lang ng label ang K1 at L1 KUNG blangko pa ang mga ito.
 * Hindi ito nag-iinsert ng bagong row at hindi ginagalaw ang mga
 * umiiral nang lead row — ligtas itong tawagin paulit-ulit.
 */
function ensureLeadSheetHeaders_(sheet) {
  const k1 = sheet.getRange(1, 11).getValue(); // column K
  const l1 = sheet.getRange(1, 12).getValue(); // column L

  if (!k1) {
    sheet.getRange(1, 11).setValue('Last Follow-up Day')
      .setFontWeight('bold').setBackground('#061b3d').setFontColor('#ffffff');
  }
  if (!l1) {
    sheet.getRange(1, 12).setValue('Follow-up Stage / Status')
      .setFontWeight('bold').setBackground('#061b3d').setFontColor('#ffffff');
  }
}

function clean_(value) {
  return value == null ? '' : String(value).trim();
}

function escapeHtml_(value) {
  return clean_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function stripHtml_(html) {
  return html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

/* =========================================================
 * 30-DAY CONSULTATION FOLLOW-UP SEQUENCE
 * =========================================================
 * Sends one email per day, per lead, starting the day after
 * they submit the inquiry form, for 30 days total.
 * Run sendDailyFollowUps() on a daily time-driven trigger.
 * ========================================================= */

const FOLLOWUP_CONTENT = [
  { day: 1, subject: 'Natanggap na namin ang iyong hiling, {{name}}',
    body: 'Salamat sa pagtitiwala mo sa amin. Natanggap na ng aming koponan ang iyong kahilingan, at susuriin ito nang mabuti sa loob ng isa hanggang dalawang araw. Ihahanda namin ang mga angkop na alok batay sa iyong sinabing kagustuhan. Kung may agarang katanungan ka, maaari mo lamang itong isagot dito.' },

  { day: 2, subject: 'Bakit kailangan ang libreng konsultasyon?',
    body: 'Marami ang nag-aakalang panahon lamang ang mauubos sa isang konsultasyon, ngunit dito talaga nalilinaw ang lahat   kasama na ang paraan ng pagbabayad, ang tunay na presyo, at ang buong takbo mula pag-book hanggang paglipat. Walang bayad ito at walang obligasyon. Layunin lamang naming magabayan ka sa tamang desisyon.' },

  { day: 3, subject: 'Ilang porsyento ng kita ang dapat ilaan sa bahay?',
    body: 'Karaniwang tuntunin sa larangan ng pabahay: hindi dapat lumagpas sa tatlumpung hanggang tatlumpu\'t limang bahagdan ng buwanang kita ang halagang ilalaan sa hulugan. Kung nais mong malaman ang eksaktong halaga base sa iyong kita, sabihin mo lamang at kami na ang gagawa ng sample na kwenta para sa iyo.' },

  { day: 4, subject: 'Paano gumagana ang Pag-IBIG na pautang pambahay?',
    body: 'Isa sa mga pinakapaboritong paraan ng pagbili ay sa pamamagitan ng Pag-IBIG dahil mababa ang tubo at mahaba ang taning ng pagbabayad, hanggang tatlumpung taon. Kailangan mo lamang ng aktibong kontribusyon dito. Kung interesado kang malaman kung magkano ang maaari mong mautang, ipaalam mo lamang sa amin.' },

  { day: 5, subject: 'Bangko o developer: saan mas maganda humiram?',
    body: 'May dalawa pang paraan bukod sa Pag-IBIG: sa bangko, na kadalasa\'y mas mabilis maaprubahan ngunit mas mataas ang tubo, at sa developer mismo, na kadalasa\'y mas kaunti ang kinakailangang dokumento ngunit mas maikli ang panahon ng pagbabayad. Sa oras ng konsultasyon, tutulungan ka naming timbangin kung alin ang tama para sa iyong kalagayan.' },

  { day: 6, subject: 'Handa nang tirhan o hihintayin pa?',
    body: 'Kung nais mong dumiretso sa paglipat, mas angkop sa iyo ang mga yunit na handa na. Kung mas malaya ka naman sa iskedyul at gusto mong makatipid sa panimulang halaga, maaari mong isaalang-alang ang mga yunit na binabayaran habang itinatayo pa. Magkaiba ang bentahe ng bawat isa   pag-usapan natin kung alin ang bagay sa iyo.' },

  { day: 7, subject: 'Isang linggo na   ano ang iyong napagpasyahan?',
    body: 'Isang linggo na mula nang makipag-ugnayan ka sa amin. Kung mayroon kang bagong naisip o nagbago ang iyong mga kagustuhan, sabihin mo lamang sa amin at aayusin namin ang mga mungkahi para sa iyo. Handa kaming maghintay hanggang sa tamang panahon mo.' },

  { day: 8, subject: 'Ano ang dapat isaalang-alang sa lokasyon?',
    body: 'Hindi lamang tungkol sa distansya sa trabaho ang lokasyon   kasama rin dito ang kalapitan sa paaralan, ospital, at daan patungong siyudad. Mahalaga rin ito sa halaga ng ari-arian sa hinaharap. Sabihin mo sa amin kung ano ang pinakamahalaga sa iyo upang mas mapapili namin ang tamang lugar.' },

  { day: 9, subject: 'Limang bagay na madalas na hindi napapansin',
    body: 'Narito ang ilang bagay na kadalasang nakakaligtaan ng mga unang beses na mamimili: ang kabuuang gastos kabilang ang mga karagdagang bayarin, ang paghahambing ng iba\'t ibang paraan ng pagbabayad, ang pagmamadali sa desisyon nang hindi muna nakikita ang lugar, at ang kasaysayan ng developer. Iiwasan natin ito nang magkasama.' },

  { day: 10, subject: 'Magkano ba talaga ang unang bayad?',
    body: 'Karaniwang mula lima hanggang dalawampung bahagdan ng kabuuang halaga ang inaasahang unang bayad, at maaari pa itong hulugan sa loob ng ilang buwan bago pumasok ang buong pautang. Kung nais mong malaman ang eksaktong halaga para sa property na gusto mo, sabihin mo lamang.' },

  { day: 11, subject: 'Para sa mga nagtatrabaho sa ibang bansa',
    body: 'Kung ikaw ay malayo sa Pilipinas, hindi mo kinakailangang umuwi agad para lamang mag-aplay. Maaari kang gumamit ng espesyal na kapangyarihang legal upang may kumatawan sa iyo sa mga kailangang pirmahan dito. Handa kaming samahan ka sa buong proseso kahit saan ka man naroroon.' },

  { day: 12, subject: 'Bakit magandang pagkakataon ang pamumuhunan sa ari-arian?',
    body: 'Ang halaga ng lupa at tahanan ay may kaugaliang tumaas sa paglipas ng panahon, at maaari pa itong maging pinagkukunan ng karagdagang kita kung paupahan. Kung interesado kang mamuhunan, maaari naming ipakita sa iyo ang mga lugar na may magandang potensyal.' },

  { day: 13, subject: 'Oras na bang mag-usap nang harapan?',
    body: 'Dalawang linggo na mula nang una kang makipag-ugnayan. Kung nais mo nang pag-usapan nang mas detalyado ang lahat, maaari tayong mag-iskedyul ng maikling tawag o pagpupulong   kahit labinlima hanggang dalawampung minuto lamang, sapat na upang masagot ang lahat ng iyong katanungan.' },

  { day: 14, subject: 'Anong mga papeles ang dapat ihanda?',
    body: 'Upang mapabilis ang proseso, mainam nang ihanda ang mga sumusunod: kilalanang dokumento, patunay ng kita gaya ng payslip o sertipiko ng trabaho, at kung may kasamang umuutang, kailangan din ang kanilang mga dokumento. Kung ikaw ay nagtatrabaho sa ibang bansa, kasama rin ang kontrata sa trabaho at pasaporte.' },

  { day: 15, subject: 'Ano ang buong takbo ng pag-apruba ng pautang?',
    body: 'Karaniwang ganito ang daloy: unang pagsusumite ng dokumento, sinusundan ng pagsusuri sa kredito, pagtatasa ng ari-arian, pag-apruba, at panghuli ang paglabas ng pondo. Karaniwang tumatagal ito ng tatlumpu hanggang animnapung araw depende sa napiling paraan ng pagbabayad. Sasamahan ka namin sa bawat hakbang.' },

  { day: 16, subject: 'Ano ang nangyayari kapag ka nag-reserba?',
    body: 'Sa pagreserba, may maliit na halagang babayaran upang mapanatili ang yunit habang inaayos ang iyong mga dokumento at pautang. Karaniwan itong nasa sampu hanggang dalawampu\'t limang libong piso depende sa developer. Ipapaliwanag namin nang malinaw ang kasunduan bago ka magpasya.' },

  { day: 17, subject: 'Gusto mo bang makita ang lugar mismo?',
    body: 'Wala talagang mas mabuti kaysa personal na pagbisita upang makita mo ang aktwal na yunit, mga pasilidad, at kapaligiran nito. Kung interesado ka, maaari kaming mag-iskedyul ng bisita sa oras na maginhawa para sa iyo, nang walang bayad at walang obligasyon.' },

  { day: 18, subject: 'Mga madalas itanong tungkol sa proseso',
    body: 'Isa sa mga karaniwang tanong: maaari pa rin bang umutang kahit may umiiral nang pautang? Oo, depende ito sa kabuuang ratio ng iyong utang laban sa kita. May mga nakatagong bayarin ba? Wala   ipapaliwanag namin nang buo ang lahat ng gastusin nang maaga. May iba ka pa bang gustong itanong?' },

  { day: 19, subject: 'Ang kwento ng isang naunang kliyente',
    body: 'Marami sa aming mga naunang kliyente ang nangamba noong una, lalo na tungkol sa kanilang kakayahang umutang. Ngunit sa tamang gabay at maliwanag na paliwanag, natagpuan nila ang tahanang bagay sa kanilang kakayahan sa loob lamang ng ilang buwan. Handa kaming samahan ka sa parehong paglalakbay.' },

  { day: 20, subject: 'Tatlong linggo na   ano pa ang maitutulong namin?',
    body: 'Tatlong linggo na ang nakalipas mula nang unang makipag-ugnayan ka. Kung may partikular kang katanungan tungkol sa ari-arian, sa pautang, o sa buong proseso, huwag mag-atubiling isagot ito. Nandito kami upang gabayan ka sa bawat hakbang ng iyong paglalakbay.' },

  { day: 21, subject: 'Mga payo para sa unang beses na mamimili',
    body: 'Una, huwag magmadali   siguraduhing komportable ka sa iyong desisyon. Ikalawa, ihambing ang iba\'t ibang paraan ng pagbabayad. Ikatlo, kung maaari, bisitahin muna ang lugar. At panghuli, magtanong nang walang pag-aalinlangan   walang katanungang masyadong simple.' },

  { day: 22, subject: 'Paano malalaman kung mapagkakatiwalaan ang developer?',
    body: 'Tingnan ang kasaysayan ng kompanya: kailan sila nagsimula, ilang proyekto na ang kanilang natapos, at ano ang karanasan ng mga naunang bumili. Kami ay nakikipagtulungan lamang sa mga developer na may mahabang panahon nang tiwala at magandang reputasyon.' },

  { day: 23, subject: 'Ano ba talaga ang kabuuang gastos sa pagbili?',
    body: 'Bukod sa presyo ng ari-arian, mayroon ding mga karagdagang bayarin gaya ng buwis sa paglipat, rehistro, taunang bayad sa asosasyon, at gastos sa paglipat mismo. Sa ating konsultasyon, ipapakita namin ang buong detalyadong listahan upang walang magulat sa huli.' },

  { day: 24, subject: 'Kailan ka maaaring lumipat?',
    body: 'Para sa mga yunit na handa nang tirhan, karaniwang isa hanggang tatlong buwan mula sa pag-apruba ng pautang bago makalipat. Para naman sa mga itinatayo pa, karaniwan itong isa hanggang tatlong taon depende sa proyekto. Ipapaalam namin sa iyo ang eksaktong iskedyul para sa property na iyong pinipili.' },

  { day: 25, subject: 'Mga terminong dapat mong malaman',
    body: 'Ilan sa mga karaniwang katawagan: ang ratio ng halagang maaaring hiramin laban sa presyo ng ari-arian, ang buwanang hulugan na tinatawag na amortisasyon, at ang bahaging binayaran mo mismo na tinatawag na equity. Kung nais mo ng buong paliwanag, sabihin mo lamang.' },

  { day: 26, subject: 'Bakit maganda ang bumili ngayon kaysa maghintay?',
    body: 'Kadalasang tumataas ang halaga ng ari-arian sa bawat pagdaan ng taon, kaya ang pag-aamin ngayon ay maaaring mas makatipid sa iyo kumpara sa maghintay pa. Kung seryoso kang mamuhunan, masaya kaming tulungan kang mahanap ang tamang pagkakataon.' },

  { day: 27, subject: 'May limitadong alok ba ngayon?',
    body: 'May ilang developer na nag-aalok ng mga espesyal na benepisyo tulad ng mas mababang unang bayad o mas maluwag na termino, ngunit limitado lamang ang bilang ng mga makikinabang dito. Kung nais mong malaman kung mayroon nito para sa property na gusto mo, sabihin mo lamang sa amin.' },

  { day: 28, subject: 'Halos isang buwan na tayong magkasama',
    body: 'Malapit na sa tatlumpung araw mula nang unang kumonekta ka sa amin. Kung interesado ka pa ring magpatuloy, masaya kaming ituloy ang usapan. Kung nagbago naman ang iyong plano, walang problema   ipaalam mo lamang upang hindi ka na namin gambalain pa.' },

  { day: 29, subject: 'Bukas na ang huling araw ng abiso',
    body: 'Bukas na ang huli sa serye ng mga mensaheng ito. Kung nais mong mag-iskedyul ng pormal na pag-uusap o kailangan mo pa ng anumang detalye, mas mabilis kung tatawagan mo kami nang direkta o isagot ang email na ito.' },

  { day: 30, subject: 'Huli naming paalala   narito pa rin kami',
    body: 'Ito na ang huling awtomatikong mensahe sa seryeng ito. Kung nais mo pa ring ipagpatuloy ang paghahanap ng iyong tahanan, hindi kami mawawala   tumawag lamang o isagot ang anumang email namin anumang oras. Maraming salamat sa iyong tiwala sa Home Partners PH.' }
];

/**
 * Bagong 7-stage na iskedyul (kapalit ng dating araw-araw na 30 email):
 *   Day 1 -> Stage 1, Day 6 -> Stage 2, Day 11 -> Stage 3, Day 16 -> Stage 4,
 *   Day 21 -> Stage 5, Day 26 -> Stage 6, Day 30 -> Stage 7 (dito tumitigil).
 * Ginagamit lang nito ang existing content sa FOLLOWUP_CONTENT sa itaas base
 * sa "day" — wala pang nabura roon.
 */
const STAGE_SCHEDULE = [
  { day: 1,  stage: 1 },
  { day: 6,  stage: 2 },
  { day: 11, stage: 3 },
  { day: 16, stage: 4 },
  { day: 21, stage: 5 },
  { day: 26, stage: 6 },
  { day: 30, stage: 7 }
];

function getFollowUpContentForDay_(day) {
  return FOLLOWUP_CONTENT.filter(function (c) { return c.day === day; })[0] || null;
}

/**
 * Runs once a day via a time-driven trigger. Checks every lead in the
 * sheet (columns A:L lang, kasabay ng existing structure), at nagpapadala
 * ng susunod na due stage email base sa STAGE_SCHEDULE. Column K = Last
 * Follow-up Day (numero), column L = "Stage N - Progressing/Completed".
 * Tumitigil kapag Stage 7 (Day 30) na, o kapag "Stopped" ang laman ng L.
 */
function sendDailyFollowUps() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 12); // columns A:L
  const values = range.getValues();
  const today = new Date();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const timestamp = row[0];
    const name = clean_(row[1]);
    const email = clean_(row[2]);
    const lastDaySent = Number(row[10]) || 0;                 // column K
    const statusText = clean_(row[11]) || 'Progressing';       // column L

    if (!email || !(timestamp instanceof Date)) continue;
    if (statusText.indexOf('Completed') !== -1 || statusText.indexOf('Stopped') !== -1) continue;

    const daysSinceSubmit = Math.floor((today - timestamp) / (1000 * 60 * 60 * 24));

    // hanapin ang pinakabagong stage na dapat nang naipadala
    let dueEntry = null;
    STAGE_SCHEDULE.forEach(function (entry) {
      if (daysSinceSubmit >= entry.day && lastDaySent < entry.day) {
        dueEntry = entry;
      }
    });
    if (!dueEntry) continue;

    const content = getFollowUpContentForDay_(dueEntry.day);
    if (!content) continue;

    try {
      const html = buildFollowUpEmail_(name, content);
      MailApp.sendEmail({
        to: email,
        subject: content.subject.replace('{{name}}', name || 'kaibigan'),
        htmlBody: html,
        body: stripHtml_(html),
        replyTo: CONFIG.AGENT_EMAIL,
        name: CONFIG.BUSINESS_NAME
      });

      const newStatus = 'Stage ' + dueEntry.stage + ' - ' + (dueEntry.stage === 7 ? 'Completed' : 'Progressing');
      sheet.getRange(i + 2, 11).setValue(dueEntry.day);   // column K
      sheet.getRange(i + 2, 12).setValue(newStatus);      // column L
    } catch (err) {
      Logger.log('Follow-up send failed for ' + email + ' (stage ' + dueEntry.stage + '): ' + err);
    }
  }
}

function buildFollowUpEmail_(name, content) {
  const greeting = name ? name.split(' ')[0] : 'there';
  const bodyText = content.body;

  return '<!doctype html><html><body style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#17233a">' +
    '<div style="width:100%;background:#f4f7fb;padding:24px 0">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;margin:0 auto;border-collapse:collapse;background:#ffffff">' +
    '<tr><td style="background:#061b3d;padding:22px 28px;color:#ffffff">' +
      '<div style="font-size:20px;font-weight:900;letter-spacing:-.5px">HOME PARTNERS PH</div>' +
      '<div style="font-size:11px;letter-spacing:2px;color:#dbe5f4;margin-top:2px">Day ' + escapeHtml_(String(content.day)) + ' of 30</div>' +
    '</td></tr>' +
    '<tr><td style="padding:28px">' +
      '<div style="font-size:17px;font-weight:800;color:#0b2148;margin-bottom:14px">Hi ' + escapeHtml_(greeting) + ',</div>' +
      '<div style="font-size:14px;line-height:1.8;color:#34445b">' + escapeHtml_(bodyText) + '</div>' +
      '<div style="margin-top:22px">' +
        '<a href="tel:' + escapeHtml_(CONFIG.PHONE.replace(/\s/g,'')) + '" style="display:inline-block;background:#061b3d;color:#ffffff;text-decoration:none;padding:12px 24px;border-radius:7px;font-size:13px;font-weight:800">Call ' + escapeHtml_(CONFIG.PHONE) + '</a>' +
      '</div>' +
    '</td></tr>' +
    '<tr><td style="background:#061b3d;padding:16px 24px;color:#ffffff;font-size:11px;text-align:center">' + escapeHtml_(CONFIG.PHONE) + ' &nbsp;|&nbsp; ' + escapeHtml_(CONFIG.AGENT_EMAIL) + '</td></tr>' +
    '<tr><td style="padding:14px 24px;text-align:center;color:#7a8798;font-size:10.5px;line-height:1.6">' +
      'Natatanggap mo ito dahil humiling ka ng konsultasyon sa Home Partners PH.<br>' +
      'Isagot ang "STOP" anumang oras upang itigil ang mga update na ito.' +
    '</td></tr>' +
    '</table></div></body></html>';
}

/**
 * Optional helper: run this manually (or wire it to an email-reply
 * automation) to stop follow-ups for a specific email address, e.g. if
 * a client replies "STOP".
 */
function stopFollowUpsFor(emailAddress) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // column C
  for (let i = 0; i < emails.length; i++) {
    if (clean_(emails[i][0]).toLowerCase() === clean_(emailAddress).toLowerCase()) {
      sheet.getRange(i + 2, 12).setValue('Stopped'); // column L
    }
  }
}


/* =========================================================
 * LEAD MAGNET: Free PDF Guide Capture
 * =========================================================
 * Saves name/email to a separate sheet tab and emails the
 * requester the free guide (embedded as base64, attached).
 * ========================================================= */

const LEADMAGNET_PDF_BASE64 =
  'JVBERi0xLjcKJeLjz9MKMSAwIG9iago8PAovUHJvZHVjZXIgKHB5cGRmKQo+PgplbmRvYmoKMiAwIG9iago8PAovVHlwZSAvUGFnZXMKL0NvdW50IDQKL0tpZHMgWyA0IDAgUiAyNSAwIFIgNTEgMCBSIDUzIDAgUiBdCj4+CmVuZG9iagozIDAgb2JqCjw8Ci9UeXBl'
  + 'IC9DYXRhbG9nCi9QYWdlcyAyIDAgUgo+PgplbmRvYmoKNCAwIG9iago8PAovVHlwZSAvUGFnZQovTWVkaWFCb3ggWyAwIDAgNTk1LjI3NTU5MSA4NDEuODg5NzY0IF0KL0NvbnRlbnRzIDUgMCBSCi9SZXNvdXJjZXMgNiAwIFIKL1RyaW1Cb3ggWyAwIDAgNTk1LjI3'
  + 'NTU5MSA4NDEuODg5NzY0IF0KL0JsZWVkQm94IFsgMCAwIDU5NS4yNzU1OTEgODQxLjg4OTc2NCBdCi9QYXJlbnQgMiAwIFIKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0ZpbHRlciAvRmxhdGVEZWNvZGUKL0xlbmd0aCAxMjUyCj4+CnN0cmVhbQp42u1Y32vkRgx+91/h'
  + '58LOSjPS/IDD0NJS2r40bY4UQh+cIXsUNmmT3P9PNd7Y1mxs392W0pfGEK8/eaTRN9KMZGxBrh3Kv0hoYkzBU5sfmqcGTOBBOv0Y4Kdm/3P/8eP982ObX5r9X9C+5MemiENyJgCEGFpEaw1j8pHb5/vm0FzJ9aSufY8G2g8vDbblev7QfHPd4DgZtiahc9G3wRnygTm2'
  + '1w/N/v3N9z/cfN0itdeH5vYdgD10u4Dvuh2SNzZyQBr/F2nudj6uSnlL6jalNm6ORZGG1bH9lhTcpt3NsZY2pdseUWdXh7rcud/b6x+b766nlfzcFUSOxiXLrlpCgmkJ+w7kRnS68fjkyp3z8CikDE9ev/oKOtAgh0rb4XQ7aWE83ax+or5S5jrr1fi+mlQ+VzpwsuCz'
  + 'BTY8+Xy7Zs1yNdm7N66vcg4mufInuZnYJsLyI1imtL4QIjUYg5OULAvx209XV7/+IrlkvPeAflyPkc9+gWt+vdXzvq+YzyOI4c1CjGpwQffsvlJKlaUzEpVOdHd2QfX45BfCoF6KMSjcgi91wLwGIVQe1dE36txgkqo45c/QUkc0LcheJzhaqBep166Mr9SLtGzwNBBh'
  + 'I9qDV4F1+zbh0klT1PydObmU2aOW/Mm0rYfzWrquDg/LebgaUIjnubnvwTiutsMk5yZE77F1NhkOFMG1D00AIS5ZtF7jxyZYLBsvuKTx4GQ7LgepoESGCYOzFeqtiY48+TY3CpdT2Hh2TIPuGY+yNSV5nUOrLEYL04k7z0+hWbmj4DJtQ1gmq2EOhihGrMzNoJpbbmZY'
  + 'uXLUsPJ7MqYoWuQ5V/SzESddOGd/hIsXUroE2UqjggOJiij/rczBGslZS1ShFfczXnM/4xGiOAWJuFUGoxgcidHcT2jN/QQfG5/Mm5fntFTWVD1VUT/BNfUzPLs92Zr5WWI4N4evJB3+j/3/NPZvmsehXlCKZrEP0cgduaWERubM0ZUi/TRISozTLla2eVQTd5AMMVv3'
  + '2h7MaoYDQUp0Q+Ci1ObQqgd5d/9HltIEubd9cpkkFVyfkuuzF697tpyx/fbPYSMdCx2EYDBYkCVhFBcdWpb4GbxwCOKvwo+NkxyzFJO1GneUDKCkgbztwUgc20gVGtBgApQMyc0yfqzwiCZBBCiZMVuURfAyxHOr5zejWbszw8cmeWGUMEUNl/g4BaYyp0A950X4qOHZ'
  + 'b2VspmiR5zyFgkwqnspNsE7qTKkypb0jV8pNlEYxDAedMjcrsb7ElgOSbTJKbUpAY5B9yfuH6Xzd6jYQ5fT1VspuIcLLVuRlR1zuGKW83SGURkcytmp00umY35DK6X+5NG9JS9e3IU2Xj7X/mkfbXEnZdTEb23a3/XWpC6tz2rZqL/dHatvLWQ6fYLkqNb+0/Y5ysEiK'
  + 'STbpdAhDrkGcmvC7bseFN04lw0EOdM8WKVGhVIhhvyaEDWH5TLIqLOu4rtZ1jte00tZAFiJxTRi2RuKlnhSb60L/DzjYmNBdh2NcgJH6kodvAJCiLLX8oCQF53pURG/kTE7eVd8C3FlYuKX+crxVn2YgVR82lgbUHSznhZ5Vktfim8Z77OBipTQsdIx+GHAHrJ7Up52h'
  + '2ZXucLBRd/yjctbKz74+1X1x0OPO+tCsB4wWKi/OO8jp+hudkXMBCmVuZHN0cmVhbQplbmRvYmoKNiAwIG9iago8PAovRXh0R1N0YXRlIDw8Ci9hMS4wIDw8Ci9jYSAxCj4+Ci9hMC4zNSA8PAovY2EgMC4zNQo+PgovYTEgPDwKL2NhIDEKPj4KL2EwLjY4IDw8Ci9j'
  + 'YSAwLjY4Cj4+Cj4+Ci9YT2JqZWN0IDw8Ci9pYzcyNTE1YTJhOTNjNDc0MDNhOTkzYWM2MzYyYTUyNWMxIDcgMCBSCj4+Ci9QYXR0ZXJuIDw8Ci9wMCA4IDAgUgo+PgovU2hhZGluZyA8PAo+PgovQ29sb3JTcGFjZSAxMyAwIFIKL0ZvbnQgMTQgMCBSCj4+CmVuZG9i'
  + 'ago3IDAgb2JqCjw8Ci9UeXBlIC9YT2JqZWN0Ci9TdWJ0eXBlIC9JbWFnZQovV2lkdGggNjU0Ci9IZWlnaHQgNTM4Ci9Db2xvclNwYWNlIC9EZXZpY2VSR0IKL0JpdHNQZXJDb21wb25lbnQgOAovSW50ZXJwb2xhdGUgdHJ1ZQovRmlsdGVyIC9EQ1REZWNvZGUKL0xl'
  + 'bmd0aCA1NjcwNgo+PgpzdHJlYW0K/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwQDAwQEAwQFBAQFBgoHBgYGBg0JCggKDw0QEA8NDw4RExgUERIXEg4PFRwVFxkZGxsbEBQdHx0aHxgaGxr/2wBDAQQFBQYFBgwHBwwaEQ8RGhoaGhoaGhoaGhoaGhoaGhoaGhoa'
  + 'GhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhoaGhr/wAARCAIaAo4DASIAAhEBAxEB/8QAHQAAAAcBAQEAAAAAAAAAAAAAAAECAwQFBgcICf/EAFAQAAIBAwMCBAMFAwcICgECBwECAwAEEQUSIQYxE0FRYQcicRQygZGhI0KxCBUzUnLB0SQ0NWJ0suHwFjZDU3OCkqKz'
  + '8WMlZHWjRFRlg9L/xAAbAQACAwEBAQAAAAAAAAAAAAAAAgEDBAUGB//EADERAAICAQQBBAEDAwMFAQAAAAABAgMRBBIhMUEFEyJRMhRhcSOB0UKRsQYVUsHwof/aAAwDAQACEQMRAD8A9uaiwXT7oscDwXyfwNZGyAdw4PA881qtVbGm3pzjED/7prB6TqDPHGJFxn/2'
  + '4NaoLMWUyeGjTCcJgEinhOAu48/jVE17G7EuQFUZUUtL9tpBwRjuDUbA3F6t4ucHFOiZTjJ4NZgagsjD6eVORXp8QKclTyPYUOsN5qEfOeeacBBFUSX6xEZJ74Iq1gmDAHjBrNJOLLU1IfLc80TNgZNIndVTdkcVDNyWOEJx500eSHwSzKF70QkFRPEJA3HB9z3pSPnP'
  + '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'
  + 'hilYoYoATihilYo8UAIxRYpeKPFADYFUsv8ATy/2z/Gr3FUU3+cTf2z/ABp4kMOH/OYf7Y/jV9mqGH/OYv7Y/jV7UsEHmhmioUpIeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGa'
  + 'KhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGaKhQAeaGa'
  + 'KhQAeaoJv85m/tn+NX1UM3+cTf2z/GmRDP/ZCmVuZHN0cmVhbQplbmRvYmoKOCAwIG9iago8PAovVHlwZSAvUGF0dGVybgovUGF0dGVyblR5cGUgMQovQkJveCBbIDAgMCA3OTMuNzAwNzg3IDExMjIuNTE5NjggXQovWFN0ZXAgNzkzLjcwMDc4NwovWVN0ZXAgMTEy'
  + 'Mi41MTk2OAovVGlsaW5nVHlwZSAxCi9QYWludFR5cGUgMQovTWF0cml4IFsgMC43NSAwIDAgLTAuNzUgMCA4NDEuODg5NzY0IF0KL1Jlc291cmNlcyA5IDAgUgovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDE0Cj4+CnN0cmVhbQp42tOvMFBwyQcABY8Bqwpl'
  + 'bmRzdHJlYW0KZW5kb2JqCjkgMCBvYmoKPDwKL0V4dEdTdGF0ZSA8PAo+PgovWE9iamVjdCA8PAoveDAgMTAgMCBSCj4+Ci9QYXR0ZXJuIDw8Cj4+Ci9TaGFkaW5nIDw8Cj4+Ci9Db2xvclNwYWNlIDEzIDAgUgovRm9udCAxNCAwIFIKPj4KZW5kb2JqCjEwIDAgb2Jq'
  + 'Cjw8Ci9UeXBlIC9YT2JqZWN0Ci9TdWJ0eXBlIC9Gb3JtCi9CQm94IFsgMCAwIDc5My43MDA3ODcgMTEyMi41MTk2OCBdCi9SZXNvdXJjZXMgMTEgMCBSCi9Hcm91cCA8PAovVHlwZSAvR3JvdXAKL1MgL1RyYW5zcGFyZW5jeQovSSB0cnVlCi9DUyAvRGV2aWNlUkdC'
  + 'Cj4+Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGggMjQKPj4Kc3RyZWFtCnjaM1QwAEJDMJmcy6VfbKBQnAEAKmkEigplbmRzdHJlYW0KZW5kb2JqCjExIDAgb2JqCjw8Ci9FeHRHU3RhdGUgPDwKPj4KL1hPYmplY3QgPDwKPj4KL1BhdHRlcm4gPDwKPj4KL1No'
  + 'YWRpbmcgPDwKL3MwIDEyIDAgUgo+PgovQ29sb3JTcGFjZSAxMyAwIFIKL0ZvbnQgMTQgMCBSCj4+CmVuZG9iagoxMiAwIG9iago8PAovU2hhZGluZ1R5cGUgMgovQ29sb3JTcGFjZSAvRGV2aWNlUkdCCi9Eb21haW4gWyAwIDEgXQovQ29vcmRzIFsgLTgyLjIwNDcy'
  + 'NCA4Mi4yMDQ3MjUgODc1LjkwNTUxMiAxMDQwLjMxNDk2IF0KL0Z1bmN0aW9uIDw8Ci9GdW5jdGlvblR5cGUgMwovRG9tYWluIFsgMCAxIF0KL0VuY29kZSBbIDAgMSAwIDEgXQovQm91bmRzIFsgMC41IF0KL0Z1bmN0aW9ucyBbIDw8Ci9GdW5jdGlvblR5cGUgMgov'
  + 'RG9tYWluIFsgMCAxIF0KL0MwIFsgMC45ODgyMzUgMC40MTU2ODYgMC4wMDc4NDMgXQovQzEgWyAwLjA2Mjc0NSAwLjIxMTc2NSAwLjQzMTM3MyBdCi9OIDEKPj4gPDwKL0Z1bmN0aW9uVHlwZSAyCi9Eb21haW4gWyAwIDEgXQovQzAgWyAwLjA2Mjc0NSAwLjIxMTc2'
  + 'NSAwLjQzMTM3MyBdCi9DMSBbIDAuMDU4ODI0IDAuNDc4NDMxIDAuMjQ3MDU5IF0KL04gMQo+PiBdCj4+Ci9FeHRlbmQgWyB0cnVlIHRydWUgXQo+PgplbmRvYmoKMTMgMCBvYmoKPDwKL2xhYi1kNTAgWyAvTGFiIDw8Ci9XaGl0ZVBvaW50IFsgMC45NjQyOTYgMSAw'
  + 'LjgyNTEwNSBdCi9SYW5nZSBbIC0xMjUgMTI1IC0xMjUgMTI1IF0KPj4gXQovbGFiLWQ2NSBbIC9MYWIgPDwKL1doaXRlUG9pbnQgWyAwLjk1MDQ1NiAxIDEuMDg5MDU4IF0KL1JhbmdlIFsgLTEyNSAxMjUgLTEyNSAxMjUgXQo+PiBdCj4+CmVuZG9iagoxNCAwIG9i'
  + 'ago8PAovVVdHSVdBIDE1IDAgUgovWEtRUVNSIDIwIDAgUgo+PgplbmRvYmoKMTUgMCBvYmoKPDwKL1R5cGUgL0ZvbnQKL1N1YnR5cGUgL1R5cGUwCi9CYXNlRm9udCAvVVdHSVdBK0RlamFWdS1TYW5zLUJvbGQKL1RvVW5pY29kZSAxNiAwIFIKL0VuY29kaW5nIC9J'
  + 'ZGVudGl0eS1ICi9EZXNjZW5kYW50Rm9udHMgWyAxNyAwIFIgXQo+PgplbmRvYmoKMTYgMCBvYmoKPDwKL0ZpbHRlciAvRmxhdGVEZWNvZGUKL0xlbmd0aCAzNTkKPj4Kc3RyZWFtCnjaXZLLaoQwFIb3PkWW08VgNJfpgAhlunHRC532ATQ5ToUaQ3QWvn1j/mEKDSh8'
  + '/Of2Jyc/Nc+NGxaWv4fJnGlh/eBsoHm6BkOso8vgsqJkdjDLjdLfjK3P8ph8XueFxsb1U1ZVLP+I4ryEle2e7NTRQ5a/BUthcBe2+zqdI5+v3v/QSG5hPKtrZqmPhV5a/9qOxPKUtm9s1Idl3cecv4jP1RMrExcYxkyWZt8aCq27UFbxeGpW9fHUGTn7TxccaV1vvtuw'
  + 'hZd9DOdcmjqRAR1BClQmEiAFKh+hKWgFiKC1oEMiLhKVHJqEViAP/RT6SWgamkQ/jX4KkQdEKg0SqMJR00I7QJOoAn8a/iSq6FsVTK0xdTSWqEckPGh4EPCg4EF20B7h7wh/GpEgBZK4JY1bKpEnkSfQT6KfwNQKUwv4UyI94u21tufctu6+K+YaQlyTtJppP7bNGBzd'
  + 't9dPfsvavl9eYMDaCmVuZHN0cmVhbQplbmRvYmoKMTcgMCBvYmoKPDwKL1R5cGUgL0ZvbnQKL1N1YnR5cGUgL0NJREZvbnRUeXBlMgovQmFzZUZvbnQgL1VXR0lXQStEZWphVnUtU2Fucy1Cb2xkCi9DSURTeXN0ZW1JbmZvIDw8Ci9SZWdpc3RyeSAoQWRvYmUpCi9P'
  + 'cmRlcmluZyAoSWRlbnRpdHkpCi9TdXBwbGVtZW50IDAKPj4KL0NJRFRvR0lETWFwIC9JZGVudGl0eQovVyBbIDMgWyAzNDggXSA5IFsgODcyIF0gMzYgWyA3NzQgNzYyIF0gNDAgWyA2ODMgXSA0MiBbIDgyMSA4MzcgMzcyIF0gNDcgWyA2MzcgOTk1IDgzNyA4NTAg'
  + 'NzMzIF0gNTMgWyA3NzAgNzIwIDY4MiBdIDU3IFsgNzc0IF0gNjAgWyA3MjQgXSA2OCBbIDY3NSA3MTYgXSA3MiBbIDY3OCBdIDc0IFsgNzE2IDcxMiAzNDMgXSA3OSBbIDM0MyBdIDgxIFsgNzEyIDY4NyBdIDg2IFsgNTk1IDQ3OCBdIDkyIFsgNjUyIF0gXQovRm9u'
  + 'dERlc2NyaXB0b3IgMTggMCBSCj4+CmVuZG9iagoxOCAwIG9iago8PAovVHlwZSAvRm9udERlc2NyaXB0b3IKL0ZvbnROYW1lIC9VV0dJV0ErRGVqYVZ1LVNhbnMtQm9sZAovRm9udEZhbWlseSAoRGVqYVZ1IFNhbnMpCi9GbGFncyA0Ci9Gb250QkJveCBbIDAgLTIz'
  + 'NiA5OTUgOTI4IF0KL0l0YWxpY0FuZ2xlIDAKL0FzY2VudCA5MjgKL0Rlc2NlbnQgLTIzNgovQ2FwSGVpZ2h0IDkyOAovU3RlbVYgODAKL1N0ZW1IIDgwCi9Gb250RmlsZTIgMTkgMCBSCj4+CmVuZG9iagoxOSAwIG9iago8PAovTGVuZ3RoMSA0MTI4Ci9GaWx0ZXIg'
  + 'L0ZsYXRlRGVjb2RlCi9MZW5ndGggMjYxMAo+PgpzdHJlYW0KeNq9V3tUlMcVv/d7QfABiwvE1McuKyAxIrLurlhjasTVIOAzhgRBl4XloQury8YQQELSk6eNpoQoxpMoRQ5Rmyr1JDZuSUvM03hMYjBJ1RpDT5oqNjYnNRHY2d759hFqbZr2j373fPPNfHNfc+c3d2YA'
  + 'AWAUPAAiLLFaV+Tuvrr7FED00/R33IKs+VaxTdxN7QPUdi1ePi1jrWH9RQDMp/ZKu9Pmkt6VzQDiJGrvL7O5XRBBBNGPUntk2bpax6sps0cDSJsBIs+Ul9pK0kanuqjvDL3mcvoxukUpJH1aak8qd9bc1/nO6Cxqkz2sWVdtty0albuJ9L9H/V6n7T4X2CCD+mZSW1dl'
  + 'c5beMnLuBhpAOf1jrmp3jf803EP2y3k/8LEJRYdzD3w+sDp69t9hYiTw59OB9zP59/PDByb7+5k1IlPREm8kCBB4SC7CycZTecXf7++PyFQ1DX8K1T+FMJ3iFpS45kGpRfCCDCAb5VZqTgh8xY/AIcSSWyMiRVGRBEG6VnqJY34J+a6DWkXLtLgjwol93/FIx8ER4hTe'
  + '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'
  + 'jn94sKtLfo71+IEl5Vn8cPBDPI2Ac7iWsYTDWGkNRIfyA0HRYOKIt4xFB04lPgdLXfDbZ0+4CpYnRElrfJHCt4PmIy0XZy2aSPtCBN1M6PQo0T6C6ep5ktcR4qkVqAsQidZgXQQd5gXr0rC6DDdiVbCuwAR8GObRGdMFtbCBzptlUE7nTR2tFDuk0jeDzlHpYKRaMXHo'
  + '6IxaQf1uejfQ2dNGZ7xb6O8ddAK1QxrV5tJpdB19l4V1udVWKX1LSeZeKkuIM+oHWDWHra4gS/eSrUqSqSJu7oeNZP47i1lUqyS5leAhDjvx2lRtpaqETR2RjrRUUekinmLSW0F8OpKvJus2te9aPctVLW7yqJr4S/5Nry7cv1L1yk26qlVLGeSbESz/JBeSmhqWguA9'
  + 'Efwf8/vpdR+B30DppiCGbnufin778Puq+uU9I8MSdIxQbyQ3AEfSJNpfkGZhMpU308wiTCPfEEzkH0ImcOnZREjezqdyIRFCNt1HUD3xIyylyCON4S4qVxHhPwA5re3LCmVuZHN0cmVhbQplbmRvYmoKMjAgMCBvYmoKPDwKL1R5cGUgL0ZvbnQKL1N1YnR5cGUgL1R5'
  + 'cGUwCi9CYXNlRm9udCAvWEtRUVNSK0RlamFWdS1TYW5zCi9Ub1VuaWNvZGUgMjEgMCBSCi9FbmNvZGluZyAvSWRlbnRpdHktSAovRGVzY2VuZGFudEZvbnRzIFsgMjIgMCBSIF0KPj4KZW5kb2JqCjIxIDAgb2JqCjw8Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5n'
  + 'dGggMzY3Cj4+CnN0cmVhbQp42l2Sy26DMBBF93yFl+kiwphXIiGkKt2w6EOl/QCwhxSpGMuQBX9f2xclUi2BdHzn2jPjiS/NS6PHlcUfdpYtrWwYtbK0zDcrifV0HXWUCKZGue4U/nLqTBQ7c7stK02NHuaoqlj86cRltRs7PKu5p6cofreK7Kiv7PB9aR23N2N+aSK9'
  + 'Mh7VNVM0uINeO/PWTcTiYDs2yunjuh2d5xHxtRliInCCZOSsaDGdJNvpK0UVd6tm1eBWHZFW//R0t/WD/OmsD099OOeZqj1lXaCiBGWgJBBPAwkeKAeVO+UgAR/B10MroWXQJLQzaABJRCYggtaDTiDcUOAGDp/wviTtBbTicW6BnRTn7voAglbuGuov9vqRa4FcE2gC'
  + 'Wn6Gr0DkCZE5NFRV4vYEdQjUweET8An0NENPU2g5tAJ5dqGnPfcVCy5QcYp+u3T9s+7v5x/Yz+F9euTNWjc4YVjDxPhZGTXd59nMxrv89we2b8SwCmVuZHN0cmVhbQplbmRvYmoKMjIgMCBvYmoKPDwKL1R5cGUgL0ZvbnQKL1N1YnR5cGUgL0NJREZvbnRUeXBlMgov'
  + 'QmFzZUZvbnQgL1hLUVFTUitEZWphVnUtU2FucwovQ0lEU3lzdGVtSW5mbyA8PAovUmVnaXN0cnkgKEFkb2JlKQovT3JkZXJpbmcgKElkZW50aXR5KQovU3VwcGxlbWVudCAwCj4+Ci9DSURUb0dJRE1hcCAvSWRlbnRpdHkKL1cgWyAzIFsgMzE4IF0gOSBbIDc4MCBd'
  + 'IDE1IFsgMzE4IDM2MSAzMTggXSAzNiBbIDY4NCBdIDQ4IFsgODYzIF0gNTEgWyA2MDMgXSA1NyBbIDY4NCBdIDY4IFsgNjEzIDYzNSA1NTAgNjM1IDYxNSBdIDc0IFsgNjM1IDYzNCAyNzggXSA3OCBbIDU3OSAyNzggOTc0IDYzNCA2MTIgNjM1IF0gODUgWyA0MTEg'
  + 'NTIxIDM5MiBdIDg5IFsgNTkyIF0gOTIgWyA1OTIgXSA5OCBbIDMxOCBdIDI4MjEgWyA1OTAgXSA1MDQyIFsgNjMwIF0gXQovRm9udERlc2NyaXB0b3IgMjMgMCBSCj4+CmVuZG9iagoyMyAwIG9iago8PAovVHlwZSAvRm9udERlc2NyaXB0b3IKL0ZvbnROYW1lIC9Y'
  + 'S1FRU1IrRGVqYVZ1LVNhbnMKL0ZvbnRGYW1pbHkgKERlamFWdSBTYW5zKQovRmxhZ3MgNAovRm9udEJCb3ggWyAwIC0yMzYgOTc0IDkyOCBdCi9JdGFsaWNBbmdsZSAwCi9Bc2NlbnQgOTI4Ci9EZXNjZW50IC0yMzYKL0NhcEhlaWdodCA5MjgKL1N0ZW1WIDgwCi9T'
  + 'dGVtSCA4MAovRm9udEZpbGUyIDI0IDAgUgo+PgplbmRvYmoKMjQgMCBvYmoKPDwKL0xlbmd0aDEgMzM4NzYKL0ZpbHRlciAvRmxhdGVEZWNvZGUKL0xlbmd0aCAyNzQwCj4+CnN0cmVhbQp42u1dfVSUVRp/7n3fmUHkIAMzUKdUhkFQMTBGQEwTy1ylZNHQkswYvoQE'
  + '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'
  + 'ZHN0cmVhbQplbmRvYmoKMjUgMCBvYmoKPDwKL1R5cGUgL1BhZ2UKL01lZGlhQm94IFsgMCAwIDU5NS4yNzU1OTEgODQxLjg4OTc2NCBdCi9Db250ZW50cyAyNiAwIFIKL1Jlc291cmNlcyAyNyAwIFIKL1RyaW1Cb3ggWyAwIDAgNTk1LjI3NTU5MSA4NDEuODg5NzY0'
  + 'IF0KL0JsZWVkQm94IFsgMCAwIDU5NS4yNzU1OTEgODQxLjg4OTc2NCBdCi9QYXJlbnQgMiAwIFIKPj4KZW5kb2JqCjI2IDAgb2JqCjw8Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGggMjE3MAo+PgpzdHJlYW0KeNrlW82L3DYUv89f4XMhir4lQ1hIoYS2p022'
  + 'bGHpwePs5LJbSHrov98n2fLoyc+yZibkUhYysiw9vW/99OSIjsPfGwH/eC2Y972zuhtfD18PnDkT3y6N2J3/Oa2Y49x51/meie71YPjUeoktG1v5qKnn8afD34e3g2C8+/IPLORVL4WFhawSyiloSGVkr7tvXwIdLTvbLzS0BhrfnokXbnpx+ulwD8xZ3inumQDmXw9O'
  + 'LcNSJzCmeiZxZ8arEpLZsmsaNeYyGS8ZltJ4x2DembrpYUjGQnweA4exBcrSMCG2QappemzEYdBKC79MDxNjcVJiPJN2PDyCer/mCu6tdjwouHfS6D40vFYiKXghclam9Ip5aYI+A7XdQSfQ+T2yqdNGhIW46G1saCGcDStGXnvKMv1iGc+sQZ1Oy7MZFJOy6DpbRq4s'
  + 'k/VY1gfTLORNLwrTiMU0IprGMG6mB5BnJhBbk3FEbhyx8DZNS8xnEo//Z8eE0ITA7DtlLDOmk9wwoNVzkEAFfc3eFpIPjykA/EZYb6WDhuaQJmR0oF0K9RGTs349SKeXd9NobWBB4bwXJUeZa9dYayfZNjKxGgSCzGqRPJopSbOYxXlgLTY8sJ20VydVH0FpL47O+b+O'
  + 'tXaSbSMz7Wnjwz6GJfLXqW+P1s4QQoHzcCyEv02DezQbhyZugYX35+wesrkPK0NO1y6m957L7uOHg+j+jQqQivUypLh8nbnz5fDpCqoEIUigUi8A4CbqZ6O9Eqq5neNEO2f4FtqBX9iKzYrhufN2jhfqOcs3UQeejfRMY46nrpv5TZQzbjHlxhT+88NBJFgMu5q1zFor'
  + 'hegeXg9v/3j88Ovj+w5GPpwOT+84V/aOw48e44/h04+fOk/Tj546RfzhKv4ogd45RGV6Jz7nE+SRIKZd3plGPhOdeL000uSdUt7xv7qH3w6/PBxiIpj0vahIGGgpQkVgIEDsVkkfdfTn7/f3nz52gIRM0pLEXA85E0blTMzvtMkVkhQ5onkUMY0k0lhnaDrdORvAEsqi'
  + 'OxVhuLT6851w2UocvTwVQwkyWLLZ1B4xg/idZTGS8DdD2T9Np+YVqw+UkWbBBoL5ad5wcmcShTpGSp04Knjp8tE9kesBmBVWzK73tFbzsaL0gYoo7JsVm625W82bNVp0+opBGngZKm5ITSfXS3qpOdyR8lCzy2cRggOhicIVsBCJX6xem5OZmWkITGPupMh82xJBBO8i'
  + 'NwKZcJ2y1s5nYSfxNnO+IkzMdqapxqOjdEbmq2E7/RjSUfyuv6SRjtjKCl8atiOgjPb96CL9FAtBr7Tvi4m03E+vKNLThFnnp+1oxqmhZcfA4eEIUQpQUMuXpHd6gDsiT40XpHscww5x4HPuyGRR+EdDnrWVZGgIh6hhABpQIAhR+P8kihA5/nFwGuoD8IylyFDBklsQ'
  + 'UUrBrPDK9DlGFGf0M27DQcy6rMCXJIirbDskzWcit7RvOxKvRykujaTExKQLiBmRPCg2YE0TVA0a55rSMJxMudVWY4QpM4RJAsZm2Fl4LE5ehURHCrk3Z1Ic3EXSwxDqREXYSOH/cTuV1SxWZHVOZSZFbMO1qBVTp5jxnz33TKtwvw/i2/NUYSic7v2l+wOZqHR1+ctD'
  + 'a5Wope2ZVv3s3E9bXlyLqyrsJHPrjYIft8EBme4LbEWicnSE1qaCZU6FEdQmNmnBCJd4W6u584A8+n7SgcsiIwkxP+GfDe921xmyUFeDQ3rPvNC5QxY65ghjy3yfxmpJZjwh6PBdXKqGiwsXJiJgMcrUKcyeUUIqKye5ZktWvLlIyfv2xDiF63Bpmtdp1FINznfRUNQP'
  + 'h2WfEs2CU/Jd9Hj3Rph38I9lSk6noKLjVHaMRYfi5QizJir7fMaAn1eLKlFSGIoO0F4xRd05WxUEFs3HuzsrLuIBZuxpYrxzhVyZ4c6Xsglp9mB2vwKYDgznLPMBX2pkOFGAn2uz+SZCxJlB+vOdC2Iow7gpeyOUUmQNTyOuy6oT8sp0wY8XoZTxCrwKBtOCM90rsCzC'
  + 'q2KJtCWTGLTqVibZNAReKCc8IIp+MzdF4YuceERzh5IE5iaIqw1TkFnWXLRtedvibRJuTdsVKduS+FpYHzyBUnl/o7CbhDVllsuENTmldmGN0JCoKgZQzZjmuSL69jL2MsxEiW7bQyyGuZrDfL4SCgEWPmIhNlRjFJNqyctLmKe8zD/HjUjrdW3H7paSjavUnpvrHtdW'
  + 'lhpq3fMKGNorTpxSikrO2FxwH5oryvh2RFFwEw2pHo5GSlkeFZRxCbkoDZI1gZrUF9+LVQrM+LRAg39HiIuJ7VVJcCiHWHCwMSJIgHWLT3nzfn91yc1KACC1khtCRLCatNn1gdt1AaWQx/rta7Zq6W2jnKdmMs6tLCTH3fyQWDSb127LrXTltnezqHs53gFzGMNktT6n'
  + 'fHNtocGpEZYrrxpMJfr2S9HTPjFBQoM2mIbC1a3lqObbygWYIiv48AltXkiiDwZjay3v2kp6Nds0m5dmCeff4qRgUW01u/tb6QqOXoyjGgec54Q1e/dRt3ls9Z4DXXEWCWD/gv8cDPvVMEuVutClI3mOS0/oOMfVds0xubwhitWmp764IAWkfN0pwZRHvj5QGnOEUn+0'
  + 'yxcH3yHPuHSBzJ7T8PJN1PnrpdZ7k/CVs1JyPzsfuVn4/Z7ronz049bFsf0D10U+eeG62Xfaocnj142AdNLDy/Sg5wc0PHWu/mPFxsdz02f/aV7+Rb9kbvrsM1s5f6/S+/l/WWSLmekDf2j0UpiwmA3lLOo6z9nwpbg1VuaK6lEdkkhGvkBo6FyFqs8pYXn0zhIIb15p'
  + '/bmSPBHguET/JOjb+QakTGWgZAFgzhdaeVrLSJfF8O2bxp61/P0HcxzfKgplbmRzdHJlYW0KZW5kb2JqCjI3IDAgb2JqCjw8Ci9FeHRHU3RhdGUgPDwKL2ExLjAgPDwKL2NhIDEKPj4KL0EwLjg1IDw8Ci9DQSAwLjg1Cj4+Ci9hMC44NSA8PAovY2EgMC44NQo+Pgov'
  + 'QTAuOSA8PAovQ0EgMC45Cj4+Ci9hMC45IDw8Ci9jYSAwLjkKPj4KL0ExLjAgPDwKL0NBIDEKPj4KPj4KL1hPYmplY3QgPDwKL3gwIDI4IDAgUgoveDEgNDIgMCBSCi94MiA0NCAwIFIKPj4KL1BhdHRlcm4gPDwKL3AwIDQ2IDAgUgo+PgovU2hhZGluZyA8PAo+Pgov'
  + 'Q29sb3JTcGFjZSAzMCAwIFIKL0ZvbnQgMzEgMCBSCj4+CmVuZG9iagoyOCAwIG9iago8PAovVHlwZSAvWE9iamVjdAovU3VidHlwZSAvRm9ybQovQkJveCBbIDAgMCA3OTMuNzAwNzg3IDExMjIuNTE5NjggXQovUmVzb3VyY2VzIDI5IDAgUgovR3JvdXAgPDwKL1R5'
  + 'cGUgL0dyb3VwCi9TIC9UcmFuc3BhcmVuY3kKL0kgdHJ1ZQovQ1MgL0RldmljZVJHQgo+PgovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDE1Mgo+PgpzdHJlYW0KeNqdj70KwjAUhff7FPcFkt6b5OYHSkBBRDch0kEcSqGdHNT3B9PRoRl6znLg4xvOG7qRNeHy'
  + 'BdLJu0Ae63BCKa6DJEaDnwWOBRipVjGKYLDaiTMpYHlBdx/Ol+GAbLDM8OiJzJxVtH1WHEj7v6x0yiqkTSotapvUxKbLTXdsUbL7XeP2/60u+81DU7ZPLFc4Fbj9AChaYQYKZW5kc3RyZWFtCmVuZG9iagoyOSAwIG9iago8PAovRXh0R1N0YXRlIDw8Ci9hMS4wIDw8'
  + 'Ci9jYSAxCj4+Cj4+Ci9YT2JqZWN0IDw8Cj4+Ci9QYXR0ZXJuIDw8Cj4+Ci9TaGFkaW5nIDw8Cj4+Ci9Db2xvclNwYWNlIDMwIDAgUgovRm9udCAzMSAwIFIKPj4KZW5kb2JqCjMwIDAgb2JqCjw8Ci9sYWItZDUwIFsgL0xhYiA8PAovV2hpdGVQb2ludCBbIDAuOTY0'
  + 'Mjk2IDEgMC44MjUxMDUgXQovUmFuZ2UgWyAtMTI1IDEyNSAtMTI1IDEyNSBdCj4+IF0KL2xhYi1kNjUgWyAvTGFiIDw8Ci9XaGl0ZVBvaW50IFsgMC45NTA0NTYgMSAxLjA4OTA1OCBdCi9SYW5nZSBbIC0xMjUgMTI1IC0xMjUgMTI1IF0KPj4gXQo+PgplbmRvYmoK'
  + 'MzEgMCBvYmoKPDwKL1VXR0lXQSAzMiAwIFIKL1hLUVFTUiAzNyAwIFIKPj4KZW5kb2JqCjMyIDAgb2JqCjw8Ci9UeXBlIC9Gb250Ci9TdWJ0eXBlIC9UeXBlMAovQmFzZUZvbnQgL1VXR0lXQStEZWphVnUtU2Fucy1Cb2xkCi9Ub1VuaWNvZGUgMzMgMCBSCi9FbmNv'
  + 'ZGluZyAvSWRlbnRpdHktSAovRGVzY2VuZGFudEZvbnRzIFsgMzQgMCBSIF0KPj4KZW5kb2JqCjMzIDAgb2JqCjw8Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGggNTIxCj4+CnN0cmVhbQp42l2UzYrbMBSF934KLaeLwZauJE8hBMrMJov+0LQP4B85NUxs4ziL'
  + 'vH1lfWIKNSRwcq7OvecovuXr6e00jZsqf6xzdw6bGsapX8Ntvq9dUG24jFOhjerHbssofXfXZinKePj8uG3hepqGuTgcVPkzkrdtfainL/3chk9F+X3twzpOF/X0+/Uc8fm+LO/hGqZNVcXxqPowRKGvzfKtuQZVpmPPpz7y4/Z4jmf+Vfx6LEGZhDXDdHMfbkvThbWZ'
  + 'LqE4VPE5qsMQn2MRpv4/3nuOtUP3p1n3ctPE8qqy9XFH1ibkNciBTEKuS6j+nFAlCZkKzsNJQlKh2cPVcBbNAc0OhKZH02lQABnQQCVzeuYUuju6G+a0zGlbKl/gQBYkaFo0zQvIgfBu8S5oOjQFDw4PwpyWOQXvDu8GfxZ/Bn8WfwJy2S0p+ZwSs9TMovuEpMER3T3d'
  + 'TUCzBeFIsiMqLZWOlGpSEhJ0JOhAdb53ND2aFZkZMtP4E/xpNCVpNsPOmf3HxOFB8FCBTEZ4N9k76db5xqj0uZIEDQlqbkW4FU3yhuQ1CRoSFFRcTpAOQofYNuXiQTiy+d8K5+A0CQoJajjJHJkJmWkcCY40mQmZeW6lyW8HjmocOTTrPAt5Wiotk3lJL29+S/fXeN82'
  + 'Hzuiu69rXA9pJaW9sG+EcQofW2uZl/3U/vkL+04lywplbmRzdHJlYW0KZW5kb2JqCjM0IDAgb2JqCjw8Ci9UeXBlIC9Gb250Ci9TdWJ0eXBlIC9DSURGb250VHlwZTIKL0Jhc2VGb250IC9VV0dJV0ErRGVqYVZ1LVNhbnMtQm9sZAovQ0lEU3lzdGVtSW5mbyA8PAov'
  + 'UmVnaXN0cnkgKEFkb2JlKQovT3JkZXJpbmcgKElkZW50aXR5KQovU3VwcGxlbWVudCAwCj4+Ci9DSURUb0dJRE1hcCAvSWRlbnRpdHkKL1cgWyAzIFsgMzQ4IF0gOCBbIDEwMDIgXSAxMSBbIDQ1NyA0NTcgXSAxNSBbIDM4MCA0MTUgMzgwIF0gMTkgWyA2OTYgNjk2'
  + 'IDY5NiA2OTYgNjk2IDY5NiA2OTYgNjk2IDY5NiA2OTYgNDAwIF0gMzQgWyA1ODAgMTAwMCA3NzQgNzYyIDczNCA4MzAgNjgzIDY4MyA4MjEgODM3IDM3MiBdIDQ2IFsgNzc1IDYzNyA5OTUgODM3IDg1MCA3MzMgXSA1MyBbIDc3MCA3MjAgNjgyIDgxMiA3NzQgMTEw'
  + 'MyBdIDYwIFsgNzI0IF0gNjggWyA2NzUgNzE2IDU5MyA3MTYgNjc4IF0gNzQgWyA3MTYgNzEyIDM0MyBdIDc4IFsgNjY1IDM0MyAxMDQyIDcxMiA2ODcgNzE2IF0gODUgWyA0OTMgNTk1IDQ3OCA3MTIgNjUyIDkyNCBdIDkyIFsgNjUyIF0gOTUgWyAzNjUgXSA5OCBb'
  + 'IDM0OCBdIDI4MDYgWyA1MDAgXSBdCi9Gb250RGVzY3JpcHRvciAzNSAwIFIKPj4KZW5kb2JqCjM1IDAgb2JqCjw8Ci9UeXBlIC9Gb250RGVzY3JpcHRvcgovRm9udE5hbWUgL1VXR0lXQStEZWphVnUtU2Fucy1Cb2xkCi9Gb250RmFtaWx5IChEZWphVnUgU2FucykK'
  + 'L0ZsYWdzIDQKL0ZvbnRCQm94IFsgMCAtMjM2IDExMDMgOTI4IF0KL0l0YWxpY0FuZ2xlIDAKL0FzY2VudCA5MjgKL0Rlc2NlbnQgLTIzNgovQ2FwSGVpZ2h0IDkyOAovU3RlbVYgODAKL1N0ZW1IIDgwCi9Gb250RmlsZTIgMzYgMCBSCj4+CmVuZG9iagozNiAwIG9i'
  + 'ago8PAovTGVuZ3RoMSAyMzE2NAovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDQ4MDAKPj4Kc3RyZWFtCnja7Tx5XFVVt2ud4YI4cRk1B7gySaQiCKhlmiLiABqZCaIyXiYZZNAQEJXKIYeMNFFLJERUKiUqIzDFLM3Pz/wMDRXNeKYZmvkzJ7ibt/a5F0S/Xu97'
  + 'f7+zFvucfc5ea+2117D3PpfzO4AA0B2WgAjTAgKmB21/sP0MQOg1utt3gv/4gB6pPZIBwiS6XjD1pSFeSa7a1wFwJl3PiE6OTOtZ6DIVYNBquq6Ki8xIAzNCCL1A193i5mXrx+4fbAPg7wlgvSo+NjJmyNwNK6iNt/vG042ev2t3kzyiAef45MxX42bqvqfrvQAeF+al'
  + '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'
  + 'Y1dqQYpsLR2tgX8neQD5FxUNEXyA/3r6HCGStoF0nART6BhM+iK8THohhBMizCVEsl8UoMVdC/5rIvw3zj6azgplbmRzdHJlYW0KZW5kb2JqCjM3IDAgb2JqCjw8Ci9UeXBlIC9Gb250Ci9TdWJ0eXBlIC9UeXBlMAovQmFzZUZvbnQgL1hLUVFTUitEZWphVnUtU2Fu'
  + 'cwovVG9Vbmljb2RlIDM4IDAgUgovRW5jb2RpbmcgL0lkZW50aXR5LUgKL0Rlc2NlbmRhbnRGb250cyBbIDM5IDAgUiBdCj4+CmVuZG9iagozOCAwIG9iago8PAovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDUzMgo+PgpzdHJlYW0KeNpdlM2q2zAQhfd+Ci1v'
  + 'FxdLI8lOIBjK7SaL/tC0D+AfOTU0tnGcRd6+tj5xLzSQwMnRnJkzI03+dv5yHodV5T+Wqb2EVfXD2C3hPj2WNqgmXIcxM6K6oV0Tir/trZ6zfAu+PO9ruJ3HfspOJ5X/3Mj7ujzVy+duasKnLP++dGEZxqt6+f122fDlMc9/wy2Mq9JZVaku9JvQ13r+Vt+CymPY67nb'
  + '+GF9vm4xHyd+PeegJGJDMe3Uhftct2Gpx2vITnr7VOrUb58qC2P3H1+UhDV9+6de9uN2P66166oduTqiogQ5kIlI24hER+RBZUIeJMQF4hq4Es7BtXBHUA9qOWlAAa4BHUBkKMigiZM9zthG4IoP3YJ/LLqJ70FwZeLwXyT/1FpQq4ETOH8kruDkgZMeDlcl2Q0+BB+a'
  + 'OCFO6KmjpxbOwxXUWceeNnp3LFpwbOm3T/1mTmWcU92X8aShak9lJZUJfXOoGLxbvBs0rUalQCVxqFhUNEiSJpNxcTLN4RjjGhwZOmhTB3FrE0ctNtWCB8tdM2hapm3oi6Uvmvsk3CfdgWr6QmU+1YmKoKKZijAVS2WeyoTsjuzCfXLcQyGfI5+g4lARVBwqFkc+zYj5'
  + 'Oe6Z4MHhwZLPJ7d4sHhwx3R/qZqeSXp1xG1p9+ec3u3+sPf987412seybAsjLqm4KfYdMYzhfY/N07xH7d9/p0UqMwplbmRzdHJlYW0KZW5kb2JqCjM5IDAgb2JqCjw8Ci9UeXBlIC9Gb250Ci9TdWJ0eXBlIC9DSURGb250VHlwZTIKL0Jhc2VGb250IC9YS1FRU1Ir'
  + 'RGVqYVZ1LVNhbnMKL0NJRFN5c3RlbUluZm8gPDwKL1JlZ2lzdHJ5IChBZG9iZSkKL09yZGVyaW5nIChJZGVudGl0eSkKL1N1cHBsZW1lbnQgMAo+PgovQ0lEVG9HSURNYXAgL0lkZW50aXR5Ci9XIFsgMyBbIDMxOCA0MDEgXSA4IFsgOTUwIDc4MCAyNzUgMzkwIDM5'
  + 'MCA1MDAgODM4IDMxOCAzNjEgMzE4IF0gMTkgWyA2MzYgNjM2IDYzNiA2MzYgNjM2IDYzNiA2MzYgNjM2IDYzNiBdIDI5IFsgMzM3IF0gMzYgWyA2ODQgNjg2IF0gMzkgWyA3NzAgXSA0MSBbIDU3NSA3NzUgNzUyIDI5NSBdIDQ2IFsgNjU2IDU1NyA4NjMgXSA1MCBb'
  + 'IDc4NyA2MDMgXSA1MyBbIDY5NSBdIDU1IFsgNjExIDczMiA2ODQgOTg5IF0gNjggWyA2MTMgNjM1IDU1MCA2MzUgNjE1IDM1MiA2MzUgNjM0IDI3OCBdIDc4IFsgNTc5IDI3OCA5NzQgNjM0IDYxMiA2MzUgXSA4NSBbIDQxMSA1MjEgMzkyIDYzNCA1OTIgODE4IF0g'
  + 'OTIgWyA1OTIgXSA5OCBbIDMxOCBdIDI4MDYgWyA1MDAgMTAwMCBdIDI4MjEgWyA1OTAgXSAyOTUzIFsgNjM2IF0gNTA0MiBbIDYzMCBdIF0KL0ZvbnREZXNjcmlwdG9yIDQwIDAgUgo+PgplbmRvYmoKNDAgMCBvYmoKPDwKL1R5cGUgL0ZvbnREZXNjcmlwdG9yCi9G'
  + 'b250TmFtZSAvWEtRUVNSK0RlamFWdS1TYW5zCi9Gb250RmFtaWx5IChEZWphVnUgU2FucykKL0ZsYWdzIDQKL0ZvbnRCQm94IFsgMCAtMjM2IDEwMDAgOTI4IF0KL0l0YWxpY0FuZ2xlIDAKL0FzY2VudCA5MjgKL0Rlc2NlbnQgLTIzNgovQ2FwSGVpZ2h0IDkyOAov'
  + 'U3RlbVYgODAKL1N0ZW1IIDgwCi9Gb250RmlsZTIgNDEgMCBSCj4+CmVuZG9iago0MSAwIG9iago8PAovTGVuZ3RoMSAzNjkwMAovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDUxNDcKPj4Kc3RyZWFtCnja7V0LXFRVt997n3NmwCcMM1BpyTg8NHwyDihWWpkf'
  + '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'
  + 'N42rjqQlhFKrpaDSwgxxIxocS7pRWT5Dn7DOu4EwSVPapbAvv2QZaZAQopjinPY7W4w6ECoYmFCSVmVZd8ZSk5nAjDkt11MN3Ye34Fv4r8J0FvcD2Rs85Fmdsr7bt9pty5fRvqSBPIM4toKqYRX7eHcFW9J+q6hg6V04a1/MPvkbW9lw2XX2GVY1CmWfyBXhvQIPUj6j'
  + 'y+oYucOVvU6QEx7tqAsd7osd6hJ6BE9w1FVIiyPR8ygaLUKxaAmah+aiKLQU3rn6oHDUF87+aBCQEWphwOGJngOepSgGyhI0B5nRQtQP7o5BbwD/AKg9i14H8kQT22XFKFdz4DwH2rwJxwjg7PQntAa0a2WfGH4TdM2HNm8AN7PDDG3+cxpHQW0+tJuClgFHOPCaFWlz'
  + 'lBZmpUeeIOUNOC4CnjCQOw/4PKF9NGg3K89+L+cVRUoMWBQNtADuMq0xwButSPIH3UbIox1btbUR7B9Otn0LI/vHgBFFxGZTOJVPPl8UbOHKueXssPYze9KlvYUOuEW4x95nMeTdHnB8HAhD3vCGoy8QBj/3g+MA8DMGG41wNAGxjB+k7NXHwXECWIrRZLAOoxlAGHwS'
  + 'hnCnXzvdQ7iz1BnWpc4pnVMQ7rq7aznCbmVuZcqnx5k9ndBFdA2FoKe6IPV5xbBEmBGzEfstsQJW74i2axzlqE/+1x/rxt1A7pt/4vPfoXZexc9t9Q4yHrof2uF6yYM6GYpQazrUn2pv2gd1Q+5Kjx+D8cdQ6wzjwHovOjiErn9nfM7LEOq6B/j7sJzSdTfblTpnP8RX'
  + '3s73STvf3nY+9A9UPbUfCmVuZHN0cmVhbQplbmRvYmoKNDIgMCBvYmoKPDwKL1R5cGUgL1hPYmplY3QKL1N1YnR5cGUgL0Zvcm0KL0JCb3ggWyAwIDAgNzkzLjcwMDc4NyAxMTIyLjUxOTY4IF0KL1Jlc291cmNlcyA0MyAwIFIKL0dyb3VwIDw8Ci9UeXBlIC9Hcm91'
  + 'cAovUyAvVHJhbnNwYXJlbmN5Ci9JIHRydWUKL0NTIC9EZXZpY2VSR0IKPj4KL0ZpbHRlciAvRmxhdGVEZWNvZGUKL0xlbmd0aCAyNjkKPj4Kc3RyZWFtCnjafZKxTsQwDIb3PIVfgJzd2EkqIQYkFpgKGZAQQ+9Eb2K44/0l2ksD8TU6dXDi2r/tLz6Z3UgW4fhj0Pae'
  + 'A3qYDyzYx+WAEmMH56N5TIYA5++OQATYiQ2uR+8gfZvd+8swvL0CeUiT+bhHdPiAs+ExG74YdBcjq5H6H39lZ8i3Q3FSqGImJUMNbd43RFlVKnm5Naw0ye27hnS5+bq1K5W1pU4NqGYRJS1+U/2/3pz3CelZA/do5Q/4glhuIOagBDfyVcuogE0qEhsYSgX9emM9YwnR'
  + 'r9cumBMJG2ylz85YE7vq3reGX1Va3Jd6W7ZRrNNslW7Z1INidGsPFH69k83dIlrbekpm+AXXosJ2CmVuZHN0cmVhbQplbmRvYmoKNDMgMCBvYmoKPDwKL0V4dEdTdGF0ZSA8PAovYTEuMCA8PAovY2EgMQo+Pgo+PgovWE9iamVjdCA8PAo+PgovUGF0dGVybiA8PAo+'
  + 'PgovU2hhZGluZyA8PAo+PgovQ29sb3JTcGFjZSAzMCAwIFIKL0ZvbnQgMzEgMCBSCj4+CmVuZG9iago0NCAwIG9iago8PAovVHlwZSAvWE9iamVjdAovU3VidHlwZSAvRm9ybQovQkJveCBbIDAgMCA3OTMuNzAwNzg3IDExMjIuNTE5NjggXQovUmVzb3VyY2VzIDQ1'
  + 'IDAgUgovR3JvdXAgPDwKL1R5cGUgL0dyb3VwCi9TIC9UcmFuc3BhcmVuY3kKL0kgdHJ1ZQovQ1MgL0RldmljZVJHQgo+PgovRmlsdGVyIC9GbGF0ZURlY29kZQovTGVuZ3RoIDE5Mwo+PgpzdHJlYW0KeNptkL8KwjAQxvc8xb2A7V1y1yQgHQQXnaoZBHHQQjs5qO8P'
  + 'pm0qCUiG7/79vuPyUvWdKoTxo7DyDVtsIAYs6N0UoDin4T2qXVAEGN+GQAQcYaWNEc8Qnqq+HLvufAIiCIO6bhENthiFeRahQvpZ0Czi80z/A5ILDzmeeokzvtU0zaSmzgl2hanNTRPfzMADJct+3qYlO4Gy7EjmYgpzyc3LE6X4DLY5t440S7HPgXVDcQXFK24QDmof'
  + 'VPcFu+ZerQplbmRzdHJlYW0KZW5kb2JqCjQ1IDAgb2JqCjw8Ci9FeHRHU3RhdGUgPDwKL2ExLjAgPDwKL2NhIDEKPj4KPj4KL1hPYmplY3QgPDwKPj4KL1BhdHRlcm4gPDwKPj4KL1NoYWRpbmcgPDwKPj4KL0NvbG9yU3BhY2UgMzAgMCBSCi9Gb250IDMxIDAgUgo+'
  + 'PgplbmRvYmoKNDYgMCBvYmoKPDwKL1R5cGUgL1BhdHRlcm4KL1BhdHRlcm5UeXBlIDEKL0JCb3ggWyAwIDAgNzkzLjcwMDc4NyAxMTIyLjUxOTY4IF0KL1hTdGVwIDc5My43MDA3ODcKL1lTdGVwIDExMjIuNTE5NjgKL1RpbGluZ1R5cGUgMQovUGFpbnRUeXBlIDEK'
  + 'L01hdHJpeCBbIDAuNzUgMCAwIC0wLjc1IDAgODQxLjg4OTc2NCBdCi9SZXNvdXJjZXMgNDcgMCBSCi9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGggMTQKPj4Kc3RyZWFtCnja068wUHDJBwAFjwGrCmVuZHN0cmVhbQplbmRvYmoKNDcgMCBvYmoKPDwKL0V4dEdT'
  + 'dGF0ZSA8PAo+PgovWE9iamVjdCA8PAoveDAgNDggMCBSCj4+Ci9QYXR0ZXJuIDw8Cj4+Ci9TaGFkaW5nIDw8Cj4+Ci9Db2xvclNwYWNlIDMwIDAgUgovRm9udCAzMSAwIFIKPj4KZW5kb2JqCjQ4IDAgb2JqCjw8Ci9UeXBlIC9YT2JqZWN0Ci9TdWJ0eXBlIC9Gb3Jt'
  + 'Ci9CQm94IFsgMCAwIDc5My43MDA3ODcgMTEyMi41MTk2OCBdCi9SZXNvdXJjZXMgNDkgMCBSCi9Hcm91cCA8PAovVHlwZSAvR3JvdXAKL1MgL1RyYW5zcGFyZW5jeQovSSB0cnVlCi9DUyAvRGV2aWNlUkdCCj4+Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGgg'
  + 'MjQKPj4Kc3RyZWFtCnjaM1QwAEJDMJmcy6VfbKBQnAEAKmkEigplbmRzdHJlYW0KZW5kb2JqCjQ5IDAgb2JqCjw8Ci9FeHRHU3RhdGUgPDwKPj4KL1hPYmplY3QgPDwKPj4KL1BhdHRlcm4gPDwKPj4KL1NoYWRpbmcgPDwKL3MwIDUwIDAgUgo+PgovQ29sb3JTcGFj'
  + 'ZSAzMCAwIFIKL0ZvbnQgMzEgMCBSCj4+CmVuZG9iago1MCAwIG9iago8PAovU2hhZGluZ1R5cGUgMgovQ29sb3JTcGFjZSAvRGV2aWNlUkdCCi9Eb21haW4gWyAwIDEgXQovQ29vcmRzIFsgLTgyLjIwNDcyNCA4Mi4yMDQ3MjUgODc1LjkwNTUxMiAxMDQwLjMxNDk2'
  + 'IF0KL0Z1bmN0aW9uIDw8Ci9GdW5jdGlvblR5cGUgMwovRG9tYWluIFsgMCAxIF0KL0VuY29kZSBbIDAgMSAwIDEgMCAxIF0KL0JvdW5kcyBbIDAuNDIgMC42MiBdCi9GdW5jdGlvbnMgWyA8PAovRnVuY3Rpb25UeXBlIDIKL0RvbWFpbiBbIDAgMSBdCi9DMCBbIDEg'
  + 'MC40MTk2MDggMCBdCi9DMSBbIDAuMDkwMTk2IDAuMjM5MjE2IDAuNDUwOTggXQovTiAxCj4+IDw8Ci9GdW5jdGlvblR5cGUgMgovRG9tYWluIFsgMCAxIF0KL0MwIFsgMC4wOTAxOTYgMC4yMzkyMTYgMC40NTA5OCBdCi9DMSBbIDAuMDIzNTI5IDAuMTY4NjI3IDAu'
  + 'NDAzOTIyIF0KL04gMQo+PiA8PAovRnVuY3Rpb25UeXBlIDIKL0RvbWFpbiBbIDAgMSBdCi9DMCBbIDAuMDIzNTI5IDAuMTY4NjI3IDAuNDAzOTIyIF0KL0MxIFsgMC4wNTg4MjQgMC40ODIzNTMgMC4yNDcwNTkgXQovTiAxCj4+IF0KPj4KL0V4dGVuZCBbIHRydWUg'
  + 'dHJ1ZSBdCj4+CmVuZG9iago1MSAwIG9iago8PAovVHlwZSAvUGFnZQovTWVkaWFCb3ggWyAwIDAgNTk1LjI3NTU5MSA4NDEuODg5NzY0IF0KL0NvbnRlbnRzIDUyIDAgUgovUmVzb3VyY2VzIDI3IDAgUgovVHJpbUJveCBbIDAgMCA1OTUuMjc1NTkxIDg0MS44ODk3'
  + 'NjQgXQovQmxlZWRCb3ggWyAwIDAgNTk1LjI3NTU5MSA4NDEuODg5NzY0IF0KL1BhcmVudCAyIDAgUgo+PgplbmRvYmoKNTIgMCBvYmoKPDwKL0ZpbHRlciAvRmxhdGVEZWNvZGUKL0xlbmd0aCAyMDAyCj4+CnN0cmVhbQp42uVbS4scNxC+z6+Ys8FtlVSlB5iFBEJI'
  + 'ctpkwxqWHGY7O77YATv/HyKppWmVWv0arzcJZsGjVqtL9ZC+r0rdhqPwf6/B/2MROmud0XjsPx4+HURnKN69NGJ3+WdQdUYIY83Rug6OHw8khtaH2NKxVY4aeu5fHf46vDlBJ47v//YTWeUkaD+RVqCM8g2pSDo8fn4f5KA8aneRgehlfH5q3DDDjfOrw61XTosjaNE5'
  + '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'
  + '6H2NAj4nsJVXGnTazokZAALxNXX5+weFB3HnCmVuZHN0cmVhbQplbmRvYmoKNTMgMCBvYmoKPDwKL1R5cGUgL1BhZ2UKL01lZGlhQm94IFsgMCAwIDU5NS4yNzU1OTEgODQxLjg4OTc2NCBdCi9Db250ZW50cyA1NCAwIFIKL1Jlc291cmNlcyAyNyAwIFIKL1RyaW1C'
  + 'b3ggWyAwIDAgNTk1LjI3NTU5MSA4NDEuODg5NzY0IF0KL0JsZWVkQm94IFsgMCAwIDU5NS4yNzU1OTEgODQxLjg4OTc2NCBdCi9QYXJlbnQgMiAwIFIKPj4KZW5kb2JqCjU0IDAgb2JqCjw8Ci9GaWx0ZXIgL0ZsYXRlRGVjb2RlCi9MZW5ndGggMjM0Mwo+PgpzdHJl'
  + 'YW0KeNq9W1uP7DQMfp9fMc9IlNyci4SQQEIIeDqw6CAhHmaqHV72IAH/X8JpkjZO3Uynszra1dZJc7G/2I7jZuVZ4M+XEv94Iwfvg7PmPH46/XMSg4Pp7UxM1fWPM3pwQjjvzj4M8vzpBCJRbxNlJ6pulWo+fnH6+/TVRQ7i/Nd/OJHXQUmLE1kttdNIKA0qmPO/f8Vx'
  + 'jDrbMI9hDI7x7yvzwqUXty9OH5A5q89S6ImrmURmxCAhl5w9S4lCS48CxoJSg0FyPFWFt6mgsb2WCpul/hErOMdxEzWeFhplh/jMReRz6R4L8yxVYcJrYSYPMHE5Mz+ePiJs/zwAXOqn7Fn7uIaITRyBfXFDzApqMPH9qaJn3FIxIgJucBNhTQJgoRNkThWwnCtgJSqB'
  + 'legZrFSMvGHH+CjjLnQCaZo4d4J5GSAvw8MQpY4IBf5WADXVNTzeLfB4R+BJxSh9VMtIKCFmeDL9lmif4VFSZyEyNcGT6QJPLiJnseP0EDM887iR8WniDI+fgffuIDyp4wqeprqCRyk9w5PpAk8uRulVyPBou8CT6ASPMQUeEwo8iUrwJHqGJxUjFtgxPsq4Cz3BkybO'
  + 'ndQMvNLH4MkdW3ja6hoeCAs8EAg8qRilt5DhcWqBJ9EJHi8KPB6KEIlK8CR6hicVI2fYMT7KuAud4Jkmzp1gBh7CQXhSxxU8TXUNT4AFngAEnlRE6bWQCR4t/AxPpid4NPKUOmlVPESmJngyXeDJReQsdoyPMu5CvyVaLvCEGfgAB+FJHVfwNNUVPFovrjnTBZ5cjNLr'
  + '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'
  + 'MDAxNSAwMDAwMCBuIAowMDAwMDAwMDU0IDAwMDAwIG4gCjAwMDAwMDAxMzQgMDAwMDAgbiAKMDAwMDAwMDE4MyAwMDAwMCBuIAowMDAwMDAwMzgyIDAwMDAwIG4gCjAwMDAwMDE3MDcgMDAwMDAgbiAKMDAwMDAwMTk1MiAwMDAwMCBuIAowMDAwMDU4ODQ1IDAwMDAw'
  + 'IG4gCjAwMDAwNTkxMTYgMDAwMDAgbiAKMDAwMDA1OTI0MiAwMDAwMCBuIAowMDAwMDU5NDg4IDAwMDAwIG4gCjAwMDAwNTk2MTUgMDAwMDAgbiAKMDAwMDA2MDA5MyAwMDAwMCBuIAowMDAwMDYwMjkzIDAwMDAwIG4gCjAwMDAwNjAzNDUgMDAwMDAgbiAKMDAwMDA2'
  + 'MDQ5OCAwMDAwMCBuIAowMDAwMDYwOTMwIDAwMDAwIG4gCjAwMDAwNjEzNzkgMDAwMDAgbiAKMDAwMDA2MTYxNiAwMDAwMCBuIAowMDAwMDY0MzE0IDAwMDAwIG4gCjAwMDAwNjQ0NjIgMDAwMDAgbiAKMDAwMDA2NDkwMiAwMDAwMCBuIAowMDAwMDY1MzU0IDAwMDAw'
  + 'IG4gCjAwMDAwNjU1ODYgMDAwMDAgbiAKMDAwMDA2ODQxNSAwMDAwMCBuIAowMDAwMDY4NjE3IDAwMDAwIG4gCjAwMDAwNzA4NjEgMDAwMDAgbiAKMDAwMDA3MTE0MSAwMDAwMCBuIAowMDAwMDcxNTE2IDAwMDAwIG4gCjAwMDAwNzE2NTAgMDAwMDAgbiAKMDAwMDA3'
  + 'MTg1MCAwMDAwMCBuIAowMDAwMDcxOTAyIDAwMDAwIG4gCjAwMDAwNzIwNTUgMDAwMDAgbiAKMDAwMDA3MjY0OSAwMDAwMCBuIAowMDAwMDczMjU1IDAwMDAwIG4gCjAwMDAwNzM0OTMgMDAwMDAgbiAKMDAwMDA3ODM4MiAwMDAwMCBuIAowMDAwMDc4NTMwIDAwMDAw'
  + 'IG4gCjAwMDAwNzkxMzUgMDAwMDAgbiAKMDAwMDA3OTc2NCAwMDAwMCBuIAowMDAwMDc5OTk3IDAwMDAwIG4gCjAwMDAwODUyMzMgMDAwMDAgbiAKMDAwMDA4NTcyNSAwMDAwMCBuIAowMDAwMDg1ODU5IDAwMDAwIG4gCjAwMDAwODYyNzUgMDAwMDAgbiAKMDAwMDA4'
  + 'NjQwOSAwMDAwMCBuIAowMDAwMDg2NjgyIDAwMDAwIG4gCjAwMDAwODY4MDkgMDAwMDAgbiAKMDAwMDA4NzA1NSAwMDAwMCBuIAowMDAwMDg3MTgyIDAwMDAwIG4gCjAwMDAwODc3NjcgMDAwMDAgbiAKMDAwMDA4Nzk2OSAwMDAwMCBuIAowMDAwMDkwMDQ1IDAwMDAw'
  + 'IG4gCjAwMDAwOTAyNDcgMDAwMDAgbiAKdHJhaWxlcgo8PAovU2l6ZSA1NQovUm9vdCAzIDAgUgovSW5mbyAxIDAgUgo+PgpzdGFydHhyZWYKOTI2NjQKJSVFT0YK';

/**
 * Returns the guide PDF as a base64 string so the client can build a
 * downloadable file. Used by the "unlock after booking" flow on the
 * website (the Download button only appears once a consultation has
 * been booked).
 */
function getGuidePdfBase64() {
  return LEADMAGNET_PDF_BASE64;
}

const LEADMAGNET_SHEET_NAME = 'Lead Magnet Leads';

function submitLeadMagnet(data) {
  data = data || {};
  const name = clean_(data.name);
  const email = clean_(data.email);

  if (!name) throw new Error('Please enter your name.');
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) throw new Error('Please enter a valid email address.');

  const timestamp = new Date();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LEADMAGNET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LEADMAGNET_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Email']);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow([timestamp, name, email]);
  } finally {
    lock.releaseLock();
  }

  const pdfBlob = Utilities.newBlob(
    Utilities.base64Decode(LEADMAGNET_PDF_BASE64),
    'application/pdf',
    'Gabay-sa-Pagbili-ng-Bahay.pdf'
  );

  const htmlBody = '<!doctype html><html><body style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#17233a">' +
    '<div style="width:100%;background:#f4f7fb;padding:24px 0">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:560px;margin:0 auto;border-collapse:collapse;background:#ffffff">' +
    '<tr><td style="background:#061b3d;padding:22px 28px;color:#ffffff">' +
      '<div style="font-size:20px;font-weight:900">HOME PARTNERS PH</div>' +
    '</td></tr>' +
    '<tr><td style="padding:28px">' +
      '<div style="font-size:17px;font-weight:800;color:#0b2148;margin-bottom:14px">Hi ' + escapeHtml_(name.split(' ')[0]) + ',</div>' +
      '<div style="font-size:14px;line-height:1.8;color:#34445b">Narito ang libreng gabay na hiniling mo: <strong>Gabay sa Matalinong Pagbili ng Bahay</strong>. Nakalakip ito bilang PDF sa email na ito. Kung may tanong ka pagkatapos mabasa, huwag mag-atubiling mag-reply o tumawag sa amin.</div>' +
      '<div style="margin-top:22px">' +
        '<a href="tel:' + escapeHtml_(CONFIG.PHONE.replace(/\s/g,'')) + '" style="display:inline-block;background:#061b3d;color:#ffffff;text-decoration:none;padding:12px 24px;border-radius:7px;font-size:13px;font-weight:800">Tawagan Kami</a>' +
      '</div>' +
    '</td></tr>' +
    '<tr><td style="background:#061b3d;padding:16px 24px;color:#ffffff;font-size:11px;text-align:center">' + escapeHtml_(CONFIG.PHONE) + ' &nbsp;|&nbsp; ' + escapeHtml_(CONFIG.AGENT_EMAIL) + '</td></tr>' +
    '</table></div></body></html>';

  MailApp.sendEmail({
    to: email,
    subject: 'Narito ang iyong libreng gabay   Home Partners PH',
    htmlBody: htmlBody,
    body: stripHtml_(htmlBody),
    attachments: [pdfBlob],
    replyTo: CONFIG.AGENT_EMAIL,
    name: CONFIG.BUSINESS_NAME
  });

  return {
    success: true,
    message: 'Naipadala na! Tingnan ang iyong email (kasama ang spam folder kung wala sa inbox).'
  };
}
