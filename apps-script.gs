// ESG Parathedstesten — Google Apps Script
// Kopiér hele denne fil ind i dit Apps Script-projekt og deploy den som ny version.
//
// Modtager GET-parametre fra index.html og:
//  1. Gemmer data i Google Sheets (uændret)
//  2. Sender rapport-email til respondenten
//  3. Sender notifikations-email til kamel@czoo.dk

const ADMIN_EMAIL    = 'kamel@czoo.dk';
const SHEET_NAME     = 'Besvarelser'; // Tilpas hvis dit sheet hedder noget andet
const BRAND_GREEN    = '#56C183';
const BRAND_DARK     = '#0F0F0F';

// ============================================================
// HJÆLPEFUNKTIONER — scorer og tekster (spejler index.html)
// ============================================================

function getLevel(score) {
  if (score <= 25) return { label: 'Tidlig fase',  color: '#fe6c7a' };
  if (score <= 50) return { label: 'På vej',        color: '#feed48' };
  if (score <= 75) return { label: 'Struktureret',  color: '#ebebeb' };
  return             { label: 'Avanceret',           color: '#314ae2' };
}

function getDesc(score) {
  if (score <= 25) return 'ESG-arbejdet er endnu ikke struktureret i jeres virksomhed. I har et godt udgangspunkt for at opbygge et solidt fundament fra bunden.';
  if (score <= 50) return 'I har taget de første skridt med ESG. Nu handler det om at systematisere indsatsen og skabe sammenhæng på tværs af dimensionerne.';
  if (score <= 75) return 'I arbejder aktivt og struktureret med ESG. Der er basis for at løfte indsatsen til næste niveau og integrere ESG endnu dybere i organisationen.';
  return 'ESG er dybt integreret i jeres organisation. I er godt positioneret til at styrke jeres position og kommunikere ESG med høj troværdighed.';
}

function getDimLevel(s) {
  if (s <= 10) return { label: 'Ikke påbegyndt',  color: '#888888' };
  if (s <= 15) return { label: 'Tidlig fase',      color: '#fe6c7a' };
  if (s <= 20) return { label: 'Under udvikling',  color: '#feed48' };
  return             { label: 'Modent niveau',      color: BRAND_GREEN };
}

function getDimText(dimIdx, score) {
  var names = ['ESG parat', 'Kommunikationsparat', 'Organisationsparat', 'Procesparat'];
  var n = names[dimIdx].toLowerCase();
  if (score <= 10) return 'Der er endnu ikke sat en struktureret ramme op for hvordan din organisation bliver ' + n + '. Dette er et oplagt startpunkt for at skabe fundamentet.';
  if (score <= 15) return 'I er begyndt at forberede jer på at blive ' + n + ', men indsatsen er stadig fragmenteret og mangler en samlet struktur.';
  if (score <= 20) return 'I arbejder aktivt med at blive ' + n + ', men der er stadig vigtige elementer som mangler at blive implementeret fuldt ud.';
  return 'Når det handler om at blive ' + n + ', er jeres virksomhed kommet langt. Derfor har I et stærkt fundament at bygge videre på.';
}

var DIM_RECS = [
  // Dim 0 — ESG parat
  [
    ['Kortlæg virksomhedens væsentligste miljø- og sociale påvirkninger som første skridt', 'Udpeg en ESG-ansvarlig og definer 3 konkrete startmål'],
    ['Formalisér jeres ESG-program med målbare indsatser baseret på væsentlighed', 'Start systematisk dataindsamling på jeres vigtigste ESG-parametre'],
    ['Styrk jeres ESG-datafundament og dokumentation til ekstern brug', 'Integrer ESG-mål i den løbende forretningsstrategi og budgettering'],
    ['Overvej ekstern validering eller certificering af jeres ESG-indsats', 'Del best practices med leverandører og branchepartnere om ESG'],
  ],
  // Dim 1 — Kommunikationsparat
  [
    ['Definer en grundlæggende ESG-kommunikationsstrategi målrettet kunder og investorer', 'Identificér jeres vigtigste stakeholders og kortlæg deres ESG-forventninger'],
    ['Strukturér ESG-kommunikationen og gør den konsistent på tværs af kanaler', 'Udarbejd et simpelt ESG-resume til brug i salgs- og kundedialoger'],
    ['Løft kvaliteten af jeres ESG-rapportering og brug anerkendte rammer som GRI eller ESRS', 'Træn salg og kundeservice i at kommunikere ESG-værdier troværdigt'],
    ['Positionér jer aktivt som ESG-foregangsvirksomhed i jeres branche', 'Integrer ESG-historier aktivt i brand- og marketingstrategien'],
  ],
  // Dim 2 — Organisationsparat
  [
    ['Sikr ledelsens formelle opbakning til ESG som strategisk prioritet', 'Forankr ESG-ansvar i organisationen med klare roller og ejerskab'],
    ['Involvér medarbejderne bredt og skab fælles forståelse for ESG-dagsordenen', 'Opbyg intern ESG-kompetence gennem uddannelse og videndeling'],
    ['Integrer ESG-hensyn i rekruttering, kultur og medarbejderudvikling', 'Etablér tværgående ESG-processer på tværs af afdelinger og funktioner'],
    ['Styrk ESG-governance med klare beslutningsstrukturer og bestyrelsesinvolvering', 'Benchmark jeres organisation mod branchens bedste ESG-praksis'],
  ],
  // Dim 3 — Procesparat
  [
    ['Kortlæg jeres forsyningskæde og identificér de væsentligste ESG-risici', 'Begynd at dokumentere de mest kritiske processer ud fra et ESG-perspektiv'],
    ['Implementér et simpelt system til ESG-dataindsamling og -opfølgning', 'Sæt ESG-krav ind i indkøbs- og leverandørpolitikker'],
    ['Automatisér ESG-dataindsamling og -rapportering for at reducere manuel indsats', 'Integrer ESG-due diligence i eksisterende risikostyringsprocesser'],
    ['Certificér relevante processer (ISO, GRI, ESRS) for øget troværdighed', 'Del proceserfaring og data med leverandører for at løfte hele værdikæden'],
  ],
];

function getSteps(total, dimScores) {
  var title = total <= 25 ? 'Byg fundamentet'
            : total <= 50 ? 'Strukturér og systematisér'
            : total <= 75 ? 'Løft til næste niveau'
            : 'Stå frem som ESG-frontløber';

  // Find de 2 lavest-scorende dimensioner
  var indexed = dimScores.map(function(s, i) { return { i: i, s: s }; });
  indexed.sort(function(a, b) { return a.s - b.s; });
  var weakest = indexed.slice(0, 2);

  var steps = [];
  weakest.forEach(function(item) {
    var lvlIdx = item.s <= 10 ? 0 : item.s <= 15 ? 1 : item.s <= 20 ? 2 : 3;
    DIM_RECS[item.i][lvlIdx].forEach(function(rec) { steps.push(rec); });
  });

  var dimNames = ['ESG parat', 'Kommunikationsparat', 'Organisationsparat', 'Procesparat'];
  var weakNames = weakest.map(function(item) { return dimNames[item.i]; });
  return { title: title, steps: steps, weakNames: weakNames };
}

// ============================================================
// HTML EMAIL BUILDER
// ============================================================

function buildProgressBar(pct, color) {
  var filled = Math.round(pct);
  return '<table width="100%" cellpadding="0" cellspacing="0" border="0">'
    + '<tr><td style="background:#2a2a2a;border-radius:4px;height:8px;padding:0;">'
    + '<table width="' + filled + '%" cellpadding="0" cellspacing="0" border="0">'
    + '<tr><td style="background:' + color + ';border-radius:4px;height:8px;font-size:0;">&nbsp;</td></tr>'
    + '</table>'
    + '</td></tr></table>';
}

function buildReportEmailHtml(data) {
  var dimScores  = [+data.dim0, +data.dim1, +data.dim2, +data.dim3];
  var total      = +data.total;
  var lvl        = getLevel(total);
  var desc       = getDesc(total);
  var steps      = getSteps(total, dimScores);
  var dimNames   = ['ESG parat', 'Kommunikationsparat', 'Organisationsparat', 'Procesparat'];

  // Dimensions HTML
  var dimsHtml = '';
  dimScores.forEach(function(s, i) {
    var dl  = getDimLevel(s);
    var txt = getDimText(i, s);
    var pct = Math.round((s / 25) * 100);
    dimsHtml += '<tr><td style="padding:0 0 16px 0;">'
      + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;">'
      + '<tr><td style="padding:24px;">'
      + '<table width="100%" cellpadding="0" cellspacing="0" border="0">'
      + '<tr>'
      + '<td><span style="font-size:11px;letter-spacing:0.08em;text-transform:uppercase;color:#666;display:block;margin-bottom:4px;">Dimension ' + (i+1) + '</span>'
      + '<strong style="font-size:18px;color:#efeeea;">' + dimNames[i] + '</strong></td>'
      + '<td align="right" style="white-space:nowrap;">'
      + '<span style="font-size:32px;font-weight:700;color:' + dl.color + ';">' + s + '</span>'
      + '<span style="font-size:14px;color:#666;">/25</span>'
      + '<div style="font-size:12px;color:' + dl.color + ';margin-top:2px;">' + dl.label + '</div>'
      + '</td>'
      + '</tr>'
      + '</table>'
      + '<div style="margin:16px 0;">' + buildProgressBar(pct, dl.color) + '</div>'
      + '<p style="font-size:14px;color:#999;line-height:1.7;margin:0;">' + txt + '</p>'
      + '</td></tr></table>'
      + '</td></tr>';
  });

  // Steps HTML
  var stepsHtml = '';
  steps.steps.forEach(function(step, i) {
    stepsHtml += '<tr><td style="padding:0 0 12px 0;">'
      + '<table cellpadding="0" cellspacing="0" border="0">'
      + '<tr>'
      + '<td style="width:28px;height:28px;border:1px solid #2a2a2a;text-align:center;vertical-align:middle;font-size:11px;color:#666;font-family:monospace;">' + (i+1) + '</td>'
      + '<td style="padding-left:12px;font-size:14px;color:#aaa;line-height:1.6;">' + step + '</td>'
      + '</tr>'
      + '</table>'
      + '</td></tr>';
  });

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#0F0F0F;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">'

    // Wrapper
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0F0F0F;">'
    + '<tr><td align="center" style="padding:40px 16px;">'
    + '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;">'

    // Header / logo area
    + '<tr><td style="padding-bottom:40px;border-bottom:1px solid #2a2a2a;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
    + '<td><span style="font-size:13px;color:#666;letter-spacing:0.06em;text-transform:uppercase;">Creative ZOO</span></td>'
    + '<td align="right"><span style="font-size:13px;color:' + BRAND_GREEN + ';letter-spacing:0.06em;text-transform:uppercase;">ESG Parathedstesten</span></td>'
    + '</tr></table>'
    + '</td></tr>'

    // Greeting
    + '<tr><td style="padding:40px 0 32px 0;">'
    + '<p style="font-size:16px;color:#999;line-height:1.7;margin:0 0 16px 0;">Hej ' + escHtml(data.name) + ',</p>'
    + '<p style="font-size:16px;color:#999;line-height:1.7;margin:0;">Tak fordi du tog ESG Parathedstesten. Herunder finder du din personlige rapport med resultater og konkrete anbefalinger til jeres næste skridt.</p>'
    + '</td></tr>'

    // Score box
    + '<tr><td style="padding-bottom:32px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;">'
    + '<tr><td style="padding:32px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
    + '<td>'
    + '<div style="font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#666;margin-bottom:8px;">Samlet score</div>'
    + '<div style="font-size:72px;font-weight:700;color:#efeeea;line-height:1;">' + total + '</div>'
    + '<div style="font-size:16px;color:#666;margin-top:4px;">/100 point</div>'
    + '</td>'
    + '<td align="right" style="vertical-align:top;">'
    + '<div style="display:inline-block;background:' + lvl.color + '22;border:1px solid ' + lvl.color + '66;border-radius:4px;padding:8px 16px;">'
    + '<span style="font-size:14px;font-weight:600;color:' + lvl.color + ';">' + lvl.label + '</span>'
    + '</div>'
    + '</td>'
    + '</tr></table>'
    + '<p style="font-size:14px;color:#999;line-height:1.7;margin:24px 0 0 0;">' + desc + '</p>'
    + '</td></tr></table>'
    + '</td></tr>'

    // Dimensions
    + '<tr><td style="padding-bottom:8px;">'
    + '<h2 style="font-size:20px;color:#efeeea;margin:0 0 24px 0;">Resultater pr. dimension</h2>'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0">' + dimsHtml + '</table>'
    + '</td></tr>'

    // Next steps
    + '<tr><td style="padding:32px 0;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;">'
    + '<tr><td style="padding:32px;">'
    + '<h2 style="font-size:20px;color:#efeeea;margin:0 0 8px 0;">' + steps.title + '</h2>'
    + '<p style="font-size:13px;color:#666;margin:0 0 24px 0;">Baseret på jeres to svageste områder: <strong style="color:#999;">' + steps.weakNames.join(' og ') + '</strong></p>'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0">' + stepsHtml + '</table>'
    + '</td></tr></table>'
    + '</td></tr>'

    // Kontakt
    + '<tr><td style="padding:0 0 40px 0;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;"><tr><td style="padding:28px 32px;">'
    + '<p style="font-size:13px;letter-spacing:0.08em;text-transform:uppercase;color:#666;margin:0 0 16px 0;">Har du spørgsmål til rapporten?</p>'
    + '<table cellpadding="0" cellspacing="0" border="0"><tr>'
    + '<td style="padding-right:20px;vertical-align:top;">'
    + '<div style="font-size:16px;font-weight:600;color:#efeeea;margin-bottom:2px;">Morten Søholm</div>'
    + '<div style="font-size:13px;color:#666;margin-bottom:12px;">Strategic Advisor, Senior Partner</div>'
    + '<a href="tel:+4560663060" style="display:block;font-size:14px;color:#999;text-decoration:none;margin-bottom:4px;">+45 60 66 30 60</a>'
    + '<a href="mailto:ms@czoo.dk" style="display:block;font-size:14px;color:' + BRAND_GREEN + ';text-decoration:none;">ms@czoo.dk</a>'
    + '</td>'
    + '</tr></table>'
    + '</td></tr></table>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="border-top:1px solid #2a2a2a;padding-top:32px;">'
    + '<p style="font-size:12px;color:#444;text-align:center;margin:0;line-height:1.7;">Creative ZOO · czoo.dk · <a href="mailto:hej@czoo.dk" style="color:#444;">hej@czoo.dk</a></p>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';

  return html;
}

function buildAdminEmailHtml(data) {
  var dimScores = [+data.dim0, +data.dim1, +data.dim2, +data.dim3];
  var total     = +data.total;
  var lvl       = getLevel(total);
  var dimNames  = ['ESG parat', 'Kommunikationsparat', 'Organisationsparat', 'Procesparat'];

  // Kontaktdata tabel
  var rows = [
    ['Navn',       data.name  || '—'],
    ['Titel',      data.title || '—'],
    ['Email',      data.email || '—'],
    ['Telefon',    data.phone || '—'],
    ['Nyhedsbrev', data.consent_newsletter === '1' ? 'Ja' : 'Nej'],
  ];
  var contactHtml = '';
  rows.forEach(function(row) {
    contactHtml += '<tr>'
      + '<td style="padding:10px 16px;font-size:13px;color:#666;white-space:nowrap;border-bottom:1px solid #2a2a2a;">' + row[0] + '</td>'
      + '<td style="padding:10px 16px;font-size:13px;color:#efeeea;border-bottom:1px solid #2a2a2a;">' + escHtml(row[1]) + '</td>'
      + '</tr>';
  });

  // Dimensions
  var dimsHtml = '';
  dimScores.forEach(function(s, i) {
    var dl  = getDimLevel(s);
    var pct = Math.round((s / 25) * 100);
    dimsHtml += '<tr><td style="padding:0 0 12px 0;">'
      + '<table width="100%" cellpadding="0" cellspacing="0" border="0">'
      + '<tr>'
      + '<td style="font-size:13px;color:#999;">' + dimNames[i] + '</td>'
      + '<td align="right" style="white-space:nowrap;padding-left:12px;">'
      + '<strong style="color:' + dl.color + ';">' + s + '/25</strong>'
      + '<span style="font-size:11px;color:' + dl.color + ';margin-left:8px;">' + dl.label + '</span>'
      + '</td>'
      + '</tr>'
      + '<tr><td colspan="2" style="padding-top:6px;">' + buildProgressBar(pct, dl.color) + '</td></tr>'
      + '</table>'
      + '</td></tr>';
  });

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#0F0F0F;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0F0F0F;">'
    + '<tr><td align="center" style="padding:40px 16px;">'
    + '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;">'

    // Header
    + '<tr><td style="padding-bottom:32px;border-bottom:1px solid #2a2a2a;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
    + '<td><span style="font-size:13px;color:#666;letter-spacing:0.06em;text-transform:uppercase;">Creative ZOO — Intern notifikation</span></td>'
    + '<td align="right"><span style="font-size:13px;color:' + BRAND_GREEN + ';letter-spacing:0.06em;text-transform:uppercase;">ESG Parathedstesten</span></td>'
    + '</tr></table>'
    + '</td></tr>'

    // Intro
    + '<tr><td style="padding:32px 0 24px 0;">'
    + '<h2 style="font-size:22px;color:#efeeea;margin:0 0 8px 0;">Ny besvarelse fra ' + escHtml(data.name) + '</h2>'
    + '<p style="font-size:14px;color:#666;margin:0;">Herunder ses kontaktdata og testresultater.</p>'
    + '</td></tr>'

    // Kontaktdata
    + '<tr><td style="padding-bottom:32px;">'
    + '<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.08em;color:#666;margin:0 0 16px 0;">Kontaktdata</h3>'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;overflow:hidden;">'
    + contactHtml
    + '</table>'
    + '</td></tr>'

    // Score
    + '<tr><td style="padding-bottom:32px;">'
    + '<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.08em;color:#666;margin:0 0 16px 0;">Samlet score</h3>'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;"><tr><td style="padding:24px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
    + '<td><span style="font-size:48px;font-weight:700;color:#efeeea;">' + total + '</span><span style="font-size:16px;color:#666;">/100</span></td>'
    + '<td align="right"><div style="background:' + lvl.color + '22;border:1px solid ' + lvl.color + '66;border-radius:4px;padding:8px 16px;display:inline-block;"><span style="font-size:14px;font-weight:600;color:' + lvl.color + ';">' + lvl.label + '</span></div></td>'
    + '</tr></table>'
    + '</td></tr></table>'
    + '</td></tr>'

    // Dimensioner
    + '<tr><td style="padding-bottom:40px;">'
    + '<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.08em;color:#666;margin:0 0 16px 0;">Dimensioner</h3>'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;"><tr><td style="padding:24px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" border="0">' + dimsHtml + '</table>'
    + '</td></tr></table>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="border-top:1px solid #2a2a2a;padding-top:24px;">'
    + '<p style="font-size:12px;color:#444;text-align:center;margin:0;">Creative ZOO · Intern — må ikke videresende</p>'
    + '</td></tr>'

    + '</table></td></tr></table></body></html>';

  return html;
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================
// SHEETS — gem data
// ============================================================

function getExpectedHeaders() {
  var headers = ['Tidsstempel', 'Navn', 'Email', 'Titel', 'Telefon',
                 'Samtykke kontakt', 'Nyhedsbrev',
                 'Total', 'Niveau',
                 'ESG parat', 'Kommunikationsparat', 'Organisationsparat', 'Procesparat'];
  var questions = [
    // Dim 0 — ESG parat
    'Har I kortlagt virksomhedens væsentligste aftryk (miljø-, klima- og sociale forhold)?',
    'Har I et ESG-program med konkrete og målbare indsatser baseret på væsentlighed?',
    "Har I fastlagt mål og KPI'er, så I kan følge jeres udvikling over tid?",
    'Har I pålidelige ESG-data, som kan dokumenteres og forklares eksternt?',
    'Udgiver eller kan I udarbejde en ESG-rapport eller -oversigt baseret på anerkendte standarder/rammer?',
    // Dim 1 — Kommunikationsparat
    'Er I tilfredse med udbyttet af jeres ESG-kommunikation i dag?',
    'Har I kortlagt jeres vigtigste målgrupper og deres forskellige behov for indsigt i jeres ESG-arbejde?',
    'Har I fastlagt, hvad ESG-kommunikation skal bidrage med (fx salg, tillid, employer branding, relationer)?',
    'Har I udviklet klare storylines eller hovedbudskaber for jeres væsentligste ESG-emner?',
    'Har I udpeget og trænet relevante talspersoner til at kommunikere om jeres ESG-indsats?',
    // Dim 2 — Organisationsparat
    'Har I strukturerede dialoger om, hvordan ESG-arbejdet kan anvendes på tværs af organisationen?',
    'Har I et fast og velfungerende samarbejde mellem ESG, kommunikation og marketing?',
    'Har virksomheden en fælles ambition for, hvordan ESG skal bidrage til jeres position i markedet?',
    'Er relevante funktioner (salg, ledelse, indkøb, marketing, kommunikation) klædt på til at forstå og anvende ESG i dialoger med eksterne?',
    'Oplever I, at ESG-arbejdet skaber værdi og ejerskab i flere dele af organisationen – og ikke kun i ESG-teamet?',
    // Dim 3 — Procesparat
    'Har I klare processer for, hvordan ESG-budskaber og -påstande bliver udviklet og godkendt?',
    'Er roller og ansvar tydeligt defineret, når der kommunikeres om ESG (hvem ejer indhold, data og godkendelse)?',
    'Har I overblik over risici og krisepotentiale i jeres væsentligste ESG-emner?',
    'Har I retningslinjer for, hvad I må og ikke må sige om ESG (fx ift. dokumentation og gennemsigtighed)?',
    'Har I faste arbejdsgange, der sikrer sammenhæng mellem ESG-data, kommunikation og markedsføring på tværs af kanaler?',
  ];
  questions.forEach(function(q) { headers.push(q); });
  headers.push('Kommentar');
  return headers; // 34 kolonner i alt
}

// Tvangsret header-række til at matche de forventede kolonner.
// Kalder som: ?action=resetHeaders
// Skyder IKKE eksisterende datarækker bort — indsætter/overskriver kun række 1.
function resetHeaders() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = getExpectedHeaders();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    // Overskriver række 1 med korrekte headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Slet eventuelle ekstra kolonner til højre hvis arket er bredere end forventet
    var currentCols = sheet.getLastColumn();
    if (currentCols > headers.length) {
      sheet.deleteColumns(headers.length + 1, currentCols - headers.length);
    }
  }
}

function saveToSheet(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Opret header-række hvis arket er tomt, eller header-kolonneantal er forkert
    var expectedHeaders = getExpectedHeaders();
    var needsHeader = sheet.getLastRow() === 0 ||
                      sheet.getLastColumn() !== expectedHeaders.length ||
                      sheet.getRange(1, 1).getValue() !== 'Tidsstempel';
    if (needsHeader) {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(expectedHeaders);
      } else {
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      }
    }

    var row = [
      new Date(),
      data.name  || '',
      data.email || '',
      data.title || '',
      data.phone || '',
      data.consent_contact    === '1' ? 'Ja' : 'Nej',
      data.consent_newsletter === '1' ? 'Ja' : 'Nej',
      data.total || '',
      data.level || '',
      data.dim0  || '',
      data.dim1  || '',
      data.dim2  || '',
      data.dim3  || '',
    ];
    for (var i = 1; i <= 20; i++) row.push(data['q' + i] || '');
    row.push(data.comment || '');
    sheet.appendRow(row);
  } catch (e) {
    Logger.log('Sheet error: ' + e);
  }
}

// ============================================================
// EMAILS — send
// ============================================================

function sendEmails(data) {
  // Email til respondenten
  if (data.email) {
    var reportHtml = buildReportEmailHtml(data);
    MailApp.sendEmail({
      to:       data.email,
      subject:  'ESG Parat: Din rapport',
      htmlBody: reportHtml,
      name:     'Creative ZOO — ESG Parathedstesten',
      replyTo:  'hej@czoo.dk',
    });
  }

  // Email til admin
  var adminHtml = buildAdminEmailHtml(data);
  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  'Ny besvarelse fra ESG Parat: ' + (data.name || '(uden navn)'),
    htmlBody: adminHtml,
    name:     'ESG Parathedstesten',
  });
}

// ============================================================
// DATA ENDPOINT — returnér alle besvarelser som JSON
// ============================================================

function getAllData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    headers.forEach(function(h, idx) {
      var v = values[i][idx];
      // Konvertér Date-objekter til ISO-streng
      row[h] = v instanceof Date ? v.toISOString() : (v !== null && v !== undefined ? String(v) : '');
    });
    rows.push(row);
  }
  return rows;
}

// ============================================================
// MAIN — GET handler
// ============================================================

function doGet(e) {
  var params = e.parameter;

  // Dashboard data endpoint
  if (params.action === 'getData') {
    var rows = getAllData();
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Ret header-rækken i Google Sheet (kald én gang for at fikse CSV-mismatch)
  if (params.action === 'resetHeaders') {
    resetHeaders();
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Headers opdateret' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Normal besvarelse fra quiz
  saveToSheet(params);

  try {
    sendEmails(params);
  } catch (err) {
    Logger.log('Email error: ' + err);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
