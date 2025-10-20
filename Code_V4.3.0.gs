/***** Exam Builder - V4.3.0 (Full Replacement)
 * New in V4.3.0
 *  - Dynamic multi-database support (any number of Notion DBs) via NOTION_DB_MAP (Script Properties, JSON).
 *  - Sync Down picker & partial filters (course, Ref ID prefix, Topic/Chapter contains, Tags, Last Edited Since).
 *  - No new HTML files required: filter dialog is rendered from inline HTML in Code.gs.
 *  - Keeps all prior features: Title/Subtitle, grouped-by-type, optional Difficulty Summary, Import auto-clear,
 *    order-agnostic 18-column schema incl. Tags, robust Topic/Chapter import from various Notion property types.
 *
 * Menus:
 *  Exam Builder V4 -> Open Exam Builder
 *  Notion Sync -> Configure... / Preview Changes / Sync Down... / Sync Up (All rows in Import)
 *  Bank Tools -> Merge / Upsert / Validate / Clear Highlights / Setup Last Practiced
 */

// =========================
// Globals & Constants
// =========================
const SHEET_BANK   = 'Bank';
const SHEET_IMPORT = 'Import';
const REF_HEADER   = 'Ref ID';

const TYPE_ORDER   = ['mcq','true/false','short answer','long answer','fill in spaces','matching'];
const DIFF_ORDER   = ['Very Easy', 'Easy', 'Medium', 'Hard', 'Very Hard'];

const PROP = PropertiesService.getScriptProperties();
const NOTION_API_BASE = 'https://api.notion.com/v1';
const NOTION_VERSION  = '2022-06-28'; // stable & widely supported
const DBMAP_KEY       = 'NOTION_DB_MAP'; // JSON: { "BIO 1130":"dbid", "ANP 1111":"dbid", ... }
const NOTION_RATE_LIMIT_MS = 400; // keep comfortably under Notion's 3 req/sec guidance

const notionDbMetadataCache = {};

// =========================
// Menus
// =========================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Exam Builder V4')
    .addItem('Open Exam Builder', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Notion Sync')
      .addItem('Configure...', 'openNotionConfig')
      .addItem('Preview Changes', 'notionPreview')
      .addSeparator()
      .addItem('Sync Down...', 'openSyncDownFilters') // NEW picker & partial filters
      .addItem('Sync Up (Import -> Notion)', 'notionSyncUp')
    )
    .addToUi();

  ui.createMenu('Bank Tools')
    .addItem('Merge Import -> Bank (append new)','mergeImportToBank')
    .addItem('Upsert Import -> Bank (update by Ref ID)','upsertImportToBank')
    .addSeparator()
    .addItem('Validate Bank Data', 'validateBankData')
    .addItem('Clear Validation Highlighting', 'clearValidationHighlighting')
    .addSeparator()
    .addItem('Setup "Last Practiced" Column', 'setupLastPracticedColumn')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Exam Builder V4.3.0')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =========================
// Sidebar data
// =========================
function getSidebarData() {
  const bankSheet = getSheet(SHEET_BANK);
  if (!bankSheet) return { courses: [], difficulties: [], types: [], correctStatuses: [] };

  const bankData = bankSheet.getDataRange().getValues();
  if (bankData.length < 2) return { courses: [], difficulties: [], types: [], correctStatuses: [] };

  const headers = bankData[0].map(h => toLower(clean(h)));
  const getUniqueValuesFromColumn = (name) => {
    const idx = headers.indexOf(name);
    if (idx === -1) return [];
    const uniq = [...new Set(bankData.slice(1).map(r => clean(r[idx])).filter(Boolean))];
    return uniq.sort();
  };

  return {
    courses: getUniqueValuesFromColumn('course'),
    difficulties: DIFF_ORDER,
    types: getUniqueValuesFromColumn('question type'),
    correctStatuses: getUniqueValuesFromColumn('correct/incorrect')
  };
}

// =========================
// Generation API
// =========================
function buildExamFromUI(config) {
  const prepared = prepareSelectionForBuild(config);
  const selected = prepared.selected;
  const renderData = prepared.renderData;
  const showTopicAtEnd = prepared.showTopicAtEnd;
  const course = prepared.course;
  const rawTitle = clean(prepared.examTitle);
  const rawSubtitle = clean(prepared.examSubtitle);

  if (selected.length === 0) return 'Quota settings resulted in 0 questions. Try adjusting quotas or filters.';

  const courseName = clean(course) || 'Course';
  const finalTitle = rawTitle ? `${courseName} - ${rawTitle}` : `${courseName} - Practice Exam`;

  const doc = DocumentApp.create(finalTitle);
  const body = doc.getBody();
  body.clear();

  // Title
  styleParagraph(body.appendParagraph(finalTitle), { size: 16, bold: true, center: true, family: 'Arial' });
  if (rawSubtitle) {
    styleParagraph(body.appendParagraph(rawSubtitle), { size: 12, italic: true, center: true, family: 'Arial' });
  }
  body.appendParagraph('');

  const mode = config.mode || 'bothOneDoc';
  if (mode === 'studentOnly') {
    renderStudent(body, renderData, clean(config.searchText), showTopicAtEnd);
  } else if (mode === 'answerOnly') {
    styleParagraph(body.appendParagraph('Answer Key'), { fontFamily: 'Georgia', size: 14, bold: true });
    body.appendParagraph('');
    renderAnswers(body, renderData, clean(config.searchText), showTopicAtEnd);
  } else {
    renderStudent(body, renderData, clean(config.searchText), showTopicAtEnd);
    body.appendPageBreak();
    styleParagraph(body.appendParagraph('Answer Key'), { fontFamily: 'Georgia', size: 14, bold: true });
    body.appendParagraph('');
    renderAnswers(body, renderData, clean(config.searchText), showTopicAtEnd);
  }

  // Optional Difficulty Summary Page
  const includeSummary = String(config.includeSummary) === 'true';
  if (includeSummary) {
    body.appendPageBreak();
    styleParagraph(body.appendParagraph('Difficulty Summary'), { size: 14, bold: true, family: 'Arial' });
    body.appendParagraph('');

    const counts = { 'Very Easy':0, 'Easy':0, 'Medium':0, 'Hard':0, 'Very Hard':0, '':0 };
    renderData.forEach(q => {
      const d = clean(q.difficulty);
      if (counts.hasOwnProperty(d)) counts[d]++; else counts['']++;
    });
    const total = renderData.length || 1;
    const rows = [['Difficulty','Count','Percent']];
    ['Very Easy','Easy','Medium','Hard','Very Hard'].forEach(d=>{
      const c = counts[d]||0;
      const pct = Math.round((c*10000)/total)/100;
      rows.push([d, String(c), pct+'%']);
    });

    const table = body.appendTable(rows);
    table.setBorderWidth(0);
    for (let i=0; i<table.getNumRows(); i++) {
      const r = table.getRow(i);
      for (let j=0; j<r.getNumCells(); j++) {
        const cell = r.getCell(j);
        const p = cell.getChild(0).asParagraph();
        styleParagraph(p, { family:'Arial', size: 11, bold: i===0 });
      }
    }
  }

  doc.saveAndClose();
  showDocLink(doc.getUrl(), 'Your document is ready.');
  return 'Exam Created Successfully!';
}

function previewCountFromUI(config) {
  const { matched, selected } = computeCountsOnly(config);
  if (matched === 0) return 'Matched: 0';
  return `Matched: ${matched}   |   Selected (after quotas & cap): ${selected}`;
}

// =========================
// Selection utilities
// =========================
function prepareSelectionForBuild(config) {
  const {
    selectedQuestions,
    orderedQuestions,
    course, examTitle, examSubtitle,
    shuffleChoices, showTopicAtEnd
  } = filterAndSelect(config);

  const renderData = orderedQuestions.map(q => ({
    QT: q.QT_raw,
    stem: stripLeadingNum(q.Q),
    choices: [],
    ansRaw: q.ANS,
    ansMapped: null,
    expl: q.EXPL || '',
    refId: q.refId,
    topics: q.topics,
    difficulty: q.Difficulty || ''
  }));

  for (let i = 0; i < renderData.length; i++) {
    const item = renderData[i];
    const isMCQ = (item.QT || '').toLowerCase() === 'mcq';
    if (!isMCQ) continue;
    const source = orderedQuestions.find(x => x.refId === item.refId);
    let choices = [
      {label:'A', text: stripChoicePrefix(source.A)},
      {label:'B', text: stripChoicePrefix(source.B)},
      {label:'C', text: stripChoicePrefix(source.C)},
      {label:'D', text: stripChoicePrefix(source.D)},
      {label:'E', text: stripChoicePrefix(source.E)}
    ].filter(ch => clean(ch.text) !== '');
    if (String(shuffleChoices) === 'true' && choices.length > 1) {
      randShuffle(choices);
    }

    const ansTokens = (source.ANS || '').split(/[,;]\s*/).map(x => x.trim().toUpperCase()).filter(Boolean);
    let mapped = [];
    ansTokens.forEach(tok => {
      const idx = choices.findIndex(ch => ch.label === tok);
      if (idx >= 0) mapped.push(['A','B','C','D','E'][idx]);
    });

    item.choices = choices.map((ch, idx) => ({ shown: ['a','b','c','d','e'][idx], text: ch.text }));
    item.ansMapped = mapped.length ? mapped.join('; ') : source.ANS;
  }

  return { selected: selectedQuestions, ordered: orderedQuestions, renderData, showTopicAtEnd, course, examTitle, examSubtitle };
}

function computeCountsOnly(config) {
  const { selectedQuestions, matchedCount } = filterAndSelect(config, true);
  return { matched: matchedCount, selected: selectedQuestions.length };
}

function filterAndSelect(config, countsOnly) {
  const bankSheet = getSheet(SHEET_BANK);
  if (!bankSheet) return { selectedQuestions: [], orderedQuestions: [], matchedCount: 0 };

  const {
    course, examTitle, examSubtitle, shuffleChoices, maxQ,
    topicsRaw, excludeTopicsRaw,
    diffList = [], typeList = [], corrList = [],
    myAnsList = [], sourceRaw, excludeSourceRaw, searchText,
    practicedWithinDays, notPracticedDays, isPracticeEmptyStr,
    typeQuotas = {}, diffQuotas = {},
    showTopicTag
  } = config;

  const topicsList        = parseList(topicsRaw).map(toLower);
  const excludeTopicsList = parseList(excludeTopicsRaw).map(toLower);
  const sourceList        = parseList(sourceRaw).map(toLower);
  const excludeSourceList = parseList(excludeSourceRaw).map(toLower);
  const maxQuestions      = toIntOrZero(maxQ);
  const isPracticeEmpty   = (String(isPracticeEmptyStr) === 'true');
  const cleanSearchText   = clean(searchText);
  const showTopicAtEnd    = (String(showTopicTag) === 'true');

  const diffSet = new Set(diffList.map(v => toLower(v)));
  const typeSet = new Set(typeList.map(v => toLower(v)));
  const corrSet = new Set(corrList.map(v => toLower(v)));
  const myAnsSet = new Set(myAnsList.map(v => clean(v)));

  const practicedWithinDate = practicedWithinDays ? daysAgo(toIntOrZero(practicedWithinDays)) : null;
  const notPracticedDate    = notPracticedDays ? daysAgo(toIntOrZero(notPracticedDays)) : null;

  const bankData = bankSheet.getDataRange().getValues();
  const headers = bankData[0].map(h => toLower(clean(h)));
  function col(name){ return headers.indexOf(name.toLowerCase()); }
  function colByAny(possibleNames, headersLower) {
    for (const name of possibleNames) {
      const i = headersLower.indexOf(toLower(name));
      if (i !== -1) return i;
    }
    return -1;
  }

  const cQ=col('question'), cT=col('question type'), cCourse=col('course'), cTopic=col('topic/chapter');
  const cDiff=col('difficulty'), cCI=col('correct/incorrect'), cMyAns=col('my answer'), cSource=col('source');
  const cA=col('choice a'), cB=col('choice b'), cC=col('choice c'), cD=col('choice d'), cE=col('choice e');
  const cAns=col('correct answer(s)'), cExpl=col('explanation'), cRef=col('ref id');
  const cLastPracticed = colByAny(['Last Practiced', 'Last Practiced Date'], headers);

  let rows = [];
  for (let r = 1; r < bankData.length; r++) {
    const row = bankData[r];

    if (toLower(row[cCourse]) !== toLower(course)) continue;

    const rowTopicLower = toLower(row[cTopic]);
    if (topicsList.length > 0 && !cellContainsAny(rowTopicLower, topicsList)) continue;
    if (excludeTopicsList.length > 0 && cellContainsAny(rowTopicLower, excludeTopicsList)) continue;

    const rowDiffLower = toLower(row[cDiff]);
    if (diffSet.size > 0 && !diffSet.has(rowDiffLower)) continue;

    const rowTypeLower = toLower(row[cT]);
    if (typeSet.size > 0 && !typeSet.has(rowTypeLower)) continue;

    const rowCorrectStatusLower = (cCI > -1) ? toLower(row[cCI]) : '';
    if (corrSet.size > 0 && !corrSet.has(rowCorrectStatusLower)) continue;

    const rowMyAns = (cMyAns > -1) ? clean(row[cMyAns]) : '';
    if (myAnsSet.size > 0 && !myAnsSet.has(rowMyAns)) continue;

    const rowSourceLower = (cSource > -1) ? toLower(row[cSource]) : '';
    if (sourceList.length > 0 && !sourceList.some(s => rowSourceLower.includes(s))) continue;
    if (excludeSourceList.length > 0 && excludeSourceList.some(s => rowSourceLower.includes(s))) continue;

    const rowQuestionText = clean(row[cQ]);
    const chA = clean(row[cA]), chB = clean(row[cB]), chC = clean(row[cC]), chD = clean(row[cD]), chE = clean(row[cE]);
    if (cleanSearchText) {
      const needle = toLower(cleanSearchText);
      const hay = toLower([rowQuestionText, chA, chB, chC, chD, chE].filter(Boolean).join(' || '));
      if (!hay.includes(needle)) continue;
    }
    const rowExplanationText = (cExpl > -1) ? clean(row[cExpl]) : '';

    const lastPracticedRaw  = (cLastPracticed > -1) ? row[cLastPracticed] : null;
    const lastPracticedDate = parseDateValue(lastPracticedRaw);

    if (isPracticeEmpty) {
      if (lastPracticedRaw !== '' && lastPracticedRaw != null) continue;
    }
    if (practicedWithinDate) {
      const cutoff = practicedWithinDate;
      if (!lastPracticedDate) continue;
      if (startOfDay(lastPracticedDate) < cutoff) continue;
    }
    if (notPracticedDate) {
      const cutoffNP = notPracticedDate;
      if (!lastPracticedDate) continue;
      if (startOfDay(lastPracticedDate) > cutoffNP) continue;
    }

    rows.push({
      Q: rowQuestionText,
      QT_raw: clean(row[cT]),
      QT_clean: rowTypeLower,
      Difficulty: clean(row[cDiff]),
      topics: clean(row[cTopic]),
      refId: clean(row[cRef]) || ('R' + (r + 1)),
      A: chA, B: chB, C: chC, D: chD, E: chE,
      ANS: clean(row[cAns]),
      EXPL: rowExplanationText
    });
  }

  const matchedCount = rows.length;
  if (matchedCount === 0) {
    return {
      selectedQuestions: [],
      orderedQuestions: [],
      matchedCount: 0,
      course, examTitle, examSubtitle, shuffleChoices, showTopicAtEnd
    };
  }

  // Quotas
  let selectedQuestions = [];
  const totalDiffQuota = Object.values(diffQuotas || {}).reduce((s, v) => s + toIntOrZero(v), 0);
  const totalTypeQuota = Object.values(typeQuotas || {}).reduce((s, v) => s + toIntOrZero(v), 0);

  if (totalDiffQuota > 0) {
    const buckets = {}; DIFF_ORDER.forEach(d => buckets[toLower(d)] = []);
    rows.forEach(q => { const k = toLower(q.Difficulty); if (buckets[k]) buckets[k].push(q); });
    DIFF_ORDER.forEach(d => {
      const key = toLower(d).replace(/[^a-z0-9]/g, '');
      const qn = toIntOrZero((diffQuotas || {})[key]);
      if (qn > 0) { randShuffle(buckets[toLower(d)]); selectedQuestions.push(...buckets[toLower(d)].slice(0, qn)); }
    });
  } else if (totalTypeQuota > 0) {
    const buckets = {}; TYPE_ORDER.forEach(t => buckets[t] = []);
    rows.forEach(q => { if (buckets[q.QT_clean]) buckets[q.QT_clean].push(q); });
    TYPE_ORDER.forEach(t => {
      const key = t.replace(/[^a-z0-9]/g, '');
      const qn = toIntOrZero((typeQuotas || {})[key]);
      if (qn > 0) { randShuffle(buckets[t]); selectedQuestions.push(...buckets[t].slice(0, qn)); }
    });
  } else {
    selectedQuestions = rows;
  }

  if (toIntOrZero(config.maxQ) > 0 && selectedQuestions.length > toIntOrZero(config.maxQ)) {
    randShuffle(selectedQuestions);
    selectedQuestions = selectedQuestions.slice(0, toIntOrZero(config.maxQ));
  }

  // Group by TYPE_ORDER for final ordering
  const orderedQuestions = [];
  const buckets2 = {}; TYPE_ORDER.forEach(t => buckets2[t] = []);
  selectedQuestions.forEach(q => { if (buckets2[q.QT_clean]) buckets2[q.QT_clean].push(q); });
  TYPE_ORDER.forEach(t => { orderedQuestions.push(...buckets2[t]); });

  return { selectedQuestions, orderedQuestions, matchedCount, course, examTitle, examSubtitle, shuffleChoices, showTopicAtEnd };
}

// =========================
// Rendering
// =========================
function renderStudent(body, render, searchText, showTopicAtEnd){
  for(let i=0;i<render.length;i++){
    const q = render[i];
    const isMCQ = (q.QT || '').toLowerCase() === 'mcq';
    const tagParts = [q.refId];
    const topicNice = prettyTopic(q.topics);
    if (showTopicAtEnd && topicNice) tagParts.push(topicNice);
    const endTag = ' (' + tagParts.join(' \u2013 ') + ')';
    const stemParagraph = body.appendParagraph('Q' + (i+1) + '. ' + q.stem + endTag);
    styleParagraph(stemParagraph,{fontFamily:'Georgia',size:12,bold:isMCQ});
    highlightText(stemParagraph, searchText);
    if(isMCQ){
      for(let j=0;j<q.choices.length;j++){
        const ch=q.choices[j];
        const p=body.appendParagraph('    '+ch.shown+') '+ch.text).setIndentStart(36);
        styleParagraph(p,{fontFamily:'Georgia',size:12,bold:false});
        highlightText(p, searchText);
      }
    } else if((q.QT || '').toLowerCase()==='true/false'){
      styleParagraph(body.appendParagraph('    a) True').setIndentStart(36),{fontFamily:'Georgia',size:12});
      styleParagraph(body.appendParagraph('    b) False').setIndentStart(36),{fontFamily:'Georgia',size:12});
    } else { body.appendParagraph(''); }
    body.appendParagraph('');
  }
}
function renderAnswers(body, render, searchText, showTopicAtEnd){
  for(let i=0;i<render.length;i++){
    const q=render[i];
    const tagParts=[q.refId]; const topicNice=prettyTopic(q.topics);
    if (showTopicAtEnd && topicNice) tagParts.push(topicNice);
    const endTag=' ('+tagParts.join(' \u2013 ')+')';
    styleParagraph(body.appendParagraph('Q'+(i+1)+endTag),{fontFamily:'Georgia',size:12,bold:true});
    if((q.QT||'').toLowerCase()==='mcq'){ styleParagraph(body.appendParagraph('Answer: '+(q.ansMapped||q.ansRaw)),{fontFamily:'Georgia',size:12}); }
    else if((q.QT||'').toLowerCase()==='true/false'){ styleParagraph(body.appendParagraph('Answer: '+(q.ansRaw||'')),{fontFamily:'Georgia',size:12}); }
    else { if(q.ansRaw) styleParagraph(body.appendParagraph('Answer: '+q.ansRaw),{fontFamily:'Georgia',size:12}); }
    if(q.expl){ const p=body.appendParagraph('Explanation: '+q.expl); styleParagraph(p,{fontFamily:'Georgia',size:12}); highlightText(p, searchText); }
    body.appendParagraph('');
  }
}

// =========================
// Bank Tools
// =========================
function mergeImportToBank(){
  const ui = SpreadsheetApp.getUi();
  const bankSheet = getSheet(SHEET_BANK);
  const importSheet = getSheet(SHEET_IMPORT);
  if (!bankSheet || !importSheet) { ui.alert('Error: Please ensure both "Bank" and "Import" sheets exist.'); return; }
  const bankData = bankSheet.getDataRange().getValues();
  const importData = importSheet.getDataRange().getValues();
  if (bankData.length === 0 || importData.length === 0) { ui.alert('Nothing to merge.'); return; }

  const importHeaders = importData[0];
  alignSheetHeaders(bankSheet, importHeaders);

  const bankHeaders = bankSheet.getDataRange().getValues()[0].map(toLower);
  const refIdx = bankHeaders.indexOf(toLower(REF_HEADER));
  if (refIdx === -1) { ui.alert(`Error: The "Bank" sheet must contain a "${REF_HEADER}" column.`); return; }
  const existingRefIds = new Set(bankData.slice(1).map(row => clean(row[refIdx])));
  const toAdd = [];
  for (let i = 1; i < importData.length; i++) {
    const row = importData[i]; const rid = clean(row[refIdx]);
    if (rid && !existingRefIds.has(rid)) { toAdd.push(row); existingRefIds.add(rid); }
  }
  if (toAdd.length > 0) {
    bankSheet.getRange(bankSheet.getLastRow() + 1, 1, toAdd.length, toAdd[0].length).setValues(toAdd);
  }
  clearImportSheet(importSheet);
  ui.alert(`Merge complete. Added ${toAdd.length} new row(s). Import cleared.`);
}

function upsertImportToBank(){
  const ui = SpreadsheetApp.getUi();
  const bankSheet = getSheet(SHEET_BANK);
  const importSheet = getSheet(SHEET_IMPORT);
  if (!bankSheet || !importSheet) { ui.alert('Error: Please ensure both "Bank" and "Import" sheets exist.'); return; }
  const importData = importSheet.getDataRange().getValues();
  if (importData.length < 2) { ui.alert('Import is empty. Nothing to upsert.'); return; }

  const importHeaders = importData[0];
  alignSheetHeaders(bankSheet, importHeaders);

  const bankHeaders = bankSheet.getDataRange().getValues()[0].map(toLower);
  const refIdIndex = bankHeaders.indexOf(toLower(REF_HEADER));
  if (refIdIndex === -1) { ui.alert(`Error: The "Bank" sheet must contain a "${REF_HEADER}" column.`); return; }
  const bankRefMap = new Map(bankSheet.getDataRange().getValues().slice(1).map((row, i) => [clean(row[refIdIndex]), i + 2]));
  let updated = 0, added = 0; const toAdd = [];
  for (let i = 1; i < importData.length; i++) {
    const row = importData[i]; const rid = clean(row[refIdIndex]); if (!rid) continue;
    if (bankRefMap.has(rid)) { bankSheet.getRange(bankRefMap.get(rid), 1, 1, row.length).setValues([row]); updated++; }
    else { toAdd.push(row); }
  }
  if (toAdd.length > 0) { bankSheet.getRange(bankSheet.getLastRow() + 1, 1, toAdd.length, toAdd[0].length).setValues(toAdd); added = toAdd.length; }
  clearImportSheet(importSheet);
  ui.alert(`Upsert Complete!\n\nUpdated: ${updated}\nAdded: ${added}\nImport has been cleared.`);
}

function alignSheetHeaders(targetSheet, newHeaderRowValues){
  const wantNames = newHeaderRowValues.map(h=>clean(h));
  targetSheet.getRange(1,1,1,wantNames.length).setValues([wantNames]);
}

function clearImportSheet(importSheet){
  const lastRow = importSheet.getLastRow();
  const lastCol = importSheet.getLastColumn();
  if (lastRow > 1) importSheet.getRange(2,1,lastRow-1,lastCol).clearContent();
}

// Validation
function validateBankData(){
  const ui=SpreadsheetApp.getUi(); const sheet=getSheet(SHEET_BANK); if(!sheet){ ui.alert('Error: "Bank" sheet not found.'); return; }
  const range=sheet.getDataRange(); const values=range.getValues(); const headers=values[0].map(h=>toLower(clean(h)));
  const backgrounds=range.getBackgrounds();
  const cQ=headers.indexOf('question'); const cT=headers.indexOf('question type'); const cAns=headers.indexOf('correct answer(s)'); const cRef=headers.indexOf(toLower(REF_HEADER));
  const choiceCols=['choice a','choice b','choice c','choice d','choice e'].map(h=>headers.indexOf(h));
  let issues=0; const refCounts=new Map();
  for(let r=1;r<values.length;r++){
    const row=values[r]; const ref=clean(row[cRef]); if(ref) refCounts.set(ref,(refCounts.get(ref)||0)+1);
    if(cQ!==-1 && !clean(row[cQ])){ backgrounds[r][cQ]='#fce8e6'; issues++; }
    if (cT!==-1 && !clean(row[cT])){ backgrounds[r][cT]='#fce8e6'; issues++; }
    if(toLower(row[cT])==='mcq'){
      const ans=clean(row[cAns]);
      if(!ans){ backgrounds[r][cAns]='#fce8e6'; issues++; }
      else {
        const keys=ans.split(/[;,]/).map(k=>toLower(k.trim()));
        for(const k of keys){ const idx='abcde'.indexOf(k); if(idx!==-1 && choiceCols[idx]!==-1 && !clean(row[choiceCols[idx]])){ backgrounds[r][choiceCols[idx]]='#fff2cc'; issues++; } }
      }
    }
  }
  for(let r=1;r<values.length;r++){ const ref=clean(values[r][cRef]); if(ref && refCounts.get(ref)>1){ backgrounds[r][cRef]='#fce8e6'; issues++; } }
  range.setBackgrounds(backgrounds); ui.alert(`Validation complete. Found ${issues} potential issues.`);
}
function clearValidationHighlighting(){ const ui=SpreadsheetApp.getUi(); const s=getSheet(SHEET_BANK); if(!s){ ui.alert('Error: "Bank" sheet not found.'); return; } s.getDataRange().setBackground(null); ui.alert('Validation highlighting has been cleared.'); }
function setupLastPracticedColumn(){
  const s=getSheet(SHEET_BANK); if(!s){ SpreadsheetApp.getUi().alert('Error: "Bank" sheet not found.'); return; }
  const data=s.getDataRange().getValues(); if(data.length<1){ SpreadsheetApp.getUi().alert('Bank sheet is empty.'); return; }
  const headers=data[0].map(h=>toLower(clean(h))); const idxLP=headers.indexOf('last practiced'); const idxLPD=headers.indexOf('last practiced date'); const c=(idxLP!==-1)?idxLP:idxLPD;
  if(c===-1){ SpreadsheetApp.getUi().alert('No "Last Practiced" (or "Last Practiced Date") column found.'); return; }
  const lastRow=Math.max(2,s.getLastRow()); const col=c+1; s.getRange(2,col,lastRow-1,1).setNumberFormat('MM/dd/yyyy');
  const rule=SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  s.getRange(2,col,s.getMaxRows()-1,1).setDataValidation(rule);
  SpreadsheetApp.getUi().alert('"Last Practiced" column formatted as MM/DD/YYYY with a date picker.');
}

// =========================
// Notion Sync: UI wiring
// =========================
function openNotionConfig(){
  const ui = SpreadsheetApp.getUi();
  const htmlFileCandidates = ['NotionConfig_V4.3.0', 'NotionConfig'];
  let html = null;
  let lastError = null;

  for (let i = 0; i < htmlFileCandidates.length; i++){
    try {
      html = HtmlService.createHtmlOutputFromFile(htmlFileCandidates[i]);
      break;
    } catch (err) {
      lastError = err;
    }
  }

  if (!html){
    if (lastError){
      throw lastError;
    }
    throw new Error('Notion configuration dialog HTML file not found.');
  }

  html.setWidth(520).setHeight(520);
  ui.showModalDialog(html, 'Notion Sync - Configuration');
}

function readNotionConfig(){
  const map = getDbMap();
  return {
    hasToken: Boolean(PROP.getProperty('NOTION_TOKEN')),
    dbMap: map
  };
}

function saveNotionConfig(payload){
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid configuration payload.');
  }

  if (payload.dbMap) {
    setDbMap(payload.dbMap);
  }

  if (payload.clearToken) {
    PROP.deleteProperty('NOTION_TOKEN');
  } else if (Object.prototype.hasOwnProperty.call(payload, 'token')) {
    const token = clean(payload.token);
    if (!token) {
      throw new Error('Token cannot be empty. Use "Clear Token" if you intend to remove it.');
    }
    PROP.setProperty('NOTION_TOKEN', token);
  }

  return 'Configuration saved successfully.';
}

// NEW: Sync Down Filters dialog (inline HTML, no extra file)
function openSyncDownFilters(){
  const dbMap = getDbMap();
  const courseOptions = Object.keys(dbMap).map(c=>`<option value="${htmlEscape(c)}">${htmlEscape(c)}</option>`).join('');
  const html = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding:14px; }
        label { display:block; font-weight:600; margin-top:10px; }
        input, select { width:100%; padding:7px; border:1px solid #ccc; border-radius:6px; }
        .row { display:grid; grid-template-columns: 1fr 1fr; gap:10px; }
        .btns { margin-top: 16px; display:flex; gap:10px; }
        button { background:#1a73e8; color:#fff; border:none; padding:10px 14px; border-radius:8px; cursor:pointer; font-weight:700; }
        button.secondary { background:#5f6368; }
        button.disabled { opacity:0.6; cursor:default; }
        small { color:#5f6368; display:block; margin-top:4px; }
        #message { margin-top:12px; font-size:12px; display:none; }
      </style>
    </head>
    <body>
      <h3>Sync Down - Filters</h3>
      <label for="course">Course</label>
      <select id="course">
        <option value="__ALL__">All configured courses</option>
        ${courseOptions}
      </select>
      <label for="refPrefix">Ref ID prefix (optional)</label>
      <input id="refPrefix" placeholder="e.g., BIO-">
      <div class="row">
        <div>
          <label for="refStart">Ref ID range start (optional)</label>
          <input id="refStart" placeholder="e.g., BIO-0001">
        </div>
        <div>
          <label for="refEnd">Ref ID range end (optional)</label>
          <input id="refEnd" placeholder="e.g., BIO-0500">
        </div>
      </div>
      <label for="topics">Topic/Chapter contains (semicolon list)</label>
      <input id="topics" placeholder="e.g., T/Ch 01; Topic 5">
      <label for="tags">Tags include ANY (semicolon list)</label>
      <input id="tags" placeholder="e.g., genetics; exam1; image-required">
      <label for="lastEdited">Last Edited Since (YYYY-MM-DD) (optional)</label>
      <input id="lastEdited" placeholder="e.g., 2025-09-01">
      <div id="message" role="alert"></div>
      <div class="btns">
        <button id="runBtn" type="button">Sync Down</button>
        <button id="cancelBtn" type="button" class="secondary">Cancel</button>
      </div>
      <script>
        (function(){
          const runBtn = document.getElementById('runBtn');
          const cancelBtn = document.getElementById('cancelBtn');
          const message = document.getElementById('message');
          const inputs = Array.from(document.querySelectorAll('input, select'));

          function setMessage(text, isError){
            if (!message) return;
            if (!text){
              message.style.display = 'none';
              message.textContent = '';
              return;
            }
            message.style.display = 'block';
            message.style.color = isError ? '#d93025' : '#188038';
            message.textContent = text;
          }

          function setBusy(isBusy){
            inputs.forEach(el => { el.disabled = Boolean(isBusy); });
            [runBtn, cancelBtn].forEach(btn => {
              if (!btn) return;
              btn.disabled = Boolean(isBusy && btn === runBtn);
              btn.classList.toggle('disabled', Boolean(isBusy && btn === runBtn));
            });
            if (isBusy){
              setMessage('Sync in progress...', false);
            } else {
              setMessage('', false);
            }
          }

          function valueOf(id){
            const el = document.getElementById(id);
            return el ? el.value.trim() : '';
          }

          function gather(){
            return {
              course: valueOf('course'),
              refPrefix: valueOf('refPrefix'),
              refStart: valueOf('refStart'),
              refEnd: valueOf('refEnd'),
              topics: valueOf('topics'),
              tags: valueOf('tags'),
              lastEdited: valueOf('lastEdited')
            };
          }

          function handleSuccess(msg){
            setBusy(false);
            alert(msg);
            google.script.host.close();
          }

          function handleFailure(err){
            setBusy(false);
            var text = 'An unexpected error occurred.';
            if (err){
              if (typeof err === 'string') {
                text = err;
              } else if (err.message) {
                text = err.message;
              } else {
                text = String(err);
              }
            }
            setMessage(text, true);
          }

          if (cancelBtn){
            cancelBtn.addEventListener('click', function(){
              google.script.host.close();
            });
          }

          if (runBtn){
            runBtn.addEventListener('click', function(){
              setBusy(true);
              google.script.run
                .withSuccessHandler(handleSuccess)
                .withFailureHandler(handleFailure)
                .runSyncDownWithFilters(gather());
            });
          }
        })();
      </script>
    </body>
    </html>
  `).setWidth(520).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sync Down - Filters');
}

function htmlEscape(s){ return String(s||'').replace(/[&<>"']/g, m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m])); }

// Called by filter dialog
function runSyncDownWithFilters(filters){
  const cfg = requireNotionConfig();
  const dbMap = getDbMap();
  const courseNames = Object.keys(dbMap);
  if (!courseNames.length) throw new Error('No courses configured in Notion Sync -> Configure...');

  const importSheet = getSheet(SHEET_IMPORT); if(!importSheet) throw new Error('Import sheet missing.');
  ensureCanonicalHeaders(importSheet);

  const targets = [];
  if (filters.course && filters.course !== '__ALL__') {
    const dbid = dbMap[filters.course];
    if (dbid) targets.push({ name: filters.course, id: dbid });
  } else {
    courseNames.forEach(name => targets.push({ name, id: dbMap[name] }));
  }

  if (!targets.length) {
    throw new Error('No matching courses found for the selected filter.');
  }

  const numCols = CANON_HEADERS.length;
  const lastRowWithData = importSheet.getLastRow();
  const rowsToRead = Math.max(1, lastRowWithData);
  const importData = importSheet.getRange(1, 1, rowsToRead, numCols).getValues();
  const existingRows = importData.length > 1 ? importData.slice(1) : [];
  const headersLower = CANON_HEADERS.map(h => toLower(h));
  const cRef = headersLower.indexOf(toLower(REF_HEADER));
  if (cRef === -1) {
    throw new Error('Import sheet is missing the Ref ID column.');
  }
  const existingRefToIndex = new Map();
  existingRows.forEach((row, idx) => {
    const ref = clean(row[cRef]);
    if (ref) existingRefToIndex.set(ref, idx);
  });

  const newRows = [];
  const seenNewRefs = new Set();
  let updates = 0, adds = 0, fetched = 0;

  targets.forEach(t => {
    const rows = fetchNotionDbAsRowsFiltered(t.id, t.name, filters, cfg);
    fetched += rows.length;
    rows.forEach(r => {
      const ref = clean(r[cRef]);
      if (!ref) return;
      if (existingRefToIndex.has(ref)) {
        existingRows[existingRefToIndex.get(ref)] = r;
        updates++;
      } else if (!seenNewRefs.has(ref)) {
        newRows.push(r);
        seenNewRefs.add(ref);
        adds++;
      }
    });
  });

  if (existingRows.length) {
    importSheet.getRange(2, 1, existingRows.length, numCols).setValues(existingRows);
  }

  if (newRows.length) {
    const lastRowBeforeInsert = importSheet.getLastRow();
    const anchorRow = Math.max(1, lastRowBeforeInsert);
    importSheet.insertRowsAfter(anchorRow, newRows.length);
    importSheet.getRange(anchorRow + 1, 1, newRows.length, numCols).setValues(newRows);
  }

  return `Sync Down complete.\nFetched: ${fetched}\nUpdated in Import: ${updates}\nAdded to Import: ${adds}`;
}

// Back-compat Preview (counts by Ref IDs vs Import) - iterates all configured DBs
function notionPreview(){
  const cfg = requireNotionConfig();
  const dbMap = getDbMap();
  const courses = Object.keys(dbMap);
  if (!courses.length) {
    SpreadsheetApp.getUi().alert('No courses configured. Use Notion Sync -> Configure... to add database IDs.');
    return;
  }
  const notionIndex = {};
  courses.forEach(c => { notionIndex[c] = indexNotionRefs(dbMap[c], cfg); });
  const importIndex = indexSheet(getSheet(SHEET_IMPORT));

  const newInNotion = {};
  const newForNotion = {};
  courses.forEach(k => {
    newInNotion[k] = diffSet(notionIndex[k], importIndex.all);
    const subset = new Set([...importIndex.all].filter(rid => (rid||'').startsWith(k.split(' ')[0]))); // heuristic subset
    newForNotion[k] = diffSet(subset, notionIndex[k]);
  });

  const totalNewInNotion = courses.reduce((sum, k) => sum + (newInNotion[k] ? newInNotion[k].size : 0), 0);
  const totalNewForNotion = courses.reduce((sum, k) => sum + (newForNotion[k] ? newForNotion[k].size : 0), 0);

  let lines = ['Preview (by Ref ID)', '', `From Notion -> Import (new in Notion): ${totalNewInNotion}`];
  courses.forEach(k => lines.push(`  ${k}: ${newInNotion[k].size}`));
  lines.push('', `From Import -> Notion (new in Import): ${totalNewForNotion}`);
  courses.forEach(k => lines.push(`  ${k}: ${newForNotion[k].size}`));
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

function notionSyncUp(){
  const cfg = requireNotionConfig();
  const importSheet = getSheet(SHEET_IMPORT); if(!importSheet) throw new Error('Import sheet missing.');
  const data = importSheet.getDataRange().getValues(); if(data.length < 2){ SpreadsheetApp.getUi().alert('Import is empty.'); return; }
  const headers = data[0]; const rows = data.slice(1);
  const dbMap = getDbMap();

  let created=0, updated=0, skipped=0, errors=0;
  rows.forEach(row=>{
    const rec = rowToRecord(headers, row);
    if (!rec.RefID || !rec.Course) { skipped++; return; }
    const dbId = courseToDbId(rec.Course, dbMap);
    if (!dbId) { skipped++; return; }
    try {
      const pageId = findPageIdByRefId(dbId, rec.RefID, cfg);
      if (pageId) { updateNotionPage(pageId, rec, cfg); updated++; }
      else { createNotionPage(dbId, rec, cfg); created++; }
    } catch (e) { errors++; Logger.log(e); }
  });

  SpreadsheetApp.getUi().alert(`Sync Up complete.\nUpdated in Notion: ${updated}\nCreated in Notion: ${created}\nSkipped: ${skipped}\nErrors: ${errors}`);
}

// =========================
// Notion helpers & mapping
// =========================
function requireNotionConfig(){
  const token = PROP.getProperty('NOTION_TOKEN');
  if(!token){ throw new Error('Notion token missing. Use Notion Sync -> Configure...'); }
  return { token };
}

function getDbMap(){
  const raw = PROP.getProperty(DBMAP_KEY) || '{}';
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      const sanitized = sanitizeCourseMap(parsed, { strict: false });
      const parsedString = JSON.stringify(parsed || {});
      const sanitizedString = JSON.stringify(sanitized);
      if (parsedString !== sanitizedString) {
        PROP.setProperty(DBMAP_KEY, sanitizedString);
      }
      return sanitized;
    }
  } catch(e){}
  return {};
}
function setDbMap(obj){
  const sanitized = sanitizeCourseMap(obj, { strict: true });
  PROP.setProperty(DBMAP_KEY, JSON.stringify(sanitized));
  const validDbIds = new Set(Object.values(sanitized));
  Object.keys(notionDbMetadataCache).forEach(dbId => {
    if (!validDbIds.has(dbId)) delete notionDbMetadataCache[dbId];
  });
}

function notionRequest(path, method, payload, cfg){
  const url = NOTION_API_BASE + path;
  const options = {
    method: method || 'get',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + (cfg.token || PROP.getProperty('NOTION_TOKEN')), 'Notion-Version': NOTION_VERSION },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  const maxAttempts = 5;
  let attempt = 0;
  let delay = NOTION_RATE_LIMIT_MS;
  let lastError = null;

  while (attempt < maxAttempts) {
    attempt++;
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const body = resp.getContentText();
      if (code >= 200 && code < 300) {
        const parsed = body ? JSON.parse(body) : {};
        Utilities.sleep(NOTION_RATE_LIMIT_MS);
        return parsed;
      }

      const errMsg = 'Notion API error ' + code + ': ' + body;
      if (code === 429 || code >= 500) {
        lastError = new Error(errMsg);
      } else {
        throw new Error(errMsg);
      }
    } catch (e) {
      lastError = e;
    }

    if (attempt >= maxAttempts) break;
    Utilities.sleep(delay);
    delay = Math.min(delay * 2, 5000);
  }

  throw lastError || new Error('Unknown Notion API error.');
}

function describeNotionDatabase(dbId, cfg){
  if (!dbId) return null;
  if (!notionDbMetadataCache[dbId]) {
    notionDbMetadataCache[dbId] = notionRequest('/databases/' + dbId, 'get', null, cfg);
  }
  return notionDbMetadataCache[dbId];
}

// Filtered fetch
function fetchNotionDbAsRowsFiltered(dbId, courseName, filters, cfg){
  const rows = [];
  let cursor = null;
  const config = cfg || requireNotionConfig();
  const dbMeta = describeNotionDatabase(dbId, config);
  const filterPlan = buildNotionFilterPlan(filters, dbMeta);

  do {
    const payload = { page_size: 100 };
    if (filterPlan.notion) payload.filter = filterPlan.notion;
    if (cursor) payload.start_cursor = cursor;
    const res = notionRequest('/databases/' + dbId + '/query', 'post', payload, config);
    (res.results || []).forEach(page => {
      const props = page.properties || {};
      const rec = notionPropsToRecord(props);
      rec.Course = courseName;
      if (!passesClientFilters(rec, filterPlan.clientFilters)) return;
      rows.push(recordToRow(rec));
    });
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);

  return applyRefRangeFilter(rows, filters);
}

function passesClientFilters(rec, fns){
  if (!fns || !fns.length) return true;
  for (let i = 0; i < fns.length; i++) {
    if (!fns[i](rec)) return false;
  }
  return true;
}

function applyRefRangeFilter(rows, filters){
  const refStart = clean(filters.refStart);
  const refEnd = clean(filters.refEnd);
  if (!refStart && !refEnd) return rows;

  const idx = CANON_HEADERS.map(h => toLower(h)).indexOf(toLower(REF_HEADER));
  if (idx === -1) return rows;

  return rows.filter(r => {
    const rid = clean(r[idx]);
    if (refStart && rid < refStart) return false;
    if (refEnd && rid > refEnd) return false;
    return true;
  });
}

function buildNotionFilterPlan(filters, dbMeta){
  const clauses = [];
  const clientFilters = [];

  const refPrefix = clean(filters.refPrefix);
  if (refPrefix) {
    clauses.push({ property: 'Ref ID', rich_text: { starts_with: refPrefix } });
  }

  const topics = parseList(filters.topics);
  if (topics.length) {
    const topicMeta = lookupPropMeta(dbMeta, ['Topic/Chapter','Topic / Chapter','Topic- Chapter','Topic','Chapter','Chapter/Topic']);
    const topicOrs = [];
    if (topicMeta) {
      topics.forEach(tok => {
        if (!tok) return;
        if (topicMeta.type === 'select') {
          topicOrs.push({ property: topicMeta.name, select: { equals: tok } });
        } else if (topicMeta.type === 'multi_select') {
          topicOrs.push({ property: topicMeta.name, multi_select: { contains: tok } });
        } else {
          topicOrs.push({ property: topicMeta.name, rich_text: { contains: tok } });
        }
      });
    }
    if (topicOrs.length) {
      clauses.push({ or: topicOrs });
    } else {
      const lowered = topics.map(t => t.toLowerCase());
      clientFilters.push(rec => {
        const hay = toLower(rec.Topic);
        if (!hay) return false;
        return lowered.some(tok => hay.indexOf(tok) > -1);
      });
    }
  }

  const tags = parseList(filters.tags);
  if (tags.length) {
    const tagsMeta = lookupPropMeta(dbMeta, ['Tags']);
    const tagOrs = [];
    if (tagsMeta) {
      tags.forEach(t => {
        if (!t) return;
        if (tagsMeta.type === 'multi_select') {
          tagOrs.push({ property: tagsMeta.name, multi_select: { contains: t } });
        } else if (tagsMeta.type === 'select') {
          tagOrs.push({ property: tagsMeta.name, select: { equals: t } });
        }
      });
    }
    if (tagOrs.length) {
      clauses.push({ or: tagOrs });
    } else {
      const loweredTags = tags.map(t => t.toLowerCase());
      clientFilters.push(rec => {
        const hay = toLower(rec.Tags);
        if (!hay) return false;
        const parts = hay.split(/[,;]+/).map(x => x.trim()).filter(Boolean);
        const partSet = new Set(parts);
        return loweredTags.some(tok => partSet.has(tok));
      });
    }
  }

  const lastEditedRaw = clean(filters.lastEdited);
  if (lastEditedRaw) {
    const iso = formatDateForNotion(lastEditedRaw);
    if (!iso || !isValidIsoDateString(iso)) {
      throw new Error('Last Edited Since must be a valid date (YYYY-MM-DD or MM/DD/YYYY).');
    }
    clauses.push({ timestamp: 'last_edited_time', last_edited_time: { on_or_after: iso } });
  }

  let notion = null;
  if (clauses.length === 1) notion = clauses[0];
  else if (clauses.length > 1) notion = { and: clauses };

  return { notion, clientFilters };
}

function lookupPropMeta(dbMeta, candidates){
  if (!dbMeta || !dbMeta.properties) return null;
  const props = dbMeta.properties;
  for (const key in props) {
    const norm = normalizeName(key);
    for (let i = 0; i < candidates.length; i++) {
      if (norm === normalizeName(candidates[i])) {
        return { name: key, type: props[key].type };
      }
    }
  }
  return null;
}

function isValidIsoDateString(value){
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(value + 'T00:00:00Z');
  if (isNaN(date.getTime())) return false;
  const [y, m, d] = value.split('-').map(Number);
  return date.getUTCFullYear() === y && (date.getUTCMonth() + 1) === m && date.getUTCDate() === d;
}

function findPageIdByRefId(dbId, refId, cfg){
  const payload = {
    filter: { property: 'Ref ID', rich_text: { equals: refId } },
    page_size: 1
  };
  const res = notionRequest('/databases/'+dbId+'/query', 'post', payload, cfg);
  return (res.results && res.results.length) ? res.results[0].id : null;
}

function createNotionPage(dbId, rec, cfg){
  const props = recordToNotionProps(rec);
  notionRequest('/pages', 'post', { parent: { database_id: dbId }, properties: props }, cfg);
}
function updateNotionPage(pageId, rec, cfg){
  const props = recordToNotionProps(rec);
  notionRequest('/pages/'+pageId, 'patch', { properties: props }, cfg);
}

function courseToDbId(course){
  const map = getDbMap();
  const exact = map[clean(course)];
  if (exact) return exact;
  // prefix match (BIO, ANP, PSY, etc.)
  const key = clean(course).split(' ')[0].toUpperCase();
  for (const name in map){
    if (name.toUpperCase().startsWith(key)) return map[name];
  }
  return null;
}

// Canonical headers (18) - order agnostic
const CANON_HEADERS = [
  'Question','Question Type','Course','Topic/Chapter','Difficulty','Correct/Incorrect',
  'Choice A','Choice B','Choice C','Choice D','Choice E',
  'Correct Answer(s)','Explanation','My Answer','Source','Last Practiced','Tags','Ref ID'
];

function ensureCanonicalHeaders(sheet){
  const out = []; CANON_HEADERS.forEach(h => out.push(h));
  sheet.getRange(1,1,1,out.length).setValues([out]);
}

function rowToRecord(headers, row){
  const idx = {}; headers.forEach((h,i)=> idx[toLower(clean(h))]=i);
  const g = n => clean(row[idx[toLower(n)]]);
  const rec = {
    Question: g('Question'),
    QuestionType: g('Question Type'),
    Course: g('Course'),
    Topic: g('Topic/Chapter'),
    Difficulty: g('Difficulty'),
    CorrectIncorrect: g('Correct/Incorrect'),
    MyAnswer: g('My Answer'),
    Source: g('Source'),
    A: g('Choice A'), B: g('Choice B'), C: g('Choice C'), D: g('Choice D'), E: g('Choice E'),
    Correct: g('Correct Answer(s)'),
    Explanation: g('Explanation'),
    LastPracticed: g('Last Practiced'),
    Tags: g('Tags'),
    RefID: g('Ref ID')
  };
  return rec;
}

function recordToRow(rec){
  return [
    rec.Question, rec.QuestionType, rec.Course, rec.Topic, rec.Difficulty, rec.CorrectIncorrect,
    rec.A, rec.B, rec.C, rec.D, rec.E, rec.Correct, rec.Explanation, rec.MyAnswer, rec.Source,
    rec.LastPracticed, rec.Tags, rec.RefID
  ];
}

function recordToNotionProps(rec){
  const rt = v => ({ rich_text: v ? [{ type:'text', text:{ content:String(v)} }] : [] });
  const sel = v => v ? { select: { name: String(v) } } : { select: null };
  const date = v => {
    const s = clean(v); if(!s) return { date: null };
    return { date: { start: formatDateForNotion(s) } };
  };
  const tags = s => {
    const parts = parseListSemicolon(s);
    if (!parts.length) return { multi_select: [] };
    return { multi_select: parts.map(x => ({ name: x })) };
  };
  return {
    'Question': { title: rec.Question ? [{ type:'text', text:{ content:String(rec.Question) } }] : [] },
    'Question Type': sel(rec.QuestionType),
    'Course': rt(rec.Course),
    'Topic/Chapter': rt(rec.Topic),
    'Difficulty': sel(rec.Difficulty),
    'Correct/Incorrect': sel(rec.CorrectIncorrect),
    'My Answer': rt(rec.MyAnswer),
    'Source': rt(rec.Source),
    'Choice A': rt(rec.A),
    'Choice B': rt(rec.B),
    'Choice C': rt(rec.C),
    'Choice D': rt(rec.D),
    'Choice E': rt(rec.E),
    'Correct Answer(s)': rt(rec.Correct),
    'Explanation': rt(rec.Explanation),
    'Last Practiced': date(rec.LastPracticed),
    'Tags': tags(rec.Tags),
    'Ref ID': rt(rec.RefID)
  };
}

// ---- Robust Notion reading helpers ----
function normalizeName(s){ return toLower(String(s||'')).replace(/[^a-z0-9]+/g,''); }

function pickProp(props, candidates){
  const keys = Object.keys(props||{});
  const normToKey = {};
  keys.forEach(k => { normToKey[normalizeName(k)] = k; });
  for (let i=0;i<candidates.length;i++){
    const norm = normalizeName(candidates[i]);
    if (normToKey[norm]) return props[normToKey[norm]];
  }
  return null;
}

function readPropAsString(p){
  if (!p) return '';
  try{
    if (p.title && Array.isArray(p.title)) {
      return p.title.map(x=>x.plain_text||'').join('');
    }
    if (p.rich_text && Array.isArray(p.rich_text)) {
      return p.rich_text.map(x=>x.plain_text||'').join('');
    }
    if (p.select && p.select.name) return p.select.name;
    if (p.multi_select && Array.isArray(p.multi_select)) {
      return p.multi_select.map(x=>x.name||'').filter(Boolean).join('; ');
    }
    if (p.date && p.date.start) {
      return formatDateFromNotion(p.date.start);
    }
    if (typeof p.number === 'number') return String(p.number);
    if (typeof p.checkbox === 'boolean') return p.checkbox ? 'true':'false';
    if (p.url) return String(p.url);
    const any = p.title || p.rich_text || [];
    if (Array.isArray(any)) return any.map(x=>x.plain_text||'').join('');
  }catch(e){}
  return '';
}

function notionPropsToRecord(props){
  const get = (nameArr) => readPropAsString(pickProp(props, nameArr));

  const tagsProp = props['Tags'];
  const tagsStr = (tagsProp && tagsProp.multi_select)
    ? tagsProp.multi_select.map(x=>x.name).join('; ')
    : readPropAsString(tagsProp);

  return {
    Question: readPropAsString(props['Question']),
    QuestionType: get(['Question Type']),
    Topic: get(['Topic/Chapter','Topic / Chapter','Topic- Chapter','Topic','Chapter','Chapter/Topic']),
    Difficulty: get(['Difficulty']),
    CorrectIncorrect: get(['Correct/Incorrect']),
    MyAnswer: get(['My Answer']),
    Source: get(['Source']),
    A: get(['Choice A']), B: get(['Choice B']), C: get(['Choice C']), D: get(['Choice D']), E: get(['Choice E']),
    Correct: get(['Correct Answer(s)']),
    Explanation: get(['Explanation']),
    LastPracticed: get(['Last Practiced','Last Practiced Date']),
    Tags: tagsStr,
    RefID: get(['Ref ID']),
    Course: '' // filled by caller
  };
}

// =========================
// Notion/date/indices helpers
// =========================
function formatDateForNotion(input){
  const s = String(input||'').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m){ const mm=('0'+m[1]).slice(-2), dd=('0'+m[2]).slice(-2); return `${m[3]}-${mm}-${dd}`; }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if(m){ const mm=('0'+m[2]).slice(-2), dd=('0'+m[3]).slice(-2); return `${m[1]}-${mm}-${dd}`; }
  const d = new Date(s); if(!isNaN(d)) return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
  return null;
}
function formatDateFromNotion(iso){
  try{
    const d = new Date(iso);
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'UTC', 'MM/dd/yyyy');
  }catch(e){ return ''; }
}

function sanitizeCourseMap(input, options){
  const opts = options || {};
  const strict = Boolean(opts.strict);
  const out = {};
  const seen = new Set();
  const invalidCourses = [];
  const duplicateCourses = [];
  const source = input && typeof input === 'object' ? input : {};

  Object.keys(source).forEach(key => {
    const name = clean(key);
    const dbIdRaw = source[key];
    if (!name) {
      if (strict) invalidCourses.push('(missing course name)');
      return;
    }

    const norm = name.toLowerCase();
    if (seen.has(norm)) {
      if (strict) duplicateCourses.push(name);
      return;
    }

    const normalizedId = normalizeNotionDatabaseId(dbIdRaw);
    if (!normalizedId) {
      if (strict) invalidCourses.push(name);
      return;
    }

    seen.add(norm);
    out[name] = normalizedId;
  });

  if (strict && (invalidCourses.length || duplicateCourses.length)) {
    const messages = [];
    if (invalidCourses.length) {
      messages.push('Invalid Notion database ID for: ' + invalidCourses.join(', '));
    }
    if (duplicateCourses.length) {
      messages.push('Duplicate course names: ' + duplicateCourses.join(', '));
    }
    throw new Error(messages.join('; '));
  }

  return out;
}

function normalizeNotionDatabaseId(value){
  const raw = clean(value);
  if (!raw) return '';

  const withoutQuery = raw.split('?')[0].trim();
  const hyphenatedMatch = withoutQuery.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  if (hyphenatedMatch) {
    return hyphenatedMatch[0].replace(/-/g, '').toLowerCase();
  }

  const allHexMatches = withoutQuery.match(/[0-9a-f]{32}/ig);
  if (allHexMatches && allHexMatches.length) {
    return allHexMatches[allHexMatches.length - 1].toLowerCase();
  }

  if (/^[0-9a-f]{32}$/i.test(raw)) {
    return raw.toLowerCase();
  }

  return '';
}

function courseToDbId(course, mapOverride){
  const map = mapOverride || getDbMap();
  if (!course) return null;
  const cleaned = clean(course);
  if (!cleaned) return null;
  const names = Object.keys(map || {});
  if (!names.length) return null;

  if (map.hasOwnProperty(cleaned)) {
    return map[cleaned];
  }

  const lower = cleaned.toLowerCase();
  const ciExact = names.find(name => name.toLowerCase() === lower);
  if (ciExact) {
    return map[ciExact];
  }

  const prefix = cleaned.split(' ')[0].toLowerCase();
  if (!prefix) return null;
  const matches = names.filter(name => name.toLowerCase().startsWith(prefix));
  if (matches.length === 1) {
    return map[matches[0]];
  }
  return null;
}

function indexSheet(sheet){
  const out = { all:new Set(), byCourse:{} };
  if(!sheet) return out;
  const data = sheet.getDataRange().getValues(); if(data.length<2) return out;
  const headers = data[0].map(h=>toLower(clean(h)));
  const cRef = headers.indexOf(toLower(REF_HEADER));
  for(let i=1;i<data.length;i++){
    const ref=clean(data[i][cRef]); if(!ref) continue;
    out.all.add(ref);
  }
  return out;
}

function indexNotionAll(){
  const cfg = requireNotionConfig();
  const dbMap = getDbMap();
  const out = {};
  Object.keys(dbMap).forEach(name => { out[name] = indexNotionRefs(dbMap[name], cfg); });
  return out;
}
function indexNotionRefs(dbId, cfg){
  const refs = new Set(); let cursor=null;
  do{
    const payload = { page_size: 100, filter: { property:'Ref ID', rich_text:{ is_not_empty:true } } };
    if (cursor) payload.start_cursor = cursor;
    const res = notionRequest('/databases/'+dbId+'/query', 'post', payload, cfg);
    (res.results||[]).forEach(page=>{
      const ridProp = page.properties && page.properties['Ref ID'];
      const rid = readPropAsString(ridProp);
      if (rid) refs.add(rid);
    });
    cursor = res.has_more ? res.next_cursor : null;
  } while(cursor);
  return refs;
}
function diffSet(aSet, bSet){
  const out = new Set(); aSet.forEach(v=>{ if(!bSet.has(v)) out.add(v); }); return out;
}

// =========================
// Generic helpers
// =========================
function getSheet(name){ return SpreadsheetApp.getActive().getSheetByName(name); }
function clean(s){ return (s == null ? '' : String(s)).trim(); }
function toLower(s){ return clean(s).toLowerCase(); }
function randShuffle(arr){ for (let i=arr.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [arr[i],arr[j]]=[arr[j]]; } return arr; }
function parseList(s){ return clean(s).split(/[;,]/).map(x=>x.trim()).filter(Boolean); }
function parseListSemicolon(s){ return clean(s).split(/[;]+/).map(x=>x.trim()).filter(Boolean); }
function toIntOrZero(s){ s = clean(s); return (s && !isNaN(s)) ? Math.max(0, parseInt(s,10)) : 0; }
function stripLeadingNum(s){ s = clean(s); return s.replace(/^\s*(?:Q\s*)?\d+\s*[\)\.\-:]\s*/i, ''); }
function stripChoicePrefix(s){ s = clean(s); return s.replace(/^\s*[A-E]\s*[\.\)\-:]\s*/i, ''); }
function cellContainsAny(cellValueLower, needlesLower){
  if(needlesLower.length === 0) return true;
  const hay = ',' + cellValueLower.replace(/[;,]/g, ',') + ',';
  for (const tok of needlesLower){ if(hay.indexOf(',' + tok + ',') > -1) return true; }
  return false;
}
function prettyTopic(s){
  s = clean(s);
  const parts = s.split(/[;,]/).map(x=>x.trim()).filter(Boolean).map(tok=>{
    let m = tok.match(/T\/?Ch\s*(\d+)/i);
    if(m) return 'Topic ' + parseInt(m[1],10);
    m = tok.match(/Topic\s*(\d+)/i);
    if(m) return 'Topic ' + parseInt(m[1],10);
    return tok;
  });
  return parts.join(', ');
}
function styleParagraph(p, opts){
  const has = fn => typeof fn === 'function';
  const t = has(p.editAsText) ? p.editAsText() : null;
  if(t){
    if(opts.bold != null) t.setBold(opts.bold);
    if(opts.italic != null) t.setItalic(opts.italic);
    if(opts.size != null) t.setFontSize(opts.size);
    const family = opts.fontFamily || opts.family;
    if(family) t.setFontFamily(family);
  }
  if(opts.center && has(p.setAlignment)) p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if(opts.lineSpacing && has(p.setLineSpacing)) p.setLineSpacing(opts.lineSpacing);
  if(opts.keepWithNext && has(p.setKeepWithNext)) p.setKeepWithNext(true);
  if(opts.keepLines && has(p.setKeepLinesTogether)) p.setKeepLinesTogether(true);
  return p;
}
function showDocLink(url, titleText){
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial,sans-serif; padding:12px 8px;">' +
    '<div style="margin-bottom:10px;">' + titleText + '</div>' +
    '<a href="'+url+'" target="_blank" style="display:inline-block; background:#1a73e8; color:#fff; padding:8px 14px; border-radius:6px; text-decoration:none;">Open Exam</a>' +
    '<div style="margin-top:10px; color:#666;">(Opens in a new tab.)</div></div>'
  ).setWidth(360).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Exam Created');
}
function highlightText(paragraph, textToHighlight) {
  if (!textToHighlight || !paragraph) return;
  const t = paragraph.editAsText();
  let f = t.findText(textToHighlight);
  while (f) { const s = f.getStartOffset(); const e = f.getEndOffsetInclusive(); if (s !== -1) t.setBold(s, e, true); f = t.findText(textToHighlight, f); }
}

// date helpers
function parseDateValue(v){
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m){ return new Date(m[3], parseInt(m[1],10)-1, parseInt(m[2],10)); }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if(m){ return new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10)); }
  const d = new Date(s); return isNaN(d) ? null : d;
}
function daysAgo(n){
  const d = new Date(); d.setDate(d.getDate()-n); return startOfDay(d);
}
function startOfDay(d){
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
