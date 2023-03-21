"use strict";
/// <reference types="google-apps-script" />
const TABULATION_FILE_URL = 'https://drive.google.com/file/d/1JVa9eE67qMy_mLosfK3thuh-gjDlD05y/view?usp=share_link';
function printTabulation() {
    const pages = loadFormPages();
    const tabulations = loadTabulations();
    const dupeTitleCountMapForInit = {};
    const tabulationMap = Object.fromEntries(tabulations.map(t => [getTitleForRef(dupeTitleCountMapForInit, t.title), t]));
    const pageIndexToQuestionNumberMap = Object.fromEntries(pages.map(page => {
        var _a, _b;
        return [
            page.index,
            (_b = (_a = page.items.find((item) => item.kind === 'question')) === null || _a === void 0 ? void 0 : _a.number) !== null && _b !== void 0 ? _b : null,
        ];
    }));
    const pageIndexToLastQuestionNumberMap = Object.fromEntries(pages.map(page => {
        var _a, _b;
        return [
            page.index,
            (_b = (_a = page.items
                .slice()
                .reverse()
                .find((item) => item.kind === 'question')) === null || _a === void 0 ? void 0 : _a.number) !== null && _b !== void 0 ? _b : null,
        ];
    }));
    const form = FormApp.getActiveForm();
    const doc = DocumentApp.create(`${form.getTitle()}-単純集計`);
    const body = doc.getBody();
    const context = {
        body,
        tabulationMap,
        pageIndexToQuestionNumberMap,
        pageIndexToLastQuestionNumberMap,
        dupeTitleCountMap: {},
    };
    body.appendParagraph('単純集計結果').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const firstPage = pages[0];
    for (const page of pages) {
        if (page === firstPage) {
            // from first question
            const startPos = page.items.findIndex(item => item.kind === 'question');
            renderPage(context, Object.assign(Object.assign({}, page), { items: page.items.slice(startPos) }), false);
            continue;
        }
        renderPage(context, page, false);
    }
    body.appendParagraph('調査票').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(form.getTitle()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(form.getDescription());
    const context2 = Object.assign(Object.assign({}, context), { dupeTitleCountMap: {} });
    for (const page of pages) {
        renderPage(context2, page, true);
    }
    Logger.log('Document created on %s', doc.getUrl());
}
function renderPage(context, page, isQuestionnaire) {
    const { body } = context;
    if (page.title)
        body.appendParagraph(page.title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (page.description)
        body.appendParagraph(page.description);
    if (page.title || page.description)
        body.appendParagraph('');
    for (const item of page.items) {
        switch (item.kind) {
            case 'sectionHeader':
                renderSectionHeader(context, item);
                break;
            case 'question':
                if (isQuestionnaire) {
                    renderQuestionItem(context, page, item);
                }
                else {
                    renderQuestionItemTabulation(context, page, item);
                }
                break;
            case 'image':
                renderImageItem(context, item);
                break;
        }
    }
}
function renderQuestionItemTabulation(context, page, item) {
    const { body, tabulationMap, dupeTitleCountMap } = context;
    const titleForRef = getTitleForRef(dupeTitleCountMap, item.title);
    const tabulation = tabulationMap[titleForRef];
    const title = body.appendParagraph(item.number + '. ' + item.title);
    title.asText().setBold(true);
    if (item.helpText)
        body.appendParagraph(item.helpText).editAsText().setBold(false);
    const styleGuard = body.appendParagraph('').editAsText().setBold(false);
    if (!tabulation) {
        const warning = body.appendParagraph('集計情報が見つかりません');
        warning.editAsText().setForegroundColor('#ff0000').setBold(true).setFontSize(20);
        return;
    }
    if (!tabulation.table[0][0].includes('自由記述')) {
        body.appendParagraph(`(n=${tabulation.n})`);
    }
    // table
    renderTable(body, tabulation.table);
    // Table always inserts next paragraph.
    const paragraphAfterTable = body.getChild(body.getNumChildren() - 1);
    renderQuestionBranching(context, page, item);
    paragraphAfterTable.removeFromParent();
    styleGuard.removeFromParent();
}
function renderQuestionItem(context, page, item) {
    var _a, _b, _c, _d;
    const { body } = context;
    const title = body.appendParagraph(item.number + '. ' + item.title);
    title.asText().setBold(true);
    if (item.helpText)
        body.appendParagraph(item.helpText).editAsText().setBold(false);
    const styleGuard = body.appendParagraph('').editAsText().setBold(false);
    // table
    const values = (_b = (_a = item.choices) === null || _a === void 0 ? void 0 : _a.map(choice => choice.value)) !== null && _b !== void 0 ? _b : [];
    if (values.length > 0 && values.every(v => isNumericString(v))) {
        const nums = values.map(v => Number(v));
        const min = nums.reduce((cur, v) => Math.min(cur, v), Infinity);
        const max = nums.reduce((cur, v) => Math.max(cur, v), -Infinity);
        if (nums.every((num, i) => num === i + min)) {
            // continuous number
            renderTable(body, [[`[   ] (${min}~${max}までの年齢を選択)`]]);
        }
        else {
            renderTable(body, (_c = values.map(value => [value])) !== null && _c !== void 0 ? _c : []);
        }
    }
    else {
        renderTable(body, (_d = values.map(value => [value])) !== null && _d !== void 0 ? _d : []);
    }
    // Table always inserts next paragraph.
    const paragraphAfterTable = body.getChild(body.getNumChildren() - 1);
    renderQuestionBranching(context, page, item);
    paragraphAfterTable.removeFromParent();
    styleGuard.removeFromParent();
}
function isNumericString(v) {
    return !!v && !Number.isNaN(Number(v));
}
function renderTable(body, tableData) {
    const paddingAttrs = {
        [DocumentApp.Attribute.PADDING_TOP]: 2,
        [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
        [DocumentApp.Attribute.PADDING_RIGHT]: 8,
        [DocumentApp.Attribute.PADDING_LEFT]: 8,
    };
    const table = body.appendTable(tableData);
    const numRows = table.getNumRows();
    for (let i = 0; i < numRows; i++) {
        const row = table.getRow(i);
        const numCells = row.getNumCells();
        for (let j = 0; j < numCells; j++) {
            const cell = table.getCell(i, j);
            cell.setAttributes(paddingAttrs);
        }
    }
}
// question branching
function renderQuestionBranching(context, page, item) {
    var _a, _b;
    const { body, pageIndexToQuestionNumberMap, pageIndexToLastQuestionNumberMap } = context;
    const getGoToQuestionNumber = (goTo) => goTo != null ? pageIndexToQuestionNumberMap[goTo] : item.number + 1;
    const isGoToNextQuestion = (goTo) => getGoToQuestionNumber(goTo) === item.number + 1;
    // There is no interface to access goToPage of other option: https://issuetracker.google.com/issues/36763602
    // Use go to page of previous choice.
    const choices = (_b = (_a = item.choices) === null || _a === void 0 ? void 0 : _a.map((choice, i) => (choice.isOther ? Object.assign(Object.assign({}, choice), { goTo: item.choices[i - 1].goTo }) : choice))) !== null && _b !== void 0 ? _b : [];
    const hasChoiceGoTo = choices.some(choice => !isGoToNextQuestion(choice.goTo));
    const isLastQuestion = pageIndexToLastQuestionNumberMap[page.index] === item.number;
    let branches = [];
    // 選択肢自体に遷移先設定があれば、すべて表示
    if (hasChoiceGoTo) {
        branches.push(...choices);
    }
    // 無回答の場合は、ページ内の最後の質問にだけ表示
    // 選択肢自体に遷移先設定がある場合は、単に次の質問へ遷移する場合も表示
    if (isLastQuestion && (hasChoiceGoTo || !isGoToNextQuestion(page.defaultGoTo))) {
        branches.push({ value: '無回答', goTo: page.defaultGoTo });
    }
    const lines = [];
    let hasError = false;
    // 選択肢自体に遷移先設定がある場合のみ、すべて一緒ならまとめて表示
    if (hasChoiceGoTo && new Set(branches.map(choice => getGoToQuestionNumber(choice.goTo))).size === 1) {
        const questionNumber = getGoToQuestionNumber(branches[0].goTo);
        if (questionNumber == null) {
            lines.push(`ページ index ${branches[0].goTo} に設問がありません`);
            hasError = true;
        }
        lines.push(`質問 ${questionNumber} に進む`);
    }
    else {
        for (const choice of branches) {
            const goTo = choice.goTo;
            const questionNumber = getGoToQuestionNumber(goTo);
            if (questionNumber == null) {
                lines.push(`ページ index ${goTo} に設問がありません`);
                hasError = true;
            }
            lines.push(`${choice.value}: 質問 ${questionNumber} に進む`);
        }
    }
    const p = body.appendParagraph(lines.join('\n'));
    if (hasError) {
        p.editAsText().setBold(true).setForegroundColor('#ff0000');
    }
    body.appendParagraph('');
}
function renderSectionHeader(context, item) {
    const { body } = context;
    const title = body.appendParagraph(item.title);
    title.asText().setBold(true);
    if (item.helpText)
        body.appendParagraph(item.helpText).editAsText().setBold(false);
    body.appendParagraph('').editAsText().setBold(false);
}
function renderImageItem(context, item) {
    const { body } = context;
    const title = body.appendParagraph(item.title);
    title.asText().setBold(true);
    if (item.helpText)
        body.appendParagraph(item.helpText).editAsText().setBold(false);
    const styleGuard = body.appendParagraph('').editAsText().setBold(false);
    const imageContainer = body.appendParagraph('');
    const image = imageContainer.appendInlineImage(item.imageBlob);
    const height = image.getHeight();
    const width = image.getWidth();
    const newWidth = Math.min(width, 480);
    image.setWidth(newWidth);
    image.setHeight((newWidth / width) * height);
    body.appendParagraph('');
    imageContainer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    styleGuard.removeFromParent();
}
function getTitleForRef(dupTitleCountMap, title) {
    var _a;
    (_a = dupTitleCountMap[title]) !== null && _a !== void 0 ? _a : (dupTitleCountMap[title] = 0);
    const titleCount = ++dupTitleCountMap[title];
    return titleCount === 1 ? title : `${title} -- ${titleCount}`;
}
function loadFormPages() {
    var _a;
    const form = FormApp.getActiveForm();
    const pages = [{ index: -1, title: '', description: '', items: [] }]; // initial page
    let lastPageBreak = undefined;
    let questionNumber = 1;
    for (const item of form.getItems()) {
        if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
            const pageBreakItem = item.asPageBreakItem();
            lastPageBreak = pageBreakItem;
            pages[pages.length - 1].defaultGoTo = (_a = pageBreakItem.getGoToPage()) === null || _a === void 0 ? void 0 : _a.getIndex();
            pages.push({
                index: item.getIndex(),
                title: item.getTitle(),
                description: item.getHelpText(),
                items: [],
            });
            continue;
        }
        const page = pages[pages.length - 1];
        if (itemAcceptsResponseSet.has(item.getType())) {
            const choices = item.getType() === FormApp.ItemType.SCALE
                ? (() => {
                    const scale = item.asScaleItem();
                    const values = [...Array(scale.getUpperBound() - scale.getLowerBound() + 1)]
                        .map((_, i) => i + scale.getLowerBound())
                        .map(num => num.toString());
                    values[0] += ` (${scale.getLeftLabel()})`;
                    values[values.length - 1] += ` (${scale.getRightLabel()})`;
                    return values.map(value => ({ value }));
                })()
                : (() => {
                    var _a;
                    const parsed = getTitleAndChoicesWithOther(lastPageBreak, item);
                    if (parsed.isText)
                        return [{ value: '(自由記述)' }];
                    const result = (_a = parsed.choices) === null || _a === void 0 ? void 0 : _a.map((c) => ({
                        value: c.value,
                        goTo: c.goTo,
                    }));
                    if (parsed.hasOther)
                        result === null || result === void 0 ? void 0 : result.push({ value: 'その他', isOther: true });
                    return result;
                })();
            page.items.push({
                kind: 'question',
                title: item.getTitle(),
                helpText: item.getHelpText(),
                choices,
                number: questionNumber++,
            });
            continue;
        }
        if (item.getType() == FormApp.ItemType.SECTION_HEADER) {
            const sectionHeaderItem = item.asSectionHeaderItem();
            page.items.push({
                kind: 'sectionHeader',
                title: sectionHeaderItem.getTitle(),
                helpText: sectionHeaderItem.getHelpText(),
            });
            continue;
        }
        if (item.getType() == FormApp.ItemType.IMAGE) {
            const imageItem = item.asImageItem();
            page.items.push({
                kind: 'image',
                title: imageItem.getTitle(),
                helpText: imageItem.getHelpText(),
                imageBlob: imageItem.getImage(),
            });
            continue;
        }
        Logger.log(`Unsupported item type ${item.getType()}`);
    }
    // exclude question
    for (const page of pages) {
        if (page.title === '本調査に関する質問') {
            page.items = page.items.filter(item => item.title !== 'ご意見等をご記入された方は、回答の公開の可否をお答えください。');
        }
    }
    return pages;
}
function loadTabulations() {
    const json = DriveApp.getFileById(getIdFromUrl(TABULATION_FILE_URL)).getBlob().getDataAsString();
    return JSON.parse(json);
}
// https://stackoverflow.com/a/16840612/1474113
function getIdFromUrl(url) {
    const m = url.match(/[-\w]{25,}/);
    if (!m)
        throw new Error('Invalid drive URL');
    return m[0];
}
