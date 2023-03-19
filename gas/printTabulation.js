"use strict";
/// <reference types="google-apps-script" />
const TABULATION_FILE_URL = 'https://drive.google.com/file/d/1JVa9eE67qMy_mLosfK3thuh-gjDlD05y/view?usp=share_link';
// Some page break item returns self as getGoToPage().getIndex() while it is actually not
const DEFAULT_GO_TO_QUESTION_NUMBER_OVERRIDES = {
    // 7. 現在、収入をともなう仕事をしていますか。 → 12. 出生時の性別と、現在自分が捉えている性別が「一致」していると思いますか。
    7: 12,
    // 12. 出生時の性別と、現在自分が捉えている性別が「一致」していると思いますか。 → 16. 特定の人と「付き合いたい」と思うことがありますか。
    12: 16,
};
function printTabulation() {
    const pages = loadFormPages();
    const tabulations = loadTabulations();
    const tabulationMap = Object.fromEntries(tabulations.map(t => [t.title, t]));
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
    const context = { body, tabulationMap, pageIndexToQuestionNumberMap, pageIndexToLastQuestionNumberMap };
    body.appendParagraph('単純集計結果').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    // TODO: fix this
    const startPos = pages.findIndex(page => page.items.some(item => item.kind === 'question'));
    for (const page of pages.slice(startPos)) {
        renderPage(context, page, false);
    }
    body.appendParagraph('調査票').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(form.getTitle()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(form.getDescription());
    for (const page of pages) {
        renderPage(context, page, true);
    }
    Logger.log('Document created on %s', doc.getUrl());
}
function renderPage(context, page, isQuestionnaire) {
    const { body } = context;
    const heading = body.appendParagraph(page.title);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (page.description)
        body.appendParagraph(page.description);
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
    const { body, tabulationMap } = context;
    const tabulation = tabulationMap[item.title];
    const title = body.appendParagraph(item.number + '. ' + item.title);
    title.asText().setBold(true);
    if (!tabulation) {
        const warning = body.appendParagraph('集計情報が見つかりません');
        warning.editAsText().setForegroundColor('#ff0000').setBold(true).setFontSize(20);
        return;
    }
    body.appendParagraph(`(n=${tabulation.n})`).editAsText().setBold(false);
    // table
    renderTable(body, tabulation.table);
    // Table always inserts next paragraph.
    const paragraphAfterTable = body.getChild(body.getNumChildren() - 1);
    renderQuestionBranching(context, page, item);
    paragraphAfterTable.removeFromParent();
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
    var _a, _b, _c;
    const { body, pageIndexToQuestionNumberMap, pageIndexToLastQuestionNumberMap } = context;
    // ページのデフォルトの遷移先は、ページ内の最後の質問にだけ表示
    const isLastQuestion = pageIndexToLastQuestionNumberMap[page.index] === item.number;
    let branches = (_b = (_a = item.choices) === null || _a === void 0 ? void 0 : _a.slice()) !== null && _b !== void 0 ? _b : [];
    if (isLastQuestion) {
        const overrideNumber = DEFAULT_GO_TO_QUESTION_NUMBER_OVERRIDES[item.number];
        if (overrideNumber != null) {
            const index = (_c = Object.entries(pageIndexToQuestionNumberMap).find(([, num]) => num === overrideNumber)) === null || _c === void 0 ? void 0 : _c[0];
            if (index == null) {
                throw new Error(`Page index for question number ${item.number} is not found.`);
            }
            branches.push({ value: '回答しない', goTo: Number(index) });
        }
        else {
            branches.push({ value: '回答しない', goTo: page.defaultGoTo });
        }
    }
    let isNewlineInserted = false;
    // 全ての選択肢が同じ遷移先の場合は一括表示
    if (new Set(branches.map(choice => choice.goTo)).size === 1 && branches[0].goTo != null) {
        const goTo = branches[0].goTo;
        const questionNumber = pageIndexToQuestionNumberMap[goTo];
        if (questionNumber == null) {
            body
                .appendParagraph(`ページ index ${goTo} に設問がありません`)
                .editAsText()
                .setBold(true)
                .setForegroundColor('#ff0000');
        }
        body.appendParagraph(`質問 ${questionNumber} に進む`);
    }
    else {
        // 単に次の質問へ進むものは省く
        branches = branches.filter(choice => choice.goTo != null && pageIndexToQuestionNumberMap[choice.goTo] !== item.number + 1);
        if (branches.length > 0) {
            const lines = [];
            let hasError = false;
            for (const choice of branches) {
                const goTo = choice.goTo;
                if (goTo != null) {
                    const questionNumber = pageIndexToQuestionNumberMap[goTo];
                    if (questionNumber == null) {
                        lines.push(`ページ index ${goTo} に設問がありません`);
                        hasError = true;
                    }
                    if (questionNumber === item.number) {
                        lines.push(`ページ index ${goTo} はこの設問自体を指しています`);
                        hasError = true;
                    }
                    lines.push(`${choice.value}: 質問 ${questionNumber} に進む`);
                }
            }
            const p = body.appendParagraph(lines.join('\n'));
            body.appendParagraph(''); // prevent from propagating style to next paragraph
            if (hasError) {
                p.editAsText().setBold(true).setForegroundColor('#ff0000');
            }
            isNewlineInserted = true;
        }
    }
    if (!isNewlineInserted) {
        body.appendParagraph('');
    }
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
            pages.push({
                index: item.getIndex(),
                title: item.getTitle(),
                description: item.getHelpText(),
                defaultGoTo: (_a = pageBreakItem.getGoToPage()) === null || _a === void 0 ? void 0 : _a.getIndex(),
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
                    values[0] += ` (${scale.getRightLabel()})`;
                    values[values.length - 1] += ` (${scale.getRightLabel()})`;
                    return values.map(value => ({ value }));
                })()
                : (() => {
                    var _a;
                    const parsed = getTitleAndChoicesWithOther(lastPageBreak, item);
                    if (parsed.isText)
                        return [{ value: '(自由記述)' }];
                    return (_a = parsed.choices) === null || _a === void 0 ? void 0 : _a.map(c => ({
                        value: c.value,
                        goTo: c.goTo,
                    }));
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
