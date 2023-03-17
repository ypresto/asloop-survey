"use strict";
/// <reference types="google-apps-script" />
const TABULATION_FILE_URL = 'https://drive.google.com/file/d/1JVa9eE67qMy_mLosfK3thuh-gjDlD05y/view?usp=share_link';
const START_FROM_INDEX = 6;
function printTabulation() {
    const form = FormApp.getActiveForm();
    const pages = loadFormPages();
    const tabulations = loadTabulations();
    const tabulationMap = Object.fromEntries(tabulations.map(t => [t.title, t]));
    const body = DocumentApp.create(`${form.getTitle()}-単純集計結果`).getBody();
    const context = { body, tabulationMap };
    for (const page of pages) {
        renderPage(context, page);
    }
}
function renderPage(context, page) {
    const { body } = context;
    const heading = body.appendParagraph(page.title);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (page.description)
        body.appendParagraph(page.description);
    body.appendParagraph('');
    for (const item of page.items) {
        switch (item.kind) {
            case 'question':
                renderQuestionItem(context, item);
                break;
            case 'image':
                renderImageItem(context, item);
                break;
        }
    }
}
function renderQuestionItem(context, item) {
    const { body, tabulationMap } = context;
    const tabulation = tabulationMap[item.title];
    const heading = body.appendParagraph(item.title);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    if (!tabulation) {
        const warning = body.appendParagraph('集計情報が見つかりません');
        warning.asText().setForegroundColor('#ff0000').setBold(true).setFontSize(20);
        return;
    }
    body.appendParagraph(`(n=${tabulation.n})`);
    body.appendTable(tabulation.table);
    body.appendParagraph('');
}
function renderImageItem(context, item) {
    const { body } = context;
    body.appendParagraph(item.title);
    body.appendImage(item.imageBlob);
    body.appendParagraph('');
}
function loadFormPages() {
    var _a, _b;
    const form = FormApp.getActiveForm();
    const pages = [{ index: -1, title: '', description: '', items: [] }]; // initial page
    let lastPageBreak = undefined;
    let questionNumber = 1;
    for (const item of form.getItems()) {
        if (item.getIndex() < START_FROM_INDEX)
            continue;
        if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
            lastPageBreak = item.asPageBreakItem();
            pages.push({
                index: item.getIndex(),
                title: item.getTitle(),
                description: item.getHelpText(),
                defaultGoTo: (_a = item.asPageBreakItem().getGoToPage()) === null || _a === void 0 ? void 0 : _a.getIndex(),
                items: [],
            });
        }
        const page = pages[pages.length - 1];
        if (itemAcceptsResponseSet.has(item.getType())) {
            page.items.push({
                kind: 'question',
                title: item.getTitle(),
                helpText: item.getHelpText(),
                choices: (_b = getTitleAndChoicesWithOther(lastPageBreak, item).choices) === null || _b === void 0 ? void 0 : _b.map(c => ({ value: c.value, goTo: c.goTo })),
                number: questionNumber++,
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
