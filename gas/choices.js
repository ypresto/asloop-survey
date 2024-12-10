"use strict";
/// <reference types="google-apps-script" />
const itemAcceptsResponseSet = new Set([
    FormApp.ItemType.CHECKBOX,
    FormApp.ItemType.CHECKBOX_GRID,
    FormApp.ItemType.DATE,
    FormApp.ItemType.DATETIME,
    FormApp.ItemType.DURATION,
    FormApp.ItemType.GRID,
    FormApp.ItemType.LIST,
    FormApp.ItemType.MULTIPLE_CHOICE,
    FormApp.ItemType.PARAGRAPH_TEXT,
    FormApp.ItemType.SCALE,
    FormApp.ItemType.TEXT,
    FormApp.ItemType.TIME,
    FormApp.ItemType.FILE_UPLOAD,
]);
function getChoices() {
    const form = FormApp.getActiveForm();
    let lastPageBreak;
    const titleAndChoicesMap = [];
    for (const item of form.getItems()) {
        if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
            lastPageBreak = item.asPageBreakItem();
            continue;
        }
        if (itemAcceptsResponseSet.has(item.getType())) {
            titleAndChoicesMap.push(getTitleAndChoicesWithOther(lastPageBreak, item));
        }
    }
    titleAndChoicesMap.forEach(({ title, choices }) => {
        if (choices === null || choices === void 0 ? void 0 : choices.some(c => c.value.includes(','))) {
            throw new Error(`Item "${title}" has a comma in one of its choices.`);
        }
    });
    const spreadSheet = SpreadsheetApp.create(`${form.getTitle()}-choices`, titleAndChoicesMap.length + 1, 6);
    const sheet = spreadSheet.getSheets()[0];
    sheet.getRange(1, 1, 1, 8).setValues([['ページ番号', '設問文章', '説明文', '選択肢', '複数回答', 'その他回答', '自由記述', 'ジャンプ先ページ番号マップ']]);
    titleAndChoicesMap.forEach(({ pageIndex, title, description, choices, hasMultiple, hasOther, isText }, i) => {
        var _a;
        const pageMap = (choices === null || choices === void 0 ? void 0 : choices.length) ? JSON.stringify(choices.map(c => { var _a; return (_a = c.goTo) !== null && _a !== void 0 ? _a : null; })) : '';
        sheet.getRange(i + 2, 1, 1, 8).setValues([[pageIndex, title, description, (_a = choices === null || choices === void 0 ? void 0 : choices.map(c => c.value).join(',')) !== null && _a !== void 0 ? _a : '', hasMultiple ? 1 : 0, hasOther ? 1 : 0, isText ? 1 : 0, pageMap]]);
    });
    sheet.setColumnWidth(2, 1200);
    Logger.log('SpreadSheet created on %s', spreadSheet.getUrl());
}
function getTitleAndChoicesWithOther(pageBreak, item) {
    let choices = undefined;
    let hasMultiple = false;
    let hasOther = false;
    let isText = false;
    switch (item.getType()) {
        case FormApp.ItemType.LIST:
            {
                const concreteItem = item.asListItem();
                choices = concreteItem.getChoices().map(c => { var _a; return ({ value: c.getValue(), goTo: (_a = c.getGotoPage()) === null || _a === void 0 ? void 0 : _a.getIndex() }); });
            }
            break;
        case FormApp.ItemType.MULTIPLE_CHOICE:
            { // Radio Buttons
                const concreteItem = item.asMultipleChoiceItem();
                choices = concreteItem.getChoices().map(c => { var _a; return ({ value: c.getValue(), goTo: (_a = c.getGotoPage()) === null || _a === void 0 ? void 0 : _a.getIndex() }); });
                hasOther = concreteItem.hasOtherOption();
            }
            break;
        case FormApp.ItemType.CHECKBOX:
            {
                const concreteItem = item.asCheckboxItem();
                choices = concreteItem.getChoices().map(c => { var _a; return ({ value: c.getValue(), goTo: (_a = c.getGotoPage()) === null || _a === void 0 ? void 0 : _a.getIndex() }); });
                hasMultiple = true;
                hasOther = concreteItem.hasOtherOption();
            }
            break;
        case FormApp.ItemType.TEXT:
        case FormApp.ItemType.PARAGRAPH_TEXT:
            isText = true;
    }
    return { pageIndex: pageBreak === null || pageBreak === void 0 ? void 0 : pageBreak.getIndex(), title: item.getTitle(), description: item.getHelpText(), choices, hasMultiple, hasOther, isText };
}
