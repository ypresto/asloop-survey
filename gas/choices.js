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
    FormApp.ItemType.SCALE,
    FormApp.ItemType.TEXT,
    FormApp.ItemType.TIME,
    FormApp.ItemType.FILE_UPLOAD,
]);
function getChoices() {
    const form = FormApp.getActiveForm();
    const titleAndChoicesMap = form.getItems()
        .filter(item => itemAcceptsResponseSet.has(item.getType()))
        .map(item => getTitleAndChoicesWithOther(item));
    titleAndChoicesMap.forEach(([title, choices]) => {
        if (choices === null || choices === void 0 ? void 0 : choices.some(c => c.includes(','))) {
            throw new Error(`Item "${title}" has a comma in one of its choices.`);
        }
    });
    const spreadSheet = SpreadsheetApp.create(`${form.getTitle()}-choices-with-other`, titleAndChoicesMap.length + 1, 2);
    const sheet = spreadSheet.getSheets()[0];
    sheet.getRange(1, 1, 1, 2).setValues([['設問文章', '選択肢']]);
    titleAndChoicesMap.forEach(([title, choices], i) => {
        var _a;
        sheet.getRange(i + 2, 1, 1, 2).setValues([[title, (_a = choices === null || choices === void 0 ? void 0 : choices.join(',')) !== null && _a !== void 0 ? _a : '']]);
    });
    sheet.setColumnWidth(2, 1200);
    Logger.log('SpreadSheet created on %s', spreadSheet.getUrl());
}
function getTitleAndChoicesWithOther(item) {
    let choices = null;
    switch (item.getType()) {
        case FormApp.ItemType.MULTIPLE_CHOICE:
            { // Radio Buttons
                const concreteItem = item.asMultipleChoiceItem();
                if (concreteItem.hasOtherOption()) {
                    choices = concreteItem.getChoices().map(c => c.getValue());
                }
            }
            break;
        case FormApp.ItemType.CHECKBOX: {
            const concreteItem = item.asCheckboxItem();
            if (concreteItem.hasOtherOption()) {
                choices = concreteItem.getChoices().map(c => c.getValue());
            }
        }
    }
    return [item.getTitle(), choices];
}
