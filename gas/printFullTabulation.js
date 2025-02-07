"use strict";
/// <reference types="google-apps-script" />
const FULL_TABULATION_ZIP_FILE_URL = 'https://drive.google.com/file/d/1WOblm8N7eh0doZxbnKYlC0kvXT-PcHN4/view?usp=drive_link';
const FULL_TABULATION_OUTPUT_DOCUMENT_URL = 'https://docs.google.com/document/d/1YZNE9xg09qxEaShnAc6XmsgXlcDBASU-Y_pTRURzz90/edit';
function printFullTabulationPage1() {
    printFullTabulationImpl(1);
}
function printFullTabulationPage2() {
    printFullTabulationImpl(2);
}
function printFullTabulationImpl(page) {
    var _a, _b, _c, _d;
    const batchSize = 200;
    const allRecords = loadFullTabulationRecords();
    const records = allRecords.slice((page - 1) * batchSize, page * batchSize);
    const prevRecord = allRecords[(page - 1) * batchSize - 1];
    const doc = DocumentApp.openByUrl(FULL_TABULATION_OUTPUT_DOCUMENT_URL);
    const body = doc.getBody();
    body.appendParagraph('Ⅱ 調査の結果').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    let currentH1 = (_a = prevRecord === null || prevRecord === void 0 ? void 0 : prevRecord.h1) !== null && _a !== void 0 ? _a : '';
    let currentH2 = (_b = prevRecord === null || prevRecord === void 0 ? void 0 : prevRecord.h2) !== null && _b !== void 0 ? _b : '';
    let currentH3 = (_c = prevRecord === null || prevRecord === void 0 ? void 0 : prevRecord.h3) !== null && _c !== void 0 ? _c : '';
    for (const tabulation of records) {
        if (tabulation.h1 !== currentH1) {
            if (tabulation.h1) {
                body.appendParagraph(tabulation.h1).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            }
            currentH1 = tabulation.h1;
        }
        if (tabulation.h2 !== currentH2) {
            if (tabulation.h2) {
                body.appendParagraph(tabulation.h2).setHeading(DocumentApp.ParagraphHeading.HEADING3);
            }
            currentH2 = tabulation.h2;
        }
        if (tabulation.h3 !== currentH3) {
            if (tabulation.h3) {
                body.appendParagraph(tabulation.h3).setHeading(DocumentApp.ParagraphHeading.HEADING4);
            }
            currentH3 = tabulation.h3;
        }
        tabulation.body.split('\n').forEach(line => body.appendParagraph('　' + line));
        body.appendParagraph('');
        for (const figure of (_d = tabulation.figures) !== null && _d !== void 0 ? _d : []) {
            if (!figure.table && !figure.imageName)
                throw new Error('Figure must have table or image');
            const figTitle = body.appendParagraph(figure.title);
            if (figure.table) {
                renderTable(body, figure.table.map(row => row.map(cell => cell.toString())));
                body.appendParagraph(''); // also work as style guard
            }
            if (figure.imageName) {
                renderFigureImage(body, figure.imageName);
                body.appendParagraph(''); // also work as style guard
            }
            figTitle.editAsText().setBold(true).setFontSize(9).setForegroundColor('#666666');
        }
    }
    Logger.log('Document created on %s', doc.getUrl());
}
function renderFigureImage(body, imageName) {
    const imageContainer = body.appendParagraph('');
    _renderFigureImageInContainer(body, imageContainer, imageName);
}
function _renderFigureImageInContainer(body, imageContainer, imageName) {
    const image = imageContainer.appendInlineImage(getFullTabBlobNamed(imageName));
    const height = image.getHeight();
    const width = image.getWidth();
    const contentWidth = Math.min(width, ((body.getPageWidth() - body.getMarginLeft() - body.getMarginRight()) / 72) * 96);
    const scale = contentWidth / width;
    const scaledHeight = height * scale;
    if (scaledHeight > 480) {
        const scale = 480 / height;
        image.setWidth(width * scale);
        image.setHeight(480);
    }
    else {
        image.setWidth(contentWidth);
        image.setHeight(scaledHeight);
    }
}
function loadFullTabulationRecords() {
    const json = getFullTabBlobNamed('full_tabulation.json').getDataAsString();
    return JSON.parse(json);
}
let cachedFileMap;
function getFullTabBlobNamed(name) {
    if (!cachedFileMap) {
        const zipFile = DriveApp.getFileById(getIdFromUrl(FULL_TABULATION_ZIP_FILE_URL));
        const blobs = Utilities.unzip(zipFile.getBlob());
        cachedFileMap = Object.fromEntries(blobs.map(blob => [blob.getName(), blob]));
    }
    const file = cachedFileMap[`full_tabulation/${name}`];
    if (!file)
        throw new Error(`File not found: ${name}`);
    return file;
}
