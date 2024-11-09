"use strict";
/// <reference types="google-apps-script" />
const imagesToReplace = [
    'figure-10-a.png',
    'figure-10-b.png',
    'figure-11-a.png',
    'figure-11-b.png',
    'figure-12-a.png',
    'figure-12-b.png',
    'figure-13-a.png',
    'figure-13-b.png',
    'figure-14-a.png',
    'figure-14-b.png',
    'figure-15-a.png',
    'figure-15-b.png',
    'figure-16-a.png',
    'figure-17-a.png',
    'figure-18-a.png',
    'figure-19-a.png',
    'figure-20-a.png',
    'figure-21-a.png',
    'figure-22-a.png',
    'figure-23-a.png',
    'figure-24-a.png',
    'figure-24-b.png',
    'figure-24-c.png',
    'figure-24-d.png',
    'figure-24-e.png',
    'figure-24-f.png',
    'figure-24-g.png',
    'figure-24.png',
    'figure-25-a.png',
    'figure-26-a.png',
    'figure-26-b.png',
    'figure-27-a.png',
    'figure-29-a.png',
    'figure-30-a.png',
    'figure-31-a.png',
    'figure-32-a.png',
    'figure-33-a.png',
    'figure-34-a.png',
    'figure-35-a.png',
    'figure-36-a.png',
    'figure-37-a.png',
    'figure-38-a.png',
    'figure-39.png',
    'figure-40-a.png',
    'figure-41-a.png',
    'figure-43-a.png',
    'figure-44.png',
    'figure-45-a.png',
    'figure-45-b.png',
    'figure-46-a.png',
    'figure-46-b.png',
    'figure-47-a.png',
    'figure-47-b.png',
    'figure-49-a.png',
    'figure-49-b.png',
    'figure-5-a.png',
    'figure-5-b.png',
    'figure-50-a.png',
    'figure-50-b.png',
    'figure-52-a.png',
    'figure-52-b.png',
    'figure-53-a.png',
    'figure-53-b.png',
    'figure-54-a.png',
    'figure-54-b.png',
    'figure-55-a.png',
    'figure-55-b.png',
    'figure-56-a.png',
    'figure-56-b.png',
    'figure-57-a.png',
    'figure-57-b.png',
    'figure-58-a.png',
    'figure-58-b.png',
    'figure-59-a.png',
    'figure-59-b.png',
    'figure-6-a.png',
    'figure-6-b.png',
    'figure-60-a.png',
    'figure-60-b.png',
    'figure-61-a.png',
    'figure-61-b.png',
    'figure-62-a.png',
    'figure-62-b.png',
    'figure-63-a.png',
    'figure-63-b.png',
    'figure-64-a.png',
    'figure-64-b.png',
    'figure-65-a.png',
    'figure-65-b.png',
    'figure-66-a.png',
    'figure-66-b.png',
    'figure-67-a.png',
    'figure-67-b.png',
    'figure-68-a.png',
    'figure-68-b.png',
    'figure-69-a.png',
    'figure-69-b.png',
    'figure-7-a.png',
    'figure-7-b.png',
    'figure-70-a.png',
    'figure-70-b.png',
    'figure-71-a.png',
    'figure-71-b.png',
    'figure-72-a.png',
    'figure-72-b.png',
    'figure-73-a.png',
    'figure-73-b.png',
    'figure-74-a.png',
    'figure-74-b.png',
    'figure-75-a.png',
    'figure-75-b.png',
    'figure-76-a.png',
    'figure-76-b.png',
    'figure-77-a.png',
    'figure-77-b.png',
    'figure-78-a.png',
    'figure-78-b.png',
    'figure-79-a.png',
    'figure-79-b.png',
    'figure-8-a.png',
    'figure-8-b.png',
    'figure-80-a.png',
    'figure-80-b.png',
    'figure-82-a.png',
    'figure-82-b.png',
    'figure-86-a.png',
    'figure-86-b.png',
    'figure-87-a.png',
    'figure-87-b.png',
    'figure-88-a.png',
    'figure-88-b.png',
    'figure-89-a.png',
    'figure-89-b.png',
    'figure-9-a.png',
    'figure-9-b.png',
    'figure-90-a.png',
    'figure-90-b.png',
    'figure-91-a.png',
    'figure-91-b.png',
    'figure-96-a.png',
    'figure-96-b.png',
    'figure-98-a.png',
    'figure-98-b.png',
    'figure-99-a.png',
    'figure-99-b.png',
    'figure-100-a.png',
    'figure-100-b.png',
    'figure-101-a.png',
    'figure-101-b.png',
    'figure-102-a.png',
    'figure-102-b.png',
    'figure-103-a.png',
    'figure-103-b.png',
    'figure-104-a.png',
    'figure-104-b.png',
    'figure-105-a.png',
    'figure-105-b.png',
    'figure-106-a.png',
    'figure-106-b.png',
    'figure-108-a.png',
    'figure-110-a.png',
    'figure-111-a.png',
    'figure-111-b.png',
    'figure-112-a.png',
    'figure-112-b.png',
];
function replaceImages() {
    const body = DocumentApp.openByUrl('https://docs.google.com/document/d/1VTRe6xdf9Jhru01QGWhcVvyWSMVlIqSTgA2Gukh2ZSw/edit').getBody();
    const map = Object.fromEntries(imagesToReplace.map(imageName => [imageName.replace(/^figure-|\.png$/g, ''), null]));
    const len = body.getNumChildren();
    for (let i = 0; i < len; i++) {
        const child = body.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.PARAGRAPH)
            continue;
        const m = child
            .asParagraph()
            .getText()
            .match(/^図表(\d+(?:-[a-z]+)?)/);
        if (!m)
            continue;
        const figNum = m[1];
        if (map[figNum]) {
            throw new Error(`figure ${figNum} is duplicated`);
        }
        if (map[figNum] === undefined)
            continue;
        if (child.getNextSibling().getType() !== DocumentApp.ElementType.PARAGRAPH) {
            throw new Error(`figure ${figNum} is not followed by a paragraph`);
        }
        const imageContainer = child.getNextSibling().asParagraph();
        const currentInline = imageContainer.getChild(0);
        if (currentInline.getType() !== DocumentApp.ElementType.INLINE_IMAGE) {
            throw new Error(`figure ${figNum} is not followed by an inline image`);
        }
        map[figNum] = { imageContainer, currentInline: currentInline.asInlineImage() };
    }
    for (const [figNum, obj] of Object.entries(map)) {
        if (!obj)
            throw new Error(`figure ${figNum} not found in document`);
    }
    Logger.log('all images found, start replacing...');
    for (const [figNum, obj] of Object.entries(map)) {
        if (!obj)
            throw new Error('assertion error');
        _renderFigureImageInContainer(body, obj.imageContainer, `figure-${figNum}.png`);
        obj.imageContainer.removeChild(obj.currentInline);
    }
}
