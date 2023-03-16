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
])

function getChoices() {
  const form = FormApp.getActiveForm()

  let lastPageBreak: GoogleAppsScript.Forms.PageBreakItem | undefined

  const titleAndChoicesMap: ReturnType<typeof getTitleAndChoicesWithOther>[] = []

  for (const item of form.getItems()) {
    if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
      lastPageBreak = item.asPageBreakItem()
      continue
    }

    if (itemAcceptsResponseSet.has(item.getType())) {
      titleAndChoicesMap.push(getTitleAndChoicesWithOther(lastPageBreak, item))
    }
  }

  titleAndChoicesMap.forEach(({ title, choices }) => {
    if (choices?.some(c => c.value.includes(','))) {
      throw new Error(`Item "${title}" has a comma in one of its choices.`)
    }
  })

  const spreadSheet = SpreadsheetApp.create(`${form.getTitle()}-choices`, titleAndChoicesMap.length + 1, 6)
  const sheet = spreadSheet.getSheets()[0]

  sheet.getRange(1, 1, 1, 7).setValues([['ページ番号', '設問文章', '説明文', '選択肢', '複数回答', '自由記述', 'ジャンプ先ページ番号マップ']])
  titleAndChoicesMap.forEach(({ pageIndex, title, description, choices, hasMultiple, hasOther }, i) => {
    const pageMap = choices?.length ? JSON.stringify(choices.map(c => c.goTo ?? null)) : ''
    sheet.getRange(i + 2, 1, 1, 7).setValues([[pageIndex, title, description, choices?.map(c => c.value).join(',') ?? '', hasMultiple ? 1 : 0, hasOther ? 1 : 0, pageMap]])
  })
  sheet.setColumnWidth(2, 1200)

  Logger.log('SpreadSheet created on %s', spreadSheet.getUrl())

}

function getTitleAndChoicesWithOther(pageBreak: GoogleAppsScript.Forms.PageBreakItem | undefined,  item: GoogleAppsScript.Forms.Item) {
  let choices: { value: string, goTo?: number }[] | null = null;
  let hasMultiple = false
  let hasOther = false

  switch (item.getType()) {
    case FormApp.ItemType.LIST: {
      const concreteItem = item.asListItem()
      choices = concreteItem.getChoices().map(c => ({ value: c.getValue(), goTo: c.getGotoPage()?.getIndex() }))
    }
    break
    case FormApp.ItemType.MULTIPLE_CHOICE: { // Radio Buttons
      const concreteItem = item.asMultipleChoiceItem()
      choices = concreteItem.getChoices().map(c => ({ value: c.getValue(), goTo: c.getGotoPage()?.getIndex() }))
      hasOther = concreteItem.hasOtherOption()
    }
    break
    case FormApp.ItemType.CHECKBOX: {
      const concreteItem = item.asCheckboxItem()
      choices = concreteItem.getChoices().map(c => ({ value: c.getValue(), goTo: c.getGotoPage()?.getIndex() }))
      hasMultiple = true
      hasOther = concreteItem.hasOtherOption()
    }
  }

  return { pageIndex: pageBreak?.getIndex(), title: item.getTitle(), description: item.getHelpText(), choices, hasMultiple, hasOther }
}
