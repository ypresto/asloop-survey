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
  const titleAndChoicesMap = form.getItems()
    .filter(item => itemAcceptsResponseSet.has(item.getType()))
    .map(item => getTitleAndChoicesWithOther(item))

  titleAndChoicesMap.forEach(({ title, choices }) => {
    if (choices?.some(c => c.includes(','))) {
      throw new Error(`Item "${title}" has a comma in one of its choices.`)
    }
  })

  const spreadSheet = SpreadsheetApp.create(`${form.getTitle()}-choices`, titleAndChoicesMap.length + 1, 4)
  const sheet = spreadSheet.getSheets()[0]

  sheet.getRange(1, 1, 1, 4).setValues([['設問文章', '選択肢', '複数回答', '自由記述']])
  titleAndChoicesMap.forEach(({ title, choices, hasMultiple, hasOther }, i) => {
    sheet.getRange(i + 2, 1, 1, 4).setValues([[title, choices?.join(',') ?? '', hasMultiple ? 1 : 0, hasOther ? 1 : 0]])
  })
  sheet.setColumnWidth(2, 1200)

  Logger.log('SpreadSheet created on %s', spreadSheet.getUrl())

}

function getTitleAndChoicesWithOther(item: GoogleAppsScript.Forms.Item) {
  let choices: string[] | null = null;
  let hasMultiple = false
  let hasOther = false

  switch (item.getType()) {
    case FormApp.ItemType.LIST: {
      const concreteItem = item.asListItem()
      choices = concreteItem.getChoices().map(c => c.getValue())
    }
    break
    case FormApp.ItemType.MULTIPLE_CHOICE: { // Radio Buttons
      const concreteItem = item.asMultipleChoiceItem()
      choices = concreteItem.getChoices().map(c => c.getValue())
      hasOther = concreteItem.hasOtherOption()
    }
    break
    case FormApp.ItemType.CHECKBOX: {
      const concreteItem = item.asCheckboxItem()
      choices =  concreteItem.getChoices().map(c => c.getValue())
      hasMultiple = true
      hasOther = concreteItem.hasOtherOption()
    }
  }

  return { title: item.getTitle(), choices, hasMultiple, hasOther }
}
