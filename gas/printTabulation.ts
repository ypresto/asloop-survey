/// <reference types="google-apps-script" />

const TABULATION_FILE_URL = 'https://drive.google.com/file/d/1JVa9eE67qMy_mLosfK3thuh-gjDlD05y/view?usp=share_link'
const START_FROM_INDEX = 6

type QuestionItemType = {
  kind: 'question'
  title: string
  helpText: string
  choices?: { value: string; goTo?: number }[]
  number: number
}
type ImageItemType = {
  kind: 'image'
  title: string
  helpText: string
  imageBlob: GoogleAppsScript.Base.Blob
}
type ItemType = QuestionItemType | ImageItemType

type PageType = {
  index: number
  title: string
  description: string
  defaultGoTo?: number
  items: ItemType[]
}

type ContextType = {
  body: GoogleAppsScript.Document.Body
  tabulationMap: Record<string, TabulationType>
}

function printTabulation() {
  const form = FormApp.getActiveForm()
  const pages = loadFormPages()
  const tabulations = loadTabulations()
  const tabulationMap = Object.fromEntries(tabulations.map(t => [t.title, t]))

  const body = DocumentApp.create(`${form.getTitle()}-単純集計結果`).getBody()
  const context = { body, tabulationMap }

  for (const page of pages) {
    renderPage(context, page)
  }
}

function renderPage(context: ContextType, page: PageType) {
  const { body } = context

  const heading = body.appendParagraph(page.title)
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  if (page.description) body.appendParagraph(page.description)
  body.appendParagraph('')

  for (const item of page.items) {
    switch (item.kind) {
      case 'question':
        renderQuestionItem(context, item)
        break
      case 'image':
        renderImageItem(context, item)
        break
    }
  }
}

function renderQuestionItem(context: ContextType, item: QuestionItemType) {
  const { body, tabulationMap } = context
  const tabulation = tabulationMap[item.title]

  const title = body.appendParagraph(item.number + '. ' + item.title)
  title.setAttributes({ [DocumentApp.Attribute.BOLD]: true })

  if (!tabulation) {
    const warning = body.appendParagraph('集計情報が見つかりません')
    warning.asText().setForegroundColor('#ff0000').setBold(true).setFontSize(20)
    return
  }
  body.appendParagraph(`(n=${tabulation.n})`)

  body.appendTable(tabulation.table)
  body.appendParagraph('')
}

function renderImageItem(context: ContextType, item: ImageItemType) {
  const { body } = context

  body.appendParagraph(item.title)
  body.appendImage(item.imageBlob)
  body.appendParagraph('')
}

function loadFormPages() {
  const form = FormApp.getActiveForm()

  const pages: PageType[] = [{ index: -1, title: '', description: '', items: [] }] // initial page

  let lastPageBreak: GoogleAppsScript.Forms.PageBreakItem | undefined = undefined
  let questionNumber = 1

  for (const item of form.getItems()) {
    if (item.getIndex() < START_FROM_INDEX) continue

    if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
      lastPageBreak = item.asPageBreakItem()
      pages.push({
        index: item.getIndex(),
        title: item.getTitle(),
        description: item.getHelpText(),
        defaultGoTo: item.asPageBreakItem().getGoToPage()?.getIndex(),
        items: [],
      })
    }

    const page = pages[pages.length - 1]
    if (itemAcceptsResponseSet.has(item.getType())) {
      page.items.push({
        kind: 'question',
        title: item.getTitle(),
        helpText: item.getHelpText(),
        choices: getTitleAndChoicesWithOther(lastPageBreak, item).choices?.map(c => ({ value: c.value, goTo: c.goTo })),
        number: questionNumber++,
      })
      continue
    }

    if (item.getType() == FormApp.ItemType.IMAGE) {
      const imageItem = item.asImageItem()
      page.items.push({
        kind: 'image',
        title: imageItem.getTitle(),
        helpText: imageItem.getHelpText(),
        imageBlob: imageItem.getImage(),
      })
    }
  }

  return pages
}

interface TabulationType {
  title: string
  n: number
  table: string[][] // rows -> cells
}

function loadTabulations() {
  const json = DriveApp.getFileById(getIdFromUrl(TABULATION_FILE_URL)).getBlob().getDataAsString()
  return JSON.parse(json) as TabulationType[]
}

// https://stackoverflow.com/a/16840612/1474113
function getIdFromUrl(url: string) {
  const m = url.match(/[-\w]{25,}/)
  if (!m) throw new Error('Invalid drive URL')
  return m[0]
}
