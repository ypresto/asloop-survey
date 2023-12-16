/// <reference types="google-apps-script" />

const FULL_TABULATION_ZIP_FILE_URL =
  'https://drive.google.com/file/d/1WOblm8N7eh0doZxbnKYlC0kvXT-PcHN4/view?usp=drive_link'

const FULL_TABULATION_OUTPUT_DOCUMENT_URL =
  'https://docs.google.com/document/d/102j2qrxxWck7KNW1mo7IioRRn-wUqMB97GmmL32pym0/edit'

interface FullTabulationFigureType {
  title: string
  imageName: string | null
  table: (string | number)[][]
}

interface FullTabulationRecordType {
  h1: string
  h2: string
  h3: string
  body: string
  description: string | null
  figures: FullTabulationFigureType[] | null
}

interface ContextFType {
  body: GoogleAppsScript.Document.Body
}

function printFullTabulationPage1() {
  printFullTabulationImpl(1)
}

function printFullTabulationPage2() {
  printFullTabulationImpl(2)
}

function printFullTabulationImpl(page: number) {
  const batchSize = 200
  const allRecords = loadFullTabulationRecords()
  const records = allRecords.slice((page - 1) * batchSize, page * batchSize)
  const prevRecord = allRecords[(page - 1) * batchSize - 1]

  const doc = DocumentApp.openByUrl(FULL_TABULATION_OUTPUT_DOCUMENT_URL)
  const body = doc.getBody()
  const context: ContextFType = { body }

  body.appendParagraph('Ⅱ 調査の結果').setHeading(DocumentApp.ParagraphHeading.HEADING1)

  let currentH1 = prevRecord?.h1 ?? ''
  let currentH2 = prevRecord?.h2 ?? ''
  let currentH3 = prevRecord?.h3 ?? ''
  for (const tabulation of records) {
    if (tabulation.h1 !== currentH1) {
      if (tabulation.h1) {
        body.appendParagraph(tabulation.h1).setHeading(DocumentApp.ParagraphHeading.HEADING2)
      }
      currentH1 = tabulation.h1
    }
    if (tabulation.h2 !== currentH2) {
      if (tabulation.h2) {
        body.appendParagraph(tabulation.h2).setHeading(DocumentApp.ParagraphHeading.HEADING3)
      }
      currentH2 = tabulation.h2
    }
    if (tabulation.h3 !== currentH3) {
      if (tabulation.h3) {
        body.appendParagraph(tabulation.h3).setHeading(DocumentApp.ParagraphHeading.HEADING4)
      }
      currentH3 = tabulation.h3
    }

    tabulation.body.split('\n').forEach(line => body.appendParagraph('　' + line))

    body.appendParagraph('')

    for (const figure of tabulation.figures ?? []) {
      body.appendParagraph(figure.title)
      if (figure.table) {
        renderTable(
          context.body,
          figure.table.map(row => row.map(cell => cell.toString()))
        )
      }
      if (figure.imageName) {
        renderFigureImage(context.body, figure.imageName)
      }
    }
  }

  Logger.log('Document created on %s', doc.getUrl())
}

function renderFigureImage(body: GoogleAppsScript.Document.Body, imageName: string) {
  const imageContainer = body.appendParagraph('')

  const image = imageContainer.appendInlineImage(getFullTabBlobNamed(imageName))
  const height = image.getHeight()
  const width = image.getWidth()
  const newWidth = Math.min(width, ((body.getPageWidth() - body.getMarginLeft() - body.getMarginRight()) / 72) * 96)
  image.setWidth(newWidth)
  image.setHeight((newWidth / width) * height)
  body.appendParagraph('')
}

function loadFullTabulationRecords() {
  const json = getFullTabBlobNamed('full_tabulation.json').getDataAsString()
  return JSON.parse(json) as FullTabulationRecordType[]
}

let cachedFileMap: Record<string, GoogleAppsScript.Base.Blob>

function getFullTabBlobNamed(name: string) {
  if (!cachedFileMap) {
    const zipFile = DriveApp.getFileById(getIdFromUrl(FULL_TABULATION_ZIP_FILE_URL))
    const blobs = Utilities.unzip(zipFile.getBlob())
    cachedFileMap = Object.fromEntries(blobs.map(blob => [blob.getName(), blob]))
  }
  const file = cachedFileMap[`full_tabulation/${name}`]
  if (!file) throw new Error(`File not found: ${name}`)
  return file
}
