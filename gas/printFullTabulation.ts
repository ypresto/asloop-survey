/// <reference types="google-apps-script" />

const FULL_TABULATION_ZIP_FILE_URL =
  'https://drive.google.com/file/d/1WOblm8N7eh0doZxbnKYlC0kvXT-PcHN4/view?usp=drive_link'

interface FullTabulationFigureType {
  title: string
  imageName: string | null
  table: (string | number)[][]
}

interface FullTabulationRecordType {
  h1: string
  h2: string
  title: string
  body: string
  description: string | null
  figures: FullTabulationFigureType[] | null
}

interface ContextFType {
  body: GoogleAppsScript.Document.Body
}

function printFullTabulation() {
  const records = loadFullTabulationRecords()

  const doc = DocumentApp.create('調査結果報告書2022')
  const body = doc.getBody()
  const context: ContextFType = { body }

  body.appendParagraph('Ⅱ 調査の結果').setHeading(DocumentApp.ParagraphHeading.HEADING1)

  let currentH1 = ''
  let currentH2 = ''
  for (const tabulation of records) {
    if (tabulation.h1 !== currentH1) {
      body.appendParagraph(tabulation.h1).setHeading(DocumentApp.ParagraphHeading.HEADING2)
      currentH1 = tabulation.h1
    }
    if (tabulation.h2 !== currentH2) {
      body.appendParagraph(tabulation.h2).setHeading(DocumentApp.ParagraphHeading.HEADING3)
      currentH2 = tabulation.h2
    }
    body.appendParagraph(tabulation.body)
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
  const newWidth = Math.min(width, 480)
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
