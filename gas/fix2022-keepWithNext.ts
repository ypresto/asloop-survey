/// <reference types="google-apps-script" />

function updateToKeepWithNext() {
  const doc = DocumentApp.openByUrl(
    'https://docs.google.com/document/d/1MNsCLjRtY9u4CJRtQZ74su7Sa1ItUD-2MSUirg6Tl_s/edit'
  )
  const body = doc.getBody()

  const indexes: number[] = []
  const len = body.getNumChildren()
  for (let i = 0; i < len; i++) {
    const child = body.getChild(i)
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue
    const m = child
      .asParagraph()
      .getText()
      .match(/^図表(\d+(?:-[a-z]+)?)/)
    if (!m) continue

    indexes.push(i)
  }

  setKeepWithNext(doc.getId(), indexes)
}

function setKeepWithNext(docId: string, elementIndexes: number[]) {
  const doc = Docs.Documents!.get(docId)
  const requests: GoogleAppsScript.Docs.Schema.Request[] = []

  for (const elementIndex of elementIndexes) {
    const paragraph = doc.body!.content![elementIndex]
    // You can construct an array of these requests and pass them all at once to batchUpdate:
    const request: GoogleAppsScript.Docs.Schema.Request = {
      updateParagraphStyle: {
        paragraphStyle: {
          keepWithNext: true,
        },
        fields: 'keepWithNext',
        range: {
          // "segmentId": "", // An empty segment ID signifies the document's body.
          startIndex: paragraph.startIndex,
          endIndex: paragraph.endIndex,
        },
      },
    }
    requests.push(request)
  }
  Docs.Documents!.batchUpdate({ requests }, docId)
}
