/**
 * Herman Hsiao（正凌）個人首頁 — Google Apps Script 後端
 * 功能：
 * 1. 訪客計數器（記錄至 Google 試算表：HermanHome_訪客計數）
 * 2. 訪客留言板（記錄名稱、留言內容、IP位址、時間至 Google 試算表：HermanHome_訪客留言）
 *
 * 【首次部署步驟】
 * 1. 前往 https://script.google.com/ 建立新專案，貼上此程式碼並儲存。
 * 2. 上方函式下拉選單選擇 `setupSheets` → 點「執行」並授權（會自動建立 2 個試算表）。
 * 3. 點右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」。
 *    - 執行身分：我
 *    - 誰可以存取：所有人
 * 4. 複製產生的「網頁應用程式 URL」，貼回 index.html 的 `GAS_ENDPOINT`。
 */

const PROP_COUNTER_ID = 'COUNTER_SHEET_ID'
const PROP_GUESTBOOK_ID = 'GUESTBOOK_SHEET_ID'

function getOrCreateSpreadsheet(propKey, title, headers) {
  const props = PropertiesService.getScriptProperties()
  let id = props.getProperty(propKey)
  let ss

  if (id) {
    try {
      ss = SpreadsheetApp.openById(id)
    } catch (err) {
      ss = null
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(title)
    id = ss.getId()
    props.setProperty(propKey, id)
    const sheet = ss.getSheets()[0]
    sheet.setName('紀錄')
    sheet.appendRow(headers)
    sheet.setFrozenRows(1)
  }

  return ss
}

/**
 * 首次設定用：手動執行一次即可建立兩個試算表並印出連結
 */
function setupSheets() {
  const counterSS = getOrCreateSpreadsheet(
    PROP_COUNTER_ID,
    'HermanHome_訪客計數',
    ['時間', '訪客IP', '累計瀏覽次數', 'UserAgent']
  )
  const guestbookSS = getOrCreateSpreadsheet(
    PROP_GUESTBOOK_ID,
    'HermanHome_訪客留言',
    ['時間', '名稱', '留言內容', 'IP位址']
  )

  Logger.log('訪客計數試算表：' + counterSS.getUrl())
  Logger.log('訪客留言試算表：' + guestbookSS.getUrl())
}

function nowTaipei() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'guestbook'

  if (action === 'visit') {
    const lock = LockService.getScriptLock()
    lock.waitLock(10000)
    try {
      const ss = getOrCreateSpreadsheet(
        PROP_COUNTER_ID,
        'HermanHome_訪客計數',
        ['時間', '訪客IP', '累計瀏覽次數', 'UserAgent']
      )
      const sheet = ss.getSheets()[0]
      const lastRow = sheet.getLastRow()
      const nextCount = Math.max(0, lastRow - 1) + 1
      const ip = (e.parameter && e.parameter.ip) || '未知'
      const ua = (e.parameter && e.parameter.ua) || ''

      sheet.appendRow([nowTaipei(), ip, nextCount, ua])
      return jsonResponse({ ok: true, count: nextCount })
    } finally {
      lock.releaseLock()
    }
  }

  if (action === 'count') {
    const ss = getOrCreateSpreadsheet(
      PROP_COUNTER_ID,
      'HermanHome_訪客計數',
      ['時間', '訪客IP', '累計瀏覽次數', 'UserAgent']
    )
    const sheet = ss.getSheets()[0]
    const count = Math.max(0, sheet.getLastRow() - 1)
    return jsonResponse({ ok: true, count: count })
  }

  if (action === 'guestbook') {
    const ss = getOrCreateSpreadsheet(
      PROP_GUESTBOOK_ID,
      'HermanHome_訪客留言',
      ['時間', '名稱', '留言內容', 'IP位址']
    )
    const sheet = ss.getSheets()[0]
    const lastRow = sheet.getLastRow()
    if (lastRow <= 1) {
      return jsonResponse({ ok: true, items: [] })
    }
    const startRow = Math.max(2, lastRow - 29)
    const numRows = lastRow - startRow + 1
    const values = sheet.getRange(startRow, 1, numRows, 3).getValues()
    const items = values.reverse().map(function (row) {
      return {
        time: String(row[0]),
        name: String(row[1]),
        message: String(row[2]),
      }
    })
    return jsonResponse({ ok: true, items: items })
  }

  return jsonResponse({ ok: false, error: 'Unknown action' })
}

function doPost(e) {
  const lock = LockService.getScriptLock()
  lock.waitLock(10000)
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}')
    const name = String(payload.name || '').trim().slice(0, 40)
    const message = String(payload.message || '').trim().slice(0, 500)
    const ip = String(payload.ip || '未知').trim().slice(0, 64)

    if (!name || !message) {
      return jsonResponse({ ok: false, error: '名稱與留言內容不可為空' })
    }

    const ss = getOrCreateSpreadsheet(
      PROP_GUESTBOOK_ID,
      'HermanHome_訪客留言',
      ['時間', '名稱', '留言內容', 'IP位址']
    )
    const sheet = ss.getSheets()[0]
    const time = nowTaipei()
    sheet.appendRow([time, name, message, ip])

    return jsonResponse({
      ok: true,
      item: { time: time, name: name, message: message },
    })
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message })
  } finally {
    lock.releaseLock()
  }
}
