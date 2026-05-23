/**
 * MIT License
 *
 * Copyright (C) 2026 Guilherme Tadashi Maeoka
 * <https://github.com/guimspace/minha-caixinha>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

/**
 * Force the authorization dialog to ask only for access to files in which
 * the add-on or script is used, rather than all of a user's spreadsheets.
 * @OnlyCurrentDoc
 */

"use strict";

const VERSION = 'v0.1.0'
const TODAY = new Date()

function isInstalled () {
  // Checks if the add-on has already been initialized in the current document
  return PropertiesService.getDocumentProperties().getProperty('is_installed') === 'true'
}

function isCurrentCalendarYear () {
  try {
    return Number(PropertiesService.getDocumentProperties().getProperty('calendar_year')) >= TODAY.getFullYear()
  } catch (e) {
    return false
  }
}

function invalidateAuth () {
  ScriptApp.invalidateAuth()
  onOpen({'authMode': ScriptApp.AuthMode.NONE})
}

function onOpen (e) {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("Minha Caixinha")

  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    // Basic menu to prompt users to authorize the script
    menu.addItem('Instalar', 'initialConfig_').addToUi()
    return
  }

  const is_installed = isInstalled()
  if (!is_installed) {
    menu.addItem('Instalar', 'initialConfig_').addToUi()
    return
  }

  const is_current = isCurrentCalendarYear()
  if (is_current) {
    menu.addItem('Sincronizar calendário', 'toolSyncCalendar_')
      .addItem('Selecionar calendário padrão', 'showDialogConfigure_')
      .addSeparator()
  }

  menu.addItem('Remover permissões', 'invalidateAuth').addToUi()
}

function showDialogConfigure_ () {
  // Show a modeless dialog using an external HTML template
  const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate()
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Selecionar calendário padrão')
}

function saveCalendar (calendar_id) {
  // Called from client-side HTML to persist the chosen calendar ID
  PropertiesService.getDocumentProperties().setProperty('default_calendar', calendar_id)
  SpreadsheetApp.getActive().toast('Calendário padrão salvo.', 'Minha Caixinha')
}

function initialConfig_ () {
  const is_installed = isInstalled()
  if (is_installed) {
    onOpen({'authMode': ScriptApp.AuthMode.FULL})
    return
  }

  const ui = SpreadsheetApp.getUi()
  // Prompt the user to supply the base year for the cash flow tracker
  const response = ui.prompt('Digite o ano-calendário', `Insira um ano a partir de ${TODAY.getFullYear()}.`, ui.ButtonSet.OK_CANCEL)

  if (response.getSelectedButton() !== ui.Button.OK) return

  const yyyy = TODAY.getFullYear()
  const calendar_year = response.getResponseText()
  const year = Number(calendar_year)
  // Validate user input to ensure it falls within a realistic bound
  if (!Number.isInteger(year) || year < yyyy || year > 2199) {
    ui.alert(`Ano inválido. Por favor, insira um número inteiro a partir de ${yyyy}.`)
    return
  }

  const documentProperties = PropertiesService.getDocumentProperties()
  documentProperties.setProperties({
      calendar_year: calendar_year,
      default_calendar: CalendarApp.getDefaultCalendar().getId(),
    })

  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale('pt_BR')
  SpreadsheetApp.flush()

  // Initializes spreadsheet layouts, formulas, balances, and protections
  new SheetCashFlow().resetWeekendColoring()
    .resetFormulas()
    .resetBalanceReference()
    .resetProtection()
  SpreadsheetApp.flush()

  documentProperties.setProperty('is_installed', 'true')
  // Rebuild the menu context after installation
  onOpen({'authMode': ScriptApp.AuthMode.FULL})
}

function toolSyncCalendar_ () {
  const ui = SpreadsheetApp.getUi()

  // Fetch the active calendar from document properties
  const calendar_id = PropertiesService.getDocumentProperties().getProperty('default_calendar')
  const calendar = CalendarApp.getAllOwnedCalendars().find(calendar => calendar.getId() === calendar_id)
  if (!calendar) {
    ui.alert('Calendário não encontrado. Selecione um novo calendário.')
    showDialogConfigure()
    return
  }

  const range = SpreadsheetApp.getActiveRange()
  const sheetName = range.getSheet().getSheetName()
  if (sheetName !== 'Fluxo de Caixa') {
    ui.alert('Página "Fluxo de Caixa" não encontrada.')
    return
  }

  new RefreshCashFlow(calendar, range).refreshCashFlow()
  SpreadsheetApp.getActive().toast('Sincronização concluída.', 'Minha Caixinha')
}

class RefreshCashFlow {
  constructor (calendar, range) {
    this.calendar = calendar
    this.calendar_year = Number(PropertiesService.getDocumentProperties().getProperty('calendar_year'))

    this.sheet = SpreadsheetApp.getActive().getSheetByName('Fluxo de Caixa')
    const specs = Object.freeze(SheetCashFlow.specs)
    // Infers the target month (mm) based on the user's currently selected column in the spreadsheet
    const width = specs.width + 1
    const column = range.getColumn() - 2
    this.mm = (column - (column % width)) / width
  }

  refreshCashFlow () {
    // Get the maximum number of days in the calculated month
    const days = new Date(this.calendar_year, this.mm + 1, 0).getDate()
    const flow = new Array(days).fill('')
    const transactions = new Array(days).fill('')

    const upcoming = this.getUpcomingMonthEvents(this.mm)
    const response = this.readCalendarTransactions_(upcoming, days)
    for (let d = 0; d < days; d++) {
      flow[d] += response.flow[d].join('')
      transactions[d] += response.transactions[d]
    }

    // Writes back generated formulas and text values to their specific columns in the sheet
    this.sheet.getRange(4, 2 + 4 * this.mm, days, 1).setFormulas(RangeUtils.transpose([flow]))
    this.sheet.getRange(4, 4 + 4 * this.mm, days, 1).setValues(RangeUtils.transpose([transactions]))

    SpreadsheetApp.flush()
  }

  getUpcomingMonthEvents () {
    if (!this.calendar) return []

    const end = new Date(this.calendar_year, this.mm + 1, 1)
    if (end <= TODAY) return []

    let start = new Date(this.calendar_year, this.mm, 1)
    if (start <= TODAY) {
      // If within the current month, only look ahead from tomorrow onwards
      start = new Date(this.calendar_year, this.mm, TODAY.getDate() + 1)
      if (start >= end) return []
    }

    return this.calendar.getEvents(start, end)
  }

  readCalendarTransactions_ (upcoming, days) {
    const response = {
      flow: new Array(days).fill(null).map(a => []),
      transactions: new Array(days).fill('')
    }

    const eventos = CalendarUtils.digestEvents(upcoming)
    if (eventos.length === 0) return response

    const startDate = new Date(this.calendar_year, this.mm, 1)
    const endDate = new Date(this.calendar_year, this.mm + 1, 1)

    for (const ev of eventos) {
      if (ev.description === '') continue

      const title = `${ev.title}, `
      // Boundaries logic to handle multi-day events that span into/out of the current month
      const first = ev.startDate < startDate ? 0 : ev.startDate.getDate() - 1
      const last = ev.endDate >= endDate ? days : ev.endDate.getDate() - 1

      for (let day = first; day < last; day++) {
        response.flow[day].push(ev.value)
        response.transactions[day] += title
      }
    }

    return response
  }
}

class CalendarUtils {
  static digestEvents (eventos) {
    const output = []

    for (const evento of eventos) {
      const description = evento.getDescription()
      if (description === '') continue

      const metadata = {
        id: evento.getId(),
        title: evento.getTitle(),
        description,
        value: 0,
        startDate: null,
        endDate: null,
      }

      // Regular expression to extract monetary values formatted as Brazilian Reais
      const matches = description.match(/(-)?R\$ ?(\d{1,3}(?:\.\d{3})*|\d+)(?:,\d{2})?/)
      if (matches) {
        metadata.value = `${matches[1] ? '-' : '+'}${matches[2]},${matches[3] || '00'}`
      } else {
        metadata.value = '+0,00'
      }

      metadata.startDate = evento.getAllDayStartDate()
      metadata.endDate = evento.getAllDayEndDate()

      output.push(metadata)
    }

    return output
  }
}

class SheetCashFlow {
  constructor () {
    this.sheet = SpreadsheetApp.getActive().getSheetByName('Fluxo de Caixa')
    this._specs = Object.freeze(SheetCashFlow.specs)
    this._calendar_year = PropertiesService.getDocumentProperties().getProperty('calendar_year')
  }

  static get specs () {
    return {
      row: 4,
      column: 2,
      width: 3,
      height: 31
    }
  }

  get specs () {
    return this._specs
  }

  removeProtection () {
    const protections = this.sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    for (const protection of protections) {
      if (protection.canEdit()) protection.remove()
    }
    return this
  }

  resetProtection () {
    this.removeProtection()

    const w = 1 + this.specs.width
    const ranges = [
      this.sheet.getRange(this.specs.row, this.specs.column, 31),
      this.sheet.getRange(this.specs.row, 2 + this.specs.column, 31)
    ]

    for (let i = 1; i < 12; i++) {
      ranges.push(
        ranges[0].offset(0, w * i),
        ranges[1].offset(0, w * i))
    }

    this.sheet
      .protect()
      .setUnprotectedRanges(ranges)
      .setWarningOnly(true)

    return this
  }

  resetBalanceReference () {
    const w = 1 + this.specs.width

    const formulas = ['0 + B4']
    for (let mm = 1; mm < 12; mm++) {
      const dd = new Date(this._calendar_year, mm, 0).getDate() - 1
      formulas.push(RangeUtils.rollA1Notation(this.specs.row + dd, 3 + w * mm - w) + ' + ' + RangeUtils.rollA1Notation(this.specs.row, 2 + w * mm))
    }

    const range = this.sheet.getRange('C4')
    for (let mm = 0; mm < 12; mm++) {
      range.offset(0, w * mm).setFormula(`ROUND(${formulas[mm]}; 2)`)
    }

    return this
  }

  resetFormulas () {
    const s = '\\'
    const w = 1 + this.specs.width

    const options = `{"charttype"${s} "column"; "color"${s} "#93c47d"; "negcolor"${s} "#e06666"; "empty"${s} "zero"; "nan"${s} "convert"}`
    const range = this.sheet.getRange('B2')
    const ranges = []

    for (let mm = 0; mm < 12; mm++) {
      const n = new Date(this._calendar_year, 1 + mm, 0).getDate()

      let formula

      formula = RangeUtils.rollA1Notation(
        this.specs.row,
        1 + this.specs.column + w * mm,
        n, 1)
      formula = `SPARKLINE(${formula}; ${options})`
      range.offset(0, w * mm).setFormula(formula)

      ranges.push(
        RangeUtils.rollA1Notation(
          1 + this.specs.row,
          1 + this.specs.column + w * mm,
          n - 1, 1))
    }

    this.sheet
      .getRangeList(ranges)
      .setFormulaR1C1('R[-1]C + RC[-1]')

    return this
  }

  resetWeekendColoring () {
    const w = 1 + this.specs.width

    const f3f3f3 = [] // Light gray background range
    const d9ead3 = [] // Light green background range

    for (let mm = 0; mm < 12; mm++) {
      const d = new Date(this._calendar_year, 1 + mm, 0).getDate()
      if (d < 31) {
        // Blacks out cells corresponding to "invalid" days at the end of shorter months
        f3f3f3.push(
          RangeUtils.rollA1Notation(
            this.specs.row + d,
            this.specs.column + w * mm,
            31 - d, this.specs.width))
      }

      let j = 0
      let s = new Date(this._calendar_year, mm, 1).getDay()
      while (j < d) {
        switch (s) {
          case 0: { // Sunday
            d9ead3.push(
              RangeUtils.rollA1Notation(
                this.specs.row + j,
                this.specs.column + w * mm,
                1, this.specs.width))
            if (mm > 0) s = 6
            else s += 6
            j += 6
            break
          }
          case 6: { // Saturday
            d9ead3.push(
              RangeUtils.rollA1Notation(
                this.specs.row + j,
                this.specs.column + w * mm,
                1, this.specs.width))
            s = 0
            j++
            break
          }
          default: { // Weekdays
            s = (s + 1) % 7
            j++
            break
          }
        }
      }
    }

    this.sheet.getRangeList(f3f3f3).setBackground('#f3f3f3')
    this.sheet.getRangeList(d9ead3).setBackground('#d9ead3')

    return this
  }
}

class RangeUtils {
  static transpose (m) {
    return m[0].map((x, i) => m.map(x => x[i]))
  }

  static filterTableRanges (ranges, specs) {
    const selected = { indexes: [], ranges: [] }
    const w = specs.width + 1

    for (const range of ranges) {
      const column = range.getColumn() - 1 - specs.columnOffset

      if (column % w === 0 && range.getNumColumns() === specs.width) {
        selected.ranges.push(range)
      } else {
        const last = range.getLastColumn() - 1 - specs.columnOffset

        const start = (column - (column % w)) / w
        const end = (last - (last % w)) / w

        for (let i = start; i <= end; i++) {
          selected.indexes.push(i)
        }
      }
    }

    return selected
  }

  static rollA1Notation (posRow, posCol,
                         height = 1, width = 1,
                         mode1 = 1, mode2 = 1) {
    if (!Number.isInteger(posRow) || posRow < 1) throw new Error('Invalid posRow.')
    if (!Number.isInteger(posCol) || posCol < 1) throw new Error('Invalid posCol.')
    if (!Number.isInteger(height) || height < -1 || height === 0) throw new Error('Invalid height.')
    if (!Number.isInteger(width) || width < 1) throw new Error('Invalid width.')
    if (!Number.isInteger(mode1) || mode1 < 1) throw new Error('Invalid mode1.')
    if (!Number.isInteger(mode2) || mode2 < 1) throw new Error('Invalid mode2.')

    posCol--
    width--
    mode1--
    mode2--

    let str, c, m

    const f_ = 26
    const s_ = 4

    m = mode1 % s_
    str = ((m === 1 || m === 3) ? '$' : '')

    c = (posCol - posCol % f_) / f_
    str += (c ? String.fromCharCode(64 + c) : '')
    str += String.fromCharCode(65 + posCol % f_)

    str += (m >= 2 ? '$' : '')
    str += posRow

    if (height === 1 && width === 0) return str

    str += ':'
    posCol += width

    m = mode2 % s_
    str += ((m === 1 || m === 3) ? '$' : '')

    c = (posCol - posCol % f_) / f_
    str += (c ? String.fromCharCode(64 + c) : '')
    str += String.fromCharCode(65 + posCol % f_)

    if (height !== -1) {
      str += (m >= 2 ? '$' : '')
      str += posRow + height - 1
    }

    return str
  }
}
