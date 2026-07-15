// excel.js — Geração do Excel formatado com ExcelJS
const ExcelJS = require('exceljs')

// ── Cores ─────────────────────────────────────────────────────────────────
const C = {
  ORANGE:      'FFF47B20',
  ORANGE_DARK: 'FFC45E0A',
  ORANGE_LITE: 'FFFEE8D6',
  WHITE:       'FFFFFFFF',
  DARK:        'FF1A1A1F',
  GREY_ALT:    'FFFFF4EA',
  GREY_LIGHT:  'FFF9FAFB',
}

const STATUS_BG = {
  'Aberto':       'FFFEE2E2',
  'Em andamento': 'FFFEF9C3',
  'Resolvido':    'FFDCFCE7',
  'Escalado':     'FFFFEDD5',
  'Sem resposta': 'FFEDE9FE',
}
const STATUS_FG = {
  'Aberto':       'FF991B1B',
  'Em andamento': 'FF854D0E',
  'Resolvido':    'FF166534',
  'Escalado':     'FF9A3412',
  'Sem resposta': 'FF5B21B6',
}

const BORDER = {
  top:    { style: 'thin', color: { argb: 'FFD1D5DB' } },
  left:   { style: 'thin', color: { argb: 'FFD1D5DB' } },
  bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
  right:  { style: 'thin', color: { argb: 'FFD1D5DB' } },
}

// ── Helpers de estilo ─────────────────────────────────────────────────────
function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } }
}

function sc(cell, { bg, fg = C.DARK, size = 10, bold = false, italic = false,
                    h = 'left', wrap = false, border = true } = {}) {
  if (bg) cell.fill = fill(bg)
  cell.font      = { name: 'Calibri', size, bold, italic, color: { argb: fg } }
  cell.alignment = { horizontal: h, vertical: 'middle', wrapText: wrap }
  if (border) cell.border = BORDER
}

// ── Helpers de dados ──────────────────────────────────────────────────────
function sum(arr) { return arr.reduce((a, b) => a + b, 0) }

function monthlyBy(items, year, { key, val, status } = {}) {
  const counts = {}
  for (const item of items) {
    if (!item.data) continue
    const parts = item.data.split('/')
    if (parts.length !== 3) continue
    const [, m, y] = parts
    if (y !== year) continue
    if (key && val && item[key] !== val) continue
    if (status && item.status !== status) continue
    const mi = parseInt(m)
    counts[mi] = (counts[mi] || 0) + 1
  }
  return Array.from({ length: 12 }, (_, i) => counts[i + 1] || 0)
}

function detectYear(items) {
  const years = new Set()
  for (const item of items) {
    if (item.data) {
      const p = item.data.split('/')
      if (p.length === 3) years.add(p[2])
    }
  }
  return years.size > 0 ? [...years].sort().pop() : String(new Date().getFullYear())
}

function countBy(items, key) {
  const counts = {}
  for (const item of items) {
    const v = item[key] || '—'
    counts[v] = (counts[v] || 0) + 1
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1])
}

// ── Aba Demandas ──────────────────────────────────────────────────────────
function buildDemandas(wb, items) {
  const ws = wb.addWorksheet('📋 Demandas')
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, showGridLines: false }]

  const colWidths = [7, 12, 24, 18, 26, 26, 20, 13, 16, 44]
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w })

  const today = new Date().toLocaleDateString('pt-BR')

  // Banner
  ws.getRow(1).height = 34
  ws.mergeCells(1, 1, 1, 10)
  sc(ws.getCell(1, 1), { bg: C.ORANGE, fg: C.WHITE, size: 14, bold: true, border: false })
  ws.getCell(1, 1).value = `🐘  KnowU  ·  Demandas CS  |  ${today}`

  // Totais
  ws.getRow(2).height = 24
  const n  = items.length
  const ab = items.filter(i => i.status === 'Aberto').length
  const an = items.filter(i => i.status === 'Em andamento').length
  const re = items.filter(i => i.status === 'Resolvido').length
  const es = items.filter(i => i.status === 'Escalado').length
  ;[
    [`Total: ${n}`,       C.ORANGE,    C.WHITE],
    [`Aberto: ${ab}`,     'FFFEE2E2',  'FF991B1B'],
    [`Andamento: ${an}`,  'FFFEF9C3',  'FF854D0E'],
    [`Resolvido: ${re}`,  'FFDCFCE7',  'FF166534'],
    [`Escalado: ${es}`,   'FFFFEDD5',  'FF9A3412'],
  ].forEach(([text, bg, fg], ci) => {
    const cell = ws.getCell(2, ci + 1)
    cell.value = text
    sc(cell, { bg, fg, bold: true, h: 'center', border: false })
  })
  for (let ci = 6; ci <= 10; ci++) sc(ws.getCell(2, ci), { bg: C.ORANGE, border: false })

  // Spacer
  ws.getRow(3).height = 6

  // Cabeçalho
  ws.getRow(4).height = 28
  const hdrs = ['Hora','Data','Operadora','Solicitante','Tipo','Beneficiário','Médica','Data EQ','Status','Observação']
  hdrs.forEach((h, ci) => {
    const cell = ws.getCell(4, ci + 1)
    cell.value = h
    sc(cell, { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' })
    cell.border = { ...BORDER, top: { style: 'medium', color: { argb: 'FFD1D5DB' } }, bottom: { style: 'medium', color: { argb: 'FFD1D5DB' } } }
  })
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: 10 } }

  // Linhas de dados
  items.forEach((item, ri) => {
    const row = ri + 5
    ws.getRow(row).height = 20
    const bg = ri % 2 === 0 ? C.WHITE : C.GREY_ALT
    const vals = [
      item.hora || '',
      item.data || '',
      item.op   || '',
      item.sol  || '',
      item.tipo || '',
      item.ben  || '',
      item.medica || '',
      item.dataeq ? new Date(item.dataeq + 'T00:00').toLocaleDateString('pt-BR') : '',
      item.status || '',
      item.obs  || '',
    ]
    vals.forEach((v, ci) => {
      const cell = ws.getCell(row, ci + 1)
      cell.value = v
      if      (ci === 0)     sc(cell, { bg, fg: C.ORANGE,    bold: true, h: 'center' })
      else if (ci === 1)     sc(cell, { bg, h: 'center' })
      else if (ci === 6 && v) sc(cell, { bg, fg: 'FF6A1B9A', bold: true })
      else if (ci === 8)     sc(cell, { bg: STATUS_BG[v] || bg, fg: STATUS_FG[v] || C.DARK, bold: true, h: 'center' })
      else if (ci === 9)     sc(cell, { bg, fg: 'FF555555',  italic: true, wrap: true })
      else                   sc(cell, { bg })
    })
  })

  // Footer
  const fr = items.length + 5
  ws.getRow(fr).height = 18
  ws.mergeCells(fr, 1, fr, 10)
  sc(ws.getCell(fr, 1), { bg: C.ORANGE_LITE, fg: 'FF999999', italic: true, h: 'right', border: false })
  ws.getCell(fr, 1).value = `Exportado em ${new Date().toLocaleString('pt-BR')}  ·  KnowU CS`
}

// ── Aba Resumo ────────────────────────────────────────────────────────────
function buildResumo(wb, items) {
  const ws    = wb.addWorksheet('📊 Resumo')
  ws.views    = [{ showGridLines: false }]
  const year  = detectYear(items)
  const today = new Date().toLocaleDateString('pt-BR')
  const total = items.length || 1

  ;[32,10,14,14,14,4,12,12,12,12,12,12,12,12].forEach((w, i) => { ws.getColumn(i + 1).width = w })

  let cr = 1

  function spacer() { ws.getRow(cr).height = 6; cr++ }

  function secTitle(text, cols = 5) {
    ws.getRow(cr).height = 26
    ws.mergeCells(cr, 1, cr, cols)
    sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
    ws.getCell(cr, 1).value = text
    cr++
  }

  function colHeader(texts) {
    ws.getRow(cr).height = 22
    texts.forEach((t, ci) => {
      const cell = ws.getCell(cr, ci + 1)
      cell.value = t
      sc(cell, { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' })
    })
    cr++
  }

  function dataRow(vals, idx, fgs = [], bolds = []) {
    ws.getRow(cr).height = 20
    const bg = idx % 2 === 0 ? C.WHITE : C.GREY_LIGHT
    vals.forEach((v, ci) => {
      const cell = ws.getCell(cr, ci + 1)
      cell.value = v
      sc(cell, { bg, fg: fgs[ci] || C.DARK, bold: bolds[ci] || false, h: ci === 0 ? 'left' : 'center' })
    })
    cr++
  }

  // Banner
  ws.getRow(cr).height = 34
  ws.mergeCells(cr, 1, cr, 7)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE, fg: C.WHITE, size: 14, bold: true, border: false })
  ws.getCell(cr, 1).value = `🐘  KnowU  ·  Resumo & Evolução ${year}  |  ${today}`
  cr++

  // ── Tabela mensal ─────────────────────────────────────────────────────────
  spacer()
  ws.getRow(cr).height = 26
  ws.mergeCells(cr, 1, cr, 7)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = `  📅  Resumo por Mês — ${year}`
  cr++

  colHeader(['Mês','Total','Aberto','Andamento','Resolvido','Escalado','Sem resp.'])

  const tot_m = monthlyBy(items, year)
  const ab_m  = monthlyBy(items, year, { status: 'Aberto' })
  const and_m = monthlyBy(items, year, { status: 'Em andamento' })
  const res_m = monthlyBy(items, year, { status: 'Resolvido' })
  const esc_m = monthlyBy(items, year, { status: 'Escalado' })
  const sr_m  = monthlyBy(items, year, { status: 'Sem resposta' })

  const MESES_EXT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
  const stBgs = [null,'FFFEE2E2','FFFEF9C3','FFDCFCE7','FFFFEDD5','FFEDE9FE']
  const stFgs = [null,'FF991B1B','FF854D0E','FF166534','FF9A3412','FF5B21B6']

  for (let mi = 0; mi < 12; mi++) {
    ws.getRow(cr).height = 20
    const bg      = mi % 2 === 0 ? C.WHITE : C.GREY_ALT
    const hasData = tot_m[mi] > 0

    const mesCell = ws.getCell(cr, 1)
    mesCell.value = MESES_EXT[mi]
    sc(mesCell, { bg: hasData ? C.GREY_ALT : bg, fg: hasData ? C.ORANGE : 'FF888888', bold: hasData })

    ;[tot_m[mi], ab_m[mi], and_m[mi], res_m[mi], esc_m[mi], sr_m[mi]].forEach((v, ci) => {
      const cell = ws.getCell(cr, ci + 2)
      cell.value = v > 0 ? v : '–'
      const cellBg = (v > 0 && stBgs[ci]) ? stBgs[ci] : bg
      const cellFg = ci === 0 && v > 0 ? C.ORANGE_DARK : ((v > 0 && stFgs[ci]) ? stFgs[ci] : (v > 0 ? C.DARK : 'FFBBBBBB'))
      sc(cell, { bg: cellBg, fg: cellFg, bold: ci === 0 && v > 0, h: 'center' })
    })
    cr++
  }

  // Total geral
  ws.getRow(cr).height = 22
  sc(ws.getCell(cr, 1), { bg: C.ORANGE, fg: C.WHITE, bold: true })
  ws.getCell(cr, 1).value = 'TOTAL GERAL'
  ;[sum(tot_m), sum(ab_m), sum(and_m), sum(res_m), sum(esc_m), sum(sr_m)].forEach((v, ci) => {
    const cell = ws.getCell(cr, ci + 2)
    cell.value = v
    sc(cell, { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' })
  })
  cr += 2

  // ── Status ────────────────────────────────────────────────────────────────
  spacer()
  secTitle('  📊  Visão geral por status')
  colHeader(['Status','Qtd','% do total','Proporção',''])
  ;['Aberto','Em andamento','Resolvido','Escalado','Sem resposta'].forEach((st, si) => {
    ws.getRow(cr).height = 20
    const qty = items.filter(i => i.status === st).length
    const bg  = si % 2 === 0 ? C.WHITE : C.GREY_LIGHT

    sc(ws.getCell(cr, 1), { bg: STATUS_BG[st], fg: STATUS_FG[st], bold: true })
    ws.getCell(cr, 1).value = st

    sc(ws.getCell(cr, 2), { bg: STATUS_BG[st], fg: STATUS_FG[st], bold: true, size: 12, h: 'center' })
    ws.getCell(cr, 2).value = qty

    sc(ws.getCell(cr, 3), { bg: STATUS_BG[st], fg: STATUS_FG[st], h: 'center' })
    ws.getCell(cr, 3).value = `${Math.round(qty / total * 100)}%`

    sc(ws.getCell(cr, 4), { bg, fg: STATUS_FG[st] })
    ws.getCell(cr, 4).value = '█'.repeat(Math.round(qty / total * 18))

    sc(ws.getCell(cr, 5), { bg })
    cr++
  })
  cr++

  // ── Por tipo ──────────────────────────────────────────────────────────────
  spacer()
  secTitle('  📋  Por tipo de demanda')
  colHeader(['Tipo','Total','Aberto','Andamento','Resolvido'])
  countBy(items, 'tipo').forEach(([tipo, totT], ti) => {
    dataRow([
      tipo,
      totT,
      items.filter(i => i.tipo === tipo && i.status === 'Aberto').length,
      items.filter(i => i.tipo === tipo && i.status === 'Em andamento').length,
      items.filter(i => i.tipo === tipo && i.status === 'Resolvido').length,
    ], ti,
    [null, C.ORANGE_DARK, 'FF991B1B', 'FF854D0E', 'FF166534'],
    [false, true, false, false, false])
  })
  cr++

  // ── Por operadora ─────────────────────────────────────────────────────────
  spacer()
  secTitle('  🏢  Por operadora')
  colHeader(['Operadora','Total','Aberto','Andamento','Resolvido'])
  countBy(items, 'op').forEach(([op, totO], oi) => {
    dataRow([
      op,
      totO,
      items.filter(i => i.op === op && i.status === 'Aberto').length,
      items.filter(i => i.op === op && i.status === 'Em andamento').length,
      items.filter(i => i.op === op && i.status === 'Resolvido').length,
    ], oi,
    [null, C.ORANGE_DARK, 'FF991B1B', 'FF854D0E', 'FF166534'],
    [false, true, false, false, false])
  })
  cr++

  // ── Por médica ────────────────────────────────────────────────────────────
  const medItems = items.filter(i => i.medica)
  if (medItems.length > 0) {
    spacer()
    secTitle('  👩‍⚕️  Por médica responsável')
    colHeader(['Médica','Total','Aberto','Andamento','Resolvido'])
    countBy(medItems, 'medica').forEach(([med, totMed], mi) => {
      dataRow([
        med,
        totMed,
        medItems.filter(i => i.medica === med && i.status === 'Aberto').length,
        medItems.filter(i => i.medica === med && i.status === 'Em andamento').length,
        medItems.filter(i => i.medica === med && i.status === 'Resolvido').length,
      ], mi,
      ['FF6A1B9A', C.ORANGE_DARK, 'FF991B1B', 'FF854D0E', 'FF166534'],
      [true, true, false, false, false])
    })
  }

  // Footer
  cr += 2
  ws.getRow(cr).height = 20
  ws.mergeCells(cr, 1, cr, 7)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_LITE, fg: 'FF999999', italic: true, h: 'right', border: false })
  ws.getCell(cr, 1).value = `Gerado em ${new Date().toLocaleString('pt-BR')}  ·  KnowU CS  |  Ano: ${year}`
}

// ── Aba Gráficos ──────────────────────────────────────────────────────────
function buildGraficos(wb, charts) {
  const ws = wb.addWorksheet('📈 Gráficos')
  ws.views = [{ showGridLines: false }]
  ws.getColumn(1).width = 130

  // Banner
  ws.getRow(1).height = 30
  const ban = ws.getCell(1, 1)
  ban.value = '🐘  KnowU  ·  Gráficos — Demandas CS'
  ban.fill = fill(C.ORANGE)
  ban.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } }
  ban.alignment = { horizontal: 'left', vertical: 'middle' }

  const defs = [
    { key: 'mensal',     label: '📅  Demandas por Mês' },
    { key: 'status',     label: '📊  Distribuição por Status' },
    { key: 'tipos',      label: '📋  Demandas por Tipo' },
    { key: 'operadoras', label: '🏢  Demandas por Operadora' },
  ]

  let cr = 2
  for (const { key, label } of defs) {
    if (!charts || !charts[key]) continue

    cr++  // espaço
    ws.getRow(cr).height = 22
    const tc = ws.getCell(cr, 1)
    tc.value = label
    tc.fill = fill(C.ORANGE_DARK)
    tc.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }
    tc.alignment = { horizontal: 'left', vertical: 'middle' }
    cr++

    const imgId = wb.addImage({ base64: charts[key], extension: 'png' })
    ws.addImage(imgId, { tl: { col: 0, row: cr - 1 }, ext: { width: 900, height: 440 } })

    // Reservar ~25 linhas (900x440px ≈ 330pt de altura a 96dpi)
    for (let r = 0; r < 25; r++) { ws.getRow(cr + r).height = 15 }
    cr += 25
  }

  cr += 2
  ws.getRow(cr).height = 18
  const fc = ws.getCell(cr, 1)
  fc.value = `Gerado em ${new Date().toLocaleString('pt-BR')}  ·  KnowU CS`
  fc.fill = fill(C.ORANGE_LITE)
  fc.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF999999' } }
  fc.alignment = { horizontal: 'right', vertical: 'middle' }
}

// ── Aba Por Operadora ─────────────────────────────────────────────────────
function buildPorOperadora(wb, items) {
  const ws = wb.addWorksheet('🏢 Por Operadora')
  ws.views = [{ showGridLines: false }]

  const year = detectYear(items)
  const MESES_ABR = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
  const TIPOS = [...new Set(items.map(i => i.tipo).filter(Boolean))].sort()
  const ops   = [...new Set(items.map(i => i.op).filter(op => op && op !== '—'))].sort()

  ws.getColumn(1).width = 32
  for (let c = 2; c <= 20; c++) ws.getColumn(c).width = 11

  let cr = 1

  // Banner
  ws.getRow(cr).height = 34
  const bannerCols = Math.max(8, TIPOS.length + 2)
  ws.mergeCells(cr, 1, cr, bannerCols)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE, fg: C.WHITE, size: 14, bold: true, border: false })
  ws.getCell(cr, 1).value = `🏢  KnowU CS  ·  Por Operadora  |  ${new Date().toLocaleDateString('pt-BR')}`
  cr += 2

  // ── SEÇÃO 1: Resumo ──────────────────────────────────────────────────────
  ws.getRow(cr).height = 26
  ws.mergeCells(cr, 1, cr, 8)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = '  🏢  Resumo por Operadora'
  cr++

  ws.getRow(cr).height = 22
  const hdrs1 = ['Operadora','Total','Resolvidas','Em Aberto','Em Andamento','Escaladas','Sem Resposta','Taxa Resolução']
  hdrs1.forEach((h, ci) => { sc(ws.getCell(cr, ci+1), { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' }); ws.getCell(cr, ci+1).value = h })
  cr++

  ops.forEach((op, oi) => {
    const opIt = items.filter(i => i.op === op)
    const total = opIt.length
    const res  = opIt.filter(i => i.status === 'Resolvido').length
    const bg = oi % 2 === 0 ? C.WHITE : C.GREY_ALT
    ws.getRow(cr).height = 20
    const row1 = [op, total, res,
      opIt.filter(i => i.status === 'Aberto').length,
      opIt.filter(i => i.status === 'Em andamento').length,
      opIt.filter(i => i.status === 'Escalado').length,
      opIt.filter(i => i.status === 'Sem resposta').length,
      total ? `${Math.round(res/total*100)}%` : '0%']
    row1.forEach((v, ci) => {
      sc(ws.getCell(cr, ci+1), { bg, fg: ci===0?C.DARK:ci===1?C.ORANGE_DARK:C.DARK, bold: ci===1, h: ci===0?'left':'center' })
      ws.getCell(cr, ci+1).value = v
    })
    cr++
  })
  cr += 2

  // ── SEÇÃO 2: Tipos por operadora ────────────────────────────────────────
  ws.getRow(cr).height = 26
  const nTC = TIPOS.length + 2
  ws.mergeCells(cr, 1, cr, nTC)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = '  📋  Tipos de Demanda por Operadora'
  cr++

  ws.getRow(cr).height = 36
  ;['Operadora', ...TIPOS, 'Total'].forEach((h, ci) => {
    sc(ws.getCell(cr, ci+1), { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center', wrap: true })
    ws.getCell(cr, ci+1).value = h
    if (ci > 0) ws.getColumn(ci+1).width = Math.max(String(h).length * 1.0 + 2, 10)
  })
  cr++

  ops.forEach((op, oi) => {
    const opIt = items.filter(i => i.op === op)
    const bg = oi % 2 === 0 ? C.WHITE : C.GREY_ALT
    ws.getRow(cr).height = 20
    ;[op, ...TIPOS.map(t => opIt.filter(i => i.tipo === t).length), opIt.length].forEach((v, ci) => {
      const isZero = typeof v === 'number' && v === 0
      sc(ws.getCell(cr, ci+1), { bg, fg: isZero?'FFBBBBBB':ci===0?C.DARK:C.ORANGE_DARK, bold: ci===nTC-1, h: ci===0?'left':'center' })
      ws.getCell(cr, ci+1).value = isZero ? '–' : v
    })
    cr++
  })
  cr += 2

  // ── SEÇÃO 3: Demandas mensais por operadora ──────────────────────────────
  ws.getRow(cr).height = 26
  ws.mergeCells(cr, 1, cr, 15)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = `  📅  Demandas Mensais por Operadora — ${year}`
  cr++

  ws.getRow(cr).height = 22
  ;['Operadora', ...MESES_ABR, 'Total'].forEach((h, ci) => {
    sc(ws.getCell(cr, ci+1), { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' })
    ws.getCell(cr, ci+1).value = h
    if (ci > 0) ws.getColumn(ci+1).width = 8
  })
  cr++

  ops.forEach((op, oi) => {
    const opIt = items.filter(i => i.op === op)
    const bg = oi % 2 === 0 ? C.WHITE : C.GREY_ALT
    ws.getRow(cr).height = 20
    const months = MESES_ABR.map((_, mi) => {
      const m = String(mi+1).padStart(2,'0')
      return opIt.filter(i => { if (!i.data) return false; const [,im,iy]=i.data.split('/'); return im===m && iy===year }).length
    })
    const yrTotal = opIt.filter(i => i.data && i.data.split('/')[2]===year).length
    ;[op, ...months, yrTotal].forEach((v, ci) => {
      const isZero = typeof v === 'number' && v === 0
      sc(ws.getCell(cr, ci+1), { bg, fg: isZero?'FFBBBBBB':ci===0?C.DARK:C.ORANGE_DARK, bold: ci===13, h: ci===0?'left':'center' })
      ws.getCell(cr, ci+1).value = isZero ? '–' : v
    })
    cr++
  })
  cr += 2

  // ── SEÇÃO 4: Tipos por mês por operadora ────────────────────────────────
  for (const op of ops) {
    const opYrIt = items.filter(i => i.op===op && i.data && i.data.split('/')[2]===year)
    if (!opYrIt.length) continue

    ws.getRow(cr).height = 22
    ws.mergeCells(cr, 1, cr, 15)
    sc(ws.getCell(cr, 1), { bg: 'FF2E2E38', fg: C.WHITE, bold: true, size: 10, border: false })
    ws.getCell(cr, 1).value = `  ${op}  —  ${year}`
    cr++

    ws.getRow(cr).height = 20
    ;['Tipo', ...MESES_ABR, 'Total'].forEach((h, ci) => {
      sc(ws.getCell(cr, ci+1), { bg: C.ORANGE_LITE, fg: C.ORANGE_DARK, bold: true, h: 'center', size: 9 })
      ws.getCell(cr, ci+1).value = h
    })
    cr++

    TIPOS.forEach((tipo, ti) => {
      const tipoOpIt = opYrIt.filter(i => i.tipo === tipo)
      if (!tipoOpIt.length) return
      const bg = ti % 2 === 0 ? C.WHITE : C.GREY_LIGHT
      const months = MESES_ABR.map((_, mi) => {
        const m = String(mi+1).padStart(2,'0')
        return tipoOpIt.filter(i => { const [,im]=i.data.split('/'); return im===m }).length
      })
      ws.getRow(cr).height = 18
      ;[tipo, ...months, tipoOpIt.length].forEach((v, ci) => {
        const isZero = typeof v === 'number' && v === 0
        sc(ws.getCell(cr, ci+1), { bg, fg: isZero?'FFBBBBBB':C.DARK, bold: ci===13, h: ci===0?'left':'center', size: 9 })
        ws.getCell(cr, ci+1).value = isZero ? '–' : v
      })
      cr++
    })
    cr++
  }

  // Footer
  cr++
  ws.getRow(cr).height = 18
  ws.mergeCells(cr, 1, cr, 15)
  const fc = ws.getCell(cr, 1)
  fc.value = `Gerado em ${new Date().toLocaleString('pt-BR')}  ·  KnowU CS`
  fc.fill = fill(C.ORANGE_LITE)
  fc.font = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF999999' } }
  fc.alignment = { horizontal: 'right', vertical: 'middle' }
}

// ── Aba Análise Dinâmica ──────────────────────────────────────────────────
// allItems = TODOS os registros (nunca filtrado), para análise correta
function buildAnalise(wb, allItems) {
  const ops   = [...new Set(allItems.map(i => i.op).filter(op => op && op !== '—'))].sort()
  const tipos = [...new Set(allItems.map(i => i.tipo).filter(Boolean))].sort()
  const mesSet = new Set()
  allItems.forEach(i => {
    if (i.data && i.data.split('/').length === 3) {
      const [, m, y] = i.data.split('/')
      mesSet.add(`${m}/${y}`)
    }
  })
  const meses = [...mesSet].sort((a, b) => {
    const [ma, ya] = a.split('/'); const [mb, yb] = b.split('/')
    return ya !== yb ? ya.localeCompare(yb) : ma.localeCompare(mb)
  })

  // ── Aba _Dados (oculta) — fonte de todos os dados para as fórmulas ───────
  // Usamos uma aba separada para que as fórmulas apontem para TODOS os registros,
  // independente do filtro aplicado na exportação da aba Demandas.
  const wsD = wb.addWorksheet('_Dados')
  wsD.state = 'hidden'
  // Cabeçalho na row 1: Op(A), Tipo(B), Mes(C)
  wsD.getCell(1, 1).value = 'Op'
  wsD.getCell(1, 2).value = 'Tipo'
  wsD.getCell(1, 3).value = 'Mes'
  allItems.forEach((item, ri) => {
    const row = ri + 2
    wsD.getCell(row, 1).value = item.op   || ''
    wsD.getCell(row, 2).value = item.tipo || ''
    // Mês no formato mm/yyyy extraído da data dd/mm/yyyy
    if (item.data && item.data.split('/').length === 3) {
      const [, m, y] = item.data.split('/')
      wsD.getCell(row, 3).value = `${m}/${y}`
    }
  })
  const DN = allItems.length + 2  // última linha de dados em _Dados

  // ── Aba _Listas (oculta) — listas para validação ─────────────────────────
  const wsL = wb.addWorksheet('_Listas')
  wsL.state = 'hidden'
  const listsData = [
    ['Todas', ...ops],
    ['Todos', ...tipos],
    ['Todos', ...meses],
  ]
  const maxLen = Math.max(...listsData.map(l => l.length))
  for (let r = 0; r < maxLen; r++)
    for (let c = 0; c < 3; c++)
      if (r < listsData[c].length) wsL.getCell(r + 1, c + 1).value = listsData[c][r]

  const opRange   = `_Listas!$A$1:$A$${listsData[0].length}`
  const tipoRange = `_Listas!$B$1:$B$${listsData[1].length}`
  const mesRange  = `_Listas!$C$1:$C$${listsData[2].length}`

  // ── Aba Analise ──────────────────────────────────────────────────────────
  const ws = wb.addWorksheet('Analise')
  ws.views = [{ showGridLines: false }]

  // Layout: filtros na col A-D (largos), tabela na col F-G, chart cols I-V
  ws.getColumn(1).width = 2   // margem
  ws.getColumn(2).width = 18  // labels filtros
  ws.getColumn(3).width = 32  // células dropdown
  ws.getColumn(4).width = 2   // sep
  ws.getColumn(5).width = 2   // sep
  ws.getColumn(6).width = 36  // tipo names
  ws.getColumn(7).width = 14  // counts
  ws.getColumn(8).width = 2   // sep
  for (let c = 9; c <= 24; c++) ws.getColumn(c).width = 7 // chart

  let cr = 1

  // ── Banner ────────────────────────────────────────────────────────────────
  ws.getRow(cr).height = 36
  ws.mergeCells(cr, 1, cr, 24)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE, fg: C.WHITE, size: 15, bold: true, border: false })
  ws.getCell(cr, 1).value = '📊  KnowU CS  ·  Análise Dinâmica'
  ws.getCell(cr, 1).alignment = { horizontal: 'center', vertical: 'middle' }
  cr++ // 2

  // ── Linha de totais gerais ────────────────────────────────────────────────
  ws.getRow(cr).height = 22
  const totalGeral = allItems.length
  const ops_count  = ops.length
  const tipos_count = tipos.length
  const statsStr = `Total geral: ${totalGeral} demandas  ·  ${ops_count} operadora${ops_count!==1?'s':''}  ·  ${tipos_count} tipo${tipos_count!==1?'s':''}`
  ws.mergeCells(cr, 1, cr, 24)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_LITE, fg: C.ORANGE_DARK, bold: true, h: 'center', size: 10, border: false })
  ws.getCell(cr, 1).value = statsStr
  cr++ // 3

  ws.getRow(cr).height = 10; cr++ // 4 spacer

  // ── Cabeçalho da seção de filtros ────────────────────────────────────────
  ws.getRow(cr).height = 24
  ws.mergeCells(cr, 1, cr, 3)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = '  🔍  Filtros Interativos'
  cr++ // 5

  // Instruções
  ws.getRow(cr).height = 18
  ws.mergeCells(cr, 1, cr, 3)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_LITE, fg: C.ORANGE_DARK, italic: true, size: 9, border: false })
  ws.getCell(cr, 1).value = '  Clique na célula laranja e escolha uma opção da lista'
  cr++ // 6

  const FILTERS = [
    { label: '🏥 Operadora', all: 'Todas', range: opRange },
    { label: '📋 Tipo',      all: 'Todos', range: tipoRange },
    { label: '📅 Mês (mm/aaaa)', all: 'Todos', range: mesRange },
  ]
  const fRows = []
  for (const { label, all, range } of FILTERS) {
    ws.getRow(cr).height = 28

    // Label
    sc(ws.getCell(cr, 2), { bg: C.WHITE, fg: C.DARK, bold: true, size: 11, border: false })
    ws.getCell(cr, 2).value = label

    // Dropdown cell
    const dc = ws.getCell(cr, 3)
    dc.value = all
    dc.fill  = fill('FFFFF0E0')
    dc.font  = { name: 'Calibri', size: 11, bold: true, color: { argb: C.ORANGE_DARK } }
    dc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 }
    dc.border = {
      top:    { style: 'medium', color: { argb: C.ORANGE } },
      left:   { style: 'medium', color: { argb: C.ORANGE } },
      bottom: { style: 'medium', color: { argb: C.ORANGE } },
      right:  { style: 'medium', color: { argb: C.ORANGE } },
    }
    dc.dataValidation = { type: 'list', allowBlank: false, formulae: [range] }

    fRows.push(cr)
    cr++
  }
  const [opRow, tipoRow, mesRow] = fRows // rows of the 3 filter cells

  ws.getRow(cr).height = 10; cr++ // spacer after filters

  // ── Cabeçalho da tabela de resultados ────────────────────────────────────
  ws.getRow(cr).height = 24
  ws.mergeCells(cr, 1, cr, 7)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11, border: false })
  ws.getCell(cr, 1).value = '  📊  Resultado por Tipo de Demanda'
  cr++

  // Cabeçalho das colunas
  ws.getRow(cr).height = 22
  sc(ws.getCell(cr, 6), { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'left' })
  ws.getCell(cr, 6).value = 'Tipo de Demanda'
  sc(ws.getCell(cr, 7), { bg: C.ORANGE, fg: C.WHITE, bold: true, h: 'center' })
  ws.getCell(cr, 7).value = 'Qtd'
  cr++

  const firstDataRow = cr

  // Fórmulas COUNTIFS usando a aba _Dados (sempre com todos os registros)
  // _Dados: col A = Op, col B = Tipo, col C = Mes (mm/yyyy)
  tipos.forEach((tipo, ti) => {
    const bg = ti % 2 === 0 ? C.WHITE : C.GREY_ALT
    ws.getRow(cr).height = 22

    sc(ws.getCell(cr, 6), { bg, fg: C.DARK, h: 'left', size: 11 })
    ws.getCell(cr, 6).value = tipo

    const cc = ws.getCell(cr, 7)
    // Quando filtro de Tipo ≠ "Todos" e linha não é o tipo filtrado → 0
    // Caso contrário: COUNTIFS em _Dados
    // Op: se "Todas" → não filtra; senão → filtra col A
    // Mes: se "Todos" → não filtra; senão → filtra col C
    cc.value = { formula:
      `=IF(AND($C$${tipoRow}<>"Todos",F${cr}<>$C$${tipoRow}),0,` +
      `IF($C$${opRow}="Todas",` +
        `IF($C$${mesRow}="Todos",` +
          `COUNTIF(_Dados!$B$2:$B$${DN},F${cr}),` +
          `COUNTIFS(_Dados!$B$2:$B$${DN},F${cr},_Dados!$C$2:$C$${DN},$C$${mesRow})),` +
        `IF($C$${mesRow}="Todos",` +
          `COUNTIFS(_Dados!$B$2:$B$${DN},F${cr},_Dados!$A$2:$A$${DN},$C$${opRow}),` +
          `COUNTIFS(_Dados!$B$2:$B$${DN},F${cr},_Dados!$A$2:$A$${DN},$C$${opRow},_Dados!$C$2:$C$${DN},$C$${mesRow}))))`
    }
    cc.fill = fill(bg)
    cc.font = { name: 'Calibri', size: 12, bold: true, color: { argb: C.ORANGE_DARK } }
    cc.alignment = { horizontal: 'center', vertical: 'middle' }
    cc.border = BORDER
    cr++
  })

  const lastDataRow = cr - 1

  // Linha TOTAL
  ws.getRow(cr).height = 26
  sc(ws.getCell(cr, 6), { bg: C.ORANGE_DARK, fg: C.WHITE, bold: true, size: 11 })
  ws.getCell(cr, 6).value = 'TOTAL'
  const tc = ws.getCell(cr, 7)
  tc.value = { formula: `=SUM(G${firstDataRow}:G${lastDataRow})` }
  tc.fill  = fill(C.ORANGE_DARK)
  tc.font  = { name: 'Calibri', size: 13, bold: true, color: { argb: C.WHITE } }
  tc.alignment = { horizontal: 'center', vertical: 'middle' }
  tc.border = BORDER
  cr++

  // Nota de rodapé
  cr++
  ws.getRow(cr).height = 16
  ws.mergeCells(cr, 1, cr, 7)
  sc(ws.getCell(cr, 1), { bg: C.ORANGE_LITE, fg: 'FF999999', italic: true, h: 'center', size: 9, border: false })
  ws.getCell(cr, 1).value =
    `Gerado em ${new Date().toLocaleString('pt-BR')}  ·  Base: ${allItems.length} demandas totais  ·  KnowU CS`

  return { firstDataRow, lastDataRow, firstDataCol: 6 }
}

// ── Injeta gráfico nativo no XLSX via JSZip ───────────────────────────────
async function injectChart(buffer, { firstDataRow, lastDataRow }) {
  let JSZip, zip
  try {
    JSZip = require('jszip')
    zip   = await JSZip.loadAsync(buffer)
  } catch (e) {
    // Se JSZip falhar, devolve o buffer original sem o gráfico nativo
    console.error('injectChart: JSZip falhou, exportando sem gráfico nativo:', e.message)
    return buffer
  }
  let files
  try {
    files = Object.keys(zip.files)
  } catch (e) {
    console.error('injectChart: zip.files falhou:', e.message)
    return buffer
  }

  // Determina índices livres para drawings e charts
  const usedD = files.map(f => { const m = f.match(/drawings\/drawing(\d+)\.xml$/); return m ? +m[1] : 0 })
  const usedC = files.map(f => { const m = f.match(/charts\/chart(\d+)\.xml$/);     return m ? +m[1] : 0 })
  const drawN  = Math.max(0, ...usedD) + 1
  const chartN = Math.max(0, ...usedC) + 1

  // Detecta o número do arquivo sheet da aba "Analise" via workbook.xml + rels
  let sheetN = 5 // fallback seguro (_Listas=4, Analise=5)
  try {
    const wbXml  = await zip.files['xl/workbook.xml'].async('string')
    const wbRels = await zip.files['xl/_rels/workbook.xml.rels'].async('string')
    const sheetMatch = wbXml.match(/<sheet\s[^>]*name="Analise"[^>]*r:id="(rId\d+)"/)
    if (sheetMatch) {
      const rid = sheetMatch[1]
      const relMatch = wbRels.match(new RegExp(`Id="${rid}"[^>]*Target="worksheets/(sheet\\d+\\.xml)"`))
      if (relMatch) sheetN = +relMatch[1].match(/\d+/)[0]
    }
  } catch (_) { /* usa fallback */ }

  // ── Chart XML — visual profissional ──────────────────────────────────────
  const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:pPr algn="ctr">
          <a:defRPr b="1" sz="1200" spc="-100">
            <a:solidFill><a:srgbClr val="1A1A2E"/></a:solidFill>
            <a:latin typeface="Calibri"/>
          </a:defRPr>
        </a:pPr>
        <a:r><a:rPr lang="pt-BR" b="1" dirty="0"/><a:t>Demandas por Tipo</a:t></a:r>
        </a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:spPr>
        <a:solidFill><a:srgbClr val="F8F9FA"/></a:solidFill>
        <a:ln w="6350"><a:solidFill><a:srgbClr val="EBEBEB"/></a:solidFill></a:ln>
      </c:spPr>
      <c:layout>
        <c:manualLayout>
          <c:layoutTarget val="inner"/>
          <c:xMode val="edge"/><c:yMode val="edge"/>
          <c:x val="0.22"/><c:y val="0.08"/>
          <c:w val="0.72"/><c:h val="0.84"/>
        </c:manualLayout>
      </c:layout>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Analise!$G$${firstDataRow - 1}</c:f>
            <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Quantidade</c:v></c:pt></c:strCache>
          </c:strRef></c:tx>
          <c:spPr>
            <a:gradFill>
              <a:gsLst>
                <a:gs pos="0"><a:srgbClr val="F47B20"/></a:gs>
                <a:gs pos="60000"><a:srgbClr val="E8650D"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="C04E00"/></a:gs>
              </a:gsLst>
              <a:lin ang="0" scaled="0"/>
            </a:gradFill>
            <a:ln><a:noFill/></a:ln>
          </c:spPr>
          <c:invertIfNegative val="0"/>
          <c:dLbls>
            <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
            <c:txPr>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:pPr>
                <a:defRPr b="1" sz="950" spc="-50">
                  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                  <a:latin typeface="Calibri"/>
                </a:defRPr>
              </a:pPr></a:p>
            </c:txPr>
            <c:dLblPos val="inEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="1"/>
            <c:showCatName val="0"/><c:showSerName val="0"/>
            <c:showPercent val="0"/><c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Analise!$F$${firstDataRow}:$F$${lastDataRow}</c:f>
            <c:strCache><c:ptCount val="0"/></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Analise!$G$${firstDataRow}:$G$${lastDataRow}</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/><c:showVal val="0"/>
          <c:showCatName val="0"/><c:showSerName val="0"/>
          <c:showPercent val="0"/><c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:gapWidth val="65"/>
        <c:axId val="10000"/><c:axId val="10001"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="10000"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D0D0D0"/></a:solidFill></a:ln></c:spPr>
        <c:txPr>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr>
            <a:defRPr b="1" sz="1050" spc="0">
              <a:solidFill><a:srgbClr val="1A1A1A"/></a:solidFill>
              <a:latin typeface="Calibri"/>
            </a:defRPr>
          </a:pPr></a:p>
        </c:txPr>
        <c:crossAx val="10001"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="10001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:majorGridlines>
          <c:spPr>
            <a:ln w="6350" cap="flat">
              <a:solidFill><a:srgbClr val="DEDEDE"/></a:solidFill>
              <a:prstDash val="dash"/>
            </a:ln>
          </c:spPr>
        </c:majorGridlines>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
        <c:txPr>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr>
            <a:defRPr b="0" sz="900">
              <a:solidFill><a:srgbClr val="888888"/></a:solidFill>
              <a:latin typeface="Calibri"/>
            </a:defRPr>
          </a:pPr></a:p>
        </c:txPr>
        <c:crossAx val="10000"/>
        <c:crossBetween val="between"/>
        <c:tickMark val="none"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
      <c:txPr>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:pPr>
          <a:defRPr b="0" sz="900">
            <a:solidFill><a:srgbClr val="555555"/></a:solidFill>
            <a:latin typeface="Calibri"/>
          </a:defRPr>
        </a:pPr></a:p>
      </c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="D8D8D8"/></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="50800" dist="25400" dir="5400000" algn="tl" rotWithShape="0">
        <a:srgbClr val="C0C0C0"><a:alpha val="35000"/></a:srgbClr>
      </a:outerShdw>
    </a:effectLst>
  </c:spPr>
</c:chartSpace>`

  // ── Drawing XML (âncora do gráfico: col 8–23, row 2–22) ──────────────────
  const drawXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor moveWithCells="0" sizeWithCells="0">
    <xdr:from><xdr:col>8</xdr:col><xdr:colOff>114300</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>76200</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>23</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>21</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Grafico Analise"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`

  const drawRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>
</Relationships>`

  const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawN}.xml"/>
</Relationships>`

  // ── Adiciona arquivos ao ZIP ──────────────────────────────────────────────
  zip.file(`xl/charts/chart${chartN}.xml`, chartXml)
  zip.file(`xl/drawings/drawing${drawN}.xml`, drawXml)
  zip.file(`xl/drawings/_rels/drawing${drawN}.xml.rels`, drawRelsXml)

  // Sheet rels — adiciona ou mescla
  const sheetRelsPath = `xl/worksheets/_rels/sheet${sheetN}.xml.rels`
  if (zip.files[sheetRelsPath]) {
    let existing = await zip.files[sheetRelsPath].async('string')
    const nextId = (existing.match(/Id="rId(\d+)"/g) || []).length + 1
    const newRel = `  <Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawN}.xml"/>`
    existing = existing.replace('</Relationships>', newRel + '\n</Relationships>')
    zip.file(sheetRelsPath, existing)
    // Adiciona drawing com o id correto na sheet XML
    const sheetXml2 = await zip.files[`xl/worksheets/sheet${sheetN}.xml`].async('string')
    if (!sheetXml2.includes('<drawing')) {
      zip.file(`xl/worksheets/sheet${sheetN}.xml`,
        sheetXml2.replace('</worksheet>', `<drawing r:id="rId${nextId}"/>\n</worksheet>`))
    }
  } else {
    zip.file(sheetRelsPath, sheetRelsXml)
    const sheetXml2 = await zip.files[`xl/worksheets/sheet${sheetN}.xml`].async('string')
    if (!sheetXml2.includes('<drawing')) {
      zip.file(`xl/worksheets/sheet${sheetN}.xml`,
        sheetXml2.replace('</worksheet>', '<drawing r:id="rId1"/>\n</worksheet>'))
    }
  }

  // ── [Content_Types].xml ───────────────────────────────────────────────────
  let ct = await zip.files['[Content_Types].xml'].async('string')
  if (!ct.includes(`chart${chartN}.xml`))
    ct = ct.replace('</Types>',
      `<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n</Types>`)
  if (!ct.includes(`drawing${drawN}.xml`))
    ct = ct.replace('</Types>',
      `<Override PartName="/xl/drawings/drawing${drawN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n</Types>`)
  zip.file('[Content_Types].xml', ct)

  try {
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 4 } })
  } catch (e) {
    console.error('injectChart: generateAsync falhou:', e.message)
    return buffer
  }
}

// ── Exportar ──────────────────────────────────────────────────────────────
async function generateExcel(items, charts, allItems) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'KnowU CS'
  wb.created = new Date()
  buildDemandas(wb, items)
  buildResumo(wb, items)
  buildPorOperadora(wb, items)
  const analiseInfo = buildAnalise(wb, allItems || items)
  if (charts) buildGraficos(wb, charts)
  const buf = await wb.xlsx.writeBuffer()
  return injectChart(buf, analiseInfo)
}

module.exports = { generateExcel }
