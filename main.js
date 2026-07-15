const { app, BrowserWindow, ipcMain, dialog, shell, globalShortcut, Notification } = require('electron')
const path = require('path')
const fs   = require('fs')
const { generateExcel } = require('./excel')

app.setAppUserModelId('br.com.knowu.demandas-cs')

// ── Diretório de dados: Documentos/Demandas CS/ ───────────────────────────
const DATA_DIR     = path.join(app.getPath('documents'), 'Demandas CS')
const DADOS_FILE   = path.join(DATA_DIR, 'dados.json')
const HISTORICO    = path.join(DATA_DIR, 'Histórico de Demandas')

fs.mkdirSync(HISTORICO, { recursive: true })

// ── Pasta com o modelo visual e as logos das operadoras (AUTOMATIZAR FECHAMENTO) ──
// Em dev, fica um nível acima de demandas-cs-electron/. Mas quando o app roda
// empacotado (app.asar via deploy.ps1), __dirname vira um caminho virtual dentro
// do asar e '..' não alcança mais a pasta do projeto no OneDrive — por isso
// caímos para o caminho absoluto conhecido nesse caso.
function resolveLogosDir() {
  const candidates = [
    path.join(__dirname, '..', 'AUTOMATIZAR FECHAMENTO'),
    'C:\\Users\\rodri\\OneDrive\\Documentos\\Agentes de IA\\AUTOMATIZAR FECHAMENTO',
  ]
  return candidates.find(c => fs.existsSync(c)) || candidates[0]
}
const LOGOS_DIR = resolveLogosDir()

function normalizeLogoName(s) {
  return (s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // remove acentos
    .toUpperCase().replace(/\s+/g, ' ').trim()
}

// Acha a logo "UNIMED [NOME]" correspondente à operadora, tolerando acentos e
// pequenas variações de grafia entre o valor cadastrado no app e o nome do arquivo.
function findOperatorLogoFile(opName) {
  if (!fs.existsSync(LOGOS_DIR)) return null
  const files = fs.readdirSync(LOGOS_DIR).filter(f => /\.(png|jpe?g)$/i.test(f))
  const target = normalizeLogoName(opName)
  const baseName = f => normalizeLogoName(f.replace(/\.(png|jpe?g)$/i, ''))

  // Só igualdade exata (após normalizar acento/caixa) — nada de "contains",
  // que causa falso positivo entre nomes que são prefixo um do outro
  // (ex.: "UNIMED SA" é prefixo de "UNIMED SANTA MARIA").
  const match = files.find(f => baseName(f) === target)
    || files.find(f => baseName(f) === `UNIMED ${target.replace(/^UNIMED\s+/, '')}`)
  return match ? path.join(LOGOS_DIR, match) : null
}

// Retorna timestamp no fuso de Brasília (America/Sao_Paulo) formatado como YYYY-MM-DD_HH-MM
function brTimestamp() {
  return new Date()
    .toLocaleString('sv-SE', { timeZone: 'America/Sao_Paulo' })
    .slice(0, 16).replace(' ', '_').replace(':', '-')
}

// ── Auto-backup silencioso (a cada 4h, mantém últimos 7) ─────────────────
function autoBackup() {
  try {
    if (!fs.existsSync(DADOS_FILE)) return
    const items = JSON.parse(fs.readFileSync(DADOS_FILE, 'utf-8'))
    if (!Array.isArray(items) || !items.length) return
    const autos = fs.readdirSync(HISTORICO).filter(f => f.startsWith('auto_')).sort()
    while (autos.length >= 7) fs.unlinkSync(path.join(HISTORICO, autos.shift()))
    const ts = brTimestamp()
    fs.writeFileSync(path.join(HISTORICO, `auto_${ts}.json`), JSON.stringify(items, null, 2), 'utf-8')
  } catch { /* silencioso */ }
}

// ── Janela principal ──────────────────────────────────────────────────────
let win

function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 900,
    minHeight: 600,
    title: 'Demandas CS · KnowU',
    backgroundColor: '#0f0f11',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })

  win.setMenuBarVisibility(false)
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'))
  win.once('ready-to-show', () => win.show())
}

app.whenReady().then(() => {
  createWindow()
  globalShortcut.register('F12', () => {
    if (win && !win.isDestroyed()) win.webContents.toggleDevTools()
  })
  autoBackup()
  setInterval(autoBackup, 4 * 60 * 60 * 1000)
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
app.on('will-quit', () => globalShortcut.unregisterAll())
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })

// ── IPC: carregar dados ───────────────────────────────────────────────────
ipcMain.handle('load-data', () => {
  try {
    if (fs.existsSync(DADOS_FILE))
      return JSON.parse(fs.readFileSync(DADOS_FILE, 'utf-8'))
    return []
  } catch { return [] }
})

// ── IPC: salvar dados ─────────────────────────────────────────────────────
ipcMain.handle('save-data', (_, items) => {
  try {
    fs.writeFileSync(DADOS_FILE, JSON.stringify(items, null, 2), 'utf-8')
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

// ── IPC: backup → Histórico de Demandas ──────────────────────────────────
ipcMain.handle('backup', (_, items) => {
  try {
    const ts    = brTimestamp()
    const fname = `backup_${ts}.json`
    fs.writeFileSync(path.join(HISTORICO, fname), JSON.stringify(items, null, 2), 'utf-8')
    return { ok: true, arquivo: fname, total: items.length }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

// ── IPC: restaurar de JSON ────────────────────────────────────────────────
ipcMain.handle('restore', async () => {
  const result = await dialog.showOpenDialog(win, {
    title: 'Selecionar backup para restaurar',
    defaultPath: HISTORICO,
    filters: [{ name: 'Backup JSON', extensions: ['json'] }],
    properties: ['openFile'],
  })
  if (result.canceled || !result.filePaths.length) return null
  try {
    const data = JSON.parse(fs.readFileSync(result.filePaths[0], 'utf-8'))
    return Array.isArray(data) ? data : null
  } catch { return null }
})

// ── Settings (pasta de exportação, etc.) ─────────────────────────────────
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json')

function readSettings() {
  try { return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8')) } catch { return {} }
}
function writeSettings(obj) {
  const cur = readSettings()
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify({ ...cur, ...obj }, null, 2), 'utf-8')
}

// ── IPC: escolher pasta de exportação ────────────────────────────────────
ipcMain.handle('choose-export-folder', async () => {
  const cur = readSettings().exportFolder
  const result = await dialog.showOpenDialog(win, {
    title: 'Escolher pasta de exportação do Excel',
    defaultPath: cur || app.getPath('documents'),
    properties: ['openDirectory', 'createDirectory'],
  })
  if (result.canceled || !result.filePaths.length) return null
  const folder = result.filePaths[0]
  writeSettings({ exportFolder: folder })
  return folder
})

// ── IPC: obter pasta de exportação salva ──────────────────────────────────
ipcMain.handle('get-export-folder', () => readSettings().exportFolder || null)

// ── IPC: exportar Excel ───────────────────────────────────────────────────
ipcMain.handle('export-excel', async (_, { items, allItems, mes, charts }) => {
  try {
    const buf  = await generateExcel(items, charts, allItems || items)
    const meses = ['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
    const suf  = mes ? `_${meses[parseInt(mes)]}` : ''
    const date = new Date().toISOString().slice(0, 10)
    const fname = `demandas_cs${suf}_${date}.xlsx`

    const savedFolder = readSettings().exportFolder
    const defaultDir  = (savedFolder && fs.existsSync(savedFolder)) ? savedFolder : DATA_DIR
    const def = path.join(defaultDir, fname)

    const save = await dialog.showSaveDialog(win, {
      title: 'Salvar Excel',
      defaultPath: def,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    })
    if (save.canceled) return { ok: false }
    fs.writeFileSync(save.filePath, buf)
    shell.openPath(save.filePath)
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

// ── IPC: abrir pasta Histórico ────────────────────────────────────────────
ipcMain.handle('open-historico', () => shell.openPath(HISTORICO))

// printToPDF pageSize custom = POLEGADAS (não microns — essa é a unidade do
// webContents.print() para impressora física; printToPDF usa inches).
const SLIDE_PAGE_SIZE_IN = { width: 1672 / 96, height: 941 / 96 }

// ── Fechamento por operadora ──────────────────────────────────────────────
let fechamentoData = null
let fechamentoWin  = null

ipcMain.handle('open-fechamento', async (_, data) => {
  fechamentoData = data
  if (fechamentoWin && !fechamentoWin.isDestroyed()) {
    fechamentoWin.webContents.reload()
    fechamentoWin.focus()
    return { ok: true }
  }
  fechamentoWin = new BrowserWindow({
    // Só para visualização na tela — o export (PNG/PDF) NÃO depende do
    // tamanho desta janela (ver createOffscreenFechamento abaixo). Em telas
    // pequenas o usuário só rola pra ver o slide inteiro; a exportação sai
    // sempre completa e no tamanho certo.
    width: 1400,
    height: 900,
    title: `Fechamento ${data.opDisplay} — ${data.mes} ${data.ano}`,
    backgroundColor: '#22252f',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })
  fechamentoWin.setMenuBarVisibility(false)
  fechamentoWin.loadFile(path.join(__dirname, 'renderer', 'fechamento.html'))
  fechamentoWin.once('ready-to-show', () => fechamentoWin.show())
  fechamentoWin.on('closed', () => { fechamentoWin = null })
  return { ok: true }
})

ipcMain.handle('get-fechamento-data', () => fechamentoData)

// ── IPC: logo da operadora (busca em AUTOMATIZAR FECHAMENTO por "UNIMED [NOME]") ──
ipcMain.handle('get-operator-logo', (_, opName) => {
  try {
    const file = findOperatorLogoFile(opName)
    if (!file) return null
    const ext = path.extname(file).slice(1).toLowerCase()
    const mime = ext === 'jpg' ? 'jpeg' : ext
    const buf = fs.readFileSync(file)
    return `data:image/${mime};base64,${buf.toString('base64')}`
  } catch {
    return null
  }
})

// ── IPC: imagem-base da capa do Fechamento Completo (texto já removido) ──
ipcMain.handle('get-cover-image', () => {
  try {
    const file = path.join(LOGOS_DIR, 'CAPA FECHAMENTO NOVO_base.png')
    if (!fs.existsSync(file)) return null
    return 'data:image/png;base64,' + fs.readFileSync(file).toString('base64')
  } catch {
    return null
  }
})

// ── Renderização do slide para exportação (PNG/PDF) ────────────────────────
// Em telas pequenas (ou com escala do Windows >100%), o work area do monitor
// é menor que o slide de 1672×941 — a janela visível fica truncada pelo SO e
// capturar "o que está na tela" sai cortado. Solução: renderizar o slide numa
// janela dedicada, oculta e posicionada bem fora de qualquer monitor, do
// tamanho exato do conteúdo — assim a exportação nunca depende da resolução
// real da tela do usuário.
async function createOffscreenWindow(htmlFile, size = { width: 1720, height: 1040 }) {
  // Windows recorta o tamanho de QUALQUER BrowserWindow (mesmo oculta, mesmo
  // posicionada fora da tela) para caber no work area do monitor — inclusive
  // se width/height forem passados direto no construtor. Confirmado testando:
  // com width/height no construtor a janela sai grudada no tamanho da tela
  // (ex.: 1280×752 numa notebook pequena) mesmo pedindo 1720×1040, e é
  // exatamente isso que cortava/"zoomava" o PNG exportado. Criar a janela sem
  // tamanho e só then chamar setContentSize()/setBounds() contorna esse
  // recorte — confirmado com debug direto (getContentSize batendo 1720×1040).
  const win = new BrowserWindow({
    show: false,
    frame: false,
    skipTaskbar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })
  win.setContentSize(size.width, size.height)
  win.setBounds({ x: -20000, y: -20000, width: size.width, height: size.height })
  await win.loadFile(path.join(__dirname, 'renderer', htmlFile))
  return win
}

async function createOffscreenFechamento() {
  const win = await createOffscreenWindow('fechamento.html')
  await win.webContents.executeJavaScript(`
    new Promise((resolve) => {
      if (window.__fechamentoReady) resolve()
      else document.addEventListener('fechamento-ready', () => resolve())
    })
  `)
  const slideRect = await win.webContents.executeJavaScript(`
    (() => {
      const r = document.getElementById('slide').getBoundingClientRect()
      return { x: Math.round(r.left), y: Math.round(r.top), width: Math.round(r.width), height: Math.round(r.height) }
    })()
  `)
  return { win, slideRect }
}

// ── IPC: exporta o slide de Fechamento como PNG ─────────────────────────────
ipcMain.handle('capture-fechamento-png', async () => {
  const data = fechamentoData
  if (!data) return { error: 'Sem dados de fechamento carregados.' }
  let off
  try {
    off = await createOffscreenFechamento()
    const image = await off.win.webContents.capturePage(off.slideRect)
    const safeName = (data.opDisplay || data.op || 'op').replace(/[\/\\\s]+/g, '_')
    const fname    = `Fechamento_${safeName}_${data.mes || ''}_${data.ano || ''}.png`
    const settings = readSettings()
    const defPath  = path.join(settings.exportFolder || DATA_DIR, fname)

    const parentWin = (fechamentoWin && !fechamentoWin.isDestroyed()) ? fechamentoWin : win
    const save = await dialog.showSaveDialog(parentWin, {
      title: 'Salvar Fechamento como PNG',
      defaultPath: defPath,
      filters: [{ name: 'PNG', extensions: ['png'] }],
    })
    if (save.canceled) return { ok: false }
    fs.writeFileSync(save.filePath, image.toPNG())
    shell.openPath(save.filePath)
    return { ok: true }
  } catch (e) {
    return { error: e.message }
  } finally {
    if (off?.win && !off.win.isDestroyed()) off.win.destroy()
  }
})

// PDF é o próprio slide renderizado (via printToPDF), não uma reconstrução
// separada — assim PDF e PNG são sempre o mesmo design, sem risco de ficarem
// dessincronizados quando o layout do relatório mudar no futuro.
ipcMain.handle('print-fechamento-pdf', async () => {
  const data = fechamentoData
  if (!data) return { ok: false, error: 'Sem dados de fechamento carregados.' }
  let off
  try {
    off = await createOffscreenFechamento()
    // pageSize custom = POLEGADAS (não microns — essa é a unidade do
    // webContents.print() para impressora física; printToPDF usa inches).
    // margins = pixels. Sem marginType:'custom' os top/bottom/left/right
    // abaixo são ignorados e ele usa a margem 'default' do Chrome.
    const pdfBuffer = await off.win.webContents.printToPDF({
      printBackground: true,
      landscape: false,
      pageSize: SLIDE_PAGE_SIZE_IN,
      margins: { marginType: 'custom', top: 0, bottom: 0, left: 0, right: 0 },
    })

    const safeName = (data.opDisplay || 'op').replace(/[\/\\\s]+/g, '_')
    const fname    = `Fechamento_${safeName}_${data.mes}_${data.ano}.pdf`
    const settings = readSettings()
    const defPath  = path.join(settings.exportFolder || DATA_DIR, fname)

    const parentWin = (fechamentoWin && !fechamentoWin.isDestroyed()) ? fechamentoWin : win
    const save = await dialog.showSaveDialog(parentWin, {
      title: 'Salvar Fechamento como PDF',
      defaultPath: defPath,
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    })
    if (save.canceled) return { ok: false }
    fs.writeFileSync(save.filePath, pdfBuffer)
    shell.openPath(save.filePath)
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  } finally {
    if (off?.win && !off.win.isDestroyed()) off.win.destroy()
  }
})

// ── Fechamento Completo (capa + todas as operadoras num único PDF) ────────
// Reaproveita a mesma fechamento.html/fechamento-shared.js — só orquestra a
// capa + N slides de operadora dentro de fechamento-completo.html.
let fechamentoCompletoData = null
let fechamentoCompletoWin  = null

ipcMain.handle('open-fechamento-completo', async (_, data) => {
  fechamentoCompletoData = data
  if (fechamentoCompletoWin && !fechamentoCompletoWin.isDestroyed()) {
    fechamentoCompletoWin.webContents.reload()
    fechamentoCompletoWin.focus()
    return { ok: true }
  }
  fechamentoCompletoWin = new BrowserWindow({
    width: 1400,
    height: 900,
    title: `Fechamento Completo — ${data.mes} ${data.ano}`,
    backgroundColor: '#22252f',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })
  fechamentoCompletoWin.setMenuBarVisibility(false)
  fechamentoCompletoWin.loadFile(path.join(__dirname, 'renderer', 'fechamento-completo.html'))
  fechamentoCompletoWin.once('ready-to-show', () => fechamentoCompletoWin.show())
  fechamentoCompletoWin.on('closed', () => { fechamentoCompletoWin = null })
  return { ok: true }
})

ipcMain.handle('get-fechamento-completo-data', () => fechamentoCompletoData)

ipcMain.handle('print-fechamento-completo-pdf', async () => {
  const data = fechamentoCompletoData
  if (!data) return { ok: false, error: 'Sem dados carregados.' }
  let offWin
  try {
    offWin = await createOffscreenWindow('fechamento-completo.html')
    await offWin.webContents.executeJavaScript(`
      new Promise((resolve) => {
        if (window.__fechamentoCompletoReady) resolve()
        else document.addEventListener('fechamento-completo-ready', () => resolve())
      })
    `)
    const pdfBuffer = await offWin.webContents.printToPDF({
      printBackground: true,
      landscape: false,
      pageSize: SLIDE_PAGE_SIZE_IN,
      margins: { marginType: 'custom', top: 0, bottom: 0, left: 0, right: 0 },
    })

    const safeName = `Consolidado_${data.mes}_${data.ano}`.replace(/[\/\\\s]+/g, '_')
    const fname    = `Fechamento_${safeName}.pdf`
    const settings = readSettings()
    const defPath  = path.join(settings.exportFolder || DATA_DIR, fname)

    const parentWin = (fechamentoCompletoWin && !fechamentoCompletoWin.isDestroyed()) ? fechamentoCompletoWin : win
    const save = await dialog.showSaveDialog(parentWin, {
      title: 'Salvar Fechamento Completo como PDF',
      defaultPath: defPath,
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    })
    if (save.canceled) return { ok: false }
    fs.writeFileSync(save.filePath, pdfBuffer)
    shell.openPath(save.filePath)
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  } finally {
    if (offWin && !offWin.isDestroyed()) offWin.destroy()
  }
})

// ── IPC: notificação Windows ──────────────────────────────────────────────
ipcMain.handle('notify', (_, { title, body }) => {
  try {
    if (Notification.isSupported()) new Notification({ title, body }).show()
  } catch { /* silencioso */ }
})
