const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const { spawn, spawnSync } = require('node:child_process')
const fsSync = require('node:fs')
const fs = require('node:fs/promises')
const path = require('node:path')

let mainWindow
let workerProcess
let cancelFile

function resolveWorkerPath(overridePath) {
  if (overridePath) return overridePath
  if (app.isPackaged) {
    const bundledExe = path.join(process.resourcesPath, 'app.asar.unpacked', 'worker', 'dist', 'occ_worker.exe')
    if (fsSync.existsSync(bundledExe)) return bundledExe
    return path.join(process.resourcesPath, 'app.asar.unpacked', 'worker', 'occ_worker.py')
  }
  return path.join(__dirname, '..', 'worker', 'occ_worker.py')
}

function resolveWorkerLaunch(workerPath, pythonPath) {
  const resolvedWorkerPath = resolveWorkerPath(workerPath)
  const isExecutable = resolvedWorkerPath.toLowerCase().endsWith('.exe')
  if (isExecutable) {
    return {
      command: resolvedWorkerPath,
      argsPrefix: [],
      workerPath: resolvedWorkerPath,
      mode: 'exe',
    }
  }

  return {
    ...resolvePythonLaunch(pythonPath),
    workerPath: resolvedWorkerPath,
    mode: 'python',
  }
}

function looksLikeWindowsAliasHint(text) {
  return /app[- ]?ausf|app execution aliases|run without arguments to install|python was not found/i.test(text)
}

function probePythonCandidate(command, prefixArgs = []) {
  const result = spawnSync(command, [...prefixArgs, '--version'], {
    windowsHide: true,
    encoding: 'utf8',
    stdio: 'pipe',
  })
  if (result.error) {
    return {
      ok: false,
      aliasHint: false,
      details: result.error.message,
    }
  }
  const output = `${result.stdout || ''}\n${result.stderr || ''}`.trim()
  if (result.status === 0) {
    return {
      ok: true,
      aliasHint: false,
      details: output,
    }
  }
  return {
    ok: false,
    aliasHint: looksLikeWindowsAliasHint(output),
    details: output || `Exit-Code ${result.status ?? 'unbekannt'}`,
  }
}

function resolvePythonLaunch(pythonPath) {
  if (pythonPath && pythonPath.trim()) {
    return { command: pythonPath.trim(), prefixArgs: [] }
  }

  const candidates = []
  if (process.platform === 'win32') {
    if (app.isPackaged) {
      const embeddedPython = path.join(process.resourcesPath, 'python', 'python.exe')
      candidates.push({ command: embeddedPython, prefixArgs: [] })
    }
    candidates.push({ command: 'py', prefixArgs: ['-3'] })
    candidates.push({ command: 'python', prefixArgs: [] })
    candidates.push({ command: 'python3', prefixArgs: [] })
  } else {
    candidates.push({ command: 'python3', prefixArgs: [] })
    candidates.push({ command: 'python', prefixArgs: [] })
  }

  let aliasHintSeen = false
  const tried = []
  for (const candidate of candidates) {
    const probe = probePythonCandidate(candidate.command, candidate.prefixArgs)
    const printable = `${candidate.command}${candidate.prefixArgs.length ? ` ${candidate.prefixArgs.join(' ')}` : ''}`
    tried.push(`${printable}${probe.details ? ` -> ${probe.details}` : ''}`)
    if (probe.ok) return candidate
    if (probe.aliasHint) aliasHintSeen = true
  }

  const aliasHint = aliasHintSeen
    ? 'Hinweis: Unter Windows blockiert oft der App-Ausfuehrungsalias fuer "python". In den Einstellungen kann der Alias deaktiviert werden.'
    : ''
  throw new Error(
    `Kein nutzbarer Python-Interpreter gefunden. Bitte Python 3.10+ installieren (inkl. py-Launcher) oder einen expliziten Interpreterpfad uebergeben. ${aliasHint} Gepruefte Kandidaten: ${tried.join(' | ')}`,
  )
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 940,
    minWidth: 1060,
    minHeight: 700,
    backgroundColor: '#edf2ef',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })
  mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
}

async function findFolders(root, onProgress) {
  const result = []
  let scannedCount = 0
  let foundCount = 0
  let skippedCount = 0
  let excludedCount = 0
  const maxWorkers = 8
  const readDirTimeoutMs = 10000
  const queue = [{ current: root, relative: path.basename(root) }]
  let lastProgressSentAt = 0

  function shouldSkipDirectory(name) {
    const lowered = name.toLowerCase()
    return lowered === 'protokollentwürfe' || lowered.includes('erledigt') || lowered.startsWith('zz_')
  }

  function emitProgress(relative, force = false) {
    const now = Date.now()
    // Drosseln, damit IPC bei sehr großen Verzeichnisbäumen nicht überläuft.
    if (!force && now - lastProgressSentAt < 120) return
    lastProgressSentAt = now
    onProgress?.({
      event: 'scan_progress',
      scannedCount,
      foundCount,
      skippedCount,
      excludedCount,
      currentPath: relative,
    })
  }

  async function readDirWithTimeout(current) {
    let timer
    try {
      const entries = await Promise.race([
        fs.readdir(current, { withFileTypes: true }),
        new Promise((_, reject) => {
          timer = setTimeout(() => reject(new Error('READDIR_TIMEOUT')), readDirTimeoutMs)
        }),
      ])
      return entries
    } finally {
      if (timer) clearTimeout(timer)
    }
  }

  async function workerLoop() {
    // eslint-disable-next-line no-constant-condition
    while (true) {
      const task = queue.shift()
      if (!task) return
      const { current, relative } = task

      scannedCount += 1
      emitProgress(relative)

      let entries
      try {
        entries = await readDirWithTimeout(current)
      } catch (error) {
        // Einzelne problematische Pfade (z. B. ENAMETOOLONG/Timeout im Netzlaufwerk)
        // sollen den gesamten Cloud-Import nicht abbrechen.
        skippedCount += 1
        console.warn(`[import-cloud] Skip unreadable path: ${current} (${error.message})`)
        emitProgress(relative, true)
        continue
      }

      const files = entries.filter((entry) => entry.isFile())
      const occFiles = files.filter((entry) => entry.name.toLowerCase().endsWith('.occ')).map((entry) => entry.name)
      const excelFiles = files
        .filter((entry) => /\.xls[xm]$/i.test(entry.name) && !entry.name.startsWith('~$') && !entry.name.toLowerCase().includes('wartung'))
        .map((entry) => entry.name)

      if (occFiles.length) {
        result.push({ path: relative, sourcePath: current, occFiles, excelFiles })
        foundCount += 1
        emitProgress(relative, true)
      }

      for (const entry of entries) {
        if (!entry.isDirectory()) continue
        if (shouldSkipDirectory(entry.name)) {
          excludedCount += 1
          continue
        }
        queue.push({
          current: path.join(current, entry.name),
          relative: path.join(relative, entry.name),
        })
      }
    }
  }

  const workerCount = Math.max(2, Math.min(maxWorkers, queue.length || 1))
  await Promise.all(Array.from({ length: workerCount }, () => workerLoop()))
  emitProgress(path.basename(root), true)
  return { folders: result, scannedCount, foundCount, skippedCount, excludedCount }
}

async function copyTree(source, destination) {
  await fs.mkdir(destination, { recursive: true })
  for (const entry of await fs.readdir(source, { withFileTypes: true })) {
    const sourcePath = path.join(source, entry.name)
    const targetPath = path.join(destination, entry.name)
    if (entry.isDirectory()) await copyTree(sourcePath, targetPath)
    else await fs.copyFile(sourcePath, targetPath, fs.constants.COPYFILE_EXCL)
  }
}

ipcMain.handle('choose-directory', async (_, title) => {
  const result = await dialog.showOpenDialog(mainWindow, { title, properties: ['openDirectory', 'createDirectory'] })
  return result.canceled ? null : result.filePaths[0]
})

ipcMain.handle('import-cloud', async (event, { source, destination }) => {
  event.sender.send('import-event', { event: 'scan_started' })
  const scan = await findFolders(source, (progress) => {
    event.sender.send('import-event', progress)
  })
  const folders = scan.folders
  const imported = []
  for (const folder of folders) {
    const relative = path.relative(source, folder.sourcePath)
    const target = path.join(destination, relative)
    const mappingState = folder.excelFiles.length === 1 ? 'bereit' : 'konflikt'
    const message = folder.excelFiles.length === 0
      ? 'Keine Excel-Datei im Fundordner gefunden.'
      : folder.excelFiles.length > 1
        ? 'Mehrere Excel-Dateien gefunden; manuelle Zuordnung erforderlich.'
        : undefined
    imported.push({ ...folder, path: relative || path.basename(source), localPath: target, state: mappingState, message })
  }
  event.sender.send('import-event', {
    event: 'scan_completed',
    scannedCount: scan.scannedCount,
    foundCount: scan.foundCount,
    skippedCount: scan.skippedCount,
    excludedCount: scan.excludedCount,
  })
  return imported
})

ipcMain.handle('prepare-local-folders', async (_, { folders }) => {
  if (!Array.isArray(folders)) throw new Error('Ungültige Kopierliste für lokale Ordner.')
  const prepared = []
  for (const folder of folders) {
    try {
      if (!folder?.sourcePath || !folder?.localPath) {
        throw new Error('sourcePath oder localPath fehlt')
      }
      await copyTree(folder.sourcePath, folder.localPath)
      prepared.push({ id: folder.id, localPath: folder.localPath })
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error)
      const label = folder?.path || folder?.id || folder?.sourcePath || 'Unbekannter Ordner'
      throw new Error(`Lokale Kopie fehlgeschlagen (${label}): ${reason}`)
    }
  }
  return prepared
})

ipcMain.handle('run-worker', async (event, { job, workerPath, pythonPath }) => {
  if (workerProcess) throw new Error('Es läuft bereits ein Verarbeitungslauf.')
  const jobFile = path.join(app.getPath('temp'), `omicron-job-${Date.now()}.json`)
  cancelFile = path.join(app.getPath('temp'), `omicron-cancel-${Date.now()}`)
  await fs.writeFile(jobFile, JSON.stringify(job, null, 2), 'utf8')
  const workerLaunch = resolveWorkerLaunch(workerPath, pythonPath)
  const workerArgs = workerLaunch.mode === 'exe'
    ? [jobFile, '--cancel-file', cancelFile]
    : [...workerLaunch.prefixArgs, workerLaunch.workerPath, jobFile, '--cancel-file', cancelFile]
  event.sender.send('worker-event', {
    event: 'worker_log',
    message: workerLaunch.mode === 'exe'
      ? `Worker-Start (EXE): ${workerLaunch.command} ${workerArgs.join(' ')}`
      : `Worker-Start (Python): ${workerLaunch.command} ${workerArgs.join(' ')}`,
  })
  workerProcess = spawn(workerLaunch.command, workerArgs, { windowsHide: false })
  workerProcess.stdout.on('data', (data) => {
    for (const line of data.toString().split(/\r?\n/).filter(Boolean)) {
      try { event.sender.send('worker-event', JSON.parse(line)) } catch { event.sender.send('worker-event', { event: 'worker_log', message: line }) }
    }
  })
  workerProcess.stderr.on('data', (data) => event.sender.send('worker-event', { event: 'worker_error', message: data.toString() }))
  workerProcess.on('error', (error) => event.sender.send('worker-event', { event: 'worker_error', message: error.message }))
  return new Promise((resolve) => workerProcess.on('close', (code) => {
    workerProcess = undefined
    resolve(code ?? 1)
  }))
})

ipcMain.handle('cancel-worker', async () => {
  if (cancelFile) await fs.writeFile(cancelFile, 'cancel', 'utf8')
})

ipcMain.handle('shutdown-computer', async () => {
  if (process.platform !== 'win32') {
    throw new Error('Automatisches Herunterfahren wird nur unter Windows unterstützt.')
  }

  const delaySeconds = 60
  const child = spawn('shutdown', ['/s', '/t', String(delaySeconds)], {
    windowsHide: true,
    detached: true,
    stdio: 'ignore',
  })
  child.unref()
  return { scheduled: true, delaySeconds }
})

app.whenReady().then(() => {
  createWindow()
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })