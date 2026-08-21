const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const { spawn } = require('node:child_process')
const fs = require('node:fs/promises')
const path = require('node:path')

let mainWindow
let workerProcess
let cancelFile

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

async function findFolders(root) {
  const result = []
  async function visit(current, relative) {
    const entries = await fs.readdir(current, { withFileTypes: true })
    const files = entries.filter((entry) => entry.isFile())
    const occFiles = files.filter((entry) => entry.name.toLowerCase().endsWith('.occ')).map((entry) => entry.name)
    const excelFiles = files
      .filter((entry) => /\.xls[xm]$/i.test(entry.name) && !entry.name.startsWith('~$') && !entry.name.toLowerCase().includes('wartung'))
      .map((entry) => entry.name)
    if (occFiles.length) result.push({ path: relative, sourcePath: current, occFiles, excelFiles })
    await Promise.all(entries
      .filter((entry) => entry.isDirectory() && entry.name !== 'Protokollentwürfe')
      .map((entry) => visit(path.join(current, entry.name), path.join(relative, entry.name))))
  }
  await visit(root, path.basename(root))
  return result
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

ipcMain.handle('import-cloud', async (_, { source, destination }) => {
  const folders = await findFolders(source)
  const imported = []
  for (const folder of folders) {
    const relative = path.relative(source, folder.sourcePath)
    const target = path.join(destination, relative)
    try {
      await copyTree(folder.sourcePath, target)
      const mappingState = folder.excelFiles.length === 1 ? 'bereit' : 'konflikt'
      const message = folder.excelFiles.length === 0
        ? 'Keine Excel-Datei im Fundordner gefunden.'
        : folder.excelFiles.length > 1
          ? 'Mehrere Excel-Dateien gefunden; manuelle Zuordnung erforderlich.'
          : undefined
      imported.push({ ...folder, path: relative || path.basename(source), localPath: target, state: mappingState, message })
    } catch (error) {
      imported.push({ ...folder, path: relative || path.basename(source), localPath: target, state: 'konflikt', message: error.message })
    }
  }
  return imported
})

ipcMain.handle('run-worker', async (event, { job, workerPath, pythonPath }) => {
  if (workerProcess) throw new Error('Es läuft bereits ein Verarbeitungslauf.')
  const jobFile = path.join(app.getPath('temp'), `omicron-job-${Date.now()}.json`)
  cancelFile = path.join(app.getPath('temp'), `omicron-cancel-${Date.now()}`)
  await fs.writeFile(jobFile, JSON.stringify(job, null, 2), 'utf8')
  const resolvedWorkerPath = workerPath || path.join(__dirname, '..', 'worker', 'occ_worker.py')
  workerProcess = spawn(pythonPath || 'python', [resolvedWorkerPath, jobFile, '--cancel-file', cancelFile], { windowsHide: false })
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

app.whenReady().then(() => {
  createWindow()
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })