import { useEffect, useState } from 'react'
import {
  Archive,
  Check,
  CircleAlert,
  Cloud,
  Copy,
  FileSpreadsheet,
  FolderOpen,
  HardDrive,
  LoaderCircle,
  Play,
  RotateCcw,
  Settings2,
  TriangleAlert,
} from 'lucide-react'
import './App.css'

type FolderState = 'bereit' | 'konflikt' | 'fertig'

type WorkFolder = {
  id: string
  path: string
  occFiles: string[]
  excelFiles: string[]
  state: FolderState
  enabled?: boolean
  localPath?: string
  message?: string
  selectedExcelByOcc?: Record<string, string>
}

type DesktopApi = {
  chooseDirectory: (title: string) => Promise<string | null>
  importCloud: (source: string, destination: string) => Promise<Array<WorkFolder & { sourcePath: string }>>
  runWorker: (job: unknown, workerPath: string, pythonPath?: string) => Promise<number>
  cancelWorker: () => Promise<void>
  onWorkerEvent: (callback: (event: { event: string; itemId?: string; message?: string; index?: number; total?: number; itemCount?: number; occPath?: string; excelPath?: string; elapsedSeconds?: number; succeededCount?: number; failedCount?: number; skippedCount?: number; reportPath?: string }) => void) => () => void
}

declare global {
  interface Window {
    desktopApi?: DesktopApi
  }
}

type DirectoryHandle = FileSystemDirectoryHandle & {
  values(): AsyncIterableIterator<FileSystemHandle>
}

type DirectoryPickerWindow = Window & {
  showDirectoryPicker?: () => Promise<FileSystemDirectoryHandle>
}

const demoFolders: WorkFolder[] = [
  {
    id: 'bensheim',
    path: '2026 / Bensheim / Station Nord',
    occFiles: ['V24_NAP_V2_Bensheim.occ', 'V9_EZE_mit_Anregung_Bensheim.occ'],
    excelFiles: ['V19m_Übergeordneter_Entkupplungsschutz.xlsm'],
    state: 'bereit',
    enabled: true,
    selectedExcelByOcc: {
      'V24_NAP_V2_Bensheim.occ': 'V19m_Übergeordneter_Entkupplungsschutz.xlsm',
      'V9_EZE_mit_Anregung_Bensheim.occ': 'V19m_Übergeordneter_Entkupplungsschutz.xlsm',
    },
  },
  {
    id: 'gross-gerau',
    path: '2026 / Groß-Gerau / Station West',
    occFiles: ['Schutzprüfung_Station_West.occ'],
    excelFiles: ['Prüfdaten_Station_West.xlsm'],
    state: 'bereit',
    enabled: true,
    selectedExcelByOcc: {
      'Schutzprüfung_Station_West.occ': 'Prüfdaten_Station_West.xlsm',
    },
  },
]

function initSelectedExcelByOcc(occFiles: string[], excelFiles: string[]) {
  if (excelFiles.length !== 1) return {}
  const excel = excelFiles[0]
  return Object.fromEntries(occFiles.map((occ) => [occ, excel]))
}

function evaluateFolderState(folder: WorkFolder): FolderState {
  if (folder.state === 'fertig') return 'fertig'
  if (folder.excelFiles.length === 0) return 'konflikt'
  if (folder.excelFiles.length === 1) return 'bereit'

  const mapping = folder.selectedExcelByOcc ?? {}
  const allAssigned = folder.occFiles.every((occ) => {
    const selected = mapping[occ]
    return Boolean(selected) && folder.excelFiles.includes(selected)
  })
  return allAssigned ? 'bereit' : 'konflikt'
}

async function entriesOf(directory: DirectoryHandle) {
  const entries: FileSystemHandle[] = []
  for await (const entry of directory.values()) entries.push(entry)
  return entries
}

async function discoverWorkFolders(
  directory: DirectoryHandle,
  relativePath = directory.name,
): Promise<Array<WorkFolder & { handle: DirectoryHandle }>> {
  const entries = await entriesOf(directory)
  const files = entries.filter((entry): entry is FileSystemFileHandle => entry.kind === 'file')
  const folders = entries.filter((entry): entry is DirectoryHandle => entry.kind === 'directory')
  const occFiles = files.filter((file) => file.name.toLowerCase().endsWith('.occ')).map((file) => file.name)
  const excelFiles = files
    .filter((file) => /\.xls[xm]$/i.test(file.name) && !file.name.startsWith('~$') && !file.name.toLowerCase().includes('wartung'))
    .map((file) => file.name)
  const current = occFiles.length
    ? [{ id: relativePath, path: relativePath, occFiles, excelFiles, state: 'bereit' as const, handle: directory }]
    : []
  const nested = await Promise.all(
    folders
      .filter((folder) => folder.name !== 'Protokollentwürfe')
      .map((folder) => discoverWorkFolders(folder, `${relativePath} / ${folder.name}`)),
  )
  return [...current, ...nested.flat()]
}

async function createTargetFolder(destination: DirectoryHandle, path: string) {
  const segments = path.split(' / ')
  let parent = destination
  for (const segment of segments.slice(0, -1)) {
    parent = await parent.getDirectoryHandle(segment, { create: true }) as DirectoryHandle
  }
  const name = segments.at(-1)!
  try {
    await parent.getDirectoryHandle(name)
    return null
  } catch {
    return parent.getDirectoryHandle(name, { create: true }) as Promise<DirectoryHandle>
  }
}

async function copyFolder(source: DirectoryHandle, destination: DirectoryHandle) {
  for await (const entry of source.values()) {
    if (entry.kind === 'file') {
      const sourceFile = await entry.getFile()
      const targetFile = await destination.getFileHandle(entry.name, { create: true })
      const writer = await targetFile.createWritable()
      await writer.write(sourceFile)
      await writer.close()
    } else {
      const targetDirectory = await destination.getDirectoryHandle(entry.name, { create: true })
      await copyFolder(entry as DirectoryHandle, targetDirectory as DirectoryHandle)
    }
  }
}

function App() {
  const [cloudPath, setCloudPath] = useState(() => localStorage.getItem('omicron-cloud-path') ?? '')
  const [localPath, setLocalPath] = useState(() => localStorage.getItem('omicron-local-path') ?? '')
  const [cloudHandle, setCloudHandle] = useState<DirectoryHandle | null>(null)
  const [localHandle, setLocalHandle] = useState<DirectoryHandle | null>(null)
  const [folders, setFolders] = useState<WorkFolder[]>([])
  const [isScanning, setIsScanning] = useState(false)
  const [notice, setNotice] = useState('Wählen Sie Cloud-Quelle und lokalen Arbeitsordner, um den Testlauf vorzubereiten.')
  const [isRunning, setIsRunning] = useState(false)
  const [progress, setProgress] = useState({ completed: 0, total: 0, current: '', detail: '' })
  const [currentStep, setCurrentStep] = useState('Vorbereitung')
  const [runStartedAt, setRunStartedAt] = useState<number | null>(null)
  const [elapsedSeconds, setElapsedSeconds] = useState(0)

  function formatDuration(totalSeconds: number) {
    const minutes = Math.floor(totalSeconds / 60)
    const seconds = totalSeconds % 60
    return `${minutes}:${String(seconds).padStart(2, '0')}`
  }

  useEffect(() => {
    if (!isRunning || runStartedAt === null) return
    const timer = window.setInterval(() => {
      setElapsedSeconds(Math.max(0, Math.floor((Date.now() - runStartedAt) / 1000)))
    }, 1000)
    return () => window.clearInterval(timer)
  }, [isRunning, runStartedAt])

  useEffect(() => window.desktopApi?.onWorkerEvent((event) => {
    if (event.event === 'run_started') {
      setRunStartedAt(Date.now())
      setElapsedSeconds(0)
      setCurrentStep('Vorbereitung')
      setProgress({ completed: 0, total: event.itemCount ?? event.total ?? 0, current: '', detail: 'Vorbereitung' })
    }
    if (event.event === 'item_started') {
      setCurrentStep('Vorbereitung')
      setProgress((current) => ({ ...current, current: event.itemId ?? 'Fundordner', detail: `Eintrag ${event.index ?? 1} von ${event.total ?? current.total}` }))
      setNotice('Verarbeitung läuft. Bitte Omicron und Excel nicht bedienen.')
    }
    if (event.event === 'occ_started') {
      setCurrentStep('Datenexport (OCC)')
      setProgress((current) => ({ ...current, detail: `Omicron-Export: ${event.occPath ?? 'OCC-Datei'}` }))
    }
    if (event.event === 'excel_started') {
      setCurrentStep('Excel-Bearbeitung')
      setProgress((current) => ({ ...current, detail: `Excel-Verarbeitung: ${event.excelPath ?? 'Arbeitsmappe'}` }))
    }
    if (event.event === 'excel_completed') {
      setCurrentStep('Excel-Bearbeitung')
      setProgress((current) => ({ ...current, detail: `Excel abgeschlossen: ${event.excelPath ?? 'Arbeitsmappe'}` }))
    }
    if (event.event === 'mashup_terminated') {
      setCurrentStep('Mashup beenden')
      setProgress((current) => ({ ...current, detail: 'Vorbereitung: Mashup-Loader beendet' }))
    }
    if (event.event === 'mashup_not_running') {
      setCurrentStep('Mashup beenden')
      setProgress((current) => ({ ...current, detail: 'Vorbereitung: Kein laufender Mashup-Loader gefunden' }))
    }
    if (event.event === 'mashup_termination_failed') {
      setCurrentStep('Mashup beenden')
      setProgress((current) => ({ ...current, detail: 'Vorbereitung: Mashup-Loader konnte nicht beendet werden' }))
      setNotice(`Hinweis: Mashup-Loader konnte nicht beendet werden: ${event.message ?? 'Unbekannter Fehler'}`)
    }
    if (event.event === 'item_completed') {
      setCurrentStep('Abschluss')
      setFolders((current) => current.map((folder) => folder.id === event.itemId ? { ...folder, state: 'fertig' } : folder))
      setProgress((current) => ({ ...current, completed: current.completed + 1, detail: 'Eintrag erfolgreich abgeschlossen' }))
    }
    if (event.event === 'item_failed') {
      setCurrentStep('Fehlerbehandlung')
      setNotice(`Verarbeitung fehlgeschlagen: ${event.message ?? 'Unbekannter Fehler'}`)
      setProgress((current) => ({ ...current, completed: current.completed + 1, detail: 'Eintrag fehlgeschlagen' }))
    }
    if (event.event === 'worker_error') setNotice(`Python-Worker-Fehler: ${event.message ?? 'Unbekannter Fehler'}`)
    if (event.event === 'run_cancelled') {
      setCurrentStep('Abbruch')
      if (typeof event.elapsedSeconds === 'number') setElapsedSeconds(event.elapsedSeconds)
      setNotice(`Verarbeitung kontrolliert abgebrochen. Laufzeit: ${formatDuration(typeof event.elapsedSeconds === 'number' ? event.elapsedSeconds : elapsedSeconds)}.`)
      setRunStartedAt(null)
    }
    if (event.event === 'run_failed') {
      setCurrentStep('Fehlerbehandlung')
      if (typeof event.elapsedSeconds === 'number') setElapsedSeconds(event.elapsedSeconds)
      setNotice(`Verarbeitung gestoppt: ${event.message ?? 'Unbekannter Fehler'}. Laufzeit: ${formatDuration(typeof event.elapsedSeconds === 'number' ? event.elapsedSeconds : elapsedSeconds)}.`)
      setRunStartedAt(null)
    }
    if (event.event === 'run_report_written') {
      setProgress((current) => ({ ...current, detail: `Fehlerbericht geschrieben: ${event.reportPath ?? 'Pfad unbekannt'}` }))
    }
    if (event.event === 'run_report_failed') {
      setProgress((current) => ({ ...current, detail: 'Fehlerbericht konnte nicht geschrieben werden' }))
      setNotice(`Hinweis: Fehlerbericht konnte nicht geschrieben werden: ${event.message ?? 'Unbekannter Fehler'}`)
    }
    if (event.event === 'run_completed') {
      setCurrentStep('Fertig')
      if (typeof event.elapsedSeconds === 'number') setElapsedSeconds(event.elapsedSeconds)
      const succeeded = event.succeededCount ?? 0
      const failed = event.failedCount ?? 0
      const skipped = event.skippedCount ?? 0
      const reportInfo = event.reportPath ? ` Fehlerbericht: ${event.reportPath}.` : ''
      setNotice(`Verarbeitung abgeschlossen. Erfolg: ${succeeded}, Fehler: ${failed}, Übersprungen: ${skipped}. Laufzeit: ${formatDuration(typeof event.elapsedSeconds === 'number' ? event.elapsedSeconds : elapsedSeconds)}.${reportInfo}`)
      setRunStartedAt(null)
    }
  }), [])

  function setOccExcelMapping(folderId: string, occFile: string, excelFile: string) {
    setFolders((current) => current.map((folder) => {
      if (folder.id !== folderId || folder.state === 'fertig') return folder
      const nextMapping = { ...(folder.selectedExcelByOcc ?? {}) }
      if (excelFile) nextMapping[occFile] = excelFile
      else delete nextMapping[occFile]
      const updated: WorkFolder = {
        ...folder,
        selectedExcelByOcc: nextMapping,
      }
      return {
        ...updated,
        state: evaluateFolderState(updated),
      }
    }))
  }

  function toggleFolderEnabled(folderId: string) {
    setFolders((current) => current.map((folder) => {
      if (folder.id !== folderId || folder.state === 'fertig') return folder
      return {
        ...folder,
        enabled: !(folder.enabled ?? true),
      }
    }))
  }

  async function chooseDirectory(kind: 'cloud' | 'local') {
    if (window.desktopApi) {
      const selected = await window.desktopApi.chooseDirectory(kind === 'cloud' ? 'Cloud-Quellordner auswählen' : 'Lokalen Arbeitsordner auswählen')
      if (!selected) return
      if (kind === 'cloud') {
        setCloudPath(selected)
        localStorage.setItem('omicron-cloud-path', selected)
      } else {
        setLocalPath(selected)
        localStorage.setItem('omicron-local-path', selected)
      }
      setNotice(`${kind === 'cloud' ? 'Cloud-Quelle' : 'Lokaler Arbeitsordner'} ausgewählt.`)
      return
    }
    const picker = (window as DirectoryPickerWindow).showDirectoryPicker
    if (!picker) {
      setNotice('Die Ordnerauswahl benötigt einen aktuellen Chromium-basierten Browser.')
      return
    }
    try {
      const handle = await picker() as DirectoryHandle
      if (kind === 'cloud') {
        setCloudHandle(handle)
        setCloudPath(handle.name)
        localStorage.setItem('omicron-cloud-path', handle.name)
      } else {
        setLocalHandle(handle)
        setLocalPath(handle.name)
        localStorage.setItem('omicron-local-path', handle.name)
      }
      setNotice(`${kind === 'cloud' ? 'Cloud-Quelle' : 'Lokaler Arbeitsordner'} ausgewählt.`)
    } catch {
      setNotice('Ordnerauswahl abgebrochen.')
    }
  }

  async function importCloudFolders() {
    if (!window.desktopApi && (!cloudHandle || !localHandle)) {
      setNotice('Bitte wählen Sie zuerst beide Ordner über die Ordnerauswahl aus.')
      return
    }
    if (window.desktopApi && (!cloudPath || !localPath)) {
      setNotice('Bitte wählen Sie zuerst beide Ordner über die Ordnerauswahl aus.')
      return
    }
    setIsScanning(true)
    try {
      if (window.desktopApi) {
        const imported = await window.desktopApi.importCloud(cloudPath, localPath)
        setFolders(imported.map((folder) => {
          const selectedExcelByOcc = initSelectedExcelByOcc(folder.occFiles, folder.excelFiles)
          const enriched: WorkFolder = { ...folder, id: folder.path, selectedExcelByOcc }
          return { ...enriched, state: evaluateFolderState(enriched) }
        }))
        setNotice(`${imported.length} OCC-Fundordner wurden lokal bereitgestellt.`)
        return
      }
      if (!cloudHandle || !localHandle) {
        setNotice('Bitte wählen Sie zuerst beide Ordner über die Ordnerauswahl aus.')
        return
      }
      const discovered = await discoverWorkFolders(cloudHandle)
      const imported: WorkFolder[] = []
      for (const folder of discovered) {
        const target = await createTargetFolder(localHandle, folder.path)
        if (!target) {
          imported.push({ ...folder, state: 'konflikt', selectedExcelByOcc: {} })
          continue
        }
        await copyFolder(folder.handle, target)
        const selectedExcelByOcc = initSelectedExcelByOcc(folder.occFiles, folder.excelFiles)
        const enriched: WorkFolder = { ...folder, selectedExcelByOcc }
        imported.push({ ...enriched, state: evaluateFolderState(enriched) })
      }
      setFolders(imported)
      setNotice(`${imported.length} OCC-Fundordner wurden geprüft. Die Cloud-Quelle wurde nur gelesen.`)
    } catch (error) {
      setNotice(`Der Import konnte nicht abgeschlossen werden: ${error instanceof Error ? error.message : 'Unbekannter Fehler'}`)
    } finally {
      setIsScanning(false)
    }
  }

  function loadDemo() {
    setCloudPath('Cloud / Schutzprüfungen')
    setLocalPath('C:\\Omicron-Arbeitsordner')
    setFolders(demoFolders)
    setNotice('Beispieldaten geladen. Der Testlauf verändert keine Dateien.')
  }

  async function runTest() {
    const conflictFolders = folders.filter((folder) => folder.state === 'konflikt')
    if (conflictFolders.length) {
      setNotice(`Es gibt ${conflictFolders.length} Ordner mit ungeklärter Zuordnung. Bitte vor Export klären.`)
      return
    }

    const readyFolders = folders.filter((folder) => folder.state === 'bereit')
    if (!readyFolders.length) {
      setNotice('Es gibt keine konfliktfreien Fundordner für den Testlauf.')
      return
    }
    if (!window.desktopApi) {
      setFolders((current) => current.map((folder) => folder.state === 'bereit' ? { ...folder, state: 'fertig' } : folder))
      setNotice(`${readyFolders.length} Fundordner wurden im Browser-Testlauf simuliert.`)
      return
    }
    const items = readyFolders.map((folder) => ({
      enabled: folder.enabled ?? true,
      ...(folder.excelFiles.length === 1
        ? {
            id: folder.id,
            mappingStatus: 'eindeutig',
            occPaths: folder.occFiles.map((file) => `${folder.localPath}\\${file}`),
            excelPath: `${folder.localPath}\\${folder.excelFiles[0]}`,
          }
        : {
            id: folder.id,
            mappingStatus: 'eindeutig',
            mappings: folder.occFiles.map((occFile) => ({
              occPath: `${folder.localPath}\\${occFile}`,
              excelPath: `${folder.localPath}\\${(folder.selectedExcelByOcc ?? {})[occFile]}`,
            })),
          }),
    }))
    const timestamp = new Date().toISOString().replace(/[:]/g, '-').replace(/\..+$/, '')
    const reportPath = `${localPath}\\Protokollentwuerfe\\Fehlerbericht_${timestamp}.json`
    setIsRunning(true)
    setRunStartedAt(Date.now())
    setElapsedSeconds(0)
    setCurrentStep('Vorbereitung')
    setProgress({ completed: 0, total: readyFolders.length, current: '', detail: 'Worker wird gestartet' })
    setNotice('Sichtbare Omicron-Automatisierung gestartet. Bitte den Arbeitsplatz nicht bedienen.')
    try {
      await window.desktopApi.runWorker({ items, reportPath }, '')
    } catch (error) {
      setNotice(`Worker konnte nicht gestartet werden: ${error instanceof Error ? error.message : 'Unbekannter Fehler'}`)
    } finally {
      setIsRunning(false)
      setRunStartedAt(null)
    }
  }

  async function cancelRun() {
    await window.desktopApi?.cancelWorker()
    setNotice('Abbruch angefordert. Der aktuelle sichere Verarbeitungsschritt wird noch beendet.')
  }

  function reset() {
    setFolders([])
    setNotice('Ansicht zurückgesetzt. Gespeicherte Ordnernamen bleiben erhalten.')
  }

  const readyCount = folders.filter((folder) => folder.state === 'bereit').length
  const completeCount = folders.filter((folder) => folder.state === 'fertig').length
  const conflictCount = folders.filter((folder) => folder.state === 'konflikt').length
  const progressPercent = progress.total ? Math.round((progress.completed / progress.total) * 100) : 0

  return (
    <div className="app-shell">
      <header className="topbar">
        <a className="brand" href="#main" aria-label="Omicron Übernahmetool">
          <span className="brand-mark">O</span>
          <span><strong>Omicron</strong><small>Übernahmetool</small></span>
        </a>
        <div className="test-mode"><CircleAlert size={16} /> {window.desktopApi ? 'Windows-Desktopversion' : 'Browser-Testversion'}</div>
      </header>

      <main id="main" className="workspace">
        <aside className="workflow-nav" aria-label="Arbeitsablauf">
          <div className="workflow-heading"><span>Arbeitsablauf</span><small>V1</small></div>
          <ol>
            <li className="active"><span>1</span><div><strong>Bereitstellen</strong><small>Cloud lokal kopieren</small></div></li>
            <li className={folders.length ? 'active' : ''}><span>2</span><div><strong>Vorschau</strong><small>Zuordnungen prüfen</small></div></li>
            <li className={completeCount ? 'active' : ''}><span>3</span><div><strong>Ablage</strong><small>Protokollentwürfe</small></div></li>
          </ol>
          <div className="side-note"><HardDrive size={18} /><p><strong>Lokale Verarbeitung</strong><span>Die Cloud-Quelle wird ausschließlich gelesen.</span></p></div>
        </aside>

        <section className="content">
          <div className="page-heading">
            <div><span className="eyebrow">Erste Testversion</span><h1>Prüfungsordner bereitstellen</h1><p>OCC-Fundordner aus der Cloud lokal kopieren und vor der Datenübernahme prüfen.</p></div>
            <button className="icon-button" type="button" title="Ansicht zurücksetzen" onClick={reset}><RotateCcw size={18} /><span className="sr-only">Ansicht zurücksetzen</span></button>
          </div>

          <section className="configuration" aria-labelledby="config-heading">
            <div className="section-heading"><div className="section-icon"><Settings2 size={19} /></div><div><h2 id="config-heading">Ordner konfigurieren</h2><p>Die Namen der gewählten Ordner bleiben für den nächsten Start gespeichert.</p></div></div>
            <div className="folder-grid">
              <div className="folder-choice"><Cloud size={21} /><div><span>Cloud-Quellordner</span><strong>{cloudPath || 'Noch nicht ausgewählt'}</strong></div><button type="button" onClick={() => chooseDirectory('cloud')}>Auswählen</button></div>
              <div className="folder-choice"><HardDrive size={21} /><div><span>Lokaler Arbeitsordner</span><strong>{localPath || 'Noch nicht ausgewählt'}</strong></div><button type="button" onClick={() => chooseDirectory('local')}>Auswählen</button></div>
            </div>
            <div className="action-row"><button className="secondary-button" type="button" onClick={loadDemo}>Beispieldaten laden</button><button className="primary-button" type="button" onClick={importCloudFolders} disabled={isScanning}>{isScanning ? <LoaderCircle className="spin" size={18} /> : <Copy size={18} />}{isScanning ? 'Durchsuche und kopiere...' : 'Cloud-Ordner bereitstellen'}</button></div>
          </section>

          <section className="preview" aria-labelledby="preview-heading">
            <div className="preview-heading"><div><h2 id="preview-heading">Lokale Vorschau</h2><p>{folders.length ? `${folders.length} Fundordner im lokalen Arbeitsbereich` : 'Noch keine Fundordner bereitgestellt'}</p></div><div className="counts"><span>{readyCount} bereit</span><span>{conflictCount} Konflikte</span><span>{completeCount} erledigt</span></div></div>
            {folders.length ? <div className="folder-list">{folders.map((folder) => <article className={`folder-row ${folder.state}${!(folder.enabled ?? true) ? ' disabled' : ''}`} key={folder.id}><input type="checkbox" checked={folder.enabled ?? true} onChange={() => toggleFolderEnabled(folder.id)} disabled={isRunning || folder.state === 'fertig'} style={{ cursor: isRunning || folder.state === 'fertig' ? 'not-allowed' : 'pointer', width: 18, height: 18 }} /><div className="folder-icon">{folder.state === 'fertig' ? <Archive size={19} /> : <FolderOpen size={19} />}</div><div className="folder-main"><strong>{folder.state === 'fertig' ? `Protokollentwürfe / ${folder.path}` : folder.path}</strong><span>{folder.occFiles.length} OCC-Datei{folder.occFiles.length === 1 ? '' : 'en'} · {folder.excelFiles.length || 'keine'} Excel-Datei{folder.excelFiles.length === 1 ? '' : 'en'}</span><small>{folder.occFiles.join(' · ')}</small>{folder.message ? <small>{folder.message}</small> : null}{folder.state !== 'fertig' && folder.excelFiles.length > 1 ? <div style={{ marginTop: 8, display: 'grid', gap: 6 }}><small>Manuelle Zuordnung vor Start:</small>{folder.occFiles.map((occFile) => <label key={occFile} style={{ display: 'grid', gridTemplateColumns: 'minmax(180px,1fr) 1fr', alignItems: 'center', gap: 8 }}><span style={{ color: 'var(--muted)' }}>{occFile}</span><select value={(folder.selectedExcelByOcc ?? {})[occFile] ?? ''} onChange={(event) => setOccExcelMapping(folder.id, occFile, event.target.value)} disabled={isRunning} style={{ padding: '6px 8px', borderRadius: 4, border: '1px solid var(--border)', background: '#fff' }}><option value="">Excel auswählen...</option>{folder.excelFiles.map((excelFile) => <option key={excelFile} value={excelFile}>{excelFile}</option>)}</select></label>)}</div> : null}</div><div className={`state-badge ${folder.state}`}>{folder.state === 'fertig' ? <Check size={15} /> : folder.state === 'konflikt' ? <TriangleAlert size={15} /> : <FileSpreadsheet size={15} />}{folder.state === 'fertig' ? 'abgelegt' : folder.state === 'konflikt' ? 'Konflikt' : 'bereit'}</div></article>)}</div> : <div className="empty-state"><FolderOpen size={28} /><p>Wählen Sie Ordner aus oder laden Sie Beispieldaten, um die Vorschau zu testen.</p></div>}
          </section>

          {isRunning ? <section aria-label="Verarbeitungsfortschritt" style={{ padding: '15px 16px', marginTop: 18, background: 'var(--accent-soft)', border: '1px solid #b9ded0', borderRadius: 6 }}><div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}><strong style={{ color: 'var(--accent-dark)', fontSize: 13 }}>{progressPercent}% abgeschlossen</strong><span style={{ color: 'var(--muted)', fontSize: 11 }}>{progress.completed} von {progress.total} Fundordnern · Laufzeit {formatDuration(elapsedSeconds)}</span></div><div style={{ marginTop: 8, color: 'var(--accent-dark)', fontSize: 12 }}><strong>Aktueller Schritt:</strong> {currentStep}</div><div style={{ height: 9, overflow: 'hidden', margin: '10px 0 9px', background: '#cfe5dc', borderRadius: 999 }}><div style={{ width: `${progressPercent}%`, height: '100%', minWidth: 2, background: 'var(--accent)', borderRadius: 'inherit', transition: 'width .35s ease' }} /></div><div style={{ display: 'flex', alignItems: 'center', gap: 7, color: 'var(--muted)', fontSize: 11 }}><LoaderCircle className="spin" size={15} /><span>{progress.detail || 'Verarbeitung wird vorbereitet'}</span></div></section> : null}
          <div className="action-bar"><div className="status-message" role="status"><CircleAlert size={17} /><span>{notice}</span></div>{isRunning ? <button className="secondary-button" type="button" onClick={cancelRun}><TriangleAlert size={18} />Abbruch anfordern</button> : <button className="primary-button" type="button" onClick={runTest} disabled={!readyCount}><Play size={18} />Verarbeitung starten</button>}</div>
        </section>
      </main>
    </div>
  )
}

export default App
