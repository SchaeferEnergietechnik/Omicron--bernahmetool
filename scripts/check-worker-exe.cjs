const fs = require('node:fs')
const path = require('node:path')

const workerExePath = path.join(__dirname, '..', 'worker', 'dist', 'occ_worker.exe')

if (!fs.existsSync(workerExePath)) {
  console.error('Fehlend: worker/dist/occ_worker.exe')
  console.error('Bitte auf Windows zuerst den Worker bauen: npm run worker:build:win')
  process.exit(1)
}

console.log(`OK: ${workerExePath}`)
