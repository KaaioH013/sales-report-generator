import { spawn } from 'child_process'
import path from 'path'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const toolkit = path.join(__dirname, '..', '..', '..', 'docs', 'portfolio-toolkit', 'capture-portfolio-screenshots.mjs')
const config = path.join(__dirname, '..', 'portfolio.capture.json')

const child = spawn(process.execPath, [toolkit, `--config=${config}`], {
  stdio: 'inherit',
  cwd: path.join(__dirname, '..'),
})
child.on('exit', (code) => process.exit(code ?? 0))
