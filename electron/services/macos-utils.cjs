const crypto = require('crypto')
const { spawnSync } = require('child_process')

function stableHash(input) {
  return crypto.createHash('sha1').update(String(input || '')).digest('hex').slice(0, 16)
}

function runSwiftJson({ scriptPath, payload, timeoutMs = 30000 }) {
  const result = spawnSync('/usr/bin/swift', [scriptPath], {
    input: JSON.stringify(payload || {}),
    encoding: 'utf8',
    maxBuffer: 32 * 1024 * 1024,
    timeout: timeoutMs,
  })

  if (result.error) {
    throw result.error
  }

  if (result.status !== 0) {
    const stderr = (result.stderr || '').trim()
    throw new Error(stderr || `Swift bridge failed with status ${result.status}`)
  }

  const stdout = (result.stdout || '').trim()
  if (!stdout) {
    throw new Error('Swift bridge returned empty output')
  }

  try {
    return JSON.parse(stdout)
  } catch (error) {
    throw new Error(`Swift bridge returned invalid JSON: ${error.message}`)
  }
}

function ensureDir(fs, dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true })
  }
}

module.exports = {
  stableHash,
  runSwiftJson,
  ensureDir,
}
