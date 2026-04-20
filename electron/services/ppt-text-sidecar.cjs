const crypto = require('crypto')
const { spawn } = require('child_process')

function sha1(content) {
  return crypto.createHash('sha1').update(content).digest('hex')
}

function createPptTextSidecarService(options = {}) {
  const { fs, path, app } = options
  const repoRoot = path.join(__dirname, '..', '..')
  const runtimeDir = path.join(app.getPath('userData'), 'ppt-text-sidecar')
  const venvDir = path.join(runtimeDir, 'venv')
  const bootstrapStatePath = path.join(runtimeDir, 'bootstrap-state.json')
  const requirementsPath = path.join(repoRoot, 'scripts', 'ppt_text_sidecar', 'requirements.txt')
  const scriptPath = path.join(repoRoot, 'scripts', 'ppt_text_sidecar', 'main.py')
  const paddleCacheDir = path.join(runtimeDir, 'cache', 'paddleocr')
  const pipCacheDir = path.join(runtimeDir, 'cache', 'pip')
  const sharedCacheDir = path.join(runtimeDir, 'cache')
  const hfCacheDir = path.join(sharedCacheDir, 'huggingface')

  function ensureDir(dirPath) {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true })
    }
  }

  function getPythonExecutable() {
    if (process.platform === 'win32') {
      return path.join(venvDir, 'Scripts', 'python.exe')
    }
    return path.join(venvDir, 'bin', 'python3')
  }

  function getPipExecutable() {
    if (process.platform === 'win32') {
      return path.join(venvDir, 'Scripts', 'pip.exe')
    }
    return path.join(venvDir, 'bin', 'pip')
  }

  function execCommand(command, args, opts = {}) {
    return new Promise((resolve, reject) => {
      const child = spawn(command, args, {
        cwd: opts.cwd || repoRoot,
        env: {
          ...process.env,
          ...(opts.env || {}),
        },
        stdio: ['pipe', 'pipe', 'pipe'],
      })

      const stdout = []
      const stderr = []
      let timeoutId = null

      if (opts.stdin) {
        child.stdin.write(opts.stdin)
      }
      child.stdin.end()

      child.stdout.on('data', (chunk) => stdout.push(chunk))
      child.stderr.on('data', (chunk) => stderr.push(chunk))
      child.on('error', reject)

      if (opts.timeoutMs) {
        timeoutId = setTimeout(() => {
          child.kill('SIGKILL')
          reject(new Error(`命令执行超时: ${command} ${args.join(' ')}`))
        }, opts.timeoutMs)
      }

      child.on('close', (code) => {
        if (timeoutId) clearTimeout(timeoutId)
        const stdoutText = Buffer.concat(stdout).toString('utf8')
        const stderrText = Buffer.concat(stderr).toString('utf8')
        if (code === 0) {
          resolve({ stdout: stdoutText, stderr: stderrText })
          return
        }
        reject(
          new Error(
            `命令执行失败 (${code}): ${command} ${args.join(' ')}\n${stderrText || stdoutText}`,
          ),
        )
      })
    })
  }

  async function ensureVenv() {
    const pythonPath = getPythonExecutable()
    if (fs.existsSync(pythonPath)) return pythonPath

    ensureDir(runtimeDir)

    let bootstrapPython = process.env.PYTHON || ''
    if (!bootstrapPython) {
      for (const candidate of ['python3', 'python']) {
        try {
          await execCommand(candidate, ['--version'], { timeoutMs: 10000 })
          bootstrapPython = candidate
          break
        } catch {
          // try next
        }
      }
    }

    if (!bootstrapPython) {
      throw new Error('未找到 Python 运行时，无法启动 PPT 文本编辑 sidecar')
    }

    await execCommand(bootstrapPython, ['-m', 'venv', venvDir], {
      timeoutMs: 120000,
    })
    return pythonPath
  }

  async function ensureDependencies() {
    ensureDir(runtimeDir)
    ensureDir(pipCacheDir)
    ensureDir(paddleCacheDir)
    ensureDir(sharedCacheDir)
    ensureDir(hfCacheDir)

    const requirementsText = fs.readFileSync(requirementsPath, 'utf8')
    const requirementsHash = sha1(Buffer.from(requirementsText, 'utf8'))

    let currentHash = ''
    if (fs.existsSync(bootstrapStatePath)) {
      try {
        const state = JSON.parse(fs.readFileSync(bootstrapStatePath, 'utf8'))
        currentHash = String(state.requirementsHash || '')
      } catch {
        currentHash = ''
      }
    }

    const pythonPath = await ensureVenv()
    if (currentHash === requirementsHash) return pythonPath

    const pipPath = getPipExecutable()
    await execCommand(
      pipPath,
      ['install', '--upgrade', 'pip', 'setuptools', 'wheel'],
      {
        timeoutMs: 300000,
        env: {
          PIP_CACHE_DIR: pipCacheDir,
        },
      },
    )

    await execCommand(
      pipPath,
      ['install', '-r', requirementsPath],
      {
        timeoutMs: 1800000,
        env: {
          PIP_CACHE_DIR: pipCacheDir,
        },
      },
    )

    fs.writeFileSync(
      bootstrapStatePath,
      JSON.stringify(
        {
          requirementsHash,
          updatedAt: new Date().toISOString(),
        },
        null,
        2,
      ),
    )

    return pythonPath
  }

  async function invoke(command, payload = {}, options = {}) {
    const pythonPath = options.bootstrap === false ? getPythonExecutable() : await ensureDependencies()
    if (!fs.existsSync(pythonPath)) {
      throw new Error('PPT 文本编辑 sidecar 尚未初始化')
    }

    const stdin = JSON.stringify({
      ...payload,
      repo_root: repoRoot,
      paddle_cache_dir: paddleCacheDir,
    })
    const result = await execCommand(
      pythonPath,
      [scriptPath, command],
      {
        stdin,
        timeoutMs: options.timeoutMs || 300000,
        env: {
          PPT_TEXT_SIDECAR_RUNTIME: runtimeDir,
          PADDLEOCR_HOME: paddleCacheDir,
          HF_HOME: hfCacheDir,
          XDG_CACHE_HOME: sharedCacheDir,
          PYTORCH_ENABLE_MPS_FALLBACK: '1',
        },
      },
    )

    const output = String(result.stdout || '').trim()
    if (!output) {
      throw new Error(`PPT 文本 sidecar 未返回结果 (${command})`)
    }

    try {
      return JSON.parse(output)
    } catch (error) {
      throw new Error(`PPT 文本 sidecar 返回非法 JSON (${command}): ${output.slice(0, 300)}`)
    }
  }

  return {
    async health(options = {}) {
      try {
        const result = await invoke('health', {}, {
          bootstrap: options.bootstrap !== false,
          timeoutMs: 120000,
        })
        return { success: true, ...result, runtimeDir }
      } catch (error) {
        return {
          success: false,
          error: error.message || String(error),
          runtimeDir,
        }
      }
    },
    async detectTextBoxes(payload) {
      return invoke('detect_text_boxes', payload, { timeoutMs: 600000 })
    },
    async cleanupTextBoxes(payload) {
      return invoke('cleanup_text_boxes', payload, { timeoutMs: 600000 })
    },
    async recognizeText(payload) {
      return invoke('recognize_text', payload, { timeoutMs: 600000 })
    },
    async recognizeTextsBatch(payload) {
      return invoke('recognize_texts_batch', payload, { timeoutMs: 600000 })
    },
    async applyTextEdits(payload) {
      return invoke('apply_text_edits', payload, { timeoutMs: 600000 })
    },
  }
}

module.exports = {
  createPptTextSidecarService,
}
