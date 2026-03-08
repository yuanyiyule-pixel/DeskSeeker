#!/usr/bin/env node

import * as child_process from 'node:child_process';
import * as fs from 'node:fs/promises';
import * as http from 'node:http';
import * as https from 'node:https';
import * as path from 'node:path';
import * as tls from 'node:tls';
import * as util from 'node:util';

const execFileAsync = util.promisify(child_process.execFile);

const OPENROUTER_API_KEY_ENV_NAMES = ['OPENROUTER_API_KEY'];

const SUPPORTED_MODELS = new Map([
  ['gpt-5.4', 'openai/gpt-5.4'],
  ['openai/gpt-5.4', 'openai/gpt-5.4'],
  ['claude-sonnet-3.5', 'anthropic/claude-3.5-sonnet'],
  ['anthropic/claude-3.5-sonnet', 'anthropic/claude-3.5-sonnet'],
  ['claude-sonnet-4.6', 'anthropic/claude-4.6-sonnet'],
  ['anthropic/claude-4.6-sonnet', 'anthropic/claude-4.6-sonnet'],
  ['gemini-3-flash', 'google/gemini-3-flash-preview'],
  ['gemini-3-flash-preview', 'google/gemini-3-flash-preview'],
  ['google/gemini-3-flash-preview', 'google/gemini-3-flash-preview'],
  ['gemini-3.1-flash-preview', 'google/gemini-3.1-flash-lite-preview'],
  ['google/gemini-3.1-flash-preview', 'google/gemini-3.1-flash-lite-preview'],
  ['gemini-3.1-flash-lite', 'google/gemini-3.1-flash-lite-preview'],
  ['gemini-3.1-flash-lite-preview', 'google/gemini-3.1-flash-lite-preview'],
  ['google/gemini-3.1-flash-lite-preview', 'google/gemini-3.1-flash-lite-preview'],
]);
const MODEL_PROVIDER_PREFERENCES = new Map();
const DEFAULT_MODEL_ALIAS = 'gemini-3-flash-preview';
const DEFAULT_REASONING_EFFORT = 'medium';
const MIN_CONFIDENCE = 0.55;
const TOOL_MAX_STEPS = 8;
const TOOL_MIN_CROP_SIZE = 48;
const FINAL_BBOX_VERIFY_MARGIN = 28;
const ROOT_REVIEW_MAX_ATTEMPTS = 3;
const COARSE_GRID_MIN = 8;
const COARSE_GRID_MAX = 12;
const FINE_GRID_MIN = 5;
const FINE_GRID_MAX = 6;
const REVIEW_CROP_MIN_SIZE = 220;
const MAJORITY_PARALLEL_CALLS = 6;
const MAJORITY_QUORUM = 4;
const MAJORITY_MIN_AGREEMENT = 2;
const MAJORITY_ROUND_MAX_ATTEMPTS = 3;
const REVIEW_REQUIRED_PASS_VOTES = MAJORITY_MIN_AGREEMENT;
const MODEL_IMAGE_MAX_DIMENSION = 1024;
const CONTEXT_IMAGE_MAX_DIMENSION = 1024;
const THIRD_STAGE_DETAIL_TARGET_DIMENSION = 1400;
const THIRD_STAGE_MODEL_MAX_DIMENSION = 1400;
const SCREENSHOT_CAPTURE_MAX_ATTEMPTS = 4;
const SCREENSHOT_CAPTURE_RETRY_BASE_DELAY_MS = 250;
const OPENROUTER_MAX_RETRIES = 5;
const OPENROUTER_RETRY_BASE_DELAY_MS = 1200;
const LOG_PREFIX = '[DeskSeeker]';
const CLICK_SUCCESS_HINT = 'Please return the coordinate most likely to succeed when clicked.';
const THIRD_GRID_MIN = 6;
const THIRD_GRID_MAX = 8;

const runtimeState = {
  activeStage: undefined,
  verbose: false,
};

function nowIso() {
  return new Date().toISOString();
}

function elapsedMs(startNs) {
  return Number(process.hrtime.bigint() - startNs) / 1e6;
}

function formatDuration(ms) {
  if (!Number.isFinite(ms)) return 'n/a';
  if (ms < 1000) return `${ms.toFixed(0)}ms`;
  if (ms < 10000) return `${(ms / 1000).toFixed(2)}s`;
  return `${(ms / 1000).toFixed(1)}s`;
}

function previewText(value, maxLength = 220) {
  const text = String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
  if (text.length <= maxLength) return text;
  return `${text.slice(0, Math.max(0, maxLength - 3))}...`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function sleepAbortable(ms, signal) {
  if (!signal) return sleep(ms);
  if (signal.aborted) {
    return Promise.reject(new Error('aborted'));
  }
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      signal.removeEventListener('abort', onAbort);
      resolve();
    }, ms);
    function onAbort() {
      clearTimeout(timer);
      signal.removeEventListener('abort', onAbort);
      reject(new Error('aborted'));
    }
    signal.addEventListener('abort', onAbort);
  });
}

function summarizeError(err) {
  if (!err) return 'unknown error';
  if (typeof err === 'string') return previewText(err, 320);
  const parts = [];
  if (err.name) parts.push(String(err.name));
  if (err.code) parts.push(`code=${String(err.code)}`);
  if (err.signal) parts.push(`signal=${String(err.signal)}`);
  if (err.message) parts.push(previewText(err.message, 240));
  return parts.length ? parts.join(' ') : previewText(String(err), 320);
}

function formatLogValue(value) {
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'string') return previewText(value, 220);
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) return previewText(value.map((item) => formatLogValue(item) ?? '').join(','), 220);
  try {
    return previewText(JSON.stringify(value), 220);
  } catch {
    return previewText(String(value), 220);
  }
}

function formatLogDetails(details) {
  if (!details || typeof details !== 'object') return '';
  const parts = [];
  for (const [key, value] of Object.entries(details)) {
    const formatted = formatLogValue(value);
    if (formatted !== undefined && formatted !== '') {
      parts.push(`${key}=${formatted}`);
    }
  }
  return parts.join(' | ');
}

function logEvent(level, message, details) {
  if (!runtimeState.verbose) return;
  const suffix = formatLogDetails(details);
  process.stderr.write(
    `${LOG_PREFIX} ${nowIso()} ${level} ${message}${suffix ? ` | ${suffix}` : ''}\n`,
  );
}

async function withTiming(label, details, fn) {
  logEvent('START', label, details);
  const startNs = process.hrtime.bigint();
  const previousStage = runtimeState.activeStage;
  runtimeState.activeStage = { label, details };
  try {
    const result = await fn();
    logEvent('DONE', label, { ...details, elapsed: formatDuration(elapsedMs(startNs)) });
    return result;
  } catch (err) {
    logEvent('FAIL', label, {
      ...details,
      elapsed: formatDuration(elapsedMs(startNs)),
      error: summarizeError(err),
    });
    throw err;
  } finally {
    runtimeState.activeStage = previousStage;
  }
}

function printHelp() {
  process.stdout.write(`
Usage:
  node scripts/run.mjs --description "..."

Requires:
  Windows PowerShell
  Node.js >= 18

Required:
  --description <text>           Natural-language target description

Options:
  --task <text>                  Alias of --description
  --model <name>                 Alias or OpenRouter model id. Built-ins: gpt-5.4 | claude-sonnet-3.5 | claude-sonnet-4.6 | gemini-3-flash (default) | gemini-3.1-flash-lite
  --reasoning-effort <level>     medium (default) | low | high | minimal | none
  --screenshot <filepath>        Reuse an existing desktop screenshot instead of capturing a new one
  --review                       Enable the final review vote stage (disabled by default)
  --verbose                      Print verbose progress logs to stderr
  --dump-raw                     Write visible raw model responses and trace to a sidecar JSON file
  --dry-run                      Capture screenshot only, no network call, no click
  --out <filepath>               Save JSON result to this path
  --help                         Show help
`);
}

function fail(message, code = 1) {
  const msg = typeof message === 'string' ? message : String(message);
  logEvent('FATAL', 'Terminating run', {
    code,
    stage: runtimeState.activeStage?.label,
    message: msg,
  });
  process.stderr.write(`${msg}\n`);
  process.exit(code);
}

function sanitizeString(s) {
  const idx = s.indexOf('data:image/');
  if (idx === -1) return s;
  const head = s.slice(0, idx);
  const prefix = s.slice(idx, idx + 64);
  return `${head}${prefix}...(omitted)`;
}

function parseArgs(argv) {
  const out = {
    description: undefined,
    model: DEFAULT_MODEL_ALIAS,
    reasoningEffort: DEFAULT_REASONING_EFFORT,
    screenshotPath: undefined,
    review: false,
    verbose: false,
    dumpRaw: false,
    dryRun: false,
    outPath: undefined,
    help: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--help') out.help = true;
    else if (arg === '--dry-run') out.dryRun = true;
    else if (arg === '--review') out.review = true;
    else if (arg === '--verbose') out.verbose = true;
    else if (arg === '--dump-raw') out.dumpRaw = true;
    else if (arg === '--description' || arg === '--task') out.description = argv[++i];
    else if (arg === '--model') out.model = argv[++i];
    else if (arg === '--reasoning-effort') out.reasoningEffort = argv[++i];
    else if (arg === '--screenshot') out.screenshotPath = argv[++i];
    else if (arg.startsWith('--description=')) out.description = arg.slice('--description='.length);
    else if (arg.startsWith('--task=')) out.description = arg.slice('--task='.length);
    else if (arg.startsWith('--model=')) out.model = arg.slice('--model='.length);
    else if (arg.startsWith('--reasoning-effort=')) {
      out.reasoningEffort = arg.slice('--reasoning-effort='.length);
    }
    else if (arg.startsWith('--screenshot=')) out.screenshotPath = arg.slice('--screenshot='.length);
    else if (arg === '--out') out.outPath = argv[++i];
    else if (arg.startsWith('--out=')) out.outPath = arg.slice('--out='.length);
    else if (arg.startsWith('-')) fail(`Unknown option: ${arg}`);
    else fail(`Unexpected arg: ${arg}`);
  }

  return out;
}

function resolveModelName(value) {
  const raw = String(value || '').trim().toLowerCase();
  const key = raw.replace(/[\s_]+/g, '-');
  const collapsedKey = key.replace(/-+/g, '-');
  const resolved = SUPPORTED_MODELS.get(raw) || SUPPORTED_MODELS.get(key) || SUPPORTED_MODELS.get(collapsedKey);
  if (resolved) return resolved;
  if (raw.includes('/')) return raw;
  fail(
    `Unsupported model: ${value}. Supported values: ${Array.from(
      new Set([
        'gpt-5.4',
        'claude-sonnet-3.5',
        'claude-sonnet-4.6',
        'gemini-3-flash',
        'gemini-3.1-flash-preview',
        'gemini-3.1-flash-lite',
      ]),
    ).join(', ')}`,
  );
}

function resolveReasoningEffort(value) {
  const effort = String(value || '')
    .trim()
    .toLowerCase();
  if (['none', 'minimal', 'low', 'medium', 'high'].includes(effort)) return effort;
  fail(`Unsupported reasoning effort: ${value}. Supported values: none, minimal, low, medium, high`);
}

function resolveProxyUrlFromEnv() {
  const env = process.env;
  const raw =
    env.HTTPS_PROXY ||
    env.https_proxy ||
    env.ALL_PROXY ||
    env.all_proxy ||
    env.HTTP_PROXY ||
    env.http_proxy ||
    undefined;
  if (!raw || typeof raw !== 'string') return undefined;
  const s = raw.trim();
  return s.length ? s : undefined;
}

function normalizeProxyUrl(raw) {
  const s = String(raw || '').trim();
  if (!s) return undefined;
  if (/^https?:\/\//i.test(s)) return s;
  return `http://${s}`;
}

function pickProxyFromProxyServerString(s, scheme) {
  const text = String(s || '').trim();
  if (!text) return undefined;

  const parts = text.split(';').map((p) => p.trim()).filter(Boolean);
  const kv = {};
  for (const p of parts) {
    const idx = p.indexOf('=');
    if (idx !== -1) {
      const k = p.slice(0, idx).trim().toLowerCase();
      const v = p.slice(idx + 1).trim();
      if (k && v) kv[k] = v;
    }
  }
  const chosen = kv[scheme] || kv.http || kv.https;
  if (chosen) return normalizeProxyUrl(chosen);
  return normalizeProxyUrl(text);
}

async function resolveProxyUrlFromWindows() {
  if (process.platform !== 'win32') return undefined;

  try {
    const { stdout } = await execFileAsync('netsh', ['winhttp', 'show', 'proxy'], {
      windowsHide: true,
      timeout: 3000,
      maxBuffer: 1024 * 1024,
    });
    const out = String(stdout || '');
    if (/Direct access \(no proxy server\)/i.test(out)) {
      return undefined;
    }
    const line = out
      .split(/\r?\n/)
      .map((l) => l.trim())
      .find((l) => /^Proxy Server\(s\)\s*:/i.test(l));
    if (line) {
      const idx = line.indexOf(':');
      const value = idx !== -1 ? line.slice(idx + 1).trim() : '';
      const picked = pickProxyFromProxyServerString(value, 'https');
      if (picked) return picked;
    }
  } catch {
    // Ignore and fall back to user Internet Settings.
  }

  try {
    const key = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings';
    const enable = await execFileAsync('reg', ['query', key, '/v', 'ProxyEnable'], {
      windowsHide: true,
      timeout: 3000,
      maxBuffer: 1024 * 1024,
    });
    const enableOut = String(enable.stdout || '');
    const enabled = /ProxyEnable\s+REG_DWORD\s+0x1\b/i.test(enableOut);
    if (!enabled) return undefined;

    const server = await execFileAsync('reg', ['query', key, '/v', 'ProxyServer'], {
      windowsHide: true,
      timeout: 3000,
      maxBuffer: 1024 * 1024,
    });
    const serverOut = String(server.stdout || '');
    const match = /ProxyServer\s+REG_SZ\s+(.+)$/im.exec(serverOut);
    const raw = match ? match[1].trim() : '';
    const picked = pickProxyFromProxyServerString(raw, 'https');
    if (picked) return picked;
  } catch {
    // Best effort only.
  }

  return undefined;
}

async function resolveProxyUrlAuto() {
  return resolveProxyUrlFromEnv() ?? (await resolveProxyUrlFromWindows());
}

class HttpsProxyAgent extends https.Agent {
  constructor(proxyUrl) {
    super({ keepAlive: true });
    this.proxyUrl = new URL(String(proxyUrl));
  }

  createConnection(options, callback) {
    const cb = typeof callback === 'function' ? callback : undefined;
    const targetHost = options.hostname || options.host;
    const targetPort = options.port || 443;
    const target = `${targetHost}:${targetPort}`;

    const proxyHost = this.proxyUrl.hostname;
    const proxyPort = this.proxyUrl.port ? Number(this.proxyUrl.port) : 80;
    const headers = { Host: target };

    if (this.proxyUrl.username || this.proxyUrl.password) {
      const user = decodeURIComponent(this.proxyUrl.username || '');
      const pass = decodeURIComponent(this.proxyUrl.password || '');
      const token = Buffer.from(`${user}:${pass}`, 'utf8').toString('base64');
      headers['Proxy-Authorization'] = `Basic ${token}`;
    }

    const req = http.request({
      host: proxyHost,
      port: proxyPort,
      method: 'CONNECT',
      path: target,
      headers,
    });

    req.once('error', (err) => {
      if (cb) cb(err);
    });

    req.once('connect', (res, socket) => {
      if (res.statusCode !== 200) {
        socket.destroy();
        if (cb) cb(new Error(`Proxy CONNECT failed with status ${res.statusCode}`));
        return;
      }
      const tlsSocket = tls.connect({
        socket,
        servername: options.servername || targetHost,
      });
      if (cb) cb(null, tlsSocket);
    });

    req.end();
    return undefined;
  }
}

function loadApiKeyFromEnv() {
  for (const envName of OPENROUTER_API_KEY_ENV_NAMES) {
    const value = process.env[envName];
    if (typeof value !== 'string') continue;
    const trimmed = value.trim();
    if (!trimmed) continue;
    return {
      apiKey: trimmed,
      source: `env:${envName}`,
    };
  }
  return undefined;
}

async function resolveApiKey() {
  return loadApiKeyFromEnv();
}

async function postOpenRouterChatCompletions({ apiKey, body, agent, signal }) {
  const endpoint = new URL('https://openrouter.ai/api/v1/chat/completions');
  const payload = JSON.stringify(body);

  const headers = {
    Authorization: `Bearer ${apiKey}`,
    'Content-Type': 'application/json',
    'Accept-Encoding': 'identity',
    'Content-Length': Buffer.byteLength(payload),
  };

  const options = {
    protocol: endpoint.protocol,
    hostname: endpoint.hostname,
    port: endpoint.port ? Number(endpoint.port) : 443,
    method: 'POST',
    path: endpoint.pathname + endpoint.search,
    headers,
    agent,
  };

  return await new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(Buffer.from(c)));
      res.on('end', () => {
        resolve({
          ok: res.statusCode >= 200 && res.statusCode < 300,
          status: res.statusCode || 0,
          text: Buffer.concat(chunks).toString('utf8'),
        });
      });
    });
    req.on('error', reject);
    if (signal) {
      if (signal.aborted) {
        req.destroy(new Error('aborted'));
        return;
      }
      const onAbort = () => {
        req.destroy(new Error('aborted'));
      };
      signal.addEventListener('abort', onAbort, { once: true });
      req.on('close', () => signal.removeEventListener('abort', onAbort));
    }
    req.write(payload);
    req.end();
  });
}

function extractAssistantText(content) {
  if (typeof content === 'string') return content;
  if (!Array.isArray(content)) return undefined;
  const parts = [];
  for (const item of content) {
    if (item && typeof item === 'object' && item.type === 'text' && typeof item.text === 'string') {
      parts.push(item.text);
    }
  }
  const joined = parts.join('');
  return joined.length ? joined : undefined;
}

function mimeFromExt(ext) {
  const e = ext.toLowerCase();
  if (e === '.png') return 'image/png';
  if (e === '.jpg' || e === '.jpeg') return 'image/jpeg';
  if (e === '.webp') return 'image/webp';
  if (e === '.gif') return 'image/gif';
  return undefined;
}

async function toDataUrlFromLocalFile(filePath, explicitMime) {
  const abs = path.resolve(filePath);
  const ext = path.extname(abs);
  const inferred = mimeFromExt(ext);
  const mime = explicitMime ?? inferred;
  if (!mime) {
    fail(`Cannot infer mime from extension (${ext || 'none'})`);
  }
  const buf = await fs.readFile(abs);
  return { dataUrl: `data:${mime};base64,${buf.toString('base64')}`, mime };
}

function computeImageScaleToFit(
  width,
  height,
  maxDimension = MODEL_IMAGE_MAX_DIMENSION,
  options = {},
) {
  const allowUpscale = options.allowUpscale === true;
  const safeWidth = Math.max(1, Number(width) || 1);
  const safeHeight = Math.max(1, Number(height) || 1);
  const longestSide = Math.max(safeWidth, safeHeight);
  if (!allowUpscale && longestSide <= maxDimension) return 1;
  if (allowUpscale && Math.abs(longestSide - maxDimension) < 1e-6) return 1;
  return maxDimension / longestSide;
}

async function resizeImageToDimensions({
  image,
  outPath,
  label,
  targetWidth,
  targetHeight,
}) {
  const safeTargetWidth = Math.max(1, Math.round(Number(targetWidth) || 1));
  const safeTargetHeight = Math.max(1, Math.round(Number(targetHeight) || 1));
  if (
    safeTargetWidth === Math.max(1, Math.round(Number(image.width) || 1)) &&
    safeTargetHeight === Math.max(1, Math.round(Number(image.height) || 1))
  ) {
    return {
      path: path.resolve(image.path),
      width: image.width,
      height: image.height,
      scaleX: 1,
      scaleY: 1,
    };
  }
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$src = '${escapePsSingleQuoted(path.resolve(image.path))}'
$out = '${escapePsSingleQuoted(path.resolve(outPath))}'
$bmp = [System.Drawing.Bitmap]::FromFile($src)
$scaled = New-Object System.Drawing.Bitmap(${safeTargetWidth}, ${safeTargetHeight})
$g = [System.Drawing.Graphics]::FromImage($scaled)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.DrawImage($bmp, 0, 0, ${safeTargetWidth}, ${safeTargetHeight})
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$scaled.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$scaled.Dispose()
$bmp.Dispose()
`;
  await runPowerShell(script, 20000, label);
  return {
    path: path.resolve(outPath),
    width: safeTargetWidth,
    height: safeTargetHeight,
    scaleX: safeTargetWidth / Math.max(1, Number(image.width) || 1),
    scaleY: safeTargetHeight / Math.max(1, Number(image.height) || 1),
  };
}

async function resizeImageForModel({ image, outPath, label, maxDimension = MODEL_IMAGE_MAX_DIMENSION }) {
  const scale = computeImageScaleToFit(image.width, image.height, maxDimension);
  if (scale >= 0.9999) {
    return {
      path: path.resolve(image.path),
      width: image.width,
      height: image.height,
      scaleX: 1,
      scaleY: 1,
    };
  }
  const targetWidth = Math.max(1, Math.round(image.width * scale));
  const targetHeight = Math.max(1, Math.round(image.height * scale));
  return await resizeImageToDimensions({
    image,
    outPath,
    label,
    targetWidth,
    targetHeight,
  });
}

async function upscaleImageForDetail({
  image,
  outPath,
  label,
  targetDimension = THIRD_STAGE_DETAIL_TARGET_DIMENSION,
}) {
  const scale = computeImageScaleToFit(image.width, image.height, targetDimension, {
    allowUpscale: true,
  });
  if (scale <= 1.0001) {
    return {
      path: path.resolve(image.path),
      width: image.width,
      height: image.height,
      scaleX: 1,
      scaleY: 1,
    };
  }
  const targetWidth = Math.max(1, Math.round(image.width * scale));
  const targetHeight = Math.max(1, Math.round(image.height * scale));
  return await resizeImageToDimensions({
    image,
    outPath,
    label,
    targetWidth,
    targetHeight,
  });
}

async function toModelDataUrl({
  image,
  outPath,
  label,
  explicitMime = 'image/png',
  maxDimension = MODEL_IMAGE_MAX_DIMENSION,
}) {
  const preparedImage = await resizeImageForModel({
    image,
    outPath,
    label,
    maxDimension,
  });
  const { dataUrl, mime } = await toDataUrlFromLocalFile(preparedImage.path, explicitMime);
  return {
    dataUrl,
    mime,
    image: preparedImage,
  };
}

function extractJsonObject(text) {
  const src = String(text || '').trim();
  if (!src) return undefined;

  const fenced = src.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const raw = fenced ? fenced[1].trim() : src;

  const start = raw.indexOf('{');
  if (start === -1) return undefined;

  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < raw.length; i += 1) {
    const ch = raw[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === '{') depth += 1;
    else if (ch === '}') {
      depth -= 1;
      if (depth === 0) return raw.slice(start, i + 1);
    }
  }
  return undefined;
}

function normalizeAction(value) {
  const s = String(value || '').trim().toLowerCase();
  if (!s) return 'none';
  if (s === 'left_click' || s === 'left-click' || s === 'left' || s === 'click') return 'left_click';
  if (s === 'right_click' || s === 'right-click' || s === 'right') return 'right_click';
  return 'none';
}

function normalizeTargetDescription(description) {
  const raw = String(description || '').trim();
  if (!raw) return raw;

  const sentenceParts = raw
    .split(/(?<=[。！？!?;；\.\n])\s*/u)
    .map((part) => part.trim())
    .filter(Boolean);

  const operationalPatterns = [
  /return.*(coordinate|center|bbox|bounding box|clickable|click area)/iu,
  /(do not|don't|avoid).*(neighbor|adjacent|nearby)/iu,
  /(whole|full).*(button|icon|control|taskbar button|hit area|clickable area)/iu,
  /(center).*(clickable area|bbox|button|icon)/iu,
  /(output|return).*(json|coordinate|bbox|center)/iu,
  ];

  const filtered = sentenceParts.filter((part) => {
    if (/coordinate most likely to succeed when clicked/i.test(part)) return true;
    return !operationalPatterns.some((pattern) => pattern.test(part));
  });

  const normalized = filtered.join(' ').replace(/\s+/g, ' ').trim() || raw;
  if (/coordinate most likely to succeed when clicked/i.test(normalized)) return normalized;
  return `${normalized} ${CLICK_SUCCESS_HINT}`.trim();
}

function fatalOrThrow(message, fatal = true) {
  if (fatal) fail(message);
  throw new Error(String(message));
}

function clampInt(value, min, max) {
  const n = Number(value);
  if (!Number.isFinite(n)) return min;
  return Math.max(min, Math.min(max, Math.round(n)));
}

function clampFloat(value, min, max) {
  const n = Number(value);
  if (!Number.isFinite(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function parseModelJson(text, fatal = true) {
  const jsonText = extractJsonObject(text);
  if (!jsonText) {
    fatalOrThrow(`Model did not return a JSON object: ${sanitizeString(String(text || ''))}`, fatal);
  }
  try {
    return JSON.parse(jsonText);
  } catch {
    fatalOrThrow(`Unable to parse model JSON: ${jsonText}`, fatal);
  }
}

function makeRequestPromptPrefix() {
  return `${Math.floor(1000 + Math.random() * 9000)}\n`;
}

function prefixPromptMessages(messages, prefix) {
  return messages.map((message) => {
    if (typeof message?.content === 'string') {
      return {
        ...message,
        content: `${prefix}${message.content}`,
      };
    }
    if (Array.isArray(message?.content)) {
      return {
        ...message,
        content: message.content.map((part) => {
          if (part?.type !== 'text' || typeof part?.text !== 'string') return part;
          return {
            ...part,
            text: `${prefix}${part.text}`,
          };
        }),
      };
    }
    return message;
  });
}

async function requestModelJson({
  apiKey,
  agent,
  model,
  reasoningEffort,
  messages,
  debugLabel = 'request',
  fatal = true,
  signal,
}) {
  const requestPromptPrefix = makeRequestPromptPrefix();
  const prefixedMessages = prefixPromptMessages(messages, requestPromptPrefix);
  const body = {
    model,
    reasoning: {
      effort: reasoningEffort,
      exclude: true,
    },
    messages: prefixedMessages,
    temperature: 0.1,
    stream: false,
  };
  const provider = MODEL_PROVIDER_PREFERENCES.get(model);
  if (provider) {
    body.provider = provider;
  }

  return await withTiming(
    `OpenRouter ${debugLabel}`,
    {
      model,
      reasoning_effort: reasoningEffort,
      messages: prefixedMessages.length,
      provider: provider ? provider.order.join(',') : 'default',
      proxy: Boolean(agent),
    },
    async () => {
      let lastFailure = 'unknown OpenRouter failure';
      for (let attempt = 1; attempt <= OPENROUTER_MAX_RETRIES; attempt += 1) {
        if (signal?.aborted) {
          throw new Error('aborted');
        }
        let resp;
        try {
          resp = await postOpenRouterChatCompletions({ apiKey, body, agent, signal });
        } catch (err) {
          lastFailure = summarizeError(err);
          if (signal?.aborted || /aborted/i.test(lastFailure)) {
            throw new Error('aborted');
          }
          if (attempt >= OPENROUTER_MAX_RETRIES) {
            fatalOrThrow(lastFailure, fatal);
          }
          const delayMs = OPENROUTER_RETRY_BASE_DELAY_MS * attempt;
          logEvent('WARN', 'Retrying OpenRouter request after transport failure', {
            label: debugLabel,
            attempt,
            next_attempt: attempt + 1,
            delay_ms: delayMs,
            reason: lastFailure,
          });
          await sleepAbortable(delayMs, signal);
          continue;
        }
        logEvent('INFO', 'OpenRouter HTTP response received', {
          label: debugLabel,
          attempt,
          status: resp.status,
          ok: resp.ok,
          bytes: Buffer.byteLength(String(resp.text || ''), 'utf8'),
        });

        const rawText = String(resp.text || '');
        const looksTransientHttpFailure = !resp.ok && resp.status >= 500;
        if (looksTransientHttpFailure) {
          lastFailure = `OpenRouter API error (${resp.status}) during ${debugLabel}: ${sanitizeString(rawText)}`;
        } else {
          let parsedResponse;
          try {
            parsedResponse = JSON.parse(rawText);
          } catch {
            fatalOrThrow(
              `Unexpected non-JSON OpenRouter response during ${debugLabel}: ${sanitizeString(rawText).slice(0, 400)}`,
              fatal,
            );
          }

          if (parsedResponse?.error) {
            lastFailure = `OpenRouter returned an error payload during ${debugLabel}: ${sanitizeString(JSON.stringify(parsedResponse.error))}`;
            const code = Number(parsedResponse.error?.code);
            const retriableErrorCode = Number.isFinite(code) ? code >= 500 : false;
            if (!retriableErrorCode) {
              fatalOrThrow(lastFailure, fatal);
            }
          } else {
            const message =
              parsedResponse && parsedResponse.choices && parsedResponse.choices[0]
                ? parsedResponse.choices[0].message
                : undefined;
            const assistantText = extractAssistantText(message?.content);
            if (String(assistantText || '').trim()) {
              const json = parseModelJson(assistantText, fatal);
              logEvent('INFO', 'Parsed model JSON', {
                label: debugLabel,
                attempt,
                type: json?.type,
                tool: json?.tool,
                action: json?.action,
                image_id: json?.image_id,
                status: json?.status,
                confidence: json?.confidence,
                assistant_preview: assistantText,
              });
              return {
                assistantText: String(assistantText || ''),
                json,
                rawResponseText: rawText,
              };
            }
            lastFailure = `Model returned no text content during ${debugLabel}. Raw response: ${sanitizeString(rawText).slice(0, 4000)}`;
          }
        }

        if (attempt >= OPENROUTER_MAX_RETRIES) {
          fatalOrThrow(lastFailure, fatal);
        }

        const delayMs = OPENROUTER_RETRY_BASE_DELAY_MS * attempt;
        logEvent('WARN', 'Retrying OpenRouter request after transient failure', {
          label: debugLabel,
          attempt,
          next_attempt: attempt + 1,
          delay_ms: delayMs,
          reason: lastFailure,
        });
        await sleepAbortable(delayMs, signal);
      }

      fatalOrThrow(lastFailure, fatal);
    },
  );
}

function summarizeMajorityCandidate(items, keyFn) {
  const groups = new Map();
  for (const item of items) {
    const key = String(keyFn(item));
    const existing = groups.get(key) || {
      key,
      count: 0,
      confidenceSum: 0,
      items: [],
    };
    existing.count += 1;
    existing.confidenceSum += Number(item?.json?.confidence || 0);
    existing.items.push(item);
    groups.set(key, existing);
  }
  const ranked = Array.from(groups.values()).sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    if (b.confidenceSum !== a.confidenceSum) return b.confidenceSum - a.confidenceSum;
    return String(a.key).localeCompare(String(b.key));
  });
  return ranked[0];
}

function majorityCount(summary) {
  return Number(summary?.count || 0);
}

function buildWorkerVariantHint(stageName, workerIndex) {
  const variants = {
    coarse: [
      'Focus on exact app identity first, then map it to one coarse grid cell in the labeled screenshot.',
      'Focus on immediate left/right neighbors and taskbar order before choosing a coarse cell.',
      'Focus on the safest clickable center of the target, not the largest visible fragment inside a cell.',
      'Double-check the row and column label directly from the labeled grid screenshot before answering.',
      'Reject similar-looking icons unless the app identity and neighboring cues both match.',
      'Do not rely on color, approximate position, or a generic icon silhouette alone; confirm the exact named target.',
      'Use the labeled coarse-grid screenshot for both semantics and coordinates; do not assume a hidden clean image exists.',
    ],
    fine: [
      'Use the single fine-grid neighborhood screenshot for both semantics and returned stage-2 cell coordinates.',
      'Focus on the exact app identity inside the cyan-outlined coarse cell before choosing a stage-2 fine cell.',
      'Verify that the cyan-outlined coarse cell really contains the target before answering.',
      'If the target is actually in a neighboring coarse cell, use outside_direction instead of forcing a wrong fine cell.',
      'Do not rely on color, approximate position, or generic shape alone; confirm the exact named target before selecting a fine cell.',
      'Choose the fine cell that contains the safest eventual click point, not merely the most colorful fragment.',
      'Return the stage-2 cell from the fine-grid image only.',
    ],
    third: [
      'Use image 1 for the stage-2 neighborhood without grid and image 2 for the final local grid refinement.',
      'Choose one third-stage grid cell and one or two safe click anchors inside it, based primarily on image 2.',
      'Confirm the target identity from image 1 before finalizing the cell and anchors from image 2.',
      'Prefer stable interior body pixels, not thin outlines, badges, or decorative corners.',
      'Do not rely on color, approximate position, or generic shape alone; confirm the exact named target before selecting anchors.',
      'If two anchors are returned, they must both lie inside the same chosen third-stage cell.',
      'Use image 2 for coordinates and image 1 for semantic disambiguation.',
    ],
    review: [
      'Pass only if the single marked screenshot clearly shows the thin red box enclosing the exact named target and not a neighbor.',
      'Focus on app identity and adjacency before approving the marker.',
      'Reject the candidate if the red box encloses a lookalike, a merely plausible nearby control, misses the target body, or sits too close to another clickable item.',
      'Judge both identity and precision from the same marked full screenshot with the thin red box; do not pass based only on color, rough position, or a reasonable-looking control.',
      'Be conservative: retry on ambiguity.',
      'Confirm the exact target named in the description, not a nearby approximate match.',
    ],
  };
  const list = variants[stageName] || variants.coarse;
  return list[(workerIndex - 1) % list.length];
}

async function requestModelJsonMajority({
  apiKey,
  agent,
  model,
  reasoningEffort,
  messages,
  debugLabel,
  voteFn,
  stageName = 'coarse',
  parallelCount = MAJORITY_PARALLEL_CALLS,
  quorum = MAJORITY_QUORUM,
}) {
  return await withTiming(
    `Majority vote ${debugLabel}`,
    {
      model,
      reasoning_effort: reasoningEffort,
      parallel_count: parallelCount,
      quorum,
    },
    async () => {
      let lastFailureMessage = undefined;
      for (let roundAttempt = 1; roundAttempt <= MAJORITY_ROUND_MAX_ATTEMPTS; roundAttempt += 1) {
        const pending = Array.from({ length: parallelCount }, (_, index) => {
          const controller = new AbortController();
          return {
            worker: index + 1,
            controller,
            promise: requestModelJson({
              apiKey,
              agent,
              model,
              reasoningEffort,
              messages: messages.map((message, messageIndex) => {
                if (messageIndex !== 0 || typeof message.content !== 'string') return message;
                return {
                  ...message,
                  content: `${message.content}\nWorker lens: ${buildWorkerVariantHint(stageName, index + 1)}`,
                };
              }),
              debugLabel: `${debugLabel} round ${roundAttempt} worker ${index + 1}`,
              fatal: false,
              signal: controller.signal,
            }),
          };
        });
        for (const entry of pending) {
          entry.promise.catch(() => {});
        }
        const successes = [];
        const failures = [];

        while (pending.length > 0 && successes.length < quorum) {
          const raced = await Promise.race(
            pending.map((entry) =>
              entry.promise
                .then((value) => ({ entry, ok: true, value }))
                .catch((error) => ({ entry, ok: false, error })),
            ),
          );
          const idx = pending.findIndex((entry) => entry.worker === raced.entry.worker);
          if (idx !== -1) pending.splice(idx, 1);
          if (raced.ok) {
            successes.push({
              worker: raced.entry.worker,
              ...raced.value,
            });
          } else {
            failures.push({
              worker: raced.entry.worker,
              error: summarizeError(raced.error),
            });
            logEvent('WARN', 'Majority worker failed', {
              label: debugLabel,
              round_attempt: roundAttempt,
              worker: raced.entry.worker,
              error: summarizeError(raced.error),
            });
            if (successes.length + pending.length < quorum) {
              break;
            }
          }
        }

        for (const entry of pending) {
          entry.controller.abort();
        }

        if (successes.length < quorum) {
          lastFailureMessage =
            `Unable to gather ${quorum} successful responses for ${debugLabel}. ` +
            `Round=${roundAttempt}, Successes=${successes.length}, failures=${failures.length}`;
          if (roundAttempt >= MAJORITY_ROUND_MAX_ATTEMPTS) {
            break;
          }
          logEvent('WARN', 'Majority vote round will retry after insufficient successful responses', {
            label: debugLabel,
            round_attempt: roundAttempt,
            successes: successes.length,
            failures: failures.length,
            required_successes: quorum,
          });
          continue;
        }

        const voted = voteFn(successes);
        if (voted.count < MAJORITY_MIN_AGREEMENT) {
          lastFailureMessage =
            `No agreement of at least ${MAJORITY_MIN_AGREEMENT} votes for ${debugLabel}. ` +
            `Round=${roundAttempt}, winner=${voted.summary}`;
          if (roundAttempt >= MAJORITY_ROUND_MAX_ATTEMPTS) {
            break;
          }
          logEvent('WARN', 'Majority vote round will retry after weak agreement', {
            label: debugLabel,
            round_attempt: roundAttempt,
            winner: voted.summary,
            min_agreement: MAJORITY_MIN_AGREEMENT,
          });
          continue;
        }

        logEvent('INFO', 'Majority vote completed', {
          label: debugLabel,
          round_attempt: roundAttempt,
          successes: successes.length,
          failures: failures.length,
          winner: voted.summary,
        });
        return {
          ...voted,
          ballots: successes,
          failures,
        };
      }
      fatalOrThrow(lastFailureMessage || `Majority vote failed for ${debugLabel}`, true);
    },
  );
}

function normalizeDecision(rawDecision, screen) {
  const widthMax = Math.max(0, screen.width - 1);
  const heightMax = Math.max(0, screen.height - 1);
  const target = rawDecision && typeof rawDecision === 'object' ? rawDecision.target : undefined;
  const bbox = rawDecision && typeof rawDecision === 'object' ? rawDecision.bbox : undefined;

  const action = normalizeAction(rawDecision?.action ?? rawDecision?.button);
  const confidenceRaw = Number(rawDecision?.confidence);
  const confidence = Number.isFinite(confidenceRaw)
    ? Math.max(0, Math.min(1, confidenceRaw))
    : 0;

  const boxX = clampInt(bbox?.x ?? target?.x ?? 0, 0, widthMax);
  const boxY = clampInt(bbox?.y ?? target?.y ?? 0, 0, heightMax);
  const boxWidth = clampInt(bbox?.width ?? 1, 1, screen.width);
  const boxHeight = clampInt(bbox?.height ?? 1, 1, screen.height);
  const x = clampInt(boxX + (boxWidth - 1) / 2, 0, widthMax);
  const y = clampInt(boxY + (boxHeight - 1) / 2, 0, heightMax);
  const reason = String(rawDecision?.reason || '').trim().slice(0, 160);
  const matchedText = String(rawDecision?.matched_text || rawDecision?.label || '').trim().slice(0, 120);

  let decision;

  if (action === 'none' || confidence < MIN_CONFIDENCE) {
    decision = {
      action: 'none',
      target: { x, y },
      bbox: { x: boxX, y: boxY, width: boxWidth, height: boxHeight },
      confidence,
      reason: reason || 'low confidence or ambiguous target',
      matched_text: matchedText,
    };
  } else {
    decision = {
      action,
      target: { x, y },
      bbox: { x: boxX, y: boxY, width: boxWidth, height: boxHeight },
      confidence,
      reason: reason || 'target located',
      matched_text: matchedText,
    };
  }

  return decision;
}

function normalizeImageRelativeBBox(raw, image) {
  return {
    x: clampFloat(raw?.bbox?.x ?? raw?.target?.x ?? 0, 0, Math.max(0, image.width - 1)),
    y: clampFloat(raw?.bbox?.y ?? raw?.target?.y ?? 0, 0, Math.max(0, image.height - 1)),
    width: clampFloat(raw?.bbox?.width ?? 1, 1, image.width),
    height: clampFloat(raw?.bbox?.height ?? 1, 1, image.height),
  };
}

function expandRectWithinImage(rect, image, margin) {
  const x0 = Math.max(0, Math.floor(rect.x - margin));
  const y0 = Math.max(0, Math.floor(rect.y - margin));
  const x1 = Math.min(image.width, Math.ceil(rect.x + rect.width + margin));
  const y1 = Math.min(image.height, Math.ceil(rect.y + rect.height + margin));
  return {
    x: x0,
    y: y0,
    width: Math.max(TOOL_MIN_CROP_SIZE, x1 - x0),
    height: Math.max(TOOL_MIN_CROP_SIZE, y1 - y0),
  };
}

function expandRectWithinImageXY(rect, image, marginX, marginY) {
  const x0 = Math.max(0, Math.floor(rect.x - marginX));
  const y0 = Math.max(0, Math.floor(rect.y - marginY));
  const x1 = Math.min(image.width, Math.ceil(rect.x + rect.width + marginX));
  const y1 = Math.min(image.height, Math.ceil(rect.y + rect.height + marginY));
  return {
    x: x0,
    y: y0,
    width: Math.max(TOOL_MIN_CROP_SIZE, x1 - x0),
    height: Math.max(TOOL_MIN_CROP_SIZE, y1 - y0),
  };
}

function encodeGridRowLabel(index) {
  let value = Number(index);
  if (!Number.isFinite(value) || value < 0) return 'A';
  let label = '';
  do {
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return label;
}

function decodeGridRowLabel(label) {
  const raw = String(label || '').trim().toUpperCase();
  if (!raw || !/^[A-Z]+$/.test(raw)) return -1;
  let value = 0;
  for (const ch of raw) {
    value = value * 26 + (ch.charCodeAt(0) - 64);
  }
  return value - 1;
}

function chooseCoarseGridSize(screen) {
  const shortestSide = Math.max(1, Math.min(Number(screen.width) || 0, Number(screen.height) || 0));
  const byResolution = Math.round(shortestSide / 150);
  return clampInt(byResolution, COARSE_GRID_MIN, COARSE_GRID_MAX);
}

function chooseFineGridSize(cellRect) {
  const shortestSide = Math.max(1, Math.min(Number(cellRect.width) || 0, Number(cellRect.height) || 0));
  const byCellSize = Math.round(shortestSide / 55);
  return clampInt(byCellSize, FINE_GRID_MIN, FINE_GRID_MAX);
}

function chooseThirdGridSize(imageRect) {
  const shortestSide = Math.max(1, Math.min(Number(imageRect.width) || 0, Number(imageRect.height) || 0));
  const byResolution = Math.round(shortestSide / 180);
  return clampInt(byResolution, THIRD_GRID_MIN, THIRD_GRID_MAX);
}

function mapPointBetweenImages(point, sourceImage, targetImage) {
  return {
    x: clampInt(
      Math.round((((Number(point.x) + 0.5) * targetImage.width) / sourceImage.width) - 0.5),
      0,
      Math.max(0, targetImage.width - 1),
    ),
    y: clampInt(
      Math.round((((Number(point.y) + 0.5) * targetImage.height) / sourceImage.height) - 0.5),
      0,
      Math.max(0, targetImage.height - 1),
    ),
  };
}

function mapRectBetweenImages(rect, sourceImage, targetImage) {
  const x0 = clampInt(
    Math.floor((Number(rect.x) * targetImage.width) / sourceImage.width),
    0,
    Math.max(0, targetImage.width - 1),
  );
  const y0 = clampInt(
    Math.floor((Number(rect.y) * targetImage.height) / sourceImage.height),
    0,
    Math.max(0, targetImage.height - 1),
  );
  const x1 = clampInt(
    Math.ceil(((Number(rect.x) + Number(rect.width)) * targetImage.width) / sourceImage.width),
    x0 + 1,
    targetImage.width,
  );
  const y1 = clampInt(
    Math.ceil(((Number(rect.y) + Number(rect.height)) * targetImage.height) / sourceImage.height),
    y0 + 1,
    targetImage.height,
  );
  return buildRect(x0, y0, Math.max(1, x1 - x0), Math.max(1, y1 - y0));
}

function buildRect(x, y, width, height) {
  return {
    x: Math.max(0, Math.round(x)),
    y: Math.max(0, Math.round(y)),
    width: Math.max(1, Math.round(width)),
    height: Math.max(1, Math.round(height)),
  };
}

function gridBoundary(origin, size, parts, index) {
  return Math.round(origin + Math.floor((Number(size) * Number(index)) / Number(parts)));
}

function makeGridCells(rect, gridSize) {
  const palette = [
    { r: 239, g: 68, b: 68 },
    { r: 249, g: 115, b: 22 },
    { r: 245, g: 158, b: 11 },
    { r: 34, g: 197, b: 94 },
    { r: 6, g: 182, b: 212 },
    { r: 59, g: 130, b: 246 },
    { r: 168, g: 85, b: 247 },
    { r: 236, g: 72, b: 153 },
  ];
  const cells = [];
  const baseRect = buildRect(rect.x, rect.y, rect.width, rect.height);
  for (let row = 0; row < gridSize; row += 1) {
    for (let col = 0; col < gridSize; col += 1) {
      const x0 = gridBoundary(baseRect.x, baseRect.width, gridSize, col);
      const y0 = gridBoundary(baseRect.y, baseRect.height, gridSize, row);
      const x1 = gridBoundary(baseRect.x, baseRect.width, gridSize, col + 1);
      const y1 = gridBoundary(baseRect.y, baseRect.height, gridSize, row + 1);
      cells.push({
        row,
        col,
        label: `${encodeGridRowLabel(row)}${col + 1}`,
        rect: buildRect(x0, y0, Math.max(1, x1 - x0), Math.max(1, y1 - y0)),
        anchor: ['tl', 'tr', 'bl', 'br'][(row + col) % 4],
        color: palette[(row * gridSize + col) % palette.length],
      });
    }
  }
  return cells;
}

function findGridCell(cells, row, col) {
  return cells.find((cell) => cell.row === row && cell.col === col);
}

function parseGridCellRef(value, gridSize) {
  const raw = String(
    value?.cell ||
      value?.grid_cell ||
      value?.cell_id ||
      value?.selected_cell ||
      value?.subcell ||
      value?.sub_cell ||
      value?.label ||
      value ||
      '',
  )
    .trim()
    .toUpperCase();
  if (!raw) return null;
  const normalizedRaw = raw.includes('|') ? raw.split('|')[0].trim() : raw;
  if (!normalizedRaw || normalizedRaw === 'NONE') return null;
  const match = normalizedRaw.match(/^([A-Z]+)\s*[-:_ ]?\s*(\d+)$/);
  if (!match) return null;
  const row = decodeGridRowLabel(match[1]);
  const col = Number(match[2]) - 1;
  if (!Number.isInteger(row) || !Number.isInteger(col)) return null;
  if (row < 0 || row >= gridSize || col < 0 || col >= gridSize) return null;
  return { row, col, label: `${encodeGridRowLabel(row)}${col + 1}` };
}

const FINE_CELL_ANCHOR_ORDER = [
  'top_left',
  'top_center',
  'top_right',
  'middle_left',
  'center',
  'middle_right',
  'bottom_left',
  'bottom_center',
  'bottom_right',
];

function normalizeAnchorToken(value) {
  const raw = String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_');
  if (!raw) return 'center';
  const patterns = [
  ['top_left', /^(top_left|upper_left|left_top|northwest)$/u],
  ['top_center', /^(top_center|top_middle|upper_center|upper_middle|middle_top|center_top)$/u],
  ['top_right', /^(top_right|upper_right|right_top|northeast)$/u],
  ['middle_left', /^(middle_left|center_left|left_middle|left_center|west)$/u],
  ['center', /^(center|middle|mid|centre)$/u],
  ['middle_right', /^(middle_right|center_right|right_middle|right_center|east)$/u],
  ['bottom_left', /^(bottom_left|lower_left|left_bottom|southwest)$/u],
  ['bottom_center', /^(bottom_center|bottom_middle|lower_center|lower_middle|middle_bottom|center_bottom)$/u],
  ['bottom_right', /^(bottom_right|lower_right|right_bottom|southeast)$/u],
    ['middle_left', /^(left)$/u],
    ['middle_right', /^(right)$/u],
    ['top_center', /^(top|up)$/u],
    ['bottom_center', /^(bottom|down)$/u],
  ];
  for (const [name, pattern] of patterns) {
    if (pattern.test(raw)) return name;
  }
  return 'center';
}

function splitAnchorCandidates(value) {
  return String(value || '')
    .split(/[|,;/]+/u)
    .map((item) => item.trim())
    .filter(Boolean);
}

function sortAnchorList(anchors) {
  return [...anchors].sort((left, right) => {
    const leftIndex = FINE_CELL_ANCHOR_ORDER.indexOf(left);
    const rightIndex = FINE_CELL_ANCHOR_ORDER.indexOf(right);
    return leftIndex - rightIndex;
  });
}

function normalizeAnchors(value, options = {}) {
  const maxCount = Math.max(1, Number(options.maxCount || 2));
  const defaultToCenter = options.defaultToCenter !== false;
  const candidates = [];
  const pushCandidate = (candidate) => {
    if (Array.isArray(candidate)) {
      for (const item of candidate) pushCandidate(item);
      return;
    }
    if (candidate === undefined || candidate === null) return;
    const parts = splitAnchorCandidates(candidate);
    if (parts.length) {
      candidates.push(...parts);
      return;
    }
    candidates.push(candidate);
  };
  if (value && typeof value === 'object') {
    pushCandidate(value.positions);
    pushCandidate(value.position);
    pushCandidate(value.position_2);
    pushCandidate(value.position2);
    pushCandidate(value.secondary_position);
    pushCandidate(value.secondaryPosition);
    pushCandidate(value.anchors);
    pushCandidate(value.anchor);
    pushCandidate(value.anchor_2);
    pushCandidate(value.anchor2);
    pushCandidate(value.secondary_anchor);
    pushCandidate(value.secondaryAnchor);
    pushCandidate(value.click_region);
    pushCandidate(value.region);
    pushCandidate(value.part);
  } else {
    pushCandidate(value);
  }
  const deduped = [];
  for (const candidate of candidates) {
    const normalized = normalizeAnchorToken(candidate);
    if (!deduped.includes(normalized)) deduped.push(normalized);
  }
  const normalizedAnchors = sortAnchorList(deduped).slice(0, maxCount);
  if (normalizedAnchors.length) return normalizedAnchors;
  return defaultToCenter ? ['center'] : [];
}

function normalizeAnchor(value) {
  return normalizeAnchors(value, { maxCount: 1, defaultToCenter: true })[0] || 'center';
}

function normalizeOutsideDirection(value) {
  const raw = String(
    value?.outside_direction ||
      value?.neighbor_direction ||
      value?.direction ||
      value ||
      '',
  )
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_');
  if (!raw || raw === 'none') return 'none';
  const patterns = [
  ['left', /^(left|west)$/u],
  ['right', /^(right|east)$/u],
  ['up', /^(up|top|north)$/u],
  ['down', /^(down|bottom|south)$/u],
  ['up_left', /^(up_left|top_left|upper_left|northwest)$/u],
  ['up_right', /^(up_right|top_right|upper_right|northeast)$/u],
  ['down_left', /^(down_left|bottom_left|lower_left|southwest)$/u],
  ['down_right', /^(down_right|bottom_right|lower_right|southeast)$/u],
  ];
  for (const [name, pattern] of patterns) {
    if (pattern.test(raw)) return name;
  }
  return 'none';
}

function shiftCellRef(cellRef, direction, gridSize) {
  const deltas = {
    left: { row: 0, col: -1 },
    right: { row: 0, col: 1 },
    up: { row: -1, col: 0 },
    down: { row: 1, col: 0 },
    up_left: { row: -1, col: -1 },
    up_right: { row: -1, col: 1 },
    down_left: { row: 1, col: -1 },
    down_right: { row: 1, col: 1 },
  };
  const delta = deltas[normalizeOutsideDirection(direction)];
  if (!delta) return null;
  const row = cellRef.row + delta.row;
  const col = cellRef.col + delta.col;
  if (row < 0 || row >= gridSize || col < 0 || col >= gridSize) return null;
  return {
    row,
    col,
    label: `${encodeGridRowLabel(row)}${col + 1}`,
  };
}

function anchorToPoint(rect, anchor) {
  const ratios = {
    top_left: { x: 0.06, y: 0.06 },
    top_center: { x: 0.56, y: 0.12 },
    top_right: { x: 0.94, y: 0.06 },
    middle_left: { x: 0.06, y: 0.5 },
    center: { x: 0.56, y: 0.5 },
    middle_right: { x: 0.94, y: 0.5 },
    bottom_left: { x: 0.06, y: 0.94 },
    bottom_center: { x: 0.56, y: 0.94 },
    bottom_right: { x: 0.94, y: 0.94 },
  };
  const ratio = ratios[normalizeAnchorToken(anchor)] || ratios.center;
  return {
    x: clampInt(rect.x + (rect.width - 1) * ratio.x, rect.x, rect.x + rect.width - 1),
    y: clampInt(rect.y + (rect.height - 1) * ratio.y, rect.y, rect.y + rect.height - 1),
  };
}

function bboxAroundPoint(point, containerRect) {
  const width = clampInt(Math.round(containerRect.width * 0.42), 10, containerRect.width);
  const height = clampInt(Math.round(containerRect.height * 0.42), 10, containerRect.height);
  const minX = containerRect.x;
  const minY = containerRect.y;
  const maxX = containerRect.x + containerRect.width - width;
  const maxY = containerRect.y + containerRect.height - height;
  return {
    x: clampInt(Math.round(point.x - width / 2), minX, maxX),
    y: clampInt(Math.round(point.y - height / 2), minY, maxY),
    width,
    height,
  };
}

function buildNeighborhoodRect(imageRect, gridSize, row, col, radius = 1) {
  const rowStart = Math.max(0, row - radius);
  const rowEnd = Math.min(gridSize - 1, row + radius);
  const colStart = Math.max(0, col - radius);
  const colEnd = Math.min(gridSize - 1, col + radius);
  const x0 = gridBoundary(imageRect.x, imageRect.width, gridSize, colStart);
  const y0 = gridBoundary(imageRect.y, imageRect.height, gridSize, rowStart);
  const x1 = gridBoundary(imageRect.x, imageRect.width, gridSize, colEnd + 1);
  const y1 = gridBoundary(imageRect.y, imageRect.height, gridSize, rowEnd + 1);
  return buildRect(x0, y0, Math.max(1, x1 - x0), Math.max(1, y1 - y0));
}

function rectRelativeTo(rect, origin) {
  return buildRect(rect.x - origin.x, rect.y - origin.y, rect.width, rect.height);
}

function psColor({ a = 255, r, g, b }) {
  return `[System.Drawing.Color]::FromArgb(${a}, ${r}, ${g}, ${b})`;
}

function buildGridLabelDrawCommands(cells, options = {}) {
  const fontScale = Number.isFinite(Number(options.fontScale)) ? Number(options.fontScale) : 1;
  const commands = [];
  commands.push(`$gridBackPen = New-Object System.Drawing.Pen (${psColor({ a: 120, r: 0, g: 0, b: 0 })}), 3`);
  commands.push(`$gridFrontPen = New-Object System.Drawing.Pen (${psColor({ a: 210, r: 255, g: 255, b: 255 })}), 1`);
  commands.push(`
function Get-RegionStats([System.Drawing.Bitmap]$bmp, [int]$x, [int]$y, [int]$w, [int]$h) {
  if ($w -le 2 -or $h -le 2) { return [double]::PositiveInfinity }
  $minX = [Math]::Max(0, $x)
  $minY = [Math]::Max(0, $y)
  $maxX = [Math]::Min($bmp.Width - 2, $x + $w - 1)
  $maxY = [Math]::Min($bmp.Height - 2, $y + $h - 1)
  if ($maxX -lt $minX -or $maxY -lt $minY) {
    return [pscustomobject]@{ Edge = [double]::PositiveInfinity; R = 127; G = 127; B = 127 }
  }
  $stepX = [Math]::Max(3, [Math]::Floor($w / 6))
  $stepY = [Math]::Max(3, [Math]::Floor($h / 4))
  $sum = 0.0
  $sumR = 0.0
  $sumG = 0.0
  $sumB = 0.0
  $samples = 0
  for ($py = $minY; $py -le $maxY; $py += $stepY) {
    for ($px = $minX; $px -le $maxX; $px += $stepX) {
      $c = $bmp.GetPixel($px, $py)
      $r = $bmp.GetPixel([Math]::Min($bmp.Width - 1, $px + 1), $py)
      $d = $bmp.GetPixel($px, [Math]::Min($bmp.Height - 1, $py + 1))
      $edge =
        [Math]::Abs([int]$c.R - [int]$r.R) +
        [Math]::Abs([int]$c.G - [int]$r.G) +
        [Math]::Abs([int]$c.B - [int]$r.B) +
        [Math]::Abs([int]$c.R - [int]$d.R) +
        [Math]::Abs([int]$c.G - [int]$d.G) +
        [Math]::Abs([int]$c.B - [int]$d.B)
      $sum += $edge / 6.0
      $sumR += [double]$c.R
      $sumG += [double]$c.G
      $sumB += [double]$c.B
      $samples += 1
    }
  }
  if ($samples -le 0) {
    return [pscustomobject]@{ Edge = [double]::PositiveInfinity; R = 127; G = 127; B = 127 }
  }
  return [pscustomobject]@{
    Edge = $sum / $samples
    R = $sumR / $samples
    G = $sumG / $samples
    B = $sumB / $samples
  }
}
function Get-ContrastingTextColor([double]$r, [double]$g, [double]$b) {
  $tr = [int][Math]::Round(255 - $r)
  $tg = [int][Math]::Round(255 - $g)
  $tb = [int][Math]::Round(255 - $b)
  $spread = [Math]::Abs($tr - $tg) + [Math]::Abs($tr - $tb) + [Math]::Abs($tg - $tb)
  $bgLuma = 0.299 * $r + 0.587 * $g + 0.114 * $b
  if ($spread -lt 42) {
    if ($bgLuma -ge 140) {
      $tr = 24; $tg = 24; $tb = 24
    } else {
      $tr = 245; $tg = 245; $tb = 245
    }
  } else {
    $tr = [int][Math]::Max(0, [Math]::Min(255, [Math]::Round((($tr - 128) * 1.22) + 128)))
    $tg = [int][Math]::Max(0, [Math]::Min(255, [Math]::Round((($tg - 128) * 1.22) + 128)))
    $tb = [int][Math]::Max(0, [Math]::Min(255, [Math]::Round((($tb - 128) * 1.22) + 128)))
  }
  return [pscustomobject]@{ R = $tr; G = $tg; B = $tb }
}
function Get-ShadowColor([int]$r, [int]$g, [int]$b) {
  $textLuma = 0.299 * $r + 0.587 * $g + 0.114 * $b
  if ($textLuma -ge 145) {
    return [pscustomobject]@{ A = 210; R = 0; G = 0; B = 0 }
  }
  return [pscustomobject]@{ A = 210; R = 255; G = 255; B = 255 }
}
function Get-QuietTileMap([System.Drawing.Bitmap]$bmp, [int]$baseX, [int]$baseY, [int]$availW, [int]$availH, [int]$tileSize, [double]$quietThreshold) {
  $cols = [Math]::Max(1, [int][Math]::Ceiling($availW / [double]$tileSize))
  $rows = [Math]::Max(1, [int][Math]::Ceiling($availH / [double]$tileSize))
  $quiet = New-Object 'int[]' ($rows * $cols)
  for ($row = 0; $row -lt $rows; $row += 1) {
    for ($col = 0; $col -lt $cols; $col += 1) {
      $tileX = $baseX + ($col * $tileSize)
      $tileY = $baseY + ($row * $tileSize)
      $tileW = [Math]::Min($tileSize, ($baseX + $availW) - $tileX)
      $tileH = [Math]::Min($tileSize, ($baseY + $availH) - $tileY)
      $stats = Get-RegionStats $bmp $tileX $tileY $tileW $tileH
      $idx = ($row * $cols) + $col
      $quiet[$idx] = if ($stats.Edge -le $quietThreshold) { 1 } else { 0 }
    }
  }
  return [pscustomobject]@{
    Rows = $rows
    Cols = $cols
    TileSize = $tileSize
    BaseX = $baseX
    BaseY = $baseY
    AvailW = $availW
    AvailH = $availH
    Quiet = $quiet
  }
}
function Get-LabelPlacement([System.Drawing.Bitmap]$bmp, [int]$cellX, [int]$cellY, [int]$cellW, [int]$cellH, [int]$labelW, [int]$labelH, [int]$padding) {
  $baseX = $cellX + $padding
  $baseY = $cellY + $padding
  $availW = [Math]::Max($labelW, $cellW - ($padding * 2))
  $availH = [Math]::Max($labelH, $cellH - ($padding * 2))
  $tileSize = [Math]::Max(4, [Math]::Min(18, [int][Math]::Floor([Math]::Min($labelW, $labelH) / 2.0)))
  $tileMap = Get-QuietTileMap $bmp $baseX $baseY $availW $availH $tileSize 24
  $heights = New-Object 'int[]' $tileMap.Cols
  $best = $null
  for ($row = 0; $row -lt $tileMap.Rows; $row += 1) {
    for ($col = 0; $col -lt $tileMap.Cols; $col += 1) {
      $idx = ($row * $tileMap.Cols) + $col
      if ($tileMap.Quiet[$idx] -eq 1) {
        $heights[$col] += 1
      } else {
        $heights[$col] = 0
      }
    }
    for ($left = 0; $left -lt $tileMap.Cols; $left += 1) {
      $minHeightTiles = [int]::MaxValue
      for ($right = $left; $right -lt $tileMap.Cols; $right += 1) {
        $minHeightTiles = [Math]::Min($minHeightTiles, $heights[$right])
        if ($minHeightTiles -le 0) { continue }
        $top = $row - $minHeightTiles + 1
        $rectX = $tileMap.BaseX + ($left * $tileMap.TileSize)
        $rectY = $tileMap.BaseY + ($top * $tileMap.TileSize)
        $rectW = [Math]::Min(($right - $left + 1) * $tileMap.TileSize, ($tileMap.BaseX + $tileMap.AvailW) - $rectX)
        $rectH = [Math]::Min($minHeightTiles * $tileMap.TileSize, ($tileMap.BaseY + $tileMap.AvailH) - $rectY)
        if ($rectW -lt $labelW -or $rectH -lt $labelH) { continue }
        $textX = $rectX + [Math]::Floor(($rectW - $labelW) / 2.0)
        $textY = $rectY + [Math]::Floor(($rectH - $labelH) / 2.0)
        $stats = Get-RegionStats $bmp $textX $textY $labelW $labelH
        $area = $rectW * $rectH
        $centerPenalty =
          [Math]::Abs(($rectX + ($rectW / 2.0)) - ($cellX + ($cellW / 2.0))) +
          [Math]::Abs(($rectY + ($rectH / 2.0)) - ($cellY + ($cellH / 2.0)))
        if (
          $null -eq $best -or
          $area -gt $best.Area -or
          ($area -eq $best.Area -and $stats.Edge -lt $best.Edge) -or
          ($area -eq $best.Area -and $stats.Edge -eq $best.Edge -and $centerPenalty -lt $best.CenterPenalty)
        ) {
          $best = [pscustomobject]@{
            X = $textX
            Y = $textY
            RectX = $rectX
            RectY = $rectY
            RectW = $rectW
            RectH = $rectH
            Area = $area
            Edge = $stats.Edge
            CenterPenalty = $centerPenalty
            R = $stats.R
            G = $stats.G
            B = $stats.B
          }
        }
      }
    }
  }
  if ($null -eq $best) {
    $fallbackX = $cellX + [Math]::Max($padding, [Math]::Floor(($cellW - $labelW) / 2.0))
    $fallbackY = $cellY + [Math]::Max($padding, [Math]::Floor(($cellH - $labelH) / 2.0))
    $stats = Get-RegionStats $bmp $fallbackX $fallbackY $labelW $labelH
    $best = [pscustomobject]@{
      X = $fallbackX
      Y = $fallbackY
      RectX = $fallbackX
      RectY = $fallbackY
      RectW = $labelW
      RectH = $labelH
      Area = $labelW * $labelH
      Edge = $stats.Edge
      CenterPenalty = 0
      R = $stats.R
      G = $stats.G
      B = $stats.B
    }
  }
  return $best
}
`);
  for (const cell of cells) {
    const fontSize = clampInt(
      Math.round(Math.min(cell.rect.width, cell.rect.height) * 0.18 * fontScale),
      9,
      24,
    );
    const labelWidth = Math.max(20, Math.round(fontSize * 0.66 * cell.label.length + 6));
    const labelHeight = Math.max(16, Math.round(fontSize + 4));
    const padding = 4;
    commands.push(`$g.DrawRectangle($gridBackPen, ${cell.rect.x}, ${cell.rect.y}, ${Math.max(1, cell.rect.width - 1)}, ${Math.max(1, cell.rect.height - 1)})`);
    commands.push(`$g.DrawRectangle($gridFrontPen, ${cell.rect.x}, ${cell.rect.y}, ${Math.max(1, cell.rect.width - 1)}, ${Math.max(1, cell.rect.height - 1)})`);
    commands.push(`$font = New-Object System.Drawing.Font('Consolas', ${fontSize}, [System.Drawing.FontStyle]::Bold)`);
    commands.push(`$labelPlacement = Get-LabelPlacement $bmp ${cell.rect.x} ${cell.rect.y} ${cell.rect.width} ${cell.rect.height} ${labelWidth} ${labelHeight} ${padding}`);
    commands.push(`$textColor = Get-ContrastingTextColor $labelPlacement.R $labelPlacement.G $labelPlacement.B`);
    commands.push(`$shadowColor = Get-ShadowColor $textColor.R $textColor.G $textColor.B`);
    commands.push(`$labelShadowBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($shadowColor.A, $shadowColor.R, $shadowColor.G, $shadowColor.B))`);
    commands.push(`$labelTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(250, $textColor.R, $textColor.G, $textColor.B))`);
    commands.push(`$textX = [int]$labelPlacement.X`);
    commands.push(`$textY = [int]$labelPlacement.Y`);
    commands.push(`$g.DrawString('${escapePsSingleQuoted(cell.label)}', $font, $labelShadowBrush, $textX + 1, $textY + 1)`);
    commands.push(`$g.DrawString('${escapePsSingleQuoted(cell.label)}', $font, $labelShadowBrush, $textX - 1, $textY + 1)`);
    commands.push(`$g.DrawString('${escapePsSingleQuoted(cell.label)}', $font, $labelTextBrush, $textX, $textY)`);
    commands.push(`$labelShadowBrush.Dispose()`);
    commands.push(`$labelTextBrush.Dispose()`);
    commands.push(`$font.Dispose()`);
  }
  commands.push(`$gridBackPen.Dispose()`);
  commands.push(`$gridFrontPen.Dispose()`);
  return commands.join('\n');
}

async function renderAnnotatedImage({ sourcePath, outPath, drawingCommands, label }) {
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$src = '${escapePsSingleQuoted(path.resolve(sourcePath))}'
$out = '${escapePsSingleQuoted(path.resolve(outPath))}'
$bmp = [System.Drawing.Bitmap]::FromFile($src)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
${drawingCommands}
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
`;
  await runPowerShell(script, 20000, label);
  return path.resolve(outPath);
}

async function annotateImageWithGrid({ image, gridRect, gridSize, outPath, label }) {
  const cells = makeGridCells(gridRect, gridSize);
  const annotatedPath = await renderAnnotatedImage({
    sourcePath: image.path,
    outPath,
    drawingCommands: buildGridLabelDrawCommands(cells),
    label,
  });
  return {
    path: annotatedPath,
    cells,
  };
}

async function annotateNeighborhoodFineGrid({
  image,
  selectedRect,
  fineGridSize,
  outPath,
  label,
}) {
  const fineCells = makeGridCells(selectedRect, fineGridSize);
  const commands = [
    `$selectedPenBack = New-Object System.Drawing.Pen (${psColor({ a: 180, r: 0, g: 0, b: 0 })}), 6`,
    `$selectedPenFront = New-Object System.Drawing.Pen (${psColor({ a: 240, r: 0, g: 255, b: 224 })}), 3`,
    `$g.DrawRectangle($selectedPenBack, ${selectedRect.x}, ${selectedRect.y}, ${Math.max(1, selectedRect.width - 1)}, ${Math.max(1, selectedRect.height - 1)})`,
    `$g.DrawRectangle($selectedPenFront, ${selectedRect.x}, ${selectedRect.y}, ${Math.max(1, selectedRect.width - 1)}, ${Math.max(1, selectedRect.height - 1)})`,
    buildGridLabelDrawCommands(fineCells, { fontScale: 0.88 }),
    `$selectedPenBack.Dispose()`,
    `$selectedPenFront.Dispose()`,
  ].join('\n');
  const annotatedPath = await renderAnnotatedImage({
    sourcePath: image.path,
    outPath,
    drawingCommands: commands,
    label,
  });
  return {
    path: annotatedPath,
    fineCells,
  };
}

async function annotateImageWithPoint({ image, point, bbox, outPath, label }) {
  const box = buildRect(
    bbox?.x ?? Math.round(point.x - 12),
    bbox?.y ?? Math.round(point.y - 8),
    bbox?.width ?? 24,
    bbox?.height ?? 16,
  );
  const commands = [
    `$reviewBoxPen = New-Object System.Drawing.Pen (${psColor({ a: 245, r: 255, g: 48, b: 48 })}), 2`,
    `$reviewBoxPen.Alignment = [System.Drawing.Drawing2D.PenAlignment]::Inset`,
    `$g.DrawRectangle($reviewBoxPen, ${box.x}, ${box.y}, ${Math.max(1, box.width - 1)}, ${Math.max(1, box.height - 1)})`,
    `$reviewBoxPen.Dispose()`,
  ].join('\n');
  const annotatedPath = await renderAnnotatedImage({
    sourcePath: image.path,
    outPath,
    drawingCommands: commands,
    label,
  });
  return path.resolve(annotatedPath);
}

function buildCoarseGridPrompt({ description, screen, gridSize, retryFeedbacks }) {
  const feedbackText =
    Array.isArray(retryFeedbacks) && retryFeedbacks.length
      ? `\nPrevious rejected candidates on the same screenshot:\n${retryFeedbacks
          .map((item, index) => `${index + 1}. ${item}`)
          .join('\n')}\nAvoid repeating those mistakes.`
      : '';
  return [
    'You are doing stage 1 of desktop target grounding.',
    `Task description: ${description}`,
    `The desktop screenshot resolution is ${screen.width}x${screen.height}.`,
    `You receive exactly one image: the full screenshot with a ${gridSize}x${gridSize} coarse grid already drawn on it.`,
    'Rows use spreadsheet-style letters (A, B, C...) and columns use 1-based numbers.',
    'Label tags may appear in different corners of cells to reduce occlusion.',
    'Return strict JSON only with this exact shape:',
    '{"cell":"A1|none","confidence":0.0,"reason":"short text","visible_anchor":"visible label or icon cue"}',
    'Rules:',
    '- Choose the single coarse cell that contains the click position most likely to succeed for the described target.',
    '- If the target spans multiple cells, choose the cell that contains the best click point, not just the largest visible fragment.',
    '- If the description names a specific app, brand, button label, or UI control, you must confirm that exact identity and reject lookalikes.',
    '- Do not choose a target based only on color, approximate position, generic shape, or the fact that it is a plausible nearby control.',
    '- Use surrounding context visible in this same labeled screenshot to disambiguate similar icons or repeated controls.',
    '- Return "none" only if the target is missing or still ambiguous after inspecting this labeled screenshot.',
    '- Do not output markdown or extra text.',
    feedbackText,
  ].join('\n');
}

function buildFineGridPrompt({ description, fineGridSize }) {
  return [
    'You are doing stage 2 of desktop target grounding.',
    `Task description: ${description}`,
    `You receive exactly one image: a crop that contains the selected coarse cell plus neighboring coarse cells, and the selected coarse cell is subdivided into a ${fineGridSize}x${fineGridSize} fine grid with a bright cyan border.`,
    'The returned cell coordinates must come from this single image only.',
    'The fine-grid labels reset inside the selected coarse cell: A1.. up to the grid size.',
    'Label tags may appear in different corners of cells to reduce occlusion.',
    'Return strict JSON only with this exact shape:',
    '{"cell":"A1|none","outside_direction":"none|left|right|up|down|up_left|up_right|down_left|down_right","confidence":0.0,"reason":"short text","visible_anchor":"visible cue"}',
    'Rules:',
    '- Choose the single fine cell inside the selected coarse cell that contains the click coordinate most likely to succeed.',
    '- If the description names a specific app, brand, button label, or UI control, you must confirm that exact identity and reject lookalikes.',
    '- Do not choose a target based only on color, approximate position, generic shape, or the fact that it is a plausible nearby control.',
    '- Do not return click anchors or positions at this stage; stage 3 will refine the final click point.',
    '- If the target is actually visible in one of the neighboring coarse cells instead of the selected coarse cell, return cell "none" and use outside_direction to point to that neighbor.',
    '- If the target is still ambiguous, return cell "none" and outside_direction "none".',
    '- Do not output markdown or extra text.',
  ].join('\n');
}

function buildThirdGridPrompt({ description, gridSize }) {
  return [
    'You are doing stage 3 of desktop target grounding.',
    `Task description: ${description}`,
    'You receive exactly two images in this order:',
    '1. The original stage-2 neighborhood crop without any fine-grid overlay.',
    `2. A local crop built from the selected stage-2 fine cell and its surrounding fine cells, with the entire local crop subdivided into a ${gridSize}x${gridSize} final grid.`,
    'The returned cell coordinates must come from image 2 only.',
    'Label tags may appear in different corners of cells to reduce occlusion.',
    'Return strict JSON only with this exact shape:',
    '{"cell":"A1|none","positions":["top_left|top_center|top_right|middle_left|center|middle_right|bottom_left|bottom_center|bottom_right"],"confidence":0.0,"reason":"short text","visible_anchor":"visible cue"}',
    'Rules:',
    '- Choose the single stage-3 grid cell in image 2 that contains the click coordinate most likely to succeed.',
    '- "positions" is used only to place the click point inside the chosen stage-3 cell after the cell is selected.',
    '- Return one anchor in "positions" when one click anchor is clearly best. Return exactly two anchors when two interior click anchors are equally safe and appropriate.',
    '- If two anchors are returned, both anchors must be inside the same chosen stage-3 cell, and the final click point will be placed at the midpoint between those two anchor coordinates.',
    '- Use these exact meanings for each value in "positions" inside the chosen cell: top_left=upper-left interior click anchor, top_center=upper-middle interior click anchor, top_right=upper-right interior click anchor, middle_left=left-middle interior click anchor, center=cell center click anchor, middle_right=right-middle interior click anchor, bottom_left=lower-left interior click anchor, bottom_center=lower-middle interior click anchor, bottom_right=lower-right interior click anchor.',
    '- These position labels describe where to click inside the chosen stage-3 cell; they do not refer to neighboring cells.',
    '- Use image 1 for semantic disambiguation and image 2 for final coordinates.',
    '- If the description names a specific app, brand, button label, or UI control, you must confirm that exact identity and reject lookalikes.',
    '- Do not choose a target based only on color, approximate position, generic shape, or the fact that it is a plausible nearby control.',
    '- Prefer stable interior body pixels of the target, not thin edges, tiny badges, or decorative corners.',
    '- If cell is not "none", "positions" must contain one or two values. If cell is "none", return an empty array for "positions".',
    '- Return "none" only if the target is still ambiguous in this local refinement view.',
    '- Do not output markdown or extra text.',
  ].join('\n');
}

function buildReviewPrompt({ description, candidatePoint, candidateBBox }) {
  return [
    'You are doing the final verification of a desktop click point.',
    `Task description: ${description}`,
    `Candidate root coordinate: x=${candidatePoint.x}, y=${candidatePoint.y}.`,
    `Candidate helper bbox on the root screenshot: x=${candidateBBox.x}, y=${candidateBBox.y}, width=${candidateBBox.width}, height=${candidateBBox.height}.`,
    'You receive exactly one image: the full screenshot with only the candidate helper bbox outlined by a thin red rectangle.',
    'No dimming mask, shaded overlay, or extra review crop is provided.',
    'Return strict JSON only with this exact shape:',
    '{"status":"pass|retry","confidence":0.0,"reason":"short text","visible_anchor":"visible cue"}',
    'Rules:',
    '- Review is not whether the red rectangle encloses a plausible clickable item; review is whether it encloses the exact target named in the task description.',
    '- Return "pass" only if the thin red rectangle clearly encloses that exact intended target and indicates the click region most likely to succeed.',
    '- Return "retry" if the red rectangle encloses the wrong target, a nearby lookalike, a merely plausible control, is too close to a neighbor, misses the target body, or is still ambiguous.',
    '- If the description names a specific app, brand, button label, or UI control, confirm that exact identity before returning "pass"; do not rely only on color, rough position, or generic shape.',
    '- Use this single marked screenshot with the thin red rectangle for both global disambiguation and point precision.',
    '- Do not output markdown or extra text.',
  ].join('\n');
}

function pickRepresentativeBallot(group) {
  return [...group.items].sort((a, b) => Number(b.json?.confidence || 0) - Number(a.json?.confidence || 0))[0];
}

function voteCoarseGridBallots(ballots, gridSize) {
  const group = summarizeMajorityCandidate(
    ballots,
    (item) => parseGridCellRef(item.json, gridSize)?.label || 'none',
  );
  const representative = pickRepresentativeBallot(group);
  return {
    summary: `${group.key} (${group.count}/${ballots.length})`,
    count: group.count,
    total: ballots.length,
    representative,
  };
}

function voteFineGridBallots(ballots, fineGridSize) {
  const parsedBallots = ballots.map((item) => {
    const cellRef = parseGridCellRef(item.json, fineGridSize);
    if (cellRef) {
      return {
        ...item,
        kind: 'cell',
        cellRef,
        outsideDirection: 'none',
      };
    }
    return {
      ...item,
      kind: 'outside',
      cellRef: null,
      outsideDirection: normalizeOutsideDirection(item.json),
    };
  });
  const group = summarizeMajorityCandidate(parsedBallots, (item) => {
    if (item.kind === 'cell') return item.cellRef.label;
    return `none|${item.outsideDirection}`;
  });
  const representative = pickRepresentativeBallot(group);
  return {
    summary: `${group.key} (${group.count}/${ballots.length}, cell_only)`,
    count: group.count,
    total: ballots.length,
    representative,
    voteMode: 'cell_only',
    fineCellRef: representative.cellRef,
    outsideDirection: representative.outsideDirection,
  };
}

function voteThirdGridBallots(ballots, gridSize) {
  const parsedBallots = ballots.map((item) => {
    const cellRef = parseGridCellRef(item.json, gridSize);
    if (cellRef) {
      return {
        ...item,
        kind: 'cell',
        cellRef,
        anchors: normalizeAnchors(item.json, { maxCount: 2, defaultToCenter: true }),
      };
    }
    return {
      ...item,
      kind: 'none',
      cellRef: null,
      anchors: [],
    };
  });
  const exactGroup = summarizeMajorityCandidate(parsedBallots, (item) => {
    if (item.kind === 'cell') return `${item.cellRef.label}|${item.anchors.join('+')}`;
    return 'none';
  });
  const exactThreshold = Math.ceil(ballots.length / 2);
  if (exactGroup && exactGroup.count >= exactThreshold) {
    const representative = pickRepresentativeBallot(exactGroup);
    return {
      summary: `${exactGroup.key} (${exactGroup.count}/${ballots.length}, paired)`,
      count: exactGroup.count,
      total: ballots.length,
      representative,
      voteMode: 'paired',
      cellRef: representative.cellRef,
      anchors: representative.anchors,
    };
  }
  const expandedVotes = [];
  for (const item of parsedBallots) {
    if (item.kind === 'cell') {
      for (const anchor of item.anchors) {
        expandedVotes.push({
          ...item,
          anchors: [anchor],
        });
      }
      continue;
    }
    expandedVotes.push(item);
  }
  const group = summarizeMajorityCandidate(expandedVotes, (item) => {
    if (item.kind === 'cell') return `${item.cellRef.label}|${item.anchors[0]}`;
    return 'none';
  });
  const representative = pickRepresentativeBallot(group);
  return {
    summary: `${group.key} (${group.count}/${expandedVotes.length}, expanded)`,
    count: group.count,
    total: expandedVotes.length,
    representative,
    voteMode: 'expanded',
    cellRef: representative.cellRef,
    anchors: representative.anchors,
  };
}

function voteReviewBallots(ballots) {
  const group = summarizeMajorityCandidate(
    ballots,
    (item) => String(item.json?.status || '').trim().toLowerCase() || 'retry',
  );
  const representative = pickRepresentativeBallot(group);
  return {
    summary: `${group.key} (${group.count}/${ballots.length})`,
    count: group.count,
    total: ballots.length,
    representative,
  };
}

async function selectCoarseCell({
  apiKey,
  agent,
  model,
  reasoningEffort,
  description,
  screen,
  rootImage,
  baseDir,
  runTag,
  attempt,
  retryFeedbacks,
}) {
  const gridSize = chooseCoarseGridSize(screen);
  const annotatedPath = path.join(baseDir, `${runTag}-coarse-grid-${String(attempt).padStart(2, '0')}.png`);
  const { path: gridImagePath, cells } = await annotateImageWithGrid({
    image: rootImage,
    gridRect: buildRect(0, 0, rootImage.width, rootImage.height),
    gridSize,
    outPath: annotatedPath,
    label: `annotate coarse grid attempt ${attempt}`,
  });
  const { dataUrl: gridDataUrl } = await toModelDataUrl({
    image: {
      path: gridImagePath,
      width: rootImage.width,
      height: rootImage.height,
    },
    outPath: path.join(baseDir, `${runTag}-coarse-grid-model-${String(attempt).padStart(2, '0')}.png`),
    label: `resize coarse grid model image attempt ${attempt}`,
  });
  const voted = await requestModelJsonMajority({
    apiKey,
    agent,
    model,
    reasoningEffort,
    messages: [
      {
        role: 'system',
        content:
          'You are a desktop UI grounding assistant. Choose one coarse grid cell only and return strict JSON.',
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: buildCoarseGridPrompt({
              description,
              screen,
              gridSize,
              retryFeedbacks,
            }),
          },
          {
            type: 'image_url',
            image_url: { url: gridDataUrl },
          },
        ],
      },
    ],
    debugLabel: `coarse grid attempt ${attempt}`,
    voteFn: (ballots) => voteCoarseGridBallots(ballots, gridSize),
    stageName: 'coarse',
  });
  const { representative } = voted;
  const assistantText = representative.assistantText;
  const json = representative.json;
  const rawResponseText = representative.rawResponseText;
  const cellRef = parseGridCellRef(json, gridSize);
  return {
    gridSize,
    gridImagePath,
    cells,
    cellRef,
    assistantText,
    rawResponseText,
    majoritySummary: voted.summary,
    majorityCount: voted.count,
    majorityTotal: voted.total,
    confidence: Number.isFinite(Number(json?.confidence)) ? Number(json.confidence) : 0,
    reason: String(json?.reason || '').trim(),
    visibleAnchor: String(json?.visible_anchor || '').trim(),
  };
}

async function selectFineCell({
  apiKey,
  agent,
  model,
  reasoningEffort,
  description,
  rootImage,
  coarseSelection,
  baseDir,
  runTag,
  attempt,
}) {
  const coarseCell = findGridCell(
    coarseSelection.cells,
    coarseSelection.cellRef.row,
    coarseSelection.cellRef.col,
  );
  if (!coarseCell) fail(`Unable to resolve coarse cell geometry for ${coarseSelection.cellRef.label}`);
  const neighborhoodRect = buildNeighborhoodRect(
    buildRect(0, 0, rootImage.width, rootImage.height),
    coarseSelection.gridSize,
    coarseSelection.cellRef.row,
    coarseSelection.cellRef.col,
  );
  const neighborhoodOriginalPath = path.join(
    baseDir,
    `${runTag}-fine-neighborhood-${String(attempt).padStart(2, '0')}.png`,
  );
  const neighborhoodImage = await cropImage({
    image: rootImage,
    x: neighborhoodRect.x,
    y: neighborhoodRect.y,
    width: neighborhoodRect.width,
    height: neighborhoodRect.height,
    outPath: neighborhoodOriginalPath,
  });
  const selectedRectLocal = rectRelativeTo(coarseCell.rect, neighborhoodRect);
  const fineGridSize = chooseFineGridSize(coarseCell.rect);
  const annotatedPath = path.join(
    baseDir,
    `${runTag}-fine-grid-${String(attempt).padStart(2, '0')}.png`,
  );
  const { path: annotatedNeighborhoodPath, fineCells } = await annotateNeighborhoodFineGrid({
    image: neighborhoodImage,
    selectedRect: selectedRectLocal,
    fineGridSize,
    outPath: annotatedPath,
    label: `annotate fine grid attempt ${attempt}`,
  });
  const { dataUrl: neighborhoodDataUrl } = await toModelDataUrl({
    image: {
      path: annotatedNeighborhoodPath,
      width: neighborhoodImage.width,
      height: neighborhoodImage.height,
    },
    outPath: path.join(baseDir, `${runTag}-fine-grid-model-${String(attempt).padStart(2, '0')}.png`),
    label: `resize fine grid model image attempt ${attempt}`,
  });
  const voted = await requestModelJsonMajority({
    apiKey,
    agent,
    model,
    reasoningEffort,
    messages: [
      {
        role: 'system',
        content:
          'You are a desktop UI grounding assistant. Choose one stage-2 fine cell and return strict JSON.',
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: buildFineGridPrompt({ description, fineGridSize }),
          },
          {
            type: 'image_url',
            image_url: { url: neighborhoodDataUrl },
          },
        ],
      },
    ],
    debugLabel: `fine grid attempt ${attempt}`,
    voteFn: (ballots) => voteFineGridBallots(ballots, fineGridSize),
    stageName: 'fine',
  });
  const { representative } = voted;
  const assistantText = representative.assistantText;
  const json = representative.json;
  const rawResponseText = representative.rawResponseText;
  const fineCellRef = parseGridCellRef(json, fineGridSize);
  return {
    coarseCell,
    neighborhoodRect,
    neighborhoodImage,
    neighborhoodOriginalPath,
    gridImagePath: annotatedNeighborhoodPath,
    selectedRectLocal,
    fineGridSize,
    fineCells,
    fineCellRef: voted.fineCellRef || fineCellRef,
    outsideDirection: voted.outsideDirection || normalizeOutsideDirection(json),
    voteMode: voted.voteMode || 'paired',
    assistantText,
    rawResponseText,
    majoritySummary: voted.summary,
    majorityCount: voted.count,
    majorityTotal: voted.total,
    confidence: Number.isFinite(Number(json?.confidence)) ? Number(json.confidence) : 0,
    reason: String(json?.reason || '').trim(),
    visibleAnchor: String(json?.visible_anchor || '').trim(),
  };
}

async function selectThirdStageCell({
  apiKey,
  agent,
  model,
  reasoningEffort,
  description,
  fineSelection,
  baseDir,
  runTag,
  attempt,
}) {
  const fineCellLocal = findGridCell(
    fineSelection.fineCells,
    fineSelection.fineCellRef.row,
    fineSelection.fineCellRef.col,
  );
  if (!fineCellLocal) fail(`Unable to resolve fine cell geometry for ${fineSelection.fineCellRef.label}`);
  const localFocusRect = buildNeighborhoodRect(
    fineSelection.selectedRectLocal,
    fineSelection.fineGridSize,
    fineSelection.fineCellRef.row,
    fineSelection.fineCellRef.col,
  );
  const focusRectRoot = buildRect(
    localFocusRect.x + fineSelection.neighborhoodRect.x,
    localFocusRect.y + fineSelection.neighborhoodRect.y,
    localFocusRect.width,
    localFocusRect.height,
  );
  const thirdOriginalPath = path.join(
    baseDir,
    `${runTag}-third-focus-${String(attempt).padStart(2, '0')}.png`,
  );
  const thirdOriginalImage = await cropImage({
    image: fineSelection.neighborhoodImage,
    x: localFocusRect.x,
    y: localFocusRect.y,
    width: localFocusRect.width,
    height: localFocusRect.height,
    outPath: thirdOriginalPath,
  });
  const detailPath = path.join(
    baseDir,
    `${runTag}-third-focus-detail-${String(attempt).padStart(2, '0')}.png`,
  );
  const detailImage = await upscaleImageForDetail({
    image: thirdOriginalImage,
    outPath: detailPath,
    label: `upscale third-stage detail image attempt ${attempt}`,
    targetDimension: THIRD_STAGE_DETAIL_TARGET_DIMENSION,
  });
  const thirdGridSize = chooseThirdGridSize(detailImage);
  const thirdGridPath = path.join(
    baseDir,
    `${runTag}-third-grid-${String(attempt).padStart(2, '0')}.png`,
  );
  const { path: thirdGridImagePath, cells } = await annotateImageWithGrid({
    image: detailImage,
    gridRect: buildRect(0, 0, detailImage.width, detailImage.height),
    gridSize: thirdGridSize,
    outPath: thirdGridPath,
    label: `annotate third-stage grid attempt ${attempt}`,
  });
  const { dataUrl: neighborhoodContextDataUrl } = await toModelDataUrl({
    image: {
      path: fineSelection.neighborhoodImage.path,
      width: fineSelection.neighborhoodImage.width,
      height: fineSelection.neighborhoodImage.height,
    },
    outPath: path.join(baseDir, `${runTag}-fine-neighborhood-context-${String(attempt).padStart(2, '0')}.png`),
    label: `resize stage-2 neighborhood context image attempt ${attempt}`,
    maxDimension: CONTEXT_IMAGE_MAX_DIMENSION,
  });
  const { dataUrl: thirdGridDataUrl } = await toModelDataUrl({
    image: {
      path: thirdGridImagePath,
      width: detailImage.width,
      height: detailImage.height,
    },
    outPath: path.join(baseDir, `${runTag}-third-grid-model-${String(attempt).padStart(2, '0')}.png`),
    label: `resize third-stage grid model image attempt ${attempt}`,
    maxDimension: THIRD_STAGE_MODEL_MAX_DIMENSION,
  });
  const voted = await requestModelJsonMajority({
    apiKey,
    agent,
    model,
    reasoningEffort,
    messages: [
      {
        role: 'system',
        content:
          'You are a desktop UI grounding assistant. Refine the final click point inside the stage-3 local crop and return strict JSON.',
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: buildThirdGridPrompt({ description, gridSize: thirdGridSize }),
          },
          {
            type: 'image_url',
            image_url: { url: neighborhoodContextDataUrl },
          },
          {
            type: 'image_url',
            image_url: { url: thirdGridDataUrl },
          },
        ],
      },
    ],
    debugLabel: `third grid attempt ${attempt}`,
    voteFn: (ballots) => voteThirdGridBallots(ballots, thirdGridSize),
    stageName: 'third',
  });
  const { representative } = voted;
  const assistantText = representative.assistantText;
  const json = representative.json;
  const rawResponseText = representative.rawResponseText;
  const cellRef = parseGridCellRef(json, thirdGridSize);
  return {
    localFocusRect,
    focusRectRoot,
    thirdOriginalPath,
    thirdOriginalImage,
    detailImage,
    thirdGridSize,
    cells,
    gridImagePath: thirdGridImagePath,
    cellRef: voted.cellRef || cellRef,
    anchors: voted.anchors || normalizeAnchors(json, { maxCount: 2, defaultToCenter: Boolean(cellRef) }),
    voteMode: voted.voteMode || 'paired',
    assistantText,
    rawResponseText,
    majoritySummary: voted.summary,
    majorityCount: voted.count,
    majorityTotal: voted.total,
    confidence: Number.isFinite(Number(json?.confidence)) ? Number(json.confidence) : 0,
    reason: String(json?.reason || '').trim(),
    visibleAnchor: String(json?.visible_anchor || '').trim(),
  };
}

function buildCandidateFromThirdSelection(thirdSelection, screen) {
  const thirdCellDetail = findGridCell(
    thirdSelection.cells,
    thirdSelection.cellRef.row,
    thirdSelection.cellRef.col,
  );
  if (!thirdCellDetail) fail(`Unable to resolve third-stage cell geometry for ${thirdSelection.cellRef.label}`);
  const thirdCellOriginalLocal = mapRectBetweenImages(
    thirdCellDetail.rect,
    thirdSelection.detailImage,
    thirdSelection.thirdOriginalImage,
  );
  const thirdCellRoot = buildRect(
    thirdCellOriginalLocal.x + thirdSelection.focusRectRoot.x,
    thirdCellOriginalLocal.y + thirdSelection.focusRectRoot.y,
    thirdCellOriginalLocal.width,
    thirdCellOriginalLocal.height,
  );
  const anchorPoints = normalizeAnchors(thirdSelection.anchors, {
    maxCount: 2,
    defaultToCenter: true,
  })
    .map((anchor) => anchorToPoint(thirdCellDetail.rect, anchor))
    .map((point) => mapPointBetweenImages(point, thirdSelection.detailImage, thirdSelection.thirdOriginalImage))
    .map((point) => ({
      x: point.x + thirdSelection.focusRectRoot.x,
      y: point.y + thirdSelection.focusRectRoot.y,
    }));
  const point =
    anchorPoints.length >= 2
      ? {
          x: clampInt(
            Math.round((anchorPoints[0].x + anchorPoints[1].x) / 2),
            thirdCellRoot.x,
            thirdCellRoot.x + thirdCellRoot.width - 1,
          ),
          y: clampInt(
            Math.round((anchorPoints[0].y + anchorPoints[1].y) / 2),
            thirdCellRoot.y,
            thirdCellRoot.y + thirdCellRoot.height - 1,
          ),
        }
      : anchorPoints[0];
  const bbox = bboxAroundPoint(point, thirdCellRoot);
  return normalizeDecision(
    {
      action: 'left_click',
      target: point,
      bbox,
      confidence: Math.min(0.99, Math.max(thirdSelection.confidence || 0, 0.6)),
      reason: thirdSelection.reason || 'stage-3 grid-based point estimate',
      matched_text: thirdSelection.visibleAnchor || thirdSelection.cellRef.label,
    },
    screen,
  );
}

async function reviewGridCandidate({
  apiKey,
  agent,
  model,
  reasoningEffort,
  description,
  rootImage,
  decision,
  baseDir,
  runTag,
  attempt,
}) {
  const reviewCropRect = expandRectWithinImageXY(
    {
      x: decision.target.x,
      y: decision.target.y,
      width: 1,
      height: 1,
    },
    rootImage,
    REVIEW_CROP_MIN_SIZE / 2,
    REVIEW_CROP_MIN_SIZE / 2,
  );
  const rootAnnotatedPath = path.join(baseDir, `${runTag}-review-root-${String(attempt).padStart(2, '0')}.png`);
  const reviewRootImagePath = await annotateImageWithPoint({
    image: rootImage,
    point: decision.target,
    bbox: decision.bbox,
    outPath: rootAnnotatedPath,
    label: `annotate review root attempt ${attempt}`,
  });
  const { dataUrl: reviewRootDataUrl } = await toModelDataUrl({
    image: {
      path: reviewRootImagePath,
      width: rootImage.width,
      height: rootImage.height,
    },
    outPath: path.join(baseDir, `${runTag}-review-root-model-${String(attempt).padStart(2, '0')}.png`),
    label: `resize review root model image attempt ${attempt}`,
  });
  const voted = await requestModelJsonMajority({
    apiKey,
    agent,
    model,
    reasoningEffort: pickReviewReasoningEffort(reasoningEffort),
    messages: [
      {
        role: 'system',
        content:
          'You are a desktop UI click-point reviewer. Check whether the marked point is safe and correct. Return strict JSON only.',
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: buildReviewPrompt({
              description,
              candidatePoint: decision.target,
              candidateBBox: decision.bbox,
            }),
          },
          {
            type: 'image_url',
            image_url: { url: reviewRootDataUrl },
          },
        ],
      },
    ],
    debugLabel: `grid review attempt ${attempt}`,
    voteFn: voteReviewBallots,
    stageName: 'review',
  });
  const { representative } = voted;
  const assistantText = representative.assistantText;
  const json = representative.json;
  const rawResponseText = representative.rawResponseText;
  return {
    status: String(json?.status || '').trim().toLowerCase(),
    confidence: Number.isFinite(Number(json?.confidence)) ? Number(json.confidence) : 0,
    reason: String(json?.reason || '').trim(),
    visibleAnchor: String(json?.visible_anchor || '').trim(),
    assistantText,
    rawResponseText,
    majoritySummary: voted.summary,
    majorityCount: voted.count,
    majorityTotal: voted.total,
    reviewCropRect,
    reviewRootImagePath,
  };
}

function buildLocatePrompt({ description, screen }) {
  return [
    'Find exactly one clickable target in this Windows desktop screenshot.',
    `Task description: ${description}`,
    `Screenshot size: ${screen.width}x${screen.height}. Origin is the screenshot top-left pixel.`,
    `The uploaded root image resolution is exactly ${screen.width}x${screen.height} pixels.`,
    'Treat that resolution as the coordinate space for image_id "root".',
    'You are allowed to inspect the image with tools before giving the final answer.',
    'Available tools:',
    '- crop: crop one image to a smaller rectangle',
    '- rotate: rotate one image by 90, 180, or 270 degrees clockwise',
    'Return strict JSON only. Two valid shapes exist:',
    '{"type":"tool","tool":"crop","image_id":"root","x":0,"y":0,"width":400,"height":300,"reason":"why crop helps"}',
    '{"type":"tool","tool":"rotate","image_id":"root","degrees":90,"reason":"why rotate helps"}',
    '{"type":"final","image_id":"root","action":"left_click|right_click|none","bbox":{"x":0,"y":0,"width":0,"height":0},"confidence":0.0,"reason":"short text","matched_text":"visible label or empty"}',
    'Rules:',
    '- Coordinates are always integers relative to the image_id you name.',
    '- Prefer the shortest reliable inspection path: use the root screenshot to identify the target region, then use one coarse crop and one precision crop when needed.',
    '- Follow structural anchors named in the description first. For example: inspect the relevant bar, panel, window, list, toolbar, or region before micro-cropping.',
    '- If you are not yet certain of the exact clickable point, call crop again on a tighter region and continue refining.',
    '- Avoid exploratory micro-crops when one larger crop can isolate the target just as well.',
    '- Use rotate only when the screenshot or target is genuinely hard to inspect without rotation.',
    '- You may call multiple tools before the final answer, but only when each additional step materially improves certainty.',
    '- In the final answer, bbox must be the safe clickable range, not a single guessed pixel.',
    '- The runner will click the center of bbox, so bbox must represent the full stable clickable hit area for the target.',
    '- bbox must fully contain the requested target body or clickable surface. Do not return a bbox that cuts off part of the target or captures only a small visual fragment.',
    '- The target may be an icon, taskbar button, toolbar item, input box, search field, dropdown, checkbox, radio, tab, list row, menu item, or another UI control.',
    '- Similar-looking icons, repeated controls, and nearby lookalikes may exist. Match the target against the task description carefully and return the most correct one only.',
    '- If multiple candidates look similar, use the description, nearby anchors, visible labels, position, and surrounding context to disambiguate before answering.',
    '- For icons, use the full visible clickable affordance, such as the full taskbar button or the full icon hit area, not just the colored glyph or one corner of it.',
    '- For buttons, use the full button body. For input boxes and search fields, use the full editable field body. For dropdowns, tabs, list rows, menu items, and toolbar items, use the full selectable surface.',
    '- Do not return a bbox that covers only the most visually salient sub-part of the target, such as one letter, one symbol stroke, one colored patch, one icon corner, or one inner decoration.',
    '- Avoid edges, thin borders, tiny status badges, decorative corners, text carets, and placeholder-text fragments when a safer center-area bbox exists.',
    '- If your bbox still includes neighboring clickable targets, crop again before answering.',
    '- If your bbox does not fully include the requested clickable area, or if it is obviously too small for the actual control, crop again before answering.',
    '- If you cannot confidently distinguish the described target from similar candidates, return a final answer with action "none".',
    '- If the target is missing, ambiguous, or blocked, return a final answer with action "none".',
    '- Do not output markdown. Do not explain outside JSON.',
  ].join('\n');
}

function buildCandidateReviewPrompt({
  description,
  verifyImage,
  verifyCandidateBBox,
  rootImage,
  rootBBox,
}) {
  return [
    'Review one candidate clickable bbox using two views of the same desktop screenshot.',
    `Task description: ${description}`,
    `Image 1 is a zoom crop around the candidate. Its resolution is exactly ${verifyImage.width}x${verifyImage.height} pixels and its image_id is ${verifyImage.id}.`,
    `Current candidate bbox in image 1: x=${Math.round(verifyCandidateBBox.x)}, y=${Math.round(verifyCandidateBBox.y)}, width=${Math.round(verifyCandidateBBox.width)}, height=${Math.round(verifyCandidateBBox.height)}.`,
    `Image 2 is the original full screenshot with the candidate highlighted in red. Its resolution is exactly ${rootImage.width}x${rootImage.height} pixels and its coordinate space is image_id root.`,
    `Current candidate bbox in image_id root: x=${Math.round(rootBBox.x)}, y=${Math.round(rootBBox.y)}, width=${Math.round(rootBBox.width)}, height=${Math.round(rootBBox.height)}.`,
    'Return strict JSON only with this shape:',
    `{"status":"confirmed|adjusted|reject","image_id":"${verifyImage.id}|root","bbox":{"x":0,"y":0,"width":0,"height":0},"reason":"short text","matched_text":"visible label or empty"}`,
    'Rules:',
    '- Use both images together: the zoom crop is for precision, and the full screenshot is for context and disambiguation.',
    '- Return status "confirmed" only if the current candidate is already correct.',
    '- Return status "adjusted" when the target is correct but the bbox should change.',
    '- If you adjust using the zoom crop, return image_id as the zoom image id. If the correction is easier to express on the full screenshot, return image_id as root.',
    '- bbox must fully contain the requested target body or clickable surface in the coordinate space named by image_id.',
    '- bbox should be tight enough to exclude neighboring clickable controls, but never so tight that it cuts off part of the target or shrinks down to only a small visual fragment of it.',
    '- The target may be an icon, button, input box, dropdown, checkbox, tab, list row, menu item, or another UI control.',
    '- Similar-looking candidates may appear nearby. Confirm that the candidate matches the task description exactly, not just approximately.',
    '- If the target is a taskbar icon, desktop icon, toolbar icon, or any other glyph inside a larger clickable affordance, prefer the full practical hit area instead of the smallest colored mark.',
    '- If the target is an input box, search field, dropdown, list row, tab, or menu item, include the full interactive body that a user would naturally click.',
    '- Exclude neighboring icons, labels, badges, separators, and background if they are not part of the target body.',
    '- If the shown candidate is the wrong target, still ambiguous, or not safe to click, return status "reject".',
    '- The runner will click the center of bbox, so optimize for a stable center click by returning the full intended hit area, not a tiny sub-region.',
  ].join('\n');
}

function createRootImageRecord({ screenshotPath, screen }) {
  return {
    id: 'root',
    path: screenshotPath,
    width: screen.width,
    height: screen.height,
    parentId: null,
    transform: null,
  };
}

function buildLocateMessages({ description, screen, rootDataUrl, retryFeedbacks }) {
  const feedbackText =
    Array.isArray(retryFeedbacks) && retryFeedbacks.length
      ? `\nPrevious rejected attempts on the same screenshot:\n${retryFeedbacks
          .map((item, index) => `${index + 1}. ${item}`)
          .join('\n')}\nFind the target again from scratch and avoid the rejected result patterns above.`
      : '';
  const userContent = [
    {
      type: 'text',
      text: `${buildLocatePrompt({ description, screen })}\nInitial image_id: root.${feedbackText}`,
    },
    {
      type: 'image_url',
      image_url: { url: rootDataUrl },
    },
  ];

  return [
    {
      role: 'system',
      content:
        'You are a desktop UI grounding assistant. You may request crop or rotate tools before you return the final coordinates. You only return strict JSON and never markdown.',
    },
    {
      role: 'user',
      content: userContent,
    },
  ];
}

function mapPointToParent(childImage, point) {
  const t = childImage.transform;
  if (!t) return point;

  if (t.type === 'crop') {
    return { x: point.x + t.x, y: point.y + t.y };
  }

  if (t.type === 'rotate') {
    if (t.degrees === 90) {
      return { x: point.y, y: t.parentHeight - 1 - point.x };
    }
    if (t.degrees === 180) {
      return { x: t.parentWidth - 1 - point.x, y: t.parentHeight - 1 - point.y };
    }
    if (t.degrees === 270) {
      return { x: t.parentWidth - 1 - point.y, y: point.x };
    }
  }

  fail(`Unsupported transform mapping for image ${childImage.id}`);
}

function mapPointToRoot(images, imageId, point) {
  let currentPoint = { x: Number(point.x), y: Number(point.y) };
  let current = images.get(imageId);
  if (!current) fail(`Unknown image_id from model: ${imageId}`);

  while (current.parentId) {
    currentPoint = mapPointToParent(current, currentPoint);
    current = images.get(current.parentId);
    if (!current) fail(`Broken image transform chain at ${imageId}`);
  }

  return {
    x: Math.round(currentPoint.x),
    y: Math.round(currentPoint.y),
  };
}

function mapRectToRoot(images, imageId, rect) {
  const x0 = Number(rect.x);
  const y0 = Number(rect.y);
  const x1 = x0 + Math.max(1, Number(rect.width)) - 1;
  const y1 = y0 + Math.max(1, Number(rect.height)) - 1;
  const corners = [
    mapPointToRoot(images, imageId, { x: x0, y: y0 }),
    mapPointToRoot(images, imageId, { x: x1, y: y0 }),
    mapPointToRoot(images, imageId, { x: x0, y: y1 }),
    mapPointToRoot(images, imageId, { x: x1, y: y1 }),
  ];
  const xs = corners.map((p) => p.x);
  const ys = corners.map((p) => p.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return {
    x: minX,
    y: minY,
    width: maxX - minX + 1,
    height: maxY - minY + 1,
  };
}

function normalizeToolCall(toolCall, image) {
  const tool = String(toolCall?.tool || '').trim().toLowerCase();
  if (tool === 'crop') {
    const width = clampInt(toolCall.width, TOOL_MIN_CROP_SIZE, image.width);
    const height = clampInt(toolCall.height, TOOL_MIN_CROP_SIZE, image.height);
    const maxX = Math.max(0, image.width - width);
    const maxY = Math.max(0, image.height - height);
    const x = clampInt(toolCall.x, 0, maxX);
    const y = clampInt(toolCall.y, 0, maxY);
    return {
      type: 'tool',
      tool: 'crop',
      image_id: image.id,
      x,
      y,
      width,
      height,
      reason: String(toolCall.reason || '').trim().slice(0, 160),
    };
  }

  if (tool === 'rotate') {
    const degrees = clampInt(toolCall.degrees, 0, 270);
    const allowed = [90, 180, 270].includes(degrees) ? degrees : 90;
    return {
      type: 'tool',
      tool: 'rotate',
      image_id: image.id,
      degrees: allowed,
      reason: String(toolCall.reason || '').trim().slice(0, 160),
    };
  }

  fail(`Unsupported tool requested by model: ${tool}`);
}

async function cropImage({ image, x, y, width, height, outPath }) {
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$src = '${escapePsSingleQuoted(path.resolve(image.path))}'
$out = '${escapePsSingleQuoted(path.resolve(outPath))}'
$bmp = [System.Drawing.Bitmap]::FromFile($src)
$rect = New-Object System.Drawing.Rectangle(${x}, ${y}, ${width}, ${height})
$clone = $bmp.Clone($rect, $bmp.PixelFormat)
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$clone.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$clone.Dispose()
$bmp.Dispose()
`;
  await runPowerShell(script, 20000, `crop ${image.id} -> ${path.basename(outPath)}`);
  return {
    id: '',
    path: outPath,
    width,
    height,
    parentId: image.id,
    transform: { type: 'crop', x, y },
  };
}

async function rotateImage({ image, degrees, outPath }) {
  const rotateFlipMap = {
    90: 'Rotate90FlipNone',
    180: 'Rotate180FlipNone',
    270: 'Rotate270FlipNone',
  };
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$src = '${escapePsSingleQuoted(path.resolve(image.path))}'
$out = '${escapePsSingleQuoted(path.resolve(outPath))}'
$bmp = [System.Drawing.Bitmap]::FromFile($src)
$bmp.RotateFlip([System.Drawing.RotateFlipType]::${rotateFlipMap[degrees]})
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
`;
  await runPowerShell(script, 20000, `rotate ${image.id} -> ${path.basename(outPath)}`);
  return {
    id: '',
    path: outPath,
    width: degrees === 180 ? image.width : image.height,
    height: degrees === 180 ? image.height : image.width,
    parentId: image.id,
    transform: {
      type: 'rotate',
      degrees,
      parentWidth: image.width,
      parentHeight: image.height,
    },
  };
}

async function annotateImageWithBBox({ image, bbox, outPath }) {
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$src = '${escapePsSingleQuoted(path.resolve(image.path))}'
$out = '${escapePsSingleQuoted(path.resolve(outPath))}'
$bmp = [System.Drawing.Bitmap]::FromFile($src)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Red), 4
$font = New-Object System.Drawing.Font('Arial', 18, [System.Drawing.FontStyle]::Bold)
$x = ${Math.round(bbox.x)}
$y = ${Math.round(bbox.y)}
$w = ${Math.max(1, Math.round(bbox.width))}
$h = ${Math.max(1, Math.round(bbox.height))}
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$g.DrawRectangle($pen, $x, $y, $w, $h)
$g.DrawString('bbox', $font, [System.Drawing.Brushes]::Red, [float][Math]::Max(0, $x - 10), [float][Math]::Max(0, $y - 28))
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
`;
  await runPowerShell(script, 20000, `annotate ${image.id} -> ${path.basename(outPath)}`);
  return {
    id: path.basename(outPath, path.extname(outPath)),
    path: path.resolve(outPath),
    width: image.width,
    height: image.height,
  };
}

async function createDerivedImage({ baseDir, images, image, toolCall, counter }) {
  const outPath = path.join(baseDir, `derived-${String(counter).padStart(2, '0')}.png`);
  logEvent('INFO', 'Creating derived image', {
    parent_image_id: image.id,
    tool: toolCall.tool,
    counter,
    x: toolCall.x,
    y: toolCall.y,
    width: toolCall.width,
    height: toolCall.height,
    degrees: toolCall.degrees,
    out_path: outPath,
  });
  let derived;
  if (toolCall.tool === 'crop') {
    derived = await cropImage({
      image,
      x: toolCall.x,
      y: toolCall.y,
      width: toolCall.width,
      height: toolCall.height,
      outPath,
    });
  } else if (toolCall.tool === 'rotate') {
    derived = await rotateImage({
      image,
      degrees: toolCall.degrees,
      outPath,
    });
  } else {
    fail(`Unsupported derived image tool: ${toolCall.tool}`);
  }

  derived.id = `img_${counter}`;
  images.set(derived.id, derived);
  logEvent('INFO', 'Derived image ready', {
    derived_image_id: derived.id,
    parent_image_id: derived.parentId,
    width: derived.width,
    height: derived.height,
    path: derived.path,
  });
  return derived;
}

function buildToolResultText({ toolCall, derivedImage }) {
  if (toolCall.tool === 'crop') {
    return [
      `Tool result: crop completed.`,
      `New image_id: ${derivedImage.id}`,
      `Parent image_id: ${derivedImage.parentId}`,
      `Crop rectangle in parent image: x=${toolCall.x}, y=${toolCall.y}, width=${toolCall.width}, height=${toolCall.height}.`,
      `The uploaded image resolution for ${derivedImage.id} is exactly ${derivedImage.width}x${derivedImage.height} pixels.`,
      `Use ${derivedImage.width}x${derivedImage.height} as the coordinate space for image_id ${derivedImage.id}.`,
      `Continue using tools if needed, or return a final answer.`,
    ].join('\n');
  }

  return [
    `Tool result: rotate completed.`,
    `New image_id: ${derivedImage.id}`,
    `Parent image_id: ${derivedImage.parentId}`,
    `Rotation: ${toolCall.degrees} degrees clockwise.`,
    `The uploaded image resolution for ${derivedImage.id} is exactly ${derivedImage.width}x${derivedImage.height} pixels.`,
    `Use ${derivedImage.width}x${derivedImage.height} as the coordinate space for image_id ${derivedImage.id}.`,
    `Continue using tools if needed, or return a final answer.`,
  ].join('\n');
}

function normalizeFinalDecision(rawFinal, images, screen) {
  const imageId = rawFinal?.image_id ? String(rawFinal.image_id) : 'root';
  const image = images.get(imageId);
  if (!image) fail(`Unknown image_id in final answer: ${imageId}`);

  const localBBox = normalizeImageRelativeBBox(rawFinal, image);

  const rootBBox = mapRectToRoot(images, imageId, localBBox);
  const rootTarget = {
    x: Math.round(rootBBox.x + (rootBBox.width - 1) / 2),
    y: Math.round(rootBBox.y + (rootBBox.height - 1) / 2),
  };
  return normalizeDecision(
    {
      action: rawFinal?.action,
      target: rootTarget,
      bbox: rootBBox,
      confidence: rawFinal?.confidence,
      reason: rawFinal?.reason,
      matched_text: rawFinal?.matched_text,
    },
    screen,
  );
}

function normalizeReviewedImageId({ imageId, verifyImage, reviewImage }) {
  const raw = String(imageId || '').trim();
  if (!raw || raw === verifyImage.id) return verifyImage.id;
  if (raw === 'root' || raw === reviewImage.id) return 'root';
  fail(`Unsupported image_id from candidate review: ${raw}`);
}

function pickReviewReasoningEffort(effort) {
  const normalized = resolveReasoningEffort(effort);
  if (normalized === 'none' || normalized === 'minimal' || normalized === 'low') return 'medium';
  return 'low';
}

function pickGridStageReasoningEffort(effort) {
  const normalized = resolveReasoningEffort(effort);
  if (normalized === 'high') return 'medium';
  if (normalized === 'medium') return 'low';
  if (normalized === 'none') return 'minimal';
  return normalized;
}

async function reviewCandidate({
  apiKey,
  agent,
  model,
  reasoningEffort,
  description,
  rawFinal,
  images,
  baseDir,
  derivedCounter,
}) {
  logEvent('INFO', 'Starting combined candidate review', {
    candidate_image_id: rawFinal?.image_id || 'root',
    derived_counter: derivedCounter,
  });
  const imageId = rawFinal?.image_id ? String(rawFinal.image_id) : 'root';
  const image = images.get(imageId);
  if (!image) fail(`Unknown image_id in final bbox verification: ${imageId}`);
  const rootImage = images.get('root');
  if (!rootImage) fail('Missing root image in candidate review');

  const localBBox = normalizeImageRelativeBBox(rawFinal, image);
  const rootBBox = mapRectToRoot(images, imageId, localBBox);
  const verifyCrop = expandRectWithinImage(localBBox, image, FINAL_BBOX_VERIFY_MARGIN);
  const verifyImage = await createDerivedImage({
    baseDir,
    images,
    image,
    toolCall: {
      tool: 'crop',
      x: Math.round(verifyCrop.x),
      y: Math.round(verifyCrop.y),
      width: Math.round(verifyCrop.width),
      height: Math.round(verifyCrop.height),
      reason: 'verify final bbox',
    },
    counter: derivedCounter,
  });
  const verifyCandidateBBox = {
    x: localBBox.x - verifyCrop.x,
    y: localBBox.y - verifyCrop.y,
    width: localBBox.width,
    height: localBBox.height,
  };
  const annotatedPath = path.join(baseDir, `review-${String(derivedCounter + 1).padStart(2, '0')}.png`);
  const reviewImage = await annotateImageWithBBox({
    image: rootImage,
    bbox: rootBBox,
    outPath: annotatedPath,
  });
  const { dataUrl: verifyDataUrl } = await toModelDataUrl({
    image: verifyImage,
    outPath: path.join(baseDir, `review-${String(derivedCounter).padStart(2, '0')}-model.png`),
    label: `resize verify model image ${derivedCounter}`,
  });
  const { dataUrl: reviewDataUrl } = await toModelDataUrl({
    image: {
      path: reviewImage.path,
      width: rootImage.width,
      height: rootImage.height,
    },
    outPath: path.join(baseDir, `review-${String(derivedCounter + 1).padStart(2, '0')}-model.png`),
    label: `resize review model image ${derivedCounter + 1}`,
  });
  const messages = [
    {
      role: 'system',
      content:
        'You are a desktop UI candidate reviewer. Use the zoomed crop for precision and the full screenshot for context. Return strict JSON only.',
    },
    {
      role: 'user',
      content: [
        {
          type: 'text',
          text: buildCandidateReviewPrompt({
            description,
            verifyImage,
            verifyCandidateBBox,
            rootImage: reviewImage,
            rootBBox,
          }),
        },
        {
          type: 'image_url',
          image_url: { url: verifyDataUrl },
        },
        {
          type: 'image_url',
          image_url: { url: reviewDataUrl },
        },
      ],
    },
  ];

  const { assistantText, json, rawResponseText } = await requestModelJson({
    apiKey,
    agent,
    model,
    reasoningEffort: pickReviewReasoningEffort(reasoningEffort),
    messages,
    debugLabel: `candidate review ${verifyImage.id}`,
  });
  const status = String(json?.status || '').trim().toLowerCase();
  if (!['confirmed', 'adjusted', 'reject'].includes(status)) {
    fail(`Unsupported candidate review status: ${JSON.stringify(json)}`);
  }
  logEvent('INFO', 'Combined candidate review result', {
    status,
    verify_image_id: verifyImage.id,
    reason: json?.reason,
    matched_text: json?.matched_text,
  });

  if (status === 'reject') {
    return {
      rawFinal: {
        action: 'none',
        bbox: { x: 0, y: 0, width: 1, height: 1 },
        confidence: 0,
        reason: String(json?.reason || 'candidate review rejected target'),
        matched_text: String(json?.matched_text || ''),
      },
      trace: {
        type: 'candidate_review',
        verify_image_id: verifyImage.id,
        review_image_id: reviewImage.id,
        status,
        reason: String(json?.reason || ''),
        assistant_text: assistantText,
        raw_response_text: rawResponseText,
      },
      derivedCounter: derivedCounter + 2,
    };
  }

  const reviewedImageId = normalizeReviewedImageId({
    imageId: json?.image_id,
    verifyImage,
    reviewImage,
  });
  const reviewedImage = reviewedImageId === 'root' ? rootImage : verifyImage;
  const reviewedBBox = json?.bbox
    ? normalizeImageRelativeBBox(json, reviewedImage)
    : reviewedImageId === 'root'
      ? rootBBox
      : verifyCandidateBBox;
  return {
    rawFinal: {
      ...rawFinal,
      image_id: reviewedImageId,
      bbox: reviewedBBox,
      reason: String(json?.reason || rawFinal?.reason || ''),
      matched_text: String(json?.matched_text || rawFinal?.matched_text || ''),
    },
    trace: {
      type: 'candidate_review',
      verify_image_id: verifyImage.id,
      review_image_id: reviewImage.id,
      status,
      reviewed_image_id: reviewedImageId,
      bbox: reviewedBBox,
      reason: String(json?.reason || ''),
      assistant_text: assistantText,
      raw_response_text: rawResponseText,
    },
    derivedCounter: derivedCounter + 2,
  };
}

async function locateTarget({
  apiKey,
  agent,
  model,
  reasoningEffort,
  dumpRaw,
  reviewEnabled,
  description,
  screenshotPath,
  screen,
  baseDir,
}) {
  const rootImage = createRootImageRecord({ screenshotPath, screen });
  const runTag = path.basename(screenshotPath, path.extname(screenshotPath));
  const trace = [];
  const retryFeedbacks = [];

  for (let attempt = 0; attempt < ROOT_REVIEW_MAX_ATTEMPTS; attempt += 1) {
    const attemptNumber = attempt + 1;
    logEvent('INFO', 'Grid locate attempt started', {
      attempt: attemptNumber,
      retry_feedbacks: retryFeedbacks,
    });

    const coarseSelection = await selectCoarseCell({
      apiKey,
      agent,
      model,
      reasoningEffort,
      description,
      screen,
      rootImage,
      baseDir,
      runTag,
      attempt: attemptNumber,
      retryFeedbacks,
    });
    logEvent('INFO', 'Coarse grid result', {
      attempt: attemptNumber,
      cell: coarseSelection.cellRef?.label || 'none',
      majority: coarseSelection.majoritySummary,
      confidence: coarseSelection.confidence,
      reason: coarseSelection.reason,
      visible_anchor: coarseSelection.visibleAnchor,
    });
    if (dumpRaw) {
      trace.push({
        type: 'coarse_grid',
        attempt: attemptNumber,
        grid_size: coarseSelection.gridSize,
        cell: coarseSelection.cellRef?.label || 'none',
        majority: coarseSelection.majoritySummary,
        confidence: coarseSelection.confidence,
        reason: coarseSelection.reason,
        visible_anchor: coarseSelection.visibleAnchor,
        assistant_text: coarseSelection.assistantText,
        raw_response_text: coarseSelection.rawResponseText,
        grid_image_path: coarseSelection.gridImagePath,
      });
    }
    if (!coarseSelection.cellRef) {
      const reason = coarseSelection.reason || 'coarse grid stage could not isolate a single candidate cell';
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Coarse grid stage failed', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }
    if (coarseSelection.majorityCount < 2) {
      const reason = `coarse grid consensus too weak: ${coarseSelection.majoritySummary}`;
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Coarse grid consensus too weak', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }

    let workingCoarseSelection = coarseSelection;
    let fineSelection = null;
    for (let fineAttempt = 0; fineAttempt < 3; fineAttempt += 1) {
      fineSelection = await selectFineCell({
        apiKey,
        agent,
        model,
        reasoningEffort,
        description,
        rootImage,
        coarseSelection: workingCoarseSelection,
        baseDir,
        runTag,
        attempt: attemptNumber,
      });
      logEvent('INFO', 'Fine grid result', {
        attempt: attemptNumber,
        fine_attempt: fineAttempt + 1,
        coarse_cell: workingCoarseSelection.cellRef.label,
        fine_cell: fineSelection.fineCellRef?.label || 'none',
        majority: fineSelection.majoritySummary,
        vote_mode: fineSelection.voteMode,
        outside_direction: fineSelection.outsideDirection,
        confidence: fineSelection.confidence,
        reason: fineSelection.reason,
        visible_anchor: fineSelection.visibleAnchor,
      });
      if (dumpRaw) {
        trace.push({
          type: 'fine_grid',
          attempt: attemptNumber,
          fine_attempt: fineAttempt + 1,
          coarse_cell: workingCoarseSelection.cellRef.label,
          fine_grid_size: fineSelection.fineGridSize,
          fine_cell: fineSelection.fineCellRef?.label || 'none',
          majority: fineSelection.majoritySummary,
          vote_mode: fineSelection.voteMode,
          outside_direction: fineSelection.outsideDirection,
          confidence: fineSelection.confidence,
          reason: fineSelection.reason,
          visible_anchor: fineSelection.visibleAnchor,
          assistant_text: fineSelection.assistantText,
          raw_response_text: fineSelection.rawResponseText,
          neighborhood_rect: fineSelection.neighborhoodRect,
          neighborhood_original_path: fineSelection.neighborhoodOriginalPath,
          grid_image_path: fineSelection.gridImagePath,
        });
      }
      if (fineSelection.fineCellRef) break;
      const shiftedCellRef = shiftCellRef(
        workingCoarseSelection.cellRef,
        fineSelection.outsideDirection,
        workingCoarseSelection.gridSize,
      );
      if (!shiftedCellRef) break;
      logEvent('INFO', 'Adjusting coarse cell based on fine-grid neighbor hint', {
        attempt: attemptNumber,
        from_cell: workingCoarseSelection.cellRef.label,
        to_cell: shiftedCellRef.label,
        outside_direction: fineSelection.outsideDirection,
      });
      workingCoarseSelection = {
        ...workingCoarseSelection,
        cellRef: shiftedCellRef,
      };
    }

    if (!fineSelection?.fineCellRef) {
      const reason = fineSelection.reason || 'fine grid stage could not isolate a precise point';
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Fine grid stage failed', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }
    if (fineSelection.majorityCount < 2) {
      const reason = `fine grid consensus too weak: ${fineSelection.majoritySummary}`;
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Fine grid consensus too weak', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }

    const thirdSelection = await selectThirdStageCell({
      apiKey,
      agent,
      model,
      reasoningEffort,
      description,
      fineSelection,
      baseDir,
      runTag,
      attempt: attemptNumber,
    });
    logEvent('INFO', 'Third-stage grid result', {
      attempt: attemptNumber,
      coarse_cell: workingCoarseSelection.cellRef.label,
      fine_cell: fineSelection.fineCellRef.label,
      third_cell: thirdSelection.cellRef?.label || 'none',
      majority: thirdSelection.majoritySummary,
      vote_mode: thirdSelection.voteMode,
      positions: thirdSelection.anchors,
      confidence: thirdSelection.confidence,
      reason: thirdSelection.reason,
      visible_anchor: thirdSelection.visibleAnchor,
    });
    if (dumpRaw) {
      trace.push({
        type: 'third_grid',
        attempt: attemptNumber,
        coarse_cell: workingCoarseSelection.cellRef.label,
        fine_cell: fineSelection.fineCellRef.label,
        third_grid_size: thirdSelection.thirdGridSize,
        third_cell: thirdSelection.cellRef?.label || 'none',
        majority: thirdSelection.majoritySummary,
        vote_mode: thirdSelection.voteMode,
        positions: thirdSelection.anchors,
        confidence: thirdSelection.confidence,
        reason: thirdSelection.reason,
        visible_anchor: thirdSelection.visibleAnchor,
        assistant_text: thirdSelection.assistantText,
        raw_response_text: thirdSelection.rawResponseText,
        neighborhood_original_path: fineSelection.neighborhoodOriginalPath,
        third_original_path: thirdSelection.thirdOriginalPath,
        third_grid_image_path: thirdSelection.gridImagePath,
        third_rect_root: thirdSelection.focusRectRoot,
      });
    }
    if (!thirdSelection.cellRef) {
      const reason = thirdSelection.reason || 'third-stage grid could not isolate a precise click point';
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Third-stage grid failed', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }
    if (thirdSelection.majorityCount < 2) {
      const reason = `third-stage consensus too weak: ${thirdSelection.majoritySummary}`;
      retryFeedbacks.push(reason);
      logEvent('WARN', 'Third-stage consensus too weak', {
        attempt: attemptNumber,
        reason,
      });
      continue;
    }

    const decision = buildCandidateFromThirdSelection(thirdSelection, screen);
    logEvent('INFO', 'Candidate point computed from grid stages', {
      attempt: attemptNumber,
      x: decision.target.x,
      y: decision.target.y,
      bbox: decision.bbox,
      confidence: decision.confidence,
    });
    if (dumpRaw) {
      trace.push({
        type: 'candidate_point',
        attempt: attemptNumber,
        decision,
      });
    }

    if (!reviewEnabled) {
      if (dumpRaw) {
        trace.push({
          type: 'grid_review',
          attempt: attemptNumber,
          status: 'skipped',
          reason: 'final review disabled',
        });
      }
      if (decision.confidence >= 0.72) {
        logEvent('INFO', 'Target confirmed without grid review', {
          attempt: attemptNumber,
          x: decision.target.x,
          y: decision.target.y,
          confidence: decision.confidence,
        });
        return { decision, trace };
      }
      const retryReason = `candidate confidence too low without review: ${decision.confidence}`;
      retryFeedbacks.push(retryReason);
      logEvent('WARN', 'Skipping low-confidence candidate because grid review is disabled', {
        attempt: attemptNumber,
        reason: retryReason,
      });
      continue;
    }

    const review = await reviewGridCandidate({
      apiKey,
      agent,
      model,
      reasoningEffort,
      description,
      rootImage,
      decision,
      baseDir,
      runTag,
      attempt: attemptNumber,
    });
    logEvent('INFO', 'Grid review result', {
      attempt: attemptNumber,
      status: review.status,
      majority: review.majoritySummary,
      confidence: review.confidence,
      reason: review.reason,
      visible_anchor: review.visibleAnchor,
    });
    if (dumpRaw) {
      trace.push({
        type: 'grid_review',
        attempt: attemptNumber,
        status: review.status,
        majority: review.majoritySummary,
        confidence: review.confidence,
        reason: review.reason,
        visible_anchor: review.visibleAnchor,
        assistant_text: review.assistantText,
        raw_response_text: review.rawResponseText,
        review_crop_rect: review.reviewCropRect,
        review_root_image_path: review.reviewRootImagePath,
      });
    }
    if (
      review.status === 'pass' &&
      review.majorityCount >= REVIEW_REQUIRED_PASS_VOTES &&
      review.confidence >= 0.6 &&
      decision.confidence >= 0.72
    ) {
      logEvent('INFO', 'Target confirmed after grid review', {
        attempt: attemptNumber,
        x: decision.target.x,
        y: decision.target.y,
        confidence: decision.confidence,
      });
      return { decision, trace };
    }

    const retryReason =
      review.reason ||
      'review rejected the grid-derived point as wrong, unsafe, or ambiguous';
    retryFeedbacks.push(retryReason);
    logEvent('WARN', 'Grid review requested retry', {
      attempt: attemptNumber,
      reason: retryReason,
    });
  }

  return {
    decision: {
      action: 'none',
      target: { x: 0, y: 0 },
      bbox: { x: 0, y: 0, width: 1, height: 1 },
      confidence: 0,
      reason:
        retryFeedbacks[retryFeedbacks.length - 1] ||
        'root screenshot review retry limit reached without a confirmed target',
      matched_text: '',
    },
    trace,
  };
}

async function runPowerShell(script, timeoutMs, label = 'PowerShell') {
  return await withTiming(
    `PowerShell ${label}`,
    {
      timeout_ms: timeoutMs,
      script_chars: script.length,
    },
    async () => {
      const command = 'powershell.exe';
      const useTempFile = script.length > 7000;
      let tempScriptPath = undefined;
      let args;
      if (useTempFile) {
        tempScriptPath = path.join(
      path.resolve('saves', 'deskseeker'),
          `ps-${defaultTimestamp()}.ps1`,
        );
        await fs.mkdir(path.dirname(tempScriptPath), { recursive: true });
        await fs.writeFile(tempScriptPath, `\uFEFF${script}`, 'utf8');
        args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', tempScriptPath];
      } else {
        const encoded = Buffer.from(script, 'utf16le').toString('base64');
        args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded];
      }
      try {
        const result = await execFileAsync(command, args, {
          windowsHide: true,
          timeout: timeoutMs,
          maxBuffer: 16 * 1024 * 1024,
        });
        const stdout = String(result.stdout || '').trim();
        const stderr = String(result.stderr || '').trim();
        logEvent('INFO', 'PowerShell command completed', {
          label,
          temp_file: tempScriptPath,
          stdout_chars: stdout.length,
          stderr_chars: stderr.length,
        });
        return stdout;
      } catch (err) {
        const stdout = String(err?.stdout || '').trim();
        const stderr = String(err?.stderr || '').trim();
        logEvent('ERROR', 'PowerShell command failed', {
          label,
          temp_file: tempScriptPath,
          exit_code: err?.code,
          signal: err?.signal,
          stdout_preview: stdout,
          stderr_preview: stderr,
        });
        throw new Error(
          `PowerShell ${label} failed` +
            (err?.code !== undefined ? ` (exit ${String(err.code)})` : '') +
            `: ${previewText(stderr || stdout || err?.message || 'unknown PowerShell failure', 1400)}`,
        );
      } finally {
        if (tempScriptPath) {
          try {
            await fs.unlink(tempScriptPath);
          } catch {
            // Best effort cleanup only.
          }
        }
      }
    },
  );
}

async function runPythonScript(script, timeoutMs, label = 'Python') {
  return await withTiming(
    `Python ${label}`,
    {
      timeout_ms: timeoutMs,
      script_chars: script.length,
    },
    async () => {
      const command = 'python';
      const tempScriptPath = path.join(
      path.resolve('saves', 'deskseeker'),
        `py-${defaultTimestamp()}.py`,
      );
      await fs.mkdir(path.dirname(tempScriptPath), { recursive: true });
      await fs.writeFile(tempScriptPath, script, 'utf8');
      try {
        const result = await execFileAsync(command, ['-X', 'utf8', tempScriptPath], {
          windowsHide: true,
          timeout: timeoutMs,
          maxBuffer: 16 * 1024 * 1024,
        });
        const stdout = String(result.stdout || '').trim();
        const stderr = String(result.stderr || '').trim();
        logEvent('INFO', 'Python command completed', {
          label,
          temp_file: tempScriptPath,
          stdout_chars: stdout.length,
          stderr_chars: stderr.length,
        });
        return stdout;
      } catch (err) {
        const stdout = String(err?.stdout || '').trim();
        const stderr = String(err?.stderr || '').trim();
        logEvent('ERROR', 'Python command failed', {
          label,
          temp_file: tempScriptPath,
          exit_code: err?.code,
          signal: err?.signal,
          stdout_preview: stdout,
          stderr_preview: stderr,
        });
        throw new Error(
          `Python ${label} failed` +
            (err?.code !== undefined ? ` (exit ${String(err.code)})` : '') +
            `: ${previewText(stderr || stdout || err?.message || 'unknown Python failure', 1400)}`,
        );
      } finally {
        try {
          await fs.unlink(tempScriptPath);
        } catch {
          // Best effort cleanup only.
        }
      }
    },
  );
}

function escapePsSingleQuoted(s) {
  return String(s).replace(/'/g, "''");
}

async function captureDesktopScreenshotWithPython(outPath) {
  const absPath = path.resolve(outPath);
  const escapedPath = absPath.replace(/\\/g, '\\\\');
  const script = `
import ctypes
import json
from pathlib import Path

out_path = Path(r"""${escapedPath}""")

SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79

user32 = ctypes.windll.user32
left = int(user32.GetSystemMetrics(SM_XVIRTUALSCREEN))
top = int(user32.GetSystemMetrics(SM_YVIRTUALSCREEN))
width = int(user32.GetSystemMetrics(SM_CXVIRTUALSCREEN))
height = int(user32.GetSystemMetrics(SM_CYVIRTUALSCREEN))

backend = None
image = None
capture_width = None
capture_height = None

try:
    from mss import mss
    from PIL import Image
    with mss() as sct:
        monitor = sct.monitors[0]
        shot = sct.grab(monitor)
        capture_width = int(shot.width)
        capture_height = int(shot.height)
        image = Image.frombytes("RGB", shot.size, shot.rgb)
        backend = "python:mss"
except Exception:
    try:
        from PIL import ImageGrab
        image = ImageGrab.grab(all_screens=True)
        capture_width, capture_height = image.size
        backend = "python:imagegrab"
    except Exception as exc:
        raise RuntimeError(f"Python screenshot backends unavailable: {exc}") from exc

if image is None:
    raise RuntimeError("Python screenshot capture returned no image")

if capture_width != width or capture_height != height:
    try:
        from PIL import Image
        resample = Image.Resampling.LANCZOS
    except Exception:
        resample = 1
    image = image.resize((width, height), resample)

out_path.parent.mkdir(parents=True, exist_ok=True)
image.save(out_path, format="PNG")
print(json.dumps({
    "path": str(out_path),
    "left": left,
    "top": top,
    "width": width,
    "height": height,
    "capture_width": int(capture_width),
    "capture_height": int(capture_height),
    "backend": backend,
}, ensure_ascii=False))
`;
  const output = await runPythonScript(script, 20000, 'capture desktop screenshot');
  try {
    return JSON.parse(output);
  } catch {
    fail(`Unable to parse Python screenshot metadata: ${output}`);
  }
}

async function captureDesktopScreenshotWithPowerShell(outPath) {
  const absPath = path.resolve(outPath);
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$path = '${escapePsSingleQuoted(absPath)}'
$bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bitmap.Size)
$bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
[pscustomobject]@{
  path = $path
  left = $bounds.Left
  top = $bounds.Top
  width = $bounds.Width
  height = $bounds.Height
} | ConvertTo-Json -Compress
`;
  let lastError = undefined;
  for (let attempt = 1; attempt <= SCREENSHOT_CAPTURE_MAX_ATTEMPTS; attempt += 1) {
    try {
      const output = await runPowerShell(
        script,
        20000,
        `capture desktop screenshot attempt ${attempt}`,
      );
      try {
        return JSON.parse(output);
      } catch {
        fail(`Unable to parse screenshot metadata: ${output}`);
      }
    } catch (err) {
      lastError = err;
      if (attempt >= SCREENSHOT_CAPTURE_MAX_ATTEMPTS) break;
      const delayMs = SCREENSHOT_CAPTURE_RETRY_BASE_DELAY_MS * attempt;
      logEvent('WARN', 'Desktop screenshot capture failed; retrying', {
        attempt,
        next_attempt: attempt + 1,
        delay_ms: delayMs,
        error: summarizeError(err),
      });
      await sleep(delayMs);
    }
  }
  throw lastError ?? new Error('desktop screenshot capture failed');
}

async function captureDesktopScreenshot(outPath) {
  try {
    const screen = await captureDesktopScreenshotWithPython(outPath);
    logEvent('INFO', 'Python screenshot backend succeeded', {
      backend: screen.backend,
      capture_width: screen.capture_width,
      capture_height: screen.capture_height,
      width: screen.width,
      height: screen.height,
    });
    return screen;
  } catch (err) {
    logEvent('WARN', 'Python screenshot backend failed; falling back to PowerShell GDI capture', {
      error: summarizeError(err),
    });
    return await captureDesktopScreenshotWithPowerShell(outPath);
  }
}

async function readScreenshotMetadata(imagePath) {
  const absPath = path.resolve(imagePath);
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$path = '${escapePsSingleQuoted(absPath)}'
$bmp = [System.Drawing.Image]::FromFile($path)
try {
  [pscustomobject]@{
    path = $path
    left = 0
    top = 0
    width = $bmp.Width
    height = $bmp.Height
  } | ConvertTo-Json -Compress
} finally {
  $bmp.Dispose()
}
`;
  const output = await runPowerShell(script, 10000, 'read screenshot metadata');
  try {
    return JSON.parse(output);
  } catch {
    fail(`Unable to parse screenshot metadata: ${output}`);
  }
}

function defaultTimestamp() {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const suffix = `${process.pid}-${crypto.randomBytes(3).toString('hex')}`;
  return `${stamp}-${suffix}`;
}

async function writeJson(outPath, data) {
  const abs = path.resolve(outPath);
  await fs.mkdir(path.dirname(abs), { recursive: true });
  await fs.writeFile(abs, `${JSON.stringify(data, null, 2)}\n`, 'utf8');
  return abs;
}

function buildLogicalDesktopCoordinateResult(target) {
  return {
    coordinate_space: 'logical_desktop',
    coordinate_note: 'This is a logical desktop coordinate. Use this coordinate directly.',
    x: target == null ? null : Number(target.x),
    y: target == null ? null : Number(target.y),
  };
}

function buildDumpRawSidecarPath(resultPath) {
  const abs = path.resolve(resultPath);
  const parsed = path.parse(abs);
  return path.join(parsed.dir, `${parsed.name}.trace${parsed.ext || '.json'}`);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  runtimeState.verbose = args.verbose === true;
  if (args.help) {
    printHelp();
    return;
  }
  logEvent('INFO', 'Run started', {
    argv: process.argv.slice(2),
    cwd: process.cwd(),
    node: process.version,
  });

  if (!args.description || typeof args.description !== 'string' || !args.description.trim()) {
    printHelp();
    fail('Missing required: --description');
  }

  const description = args.description.trim();
  const modelDescription = normalizeTargetDescription(description);
  const model = resolveModelName(args.model);
  const reasoningEffort = resolveReasoningEffort(args.reasoningEffort);
  const stamp = defaultTimestamp();
  const baseDir = path.resolve('saves', 'deskseeker');
  logEvent('INFO', 'Resolved run configuration', {
    description,
    normalized_description: modelDescription,
    model,
    reasoning_effort: reasoningEffort,
    screenshot_path: args.screenshotPath ? path.resolve(args.screenshotPath) : undefined,
    review_enabled: args.review,
    verbose: args.verbose,
    dry_run: args.dryRun,
    dump_raw: args.dumpRaw,
    out_path: args.outPath ? path.resolve(args.outPath) : undefined,
  });
  await fs.mkdir(baseDir, { recursive: true });
  const sourceScreenshotPath = args.screenshotPath ? path.resolve(args.screenshotPath) : undefined;
  const screenshotPath = path.join(baseDir, `desktop-${stamp}.png`);
  const screen = sourceScreenshotPath
    ? await withTiming(
        'Load existing desktop screenshot',
        { screenshot_path: sourceScreenshotPath },
        async () => await readScreenshotMetadata(sourceScreenshotPath),
      )
    : await withTiming(
        'Capture and parse desktop screenshot',
        { screenshot_path: screenshotPath },
        async () => await captureDesktopScreenshot(screenshotPath),
      );
  if (sourceScreenshotPath) {
    await withTiming(
      'Copy existing desktop screenshot into run workspace',
      {
        source_path: sourceScreenshotPath,
        screenshot_path: screenshotPath,
      },
      async () => {
        await fs.copyFile(sourceScreenshotPath, screenshotPath);
      },
    );
  }
  screen.path = screenshotPath;
  screen.left = Number(screen.left);
  screen.top = Number(screen.top);
  screen.width = Number(screen.width);
  screen.height = Number(screen.height);

  if (!screen || !Number.isFinite(Number(screen.width)) || !Number.isFinite(Number(screen.height))) {
    fail('Screenshot capture did not return valid screen metadata');
  }

  const resultPath = args.outPath
    ? path.resolve(args.outPath)
    : path.join(baseDir, `result-${stamp}.json`);

  if (args.dryRun) {
    const result = buildLogicalDesktopCoordinateResult(null);
    await writeJson(resultPath, result);
    logEvent('INFO', 'Dry run completed', {
      result_path: resultPath,
      screenshot_path: screenshotPath,
    });
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    return;
  }

  const apiKeyResolution = await withTiming(
    'Load OpenRouter API key',
    { env_names: OPENROUTER_API_KEY_ENV_NAMES },
    async () => await resolveApiKey(),
  );
  if (!apiKeyResolution?.apiKey) {
    fail(
      `Missing OpenRouter API key. Set ${OPENROUTER_API_KEY_ENV_NAMES.join(', ')} before running the script.`,
    );
  }
  const apiKey = apiKeyResolution.apiKey;
  logEvent('INFO', 'OpenRouter API key loaded', { source: apiKeyResolution.source });

  const proxyUrl = await withTiming(
    'Resolve proxy settings',
    {},
    async () => await resolveProxyUrlAuto(),
  );
  let agent = undefined;
  if (proxyUrl) {
    try {
      agent = new HttpsProxyAgent(proxyUrl);
      logEvent('INFO', 'Proxy agent created', { proxy_url: proxyUrl });
    } catch {
      agent = undefined;
      logEvent('WARN', 'Proxy agent creation failed, falling back to direct connection', {
        proxy_url: proxyUrl,
      });
    }
  } else {
    logEvent('INFO', 'No proxy detected, using direct connection');
  }

  const { decision, trace } = await withTiming(
    'Locate target on desktop screenshot',
    { screenshot_path: screenshotPath },
    async () => await locateTarget({
      apiKey,
      agent,
      model,
      reasoningEffort,
      dumpRaw: args.dumpRaw,
      reviewEnabled: args.review,
      description: modelDescription,
      screenshotPath,
      screen,
      baseDir,
    }),
  );
  const logicalDesktopTarget =
    decision.action === 'none'
      ? null
      : {
          x: Number(screen.left) + Number(decision.target.x),
          y: Number(screen.top) + Number(decision.target.y),
        };
  if (!logicalDesktopTarget) {
    fail(
      `Unable to determine a reliable logical desktop coordinate. ${decision.reason || 'Target remained ambiguous.'}`,
    );
  }
  const result = buildLogicalDesktopCoordinateResult(logicalDesktopTarget);
  if (args.dumpRaw) {
    const dumpRawPath = buildDumpRawSidecarPath(resultPath);
    await writeJson(dumpRawPath, {
      description,
      model,
      reasoning_effort: reasoningEffort,
      screenshot: screen,
      decision,
      logical_desktop_coordinate: result,
      mode: 'locate',
      trace,
    });
    logEvent('INFO', 'Dump-raw sidecar written', {
      dump_raw_path: dumpRawPath,
    });
  }

  await writeJson(resultPath, result);
  logEvent('INFO', 'Run completed', {
    result_path: resultPath,
    logical_desktop_coordinate: result,
  });
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
}

main().catch((err) => {
  const msg = err && typeof err === 'object' && 'stack' in err ? String(err.stack) : String(err);
  fail(msg);
});

