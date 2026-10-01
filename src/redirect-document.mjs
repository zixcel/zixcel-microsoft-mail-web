import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

const require = createRequire(import.meta.url)

/** Build-time document for the MSAL popup return; no application or router code runs here. */
export function microsoftRedirectDocument() {
  const packageDirectory = dirname(require.resolve('@azure/msal-browser/package.json'))
  const bridge = readFileSync(join(packageDirectory, 'lib/redirect-bridge/msal-redirect-bridge.min.js'), 'utf8')
  if (bridge.includes('</script')) throw new Error('MSAL redirect bridge cannot be embedded safely')
  return `<!doctype html>\n<html lang="en"><head><meta charset="utf-8"><title>Microsoft authentication</title></head><body><p>Completing sign-in…</p><script>${bridge}\nwindow.msalRedirectBridge.broadcastResponseToMainFrame().catch(function(error){document.body.textContent = 'Microsoft authentication could not complete: ' + String(error && error.message || error)});</script></body></html>\n`
}
