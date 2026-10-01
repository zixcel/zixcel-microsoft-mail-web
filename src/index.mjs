import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser'
import { createProviderHttpTransport } from '@crowsi/provider-http-transport'

const GRAPH = 'https://graph.microsoft.com/v1.0'
const READ = ['User.Read', 'Mail.Read']
const WRITE = ['Mail.ReadWrite']
const SEND = ['Mail.Send']
const FOLDERS = [
  ['inbox', 'inbox'],
  ['sent', 'sentitems'],
  ['archive', 'archive']
]
const FOLDER_ALIASES = new Set(FOLDERS.map(([, alias]) => alias))
const MAX_PAGES_PER_FOLDER = 100
const MAX_MESSAGES = 5000

export class MicrosoftMailSource {
  #client
  #account = null
  #messages = new Map()
  #cursors = new Map()
  #collections = new Map()
  #pending = null
  #connectPending = null
  #abort = null
  #generation = 0
  #transport

  constructor(clientId, tenant = 'organizations', redirectUri = globalThis.location && `${globalThis.location.origin}/auth/redirect.html`, authClient = null) {
    if (!/^[0-9a-f-]{36}$/i.test(clientId)) throw new Error('Microsoft client ID is not configured')
    if (!/^(organizations|consumers|common|[0-9a-f-]{36})$/i.test(tenant)) throw new Error('Invalid Microsoft tenant')
    const redirect = redirectUri && URL.canParse(redirectUri) ? new URL(redirectUri) : null
    if (!redirect || redirect.protocol !== 'http:' || !['localhost', '127.0.0.1'].includes(redirect.hostname) ||
      redirect.pathname !== '/auth/redirect.html' || redirect.search || redirect.hash) {
      throw new Error('Microsoft redirect must be the dedicated local /auth/redirect.html page')
    }
    this.#client = authClient ?? new PublicClientApplication({
      auth: { clientId, authority: `https://login.microsoftonline.com/${tenant}`, redirectUri, postLogoutRedirectUri: redirectUri },
      cache: { cacheLocation: 'sessionStorage' },
      system: { navigatePopups: false }
    })
    this.#transport = createProviderHttpTransport({
      allowedOrigins: ['https://graph.microsoft.com'],
      maximumRequestBytes: 64 * 1024,
      maximumResponseBytes: 4 * 1024 * 1024,
      timeoutMs: 30000
    })
  }

  async initialize() {
    await this.#client.initialize()
    this.#account = this.#client.getActiveAccount() ?? this.#client.getAllAccounts()[0] ?? null
    if (this.#account) this.#client.setActiveAccount(this.#account)
    return this.#account
  }

  connect({ recoverInterrupted = false } = {}) {
    if (this.#connectPending) return this.#connectPending
    const pending = this.#connect(recoverInterrupted).finally(() => {
      if (this.#connectPending === pending) this.#connectPending = null
    })
    this.#connectPending = pending
    return pending
  }

  async #connect(recoverInterrupted) {
    const result = await this.#client.loginPopup({
      scopes: READ, prompt: 'select_account',
      ...(recoverInterrupted ? { overrideInteractionInProgress: true } : {})
    })
    if (this.#account?.homeAccountId !== result.account.homeAccountId) {
      this.#generation += 1
      this.#abort?.abort()
      this.#pending = null
      this.#messages.clear()
      this.#cursors.clear()
      this.#collections.clear()
    }
    this.#account = result.account
    this.#client.setActiveAccount(result.account)
    return result.account
  }

  async disconnect() {
    if (this.#account) await this.#client.logoutPopup({ account: this.#account })
    this.#generation += 1
    this.#abort?.abort()
    this.#pending = null
    this.#account = null
    this.#messages.clear()
    this.#cursors.clear()
    this.#collections.clear()
  }

  getAccount() { return this.#account }

  async #token(scopes, interactive = false, account = this.#account) {
    if (!account) throw new Error('Microsoft account is not connected')
    try {
      return (await this.#client.acquireTokenSilent({ account, scopes })).accessToken
    } catch (error) {
      if (!interactive || !(error instanceof InteractionRequiredAuthError)) throw error
      return (await this.#client.acquireTokenPopup({ account, scopes })).accessToken
    }
  }

  async #request(path, scopes, init = {}, interactive = false, account = this.#account) {
    const accessToken = await this.#token(scopes, interactive, account)
    const response = await this.#transport.request({
      url: path.startsWith('https://') ? path : `${GRAPH}${path}`,
      method: init.method ?? 'GET',
      body: init.body ?? null,
      signal: init.signal,
      headers: {
        authorization: `Bearer ${accessToken}`,
        prefer: 'IdType="ImmutableId"',
        ...(init.body ? { 'content-type': 'application/json' } : {}),
        ...init.headers
      }
    })
    const responseText = new TextDecoder().decode(response.body)
    if (response.status < 200 || response.status >= 300) {
      const error = new Error(`Microsoft Graph ${response.status}: ${responseText.slice(0, 300)}`)
      error.status = response.status
      throw error
    }
    if (response.status === 204 || response.status === 202) return null
    try { return JSON.parse(responseText) }
    catch (error) { throw new Error('Microsoft Graph response is not valid JSON', { cause: error }) }
  }

  #deltaPath(folder) {
    const query = new URLSearchParams({
      '$top': '50',
      '$select': 'id,changeKey,conversationId,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,flag'
    })
    return `/me/mailFolders/${folder}/messages/delta?${query}`
  }

  #checkedCursor(link, mailbox, folder, collections) {
    let url
    try { url = new URL(link) } catch { throw new Error('Microsoft Graph delta cursor URL is invalid') }
    const direct = /^\/v1\.0\/me\/mailfolders\/([^/]+)\/messages\/delta$/i.exec(url.pathname)
    const odata = /^\/v1\.0\/me\/mailfolders\('([^']+)'\)\/messages\/delta$/i.exec(url.pathname)
    let collection
    try { collection = decodeURIComponent(direct?.[1] ?? odata?.[1] ?? '') }
    catch { throw new Error('Microsoft Graph delta cursor folder ID is invalid') }
    if (url.origin !== 'https://graph.microsoft.com' || url.username || url.password || url.hash ||
      !collection || (FOLDER_ALIASES.has(collection.toLowerCase()) && collection.toLowerCase() !== folder)) {
      throw new Error('Microsoft Graph delta cursor was not issued for this mailbox')
    }
    const pinned = collections.get(mailbox)
    if (pinned && pinned !== collection && pinned !== folder) {
      throw new Error('Microsoft Graph delta cursor changed mail folders')
    }
    for (const [otherMailbox, otherCollection] of collections) {
      if (otherMailbox !== mailbox && otherCollection === collection) {
        throw new Error('Microsoft Graph delta cursor reused another mail folder')
      }
    }
    collections.set(mailbox, collection)
    return url.href
  }

  async list() {
    if (this.#pending) return this.#pending
    const abort = new AbortController()
    const generation = this.#generation
    const account = this.#account
    this.#abort = abort
    const pending = this.#sync(generation, account, abort.signal).finally(() => {
      if (this.#pending === pending) {
        this.#pending = null
        this.#abort = null
      }
    })
    this.#pending = pending
    return pending
  }

  async #sync(generation, account, signal) {
    if (!account) throw new Error('Microsoft account is not connected')
    const messages = new Map(this.#messages)
    const cursors = new Map(this.#cursors)
    const collections = new Map(this.#collections)
    for (const [mailbox, folder] of FOLDERS) {
      let path = cursors.get(mailbox) ?? this.#deltaPath(folder)
      let complete = false
      let restarted = false
      let page = 0
      while (page < MAX_PAGES_PER_FOLDER) {
        if (signal.aborted || generation !== this.#generation) throw new Error('Microsoft account changed during synchronization')
        let result
        try {
          result = await this.#request(path, READ, { signal }, false, account)
        } catch (error) {
          if (signal.aborted || generation !== this.#generation) {
            throw new Error('Microsoft account changed during synchronization', { cause: error })
          }
          if (error?.status === 410 && cursors.has(mailbox) && !restarted) {
            for (const [id, message] of messages) if (message.mailbox === mailbox) messages.delete(id)
            cursors.delete(mailbox)
            collections.delete(mailbox)
            path = this.#deltaPath(folder)
            restarted = true
            page = 0
            continue
          }
          throw error
        }
        page += 1
        if (signal.aborted || generation !== this.#generation) throw new Error('Microsoft account changed during synchronization')
        if (!Array.isArray(result?.value)) throw new Error('Microsoft Graph delta response is invalid')
        for (const item of result.value) {
          if (typeof item.id !== 'string') throw new Error('Microsoft Graph message ID is missing')
          if (item['@removed']) {
            if (messages.get(item.id)?.mailbox === mailbox) messages.delete(item.id)
          } else {
            messages.set(item.id, this.#normalize(item, mailbox, messages.get(item.id)))
          }
        }
        if (messages.size > MAX_MESSAGES) throw new Error('Microsoft mailbox exceeds the local message limit')
        if (result['@odata.nextLink']) {
          path = this.#checkedCursor(result['@odata.nextLink'], mailbox, folder, collections)
        } else if (result['@odata.deltaLink']) {
          cursors.set(mailbox, this.#checkedCursor(result['@odata.deltaLink'], mailbox, folder, collections))
          complete = true
          break
        } else {
          throw new Error('Microsoft Graph delta response has no continuation or final cursor')
        }
      }
      if (!complete) throw new Error('Microsoft Graph delta page limit exceeded')
    }
    if (generation !== this.#generation) throw new Error('Microsoft account changed during synchronization')
    this.#messages = messages
    this.#cursors = cursors
    this.#collections = collections
    return Array.from(messages.values())
  }

  async detail(message) {
    const query = new URLSearchParams({ '$select': 'id,body,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,isRead,hasAttachments' })
    const value = await this.#request(`/me/messages/${encodeURIComponent(message.id)}?${query}`, READ, {
      headers: { prefer: 'IdType="ImmutableId", outlook.body-content-type="html"' }
    })
    const bodyContentType = value.body?.contentType?.toLocaleLowerCase()
    if (bodyContentType !== undefined && bodyContentType !== 'html' && bodyContentType !== 'text') {
      throw new Error('Microsoft Graph message body content type is invalid')
    }
    return {
      ...message,
      subject: value.subject ?? message.subject,
      from: value.from?.emailAddress?.name || value.from?.emailAddress?.address || message.from,
      to: Array.isArray(value.toRecipients) ? value.toRecipients.map(item => item.emailAddress?.address).filter(Boolean).join(', ') : message.to,
      cc: Array.isArray(value.ccRecipients) ? value.ccRecipients.map(item => item.emailAddress?.address).filter(Boolean).join(', ') : '',
      bcc: Array.isArray(value.bccRecipients) ? value.bccRecipients.map(item => item.emailAddress?.address).filter(Boolean).join(', ') : '',
      body: value.body?.content ?? message.body,
      bodyContentType: bodyContentType ?? message.bodyContentType
    }
  }

  async archive(message) {
    await this.#request(`/me/messages/${encodeURIComponent(message.id)}/move`, WRITE,
      { method: 'POST', body: JSON.stringify({ destinationId: 'archive' }) }, true)
  }

  async markRead(message, isRead) {
    await this.#request(`/me/messages/${encodeURIComponent(message.id)}`, WRITE,
      { method: 'PATCH', body: JSON.stringify({ isRead }) }, true)
  }

  async reply(message, text) {
    if (!text.trim() || text.length > 8192) throw new Error('Reply text is required and must be under 8192 characters')
    await this.#request(`/me/messages/${encodeURIComponent(message.id)}/reply`, SEND,
      { method: 'POST', body: JSON.stringify({ comment: text }) }, true)
  }

  #normalize(message, mailbox, previous) {
    return {
      id: message.id, source: 'microsoft', mailbox,
      changeKey: message.changeKey ?? previous?.changeKey ?? null,
      conversationId: message.conversationId ?? previous?.conversationId ?? null,
      fromAddress: message.from?.emailAddress?.address ?? previous?.fromAddress ?? null,
      toRecipientCount: Array.isArray(message.toRecipients) ? message.toRecipients.length : previous?.toRecipientCount ?? null,
      readObserved: typeof message.isRead === 'boolean' || previous?.readObserved === true,
      hasNonInlineAttachments: typeof message.hasAttachments === 'boolean' ? message.hasAttachments : previous?.hasNonInlineAttachments ?? null,
      subjectObserved: typeof message.subject === 'string' || previous?.subjectObserved === true,
      hasSubject: typeof message.subject === 'string' ? Boolean(message.subject.trim()) : previous?.hasSubject ?? null,
      from: message.from?.emailAddress?.name || message.from?.emailAddress?.address || previous?.from || 'Unknown sender',
      to: message.toRecipients?.map(item => item.emailAddress?.address).filter(Boolean).join(', ') ?? previous?.to ?? '',
      subject: message.subject ?? previous?.subject ?? '(no subject)',
      preview: '', body: previous?.body ?? '', bodyContentType: previous?.bodyContentType ?? null,
      receivedAt: message.receivedDateTime ?? previous?.receivedAt ?? '',
      read: message.isRead ?? previous?.read ?? false,
      starred: message.flag?.flagStatus === undefined ? (previous?.starred ?? false) : message.flag.flagStatus === 'flagged',
      attachments: message.hasAttachments === undefined ? (previous?.attachments ?? [])
        : message.hasAttachments ? ['Attachments available in Outlook'] : []
    }
  }
}
