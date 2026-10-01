import assert from 'node:assert/strict'
import { afterEach, test } from 'node:test'
import { MicrosoftMailSource } from '../src/index.mjs'
import { microsoftRedirectDocument } from '../src/redirect-document.mjs'

const account = { username: 'person@example.invalid', homeAccountId: 'account-1' }
const requestedScopes = []
const fakeClient = {
  async initialize() {},
  getActiveAccount: () => null,
  getAllAccounts: () => [],
  setActiveAccount() {},
  async loginPopup() { return { account } },
  async acquireTokenSilent({ scopes }) {
    requestedScopes.push(scopes)
    return { accessToken: 'test-token' }
  }
}
const clientId = '00000000-0000-4000-8000-000000000001'
const redirectUri = 'http://localhost:4311/auth/redirect.html'
const graph = 'https://graph.microsoft.com/v1.0'
const savedFetch = globalThis.fetch

afterEach(() => {
  globalThis.fetch = savedFetch
  requestedScopes.length = 0
})

function page(value, linkName, link) {
  return Response.json({ value, [`@odata.${linkName}`]: link })
}

test('MSAL login, paged delta, update/removal and least-privilege operation scopes', async () => {
  const calls = []
  let cycle = 0
  globalThis.fetch = async (input, init) => {
    const url = String(input)
    calls.push({ url, init })
    if (url.includes('/mailFolders/inbox/messages/delta')) {
      if (url.includes('$skiptoken=1')) return page([{ id: 'm2', changeKey: 'revision-1', conversationId: 'thread-1', subject: 'Second', from: { emailAddress: { address: 'sender@example.invalid' } }, toRecipients: [{ emailAddress: { address: 'recipient@example.invalid' } }], isRead: true, hasAttachments: false }], 'deltaLink', `${graph}/me/mailFolders/inbox/messages/delta?$deltatoken=one`)
      if (url.includes('$deltatoken=one')) {
        cycle += 1
        return page([{ id: 'm1', '@removed': { reason: 'deleted' } }, { id: 'm2', subject: 'Updated' }], 'deltaLink', `${graph}/me/mailFolders/inbox/messages/delta?$deltatoken=two`)
      }
      return page([{ id: 'm1', subject: 'First' }], 'nextLink', `${graph}/me/mailFolders/inbox/messages/delta?$skiptoken=1`)
    }
    if (url.includes('/mailFolders/')) {
      const folder = url.includes('/sentitems/') ? 'sentitems' : 'archive'
      return page([], 'deltaLink', `${graph}/me/mailFolders/${folder}/messages/delta?$deltatoken=one`)
    }
    if (url.endsWith('/move')) return Response.json({ id: 'moved' }, { status: 201 })
    if (url.endsWith('/reply')) return new Response(null, { status: 202 })
    if (init.method === 'PATCH') return new Response(null, { status: 204 })
    return Response.json({ id: 'm2', body: { contentType: 'html', content: '<p>HTML body</p>' }, ccRecipients: [{ emailAddress: { address: 'copy@example.invalid' } }] })
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  assert.equal(await source.initialize(), null)
  assert.equal((await source.connect()).username, account.username)
  const first = await source.list()
  assert.deepEqual(first.map(item => item.id), ['m1', 'm2'])
  assert.equal(first[1].mailbox, 'inbox')
  const detail = await source.detail(first[1])
  assert.equal(detail.body, '<p>HTML body</p>')
  assert.equal(detail.bodyContentType, 'html')
  assert.equal(detail.cc, 'copy@example.invalid')
  assert.equal(detail.bcc, '')
  const second = await source.list()
  assert.equal(cycle, 1)
  assert.deepEqual(second.map(item => item.id), ['m2'])
  assert.equal(second[0].subject, 'Updated')
  assert.equal(second[0].from, 'sender@example.invalid')
  assert.equal(second[0].fromAddress, 'sender@example.invalid')
  assert.equal(second[0].toRecipientCount, 1)
  assert.equal(second[0].changeKey, 'revision-1')
  assert.equal(second[0].conversationId, 'thread-1')
  assert.equal(second[0].readObserved, true)
  assert.equal(second[0].hasNonInlineAttachments, false)
  assert.equal(second[0].subjectObserved, true)
  assert.equal(second[0].hasSubject, true)
  assert.equal(second[0].read, true)
  assert.ok(calls[0].url.includes('changeKey'))
  assert.ok(!calls[0].url.includes('bodyPreview'))
  assert.equal(new Headers(calls[0].init.headers).get('prefer'), 'IdType="ImmutableId"')
  assert.equal(second[0].preview, '')
  await source.archive(second[0])
  await source.markRead(second[0], true)
  await source.reply(second[0], 'Thank you')
  assert.deepEqual(requestedScopes.slice(-3), [['Mail.ReadWrite'], ['Mail.ReadWrite'], ['Mail.Send']])
  assert.ok(calls.every(call => new Headers(call.init.headers).get('authorization') === 'Bearer test-token'))
  assert.equal(new Headers(calls.find(call => call.url.includes('/me/messages/'))?.init.headers).get('prefer'), 'IdType="ImmutableId", outlook.body-content-type="html"')
})

test('a foreign delta cursor is rejected before a token is sent to it', async () => {
  const calls = []
  globalThis.fetch = async input => {
    calls.push(String(input))
    return page([], 'nextLink', 'https://example.invalid/steal-token')
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  await source.initialize()
  await source.connect()
  await assert.rejects(source.list(), /delta cursor was not issued/)
  assert.equal(calls.length, 1)
  assert.ok(calls[0].startsWith(graph))
})

test('Graph canonical folder-ID delta links are followed and pinned to their collection', async () => {
  const calls = []
  const inbox = `${graph}/me/mailfolders('AQMk%2Bone%3D')/messages/delta`
  globalThis.fetch = async input => {
    const url = String(input)
    calls.push(url)
    if (url.includes('$skiptoken=one')) return page([{ id: 'm2', subject: 'Second' }], 'deltaLink', `${inbox}?$deltatoken=done`)
    if (url.includes('/mailFolders/inbox/messages/delta')) return page([{ id: 'm1', subject: 'First' }], 'nextLink', `${inbox}?$skiptoken=one`)
    if (url.includes('/mailFolders/sentitems/messages/delta')) return page([], 'deltaLink', `${graph}/me/mailFolders/sent-id/messages/delta?$deltatoken=done`)
    if (url.includes('/mailFolders/archive/messages/delta')) return page([], 'deltaLink', `${graph}/me/mailfolders('archive-id')/messages/delta?$deltatoken=done`)
    throw new Error(`Unexpected delta request: ${url}`)
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  await source.initialize()
  await source.connect()
  assert.deepEqual((await source.list()).map(item => item.id), ['m1', 'm2'])
  assert.equal(calls.length, 4)
  assert.equal(calls[1], `${inbox}?$skiptoken=one`)
})

test('a canonical delta link cannot switch mail folders within one synchronization', async () => {
  const calls = []
  globalThis.fetch = async input => {
    const url = String(input)
    calls.push(url)
    if (url.includes('$skiptoken=one')) {
      return page([], 'deltaLink', `${graph}/me/mailfolders('other-folder')/messages/delta?$deltatoken=done`)
    }
    return page([], 'nextLink', `${graph}/me/mailfolders('inbox-id')/messages/delta?$skiptoken=one`)
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  await source.initialize()
  await source.connect()
  await assert.rejects(source.list(), /cursor changed mail folders/)
  assert.equal(calls.length, 2)
})

test('invalid reply is rejected before contacting Graph', async () => {
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  await assert.rejects(source.reply({ id: 'm1' }, ' '), /Reply text is required/)
})

test('missing public registration is rejected before starting authentication', () => {
  assert.throws(() => new MicrosoftMailSource('', 'organizations', redirectUri),
    /client ID is not configured/)
  assert.throws(() => new MicrosoftMailSource(clientId, 'organizations', 'http://localhost:4311'),
    /dedicated local/)
})

test('overlapping sign-in is one popup and only an explicit recovery cancels an interrupted interaction', async () => {
  const requests = []
  let resolveLogin
  const client = {
    ...fakeClient,
    loginPopup(request) {
      requests.push(request)
      if (requests.length === 1) return new Promise(resolve => { resolveLogin = resolve })
      return Promise.resolve({ account })
    }
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, client)
  await source.initialize()
  const first = source.connect()
  assert.equal(source.connect(), first)
  assert.equal(requests.length, 1)
  assert.equal(requests[0].overrideInteractionInProgress, undefined)
  resolveLogin({ account })
  await first
  await source.connect({ recoverInterrupted: true })
  assert.equal(requests.length, 2)
  assert.equal(requests[1].overrideInteractionInProgress, true)
})

test('the packaged redirect is a dedicated bridge document, not the application page', () => {
  const html = microsoftRedirectDocument()
  assert.match(html, /msalRedirectBridge\.broadcastResponseToMainFrame\(\)/)
  assert.doesNotMatch(html, /nuxt|Reference App|loginPopup/i)
})

test('switching Microsoft accounts cannot commit the earlier mailbox response', async () => {
  let resolveFirst
  let login = 0
  const switchingClient = {
    ...fakeClient,
    async loginPopup() {
      login += 1
      return { account: { username: `account${login}@example.invalid`, homeAccountId: `account-${login}` } }
    }
  }
  globalThis.fetch = async () => new Promise(resolve => { resolveFirst = resolve })
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, switchingClient)
  await source.initialize()
  await source.connect()
  const earlier = source.list()
  await new Promise(resolve => setImmediate(resolve))
  await source.connect()
  resolveFirst(page([{ id: 'old-account-mail', subject: 'Old' }], 'deltaLink', `${graph}/me/mailFolders/inbox/messages/delta?$deltatoken=one`))
  await assert.rejects(earlier, /account changed/)
  assert.equal(source.getAccount().homeAccountId, 'account-2')
})

test('an expired delta token restarts only its folder and replaces stale messages atomically', async () => {
  let inboxInitial = 0
  globalThis.fetch = async input => {
    const url = String(input)
    if (url.includes('/mailFolders/inbox/messages/delta')) {
      if (url.includes('$deltatoken=one')) return new Response('expired', { status: 410 })
      inboxInitial += 1
      return page([{ id: inboxInitial === 1 ? 'old' : 'new', subject: 'Invoice' }],
        'deltaLink', `${graph}/me/mailFolders/inbox/messages/delta?$deltatoken=one`)
    }
    const folder = url.includes('/sentitems/') ? 'sentitems' : 'archive'
    return page([], 'deltaLink', `${graph}/me/mailFolders/${folder}/messages/delta?$deltatoken=one`)
  }
  const source = new MicrosoftMailSource(clientId, 'organizations', redirectUri, fakeClient)
  await source.initialize()
  await source.connect()
  assert.deepEqual((await source.list()).map(item => item.id), ['old'])
  assert.deepEqual((await source.list()).map(item => item.id), ['new'])
  assert.equal(inboxInitial, 2)
})
