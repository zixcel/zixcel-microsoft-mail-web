# @zixcel/microsoft-mail-web

Browser-side Microsoft 365 mailbox connector. It owns MSAL public-client authentication, delegated Graph scopes, plain-text mailbox reads and bounded folder delta synchronization. It does not classify mail, store tokens on a server or own UI controls. The application must supply a public Entra client ID, tenant and register the exact `http://localhost:<port>/auth/redirect.html` SPA redirect URI. Never supply a client secret. At build time, `./redirect-document` provides a standalone redirect-bridge page for the application to publish at that path.

Microsoft endpoint semantics, scopes and response interpretation remain in this
package. Browser HTTP execution is delegated to
`@crowsi/provider-http-transport`, which enforces the exact Graph origin, bounded
request/response bytes, deadlines and cancellation. Product applications do not
own a second HTTP path.

`connect()` coalesces overlapping calls. The popup opens directly at the Microsoft sign-in URL instead of leaving an intermediate `about:blank` window, and logout returns to the same bridge page. An interrupted popup is not silently replaced: only a user-requested retry may pass `{ recoverInterrupted: true }`, which invokes MSAL's popup override to cancel the pending interaction.

`MicrosoftMailSource.list()` performs a complete delta round for Inbox, Sent and Archive. Cursors and a maximum of 5,000 message summaries are held in browser memory; no message bodies are downloaded until `detail()` is called. The first round can be expensive for a large mailbox and explicitly fails after 100 pages per folder rather than pretending it is complete. The caller schedules synchronization and renders stable IDs. `@odata` links may use Graph's canonical folder-ID syntax after an initial well-known folder request; the first returned collection is pinned, and later links cannot switch folders or hosts.

Graph `Mail.Read` is requested at login; `Mail.ReadWrite` and `Mail.Send` are requested only for the corresponding operation. This package uses Graph, not IMAP and not the offline planning crate `zixcel-microsoft`. Local Node tests use an injected fake public client and never contact Microsoft.

## Package integration

The package is an independently consumable unit. Callers reference its documented
interface through a versioned dependency and own application-specific composition
and integration.
