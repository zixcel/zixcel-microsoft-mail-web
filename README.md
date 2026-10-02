# @zixcel/microsoft-mail-web

Connect a browser mail interface to Microsoft 365 with explicit delegated consent.

## What you can do

- Read supported folders and incremental mailbox changes.
- Request additional consent for reply, archive and read-state actions.

## Current scope

The application owns its Entra registration and session. Network access uses the configured bounded transport; mailbox content is not package data.

Package distribution is not activated by this documentation. Use the checked-in source and the declared dependency versions; published availability must be verified separately.

## Getting started

Use the package manager matching the checked-in lockfile and the Node.js version declared in `package.json` or the development configuration. Run from this repository:

```sh
pnpm install --frozen-lockfile
pnpm test
```

## Documentation and source

[Usage guide](docs/getting-started.md)

[Implementation and public interfaces](src) · [Verification cases](test) · [Contributing](CONTRIBUTING.md) · [Security reporting](SECURITY.md) · [License](LICENSE) · [Attribution notices](NOTICE)
