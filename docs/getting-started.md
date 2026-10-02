# Using @zixcel/microsoft-mail-web

Connect a browser mail interface to Microsoft 365 with explicit delegated consent.

## Before you start

The application owns its Entra registration and session. Network access uses the configured bounded transport; mailbox content is not package data.

## First steps

Run from the repository root:

```sh
pnpm install --frozen-lockfile
pnpm test
```

## How to assess the result

- Read supported folders and incremental mailbox changes.
- Request additional consent for reply, archive and read-state actions.

A passing source-level check establishes only what that check observes. Keep missing configuration, unavailable services and unverified deployment paths visible.

## Continue reading

[Repository overview](../README.md)
