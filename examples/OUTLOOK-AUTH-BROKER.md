# Outlook Auth Broker P1

This example documents the P1 flow for remote-managed Outlook and Hotmail Graph accounts in EmailEngine.

## Scope

- P1 is Graph-only.
- Use an EmailEngine OAuth app with `provider=outlook` and `baseScopes=api`.
- Refresh tokens stay in the remote broker, not in EmailEngine account records.

## 1. Create the EmailEngine OAuth app

Create or update an Outlook OAuth app in EmailEngine with:

- `provider=outlook`
- `baseScopes=api`
- `authority=consumers` or your tenant
- `cloud=global` unless you need another Microsoft cloud

For remote-managed Graph accounts, EmailEngine now allows `clientSecret` and `redirectUrl` to stay empty for this `baseScopes=api` app shape.

## 2. Start the broker

Run the example broker from the `emailengine` directory:

```bash
set OUTLOOK_AUTH_BROKER_USERNAME=broker
set OUTLOOK_AUTH_BROKER_PASSWORD=secret
set OUTLOOK_AUTH_BROKER_HOST=127.0.0.1
set OUTLOOK_AUTH_BROKER_PORT=3081
node examples/outlook-auth-broker.js
```

Optional:

- `OUTLOOK_AUTH_BROKER_STORE` sets the JSON file used for account storage.

## 3. Point EmailEngine to the broker

Set EmailEngine `authServer` to the broker root URL, including Basic Auth credentials if needed:

```text
http://broker:secret@127.0.0.1:3081/
```

EmailEngine uses this root endpoint for `GET ?account=<id>&proto=api`.
The batch import route also uses the same base URL to call the broker import endpoint automatically.

## 4. Import accounts

Use the new EmailEngine API route:

```text
POST /v1/accounts/import/outlook-auth-server
```

Payload fields:

- `app`: EmailEngine OAuth app ID
- `accounts`: multi-line `email----password----client_id----refresh_token`
- `accountIdPrefix`: optional account prefix
- `dryRun`: validate only
- `proxy`: optional proxy URL shared with the broker import metadata

Example body:

```json
{
  "app": "outlook-api-app",
  "accountIdPrefix": "remote:",
  "accounts": "user.one@outlook.com----unused-password----client-id-1----refresh-token-1\nuser.two@outlook.com----unused-password----client-id-2----refresh-token-2"
}
```

What happens during import:

- EmailEngine parses the 4-part input.
- Valid entries are pushed to the remote broker.
- EmailEngine creates or updates local accounts with `oauth2.useAuthServer=true`.

## Notes

- The `password` column is preserved only to keep compatibility with the existing 4-part Outlook import format. P1 does not use it.
- The example broker currently accepts `baseScopes=api` only.
- The broker caches access tokens in memory and persists rotated refresh tokens back to its JSON store.
