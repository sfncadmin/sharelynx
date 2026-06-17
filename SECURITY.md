# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in ShareLynx, please report it privately using [GitHub Security Advisories](https://github.com/sfncadmin/sharelynx/security/advisories/new).

**Do not open a public issue for security vulnerabilities.**

We will acknowledge your report within 3 business days and aim to release a fix within 14 days of confirmation.

## Scope

ShareLynx is a client-side Outlook add-in with no backend. Security concerns include:

- Token handling and authentication flows (MSAL.js)
- Microsoft Graph permission scopes (least-privilege)
- Sharing link creation and access management
- Tenant policy enforcement
- Browser storage of preferences and cached data

## What is not a vulnerability

- The `clientId` and `tenantId` in `src/config.js` are not secrets. This is a public-client SPA with no client secret; these values identify the app registration but cannot authenticate without a user's interactive sign-in.
- Hosting URL exposure in the manifest template. Each deployment generates its own manifest from the template.
