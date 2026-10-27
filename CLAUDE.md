# Microsoft 365 discovery

Reads (and compares) configuration across multiple M365 tenants through one multi-tenant
app registration, "Baseline M365 Reader", which uses certificate auth. Credentials come from the
environment: `M365_CLIENT_ID`, `M365_CERT_PFX_B64`, `M365_CERT_PASSWORD`. Never print or commit them.

Setup: `pip install msal cryptography requests`

## Onboarding a new tenant

When the user wants to add a tenant:
1. Ask only for its `.onmicrosoft.com` domain.
2. Run `python3 m365.py onboard <domain>`. It looks up the tenant ID, records the tenant in
   `tenants.json`, and prints the admin consent URL. Give the user the tenant name, tenant ID
   and consent URL. Remind them to also assign the **Global Reader** role to
   "Baseline M365 Reader" in that tenant (Entra ID → Roles & admins → Global Reader → Add
   assignments).
3. Wait for the user to say consent is done, then run `python3 m365.py verify <domain>`
   and report the results. Graph should work right away. Exchange can take 30–60 min after
   the role assignment. Work in the tenant once verify passes.

`python3 m365.py list` shows registered tenants. `python3 m365.py token <domain> [graph|exchange|sharepoint|mde]`
prints an app-only token for direct API calls.

## Access model

- Graph (Entra, Intune, Defender, parts of Purview/SharePoint): app-only token, read-only permissions.
- Exchange Online + Defender for Office 365 policies: the `adminapi/beta/<tenantId>/InvokeCommand` REST
  endpoint with an `outlook.office365.com` token (see `cmd_verify` for the request shape), or the
  ExchangeOnlineManagement module with the certificate. Requires Global Reader on the app.
- Purview DLP/retention/labels: `Connect-IPPSSession` with the certificate.
- Full SharePoint tenant settings and any writes: delegated sign-in by the user (device code).
  The app has no write permissions.
