#!/usr/bin/env python3
"""Tenant onboarding and token helper for the Baseline M365 Reader app.

Usage:
  m365.py onboard <domain.onmicrosoft.com>   look up tenant ID, register it, print the consent URL
  m365.py verify  <domain>                   confirm consent + Global Reader role after the admin approves
  m365.py token   <domain> [graph|exchange|sharepoint|mde]   print an app-only access token
  m365.py list                               show registered tenants

Reads M365_CLIENT_ID, M365_CERT_PFX_B64 and M365_CERT_PASSWORD from the environment.
"""
import base64
import datetime
import json
import os
import sys
from pathlib import Path

import msal
import requests
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs12

REGISTRY = Path(__file__).resolve().parent / "tenants.json"
LOGIN = "https://login.microsoftonline.com"
GLOBAL_READER_ROLE_TEMPLATE = "f2ef992c-3afb-46b9-b7cf-a126ee74c451"
EXO_SYSTEM_MAILBOX = "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}"


def load_registry():
    return json.loads(REGISTRY.read_text()) if REGISTRY.exists() else {}


def save_registry(reg):
    REGISTRY.write_text(json.dumps(reg, indent=2, sort_keys=True) + "\n")


def tenant(domain):
    domain = domain.lower()
    reg = load_registry()
    if domain not in reg:
        sys.exit(f"{domain} is not registered; run: m365.py onboard {domain}")
    return reg[domain]


def lookup_tenant(domain):
    r = requests.get(f"{LOGIN}/{domain}/v2.0/.well-known/openid-configuration", timeout=15)
    if r.status_code != 200:
        sys.exit(f"Tenant lookup failed for {domain}: {r.json().get('error_description', r.text)[:200]}")
    cfg = r.json()
    tenant_id = cfg["issuer"].split("/")[3]
    realm = requests.get(f"{LOGIN}/common/userrealm/probe@{domain}?api-version=2.1", timeout=15).json()
    return {
        "tenantId": tenant_id,
        "name": realm.get("FederationBrandName") or domain,
        "region": cfg.get("tenant_region_scope"),
        "cloud": cfg.get("cloud_instance_name"),
    }


def consent_url(tenant_id):
    return f"{LOGIN}/{tenant_id}/adminconsent?client_id={os.environ['M365_CLIENT_ID']}"


def credential():
    key, cert, _ = pkcs12.load_key_and_certificates(
        base64.b64decode(os.environ["M365_CERT_PFX_B64"]),
        os.environ.get("M365_CERT_PASSWORD", "").encode() or None,
    )
    pem = key.private_bytes(
        serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()
    ).decode()
    return {"private_key": pem, "thumbprint": cert.fingerprint(hashes.SHA1()).hex()}


def get_token(t, resource="graph"):
    scopes = {
        "graph": "https://graph.microsoft.com/.default",
        "exchange": "https://outlook.office365.com/.default",
        "sharepoint": f"https://{t['domain'].split('.')[0]}-admin.sharepoint.com/.default",
        "mde": "https://api.securitycenter.microsoft.com/.default",
    }
    app = msal.ConfidentialClientApplication(
        os.environ["M365_CLIENT_ID"], authority=f"{LOGIN}/{t['tenantId']}", client_credential=credential()
    )
    result = app.acquire_token_for_client(scopes=[scopes[resource]])
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result).splitlines()[0])
    return result["access_token"]


def cmd_onboard(domain):
    domain = domain.lower()
    info = lookup_tenant(domain)
    if info["cloud"] != "microsoftonline.com" or info["region"] not in ("WW", "NA", "EU", "AS", "OC", "AF", "SA", "JP"):
        print(f"WARNING: {domain} is in a sovereign/government cloud ({info['cloud']}, {info['region']}); "
              "these endpoints and this app will not work there.")
    reg = load_registry()
    reg[domain] = {**reg.get(domain, {}), **info, "domain": domain,
                   "status": reg.get(domain, {}).get("status", "pending-consent")}
    save_registry(reg)
    print(f"Tenant:    {info['name']}")
    print(f"Domain:    {domain}")
    print(f"Tenant ID: {info['tenantId']}")
    print(f"Consent:   {consent_url(info['tenantId'])}")


def cmd_verify(domain):
    t = tenant(domain)
    ok = True
    try:
        graph = {"Authorization": f"Bearer {get_token(t)}"}
    except RuntimeError as e:
        hint = ""
        if any(code in str(e) for code in ("AADSTS700016", "AADSTS7000229", "AADSTS65001")):
            hint = f"\nAdmin consent has not been granted yet: {consent_url(t['tenantId'])}"
        elif "AADSTS53003" in str(e):
            hint = "\nA Conditional Access policy for workload identities blocks this app; exclude it in that tenant."
        sys.exit(f"FAIL graph token: {e}{hint}")

    org = requests.get("https://graph.microsoft.com/v1.0/organization", headers=graph, timeout=30).json()["value"][0]
    print(f"OK   graph: {org['displayName']} ({org['id']})")

    sp = requests.get(
        f"https://graph.microsoft.com/v1.0/servicePrincipals(appId='{os.environ['M365_CLIENT_ID']}')"
        "?$select=id,displayName", headers=graph, timeout=30).json()
    roles = requests.get(
        "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments"
        f"?$filter=principalId eq '{sp['id']}'&$expand=roleDefinition", headers=graph, timeout=30).json()
    role_names = sorted(a["roleDefinition"]["displayName"] for a in roles.get("value", []))
    if any(a["roleDefinition"]["templateId"] == GLOBAL_READER_ROLE_TEMPLATE for a in roles.get("value", [])):
        print(f"OK   roles: {', '.join(role_names)}")
    else:
        ok = False
        print(f"FAIL roles: Global Reader not assigned to '{sp['displayName']}' (has: {', '.join(role_names) or 'none'})")

    try:
        r = requests.post(
            f"https://outlook.office365.com/adminapi/beta/{t['tenantId']}/InvokeCommand",
            headers={"Authorization": f"Bearer {get_token(t, 'exchange')}",
                     "X-AnchorMailbox": f"UPN:{EXO_SYSTEM_MAILBOX}@{t['domain']}"},
            json={"CmdletInput": {"CmdletName": "Get-OrganizationConfig", "Parameters": {}}}, timeout=60)
        if r.ok:
            print(f"OK   exchange: Get-OrganizationConfig ({r.json()['value'][0].get('Name')})")
        else:
            ok = False
            print(f"FAIL exchange: HTTP {r.status_code} (role assignments can take 30-60 min to reach Exchange)")
    except RuntimeError as e:
        ok = False
        print(f"FAIL exchange token: {e}")

    reg = load_registry()
    reg[t["domain"]].update(name=org["displayName"], status="active" if ok else "partial",
                            verifiedAt=datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"))
    save_registry(reg)


def cmd_list():
    for domain, t in sorted(load_registry().items()):
        print(f"{t['status']:<16} {domain:<40} {t['tenantId']}  {t['name']}")


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args or args[0] not in ("onboard", "verify", "token", "list"):
        sys.exit(__doc__)
    if args[0] == "onboard":
        cmd_onboard(args[1])
    elif args[0] == "verify":
        cmd_verify(args[1])
    elif args[0] == "token":
        print(get_token(tenant(args[1]), args[2] if len(args) > 2 else "graph"))
    else:
        cmd_list()
