from __future__ import annotations

import requests
from urllib.parse import quote

import urllib3


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def _get_pubchem_json(url: str, timeout: int):
    # Some organization proxies resign HTTPS certificates. PubChem CAS lookup
    # is read-only, so use verify=False to keep the GUI usable in those networks.
    resp = requests.get(url, timeout=timeout, verify=False)

    if resp.status_code != 200:
        raise RuntimeError(f"PubChem request failed. HTTP {resp.status_code}")

    return resp.json()


def _extract_first_property(data: dict):
    props = data.get("PropertyTable", {}).get("Properties", [])
    if not props:
        return None
    return props[0]


def _extract_smiles(prop: dict) -> str:
    if not prop:
        return ""
    for key in ("CanonicalSMILES", "ConnectivitySMILES", "IsomericSMILES", "SMILES"):
        value = prop.get(key)
        if value:
            return str(value).strip()
    for key, value in prop.items():
        if "SMILES" in str(key).upper() and value:
            return str(value).strip()
    return ""


def cas_to_smiles(cas: str, timeout: int = 15):
    """Search PubChem by CAS and return dict with CID and CanonicalSMILES.

    PubChem name search often resolves CAS RN. Internet connection is required.
    """
    cas = str(cas or "").strip()
    if not cas:
        raise ValueError("CAS is empty.")

    base = "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound"
    name = quote(cas, safe="")
    url = f"{base}/name/{name}/property/CanonicalSMILES,IsomericSMILES,ConnectivitySMILES/JSON"
    try:
        data = _get_pubchem_json(url, timeout=timeout)
    except Exception as e:
        raise RuntimeError(f"PubChem search failed for CAS {cas}: {e}") from e

    first = _extract_first_property(data)
    smiles = _extract_smiles(first)
    cid = first.get("CID", "") if first else ""

    if not smiles:
        cid_url = f"{base}/name/{name}/cids/JSON"
        cid_data = _get_pubchem_json(cid_url, timeout=timeout)
        cids = cid_data.get("IdentifierList", {}).get("CID", [])
        if cids:
            cid = cids[0]
            for prop_names in (
                "CanonicalSMILES,IsomericSMILES,ConnectivitySMILES",
                "IsomericSMILES",
                "CanonicalSMILES",
                "ConnectivitySMILES",
            ):
                prop_url = f"{base}/cid/{cid}/property/{prop_names}/JSON"
                try:
                    prop_data = _get_pubchem_json(prop_url, timeout=timeout)
                    first = _extract_first_property(prop_data)
                    smiles = _extract_smiles(first)
                    if smiles:
                        break
                except Exception:
                    continue

    if not first:
        raise RuntimeError(f"No PubChem result found for CAS {cas}.")
    if not smiles:
        raise RuntimeError(f"PubChem found CID {cid or first.get('CID', '')}, but no SMILES property was returned.")

    return {
        "CAS": cas,
        "PubChem_CID": cid or first.get("CID", ""),
        "CanonicalSMILES": smiles,
        "IsomericSMILES": first.get("IsomericSMILES", ""),
        "PubChem_status": "Found",
    }
