import requests

DEFAULT_BASE_URL = "https://api.rcsaero.com"
MAX_PAGES = 50


class RcsAeroError(Exception):
    pass


def fetch_all_items(api_id, api_key, base_url=DEFAULT_BASE_URL, last_modified_dt="1980-01-01T00:00:00"):
    headers = {"x-api-id": api_id, "x-api-key": api_key}
    items = []
    page = 1
    while page <= MAX_PAGES:
        try:
            resp = requests.get(
                f"{base_url.rstrip('/')}/api/items",
                headers=headers,
                params={"last_modified_dt": last_modified_dt, "page": page},
                timeout=30,
            )
        except requests.RequestException as e:
            raise RcsAeroError(f"Could not reach {base_url}: {e}")

        if resp.status_code == 401 or resp.status_code == 403:
            raise RcsAeroError("Authentication failed — check the API ID and API Key for this company.")
        if resp.status_code != 200:
            raise RcsAeroError(f"rcsaero API returned HTTP {resp.status_code}: {resp.text[:300]}")

        data = resp.json()
        if not isinstance(data, list) or not data:
            break
        items.extend(data)
        page += 1

    return items
