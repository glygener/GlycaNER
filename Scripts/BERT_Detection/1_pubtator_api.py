import requests
import os
import time

STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")

# Paths inside the container
pmid_file_path = os.environ.get("GLYCANER_PMID_FILE", "/data/input/pmids.txt")
output_dir = os.path.join(STAGE1_DIR, 'abstract_xml')

os.makedirs(output_dir, exist_ok=True)

with open(pmid_file_path, 'r') as file:
    pmids = [line.strip() for line in file if line.strip()]

base_url = "https://www.ncbi.nlm.nih.gov/research/pubtator3-api/publications/export/biocxml?pmids="

MAX_ATTEMPTS = 6      # per PMID
PAUSE = 0.4           # keeps us under the 3 requests/sec limit
failed_log = os.path.join(output_dir, "failed_pmids.txt")

session = requests.Session()
CONTACT_EMAIL = os.environ.get("GLYCANER_CONTACT_EMAIL", "")
user_agent = f"GlycaNER/1.0 (mailto:{CONTACT_EMAIL})" if CONTACT_EMAIL else "GlycaNER/1.0"
session.headers.update({"User-Agent": user_agent})

done = skipped = failed = 0

for idx, pmid in enumerate(pmids, 1):
    output_file = os.path.join(output_dir, f"{pmid}.xml")

    # Resume: anything already downloaded is left alone
    if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
        skipped += 1
        continue

    url = f"{base_url}{pmid}"
    saved = False

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = session.get(url, timeout=120)
        except requests.exceptions.RequestException as e:
            wait = min(2 ** attempt, 60)
            print(f"[{idx}/{len(pmids)}] PMID {pmid}: connection error ({e}). "
                  f"Retry {attempt}/{MAX_ATTEMPTS} in {wait}s", flush=True)
            time.sleep(wait)
            continue

        if response.status_code == 200:
            with open(output_file, "wb") as fh:
                fh.write(response.content)
            print(f"[{idx}/{len(pmids)}] Saved PMID {pmid}", flush=True)
            saved = True
            break

        elif response.status_code == 429:
            retry_after = int(response.headers.get('Retry-After', 2 ** attempt))
            print(f"[{idx}/{len(pmids)}] PMID {pmid}: rate limited, waiting {retry_after}s", flush=True)
            time.sleep(retry_after)

        elif response.status_code in (500, 502, 503, 504):
            wait = min(2 ** attempt, 60)
            print(f"[{idx}/{len(pmids)}] PMID {pmid}: HTTP {response.status_code}, "
                  f"retry {attempt}/{MAX_ATTEMPTS} in {wait}s", flush=True)
            time.sleep(wait)

        else:
            print(f"[{idx}/{len(pmids)}] PMID {pmid}: HTTP {response.status_code}, not retrying", flush=True)
            break

    if saved:
        done += 1
    else:
        failed += 1
        with open(failed_log, "a") as fh:
            fh.write(f"{pmid}\n")

    time.sleep(PAUSE)

print(f"\nDownloaded {done}, already present {skipped}, failed {failed} of {len(pmids)} PMIDs", flush=True)
if failed:
    print(f"Failed PMIDs listed in {failed_log}. Rerun the container to retry only those.", flush=True)