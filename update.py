import subprocess, sys, traceback, requests, json, os, threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

URLS = {
    "Win10": "https://go.microsoft.com/fwlink/?LinkId=841361",
    "Win11": "https://go.microsoft.com/fwlink/?LinkId=2156292",
}

HEADERS = {"Connection": "Keep-Alive", "User-Agent": "Windows Dlp Manager"}
JSON_FILE, BASE_DIR, MAX_WORKERS = "locations.json", "products", 5
print_lock = threading.Lock()


def load_existing_data():
    if os.path.exists(JSON_FILE):
        try:
            with open(JSON_FILE, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            with print_lock:
                print(
                    f"Warning: {JSON_FILE} contains invalid JSON. Starting with empty data."
                )
    return {}


def save_data(data):
    with open(JSON_FILE, "w") as f:
        json.dump(data, f, indent=4)


def archive_to_wayback(url):
    try:
        with print_lock:
            print(f"Archiving to Wayback: {url}")
        requests.get(
            f"http://web.archive.org/save/{url}", timeout=30
        ).raise_for_status()
    except requests.RequestException:
        with print_lock:
            print(f"Failed to archive {url} to Wayback Machine")
            traceback.print_exc()


def download_cab(url, output_path):
    try:
        with print_lock:
            print(f"Downloading CAB from: {url}")
        with requests.get(url, stream=True, headers=HEADERS, timeout=60) as response:
            response.raise_for_status()
            with open(output_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
        return True
    except requests.RequestException as e:
        with print_lock:
            print(f"Error downloading {url}: {e}")
        return False


def extract_cab_windows(cab_path, output_dir):
    try:
        output_file = os.path.join(output_dir, "products.xml")
        subprocess.run(
            ["expand", cab_path, "-f:products.xml", output_file],
            check=True,
            capture_output=True,
            text=True,
            timeout=300,
        )
        return True
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
        with print_lock:
            print(f"Error extracting {cab_path} on Windows: {e}")
            if hasattr(e, "stderr") and e.stderr:
                print(f"stderr: {e.stderr}")
        return False


def extract_cab_unix(cab_path, output_dir):
    try:
        # Use tar to extract only products.xml
        subprocess.run(
            ["cabextract", "-d",output_dir, cab_path],
            check=True,
            capture_output=True,
            text=True,
            timeout=300,
        )
        return True
    except (
        subprocess.CalledProcessError,
        FileNotFoundError,
        subprocess.TimeoutExpired,
    ) as e:
        with print_lock:
            print(f"Error extracting {cab_path} on unix: {e}")
            if hasattr(e, "stderr") and e.stderr:
                print(f"stderr: {e.stderr}")


def process_url(args):
    os_name, url = args
    try:
        path_suffix = "_".join(url.split('/')[-5:])
        if path_suffix.endswith(".cab"):
            path_suffix = path_suffix[:-4]

        output_dir = os.path.join(BASE_DIR, os_name, path_suffix)
        os.makedirs(output_dir, exist_ok=True)
        cab_path = os.path.join(output_dir, "archive.cab")

        # archive_to_wayback(url)  # Uncomment if needed

        if not download_cab(url, cab_path):
            return False

        success = (
            extract_cab_windows(cab_path, output_dir)
            if sys.platform.startswith("win")
            else extract_cab_unix(cab_path, output_dir)
        )

        if success and os.path.exists(cab_path):
            os.remove(cab_path)

        return success

    except Exception as e:
        with print_lock:
            print(f"Error processing {url}: {e}")
            traceback.print_exc()
        return False


def main():
    data = load_existing_data()

    for os_name, url in URLS.items():
        try:
            response = requests.head(
                url, headers=HEADERS, allow_redirects=False, timeout=30
            )
            if location := response.headers.get("Location"):
                if os_name not in data:
                    data[os_name] = {}
                if location not in data[os_name]:
                    data[os_name][location] = datetime.now().isoformat()
                    with print_lock:
                        print(f"Added new URL for {os_name}: {location}")
            else:
                with print_lock:
                    print(f"No Location header found for {os_name}.")
        except requests.RequestException as e:
            with print_lock:
                print(f"Error fetching {os_name} URL: {e}")

    save_data(data)

    url_list = [
        (os_name, url) for os_name, urls_dict in data.items() for url in urls_dict
    ]

    with print_lock:
        print(f"Processing {len(url_list)} URLs with {MAX_WORKERS} workers")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        results = list(executor.map(process_url, url_list))

    successful = sum(results)
    with print_lock:
        print(f"Processing complete. Successful: {successful}/{len(url_list)}")


if __name__ == "__main__":
    main()
