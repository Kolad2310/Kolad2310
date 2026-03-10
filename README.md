```
import os
import time
import logging
import requests
from urllib.parse import urlparse

# ---------- Logging Configuration ----------
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(LOG_DIR, "media_download.log"),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def download_media(media_links, parent_folder="media", retries=3, delay=5):
    """
    Download media files from URLs.

    Parameters
    ----------
    media_links : list
        List of media URLs (Zoom recordings, video links etc.)

    parent_folder : str
        Folder where files will be downloaded

    retries : int
        Number of retry attempts

    delay : int
        Seconds to wait between retries

    Returns
    -------
    list
        List of downloaded file paths
    """

    os.makedirs(parent_folder, exist_ok=True)

    downloaded_files = []

    for link in media_links:

        attempt = 0
        success = False

        while attempt < retries and not success:

            try:
                logging.info(f"Attempt {attempt+1} downloading {link}")

                response = requests.get(link, stream=True, timeout=60)

                if response.status_code == 200:

                    parsed = urlparse(link)
                    filename = os.path.basename(parsed.path)

                    if not filename:
                        filename = f"media_{int(time.time())}.mp4"

                    filepath = os.path.join(parent_folder, filename)

                    with open(filepath, "wb") as file:
                        for chunk in response.iter_content(chunk_size=8192):
                            if chunk:
                                file.write(chunk)

                    logging.info(f"Download successful: {filepath}")
                    downloaded_files.append(filepath)
                    success = True

                else:
                    logging.warning(f"Failed with status {response.status_code} for {link}")
                    raise Exception("Download failed")

            except Exception as e:

                attempt += 1
                logging.error(f"Error downloading {link}: {str(e)}")

                if attempt < retries:
                    logging.info(f"Retrying in {delay} seconds...")
                    time.sleep(delay)

        if not success:
            logging.error(f"All retries failed for {link}")

    return downloaded_files


if __name__ == "__main__":

    # Example run
    sample_links = [
        "https://example.com/video1.mp4",
        "https://example.com/video2.mp4"
    ]

    result = download_media(sample_links)

    print("Downloaded files:")
    for r in result:
        print(r)
