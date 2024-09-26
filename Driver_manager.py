import requests, os, zipfile, json
from win32com.client import Dispatch

class DriverManager:

    """A class to manage downloads of chromedrivers for Windows machines

    Attributes
    ----------
    chrome_path: str
        The path to chrome.exe
    driver_folder: str
        The path to store downloaded drivers
    proxies: dict
        A dict of proxies for http, https and ftp to use for api requests
    desired_version: str
        The version of chromedriver required
    stable_url: str
        URL to determine latest stable chromedriver
    """

    def __init__(self,
                 chrome_path: str = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
                 driver_folder: str = 'C:\\ChromeDriver',
                 proxies: dict = None,
                 desired_version: str = None,
                 stable_url: str = None):

        self.chrome_path = chrome_path
        self.driver_folder = driver_folder
        self.proxy = proxies
        self.desired_version = desired_version
        self.stable_url = 'https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json' if stable_url is None else stable_url
        self.version_urls = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'
        self.version_file = f"{self.driver_folder}\\version.json"

    def get_chrome_version(self) -> str:

        """Get the current version of Chrome"""

        parser = Dispatch('Scripting.FilesystemObject')

        if not os.path.isfile(self.chrome_path):
            raise FileNotFoundError(f"{self.chrome_path} does not exist")

        return parser.GetFileVersion(self.chrome_path)

    def get_stable_chromedriver_version(self) -> str:

        """Get the latest stable version of chromedriver"""

        with requests.get(self.stable_url, proxies=self.proxy) as response:

            if response.status_code != '200': response.raise_for_status()

            res_json = response.json()
            return res_json['channels']['Stable']['version']

    def _url_builder(self, driver_version:str) -> str:

        """Build API URL to for provided driver version"""

        return f'https://storage.googleapis.com/chrome-for-testing-public/{driver_version}/win32/chromedriver-win32.zip'

    def _folder_check(self):

        """Checks if a folder exists for the drivers and makes one if it does not exist"""

        if not os.path.isdir(self.driver_folder):
            os.mkdir(self.driver_folder)

    def check_driver_version(self) -> str:

        """Get the current version of chromedriver downloaded"""

        if not os.path.isfile(self.version_file): return '0'

        with open(self.version_file, 'r') as file:
            return json.loads(file.read())['version']

    def _get_driver_response(self, url:str) -> requests.Response:

        """Get response from provided URL"""

        with requests.get(url, proxies=self.proxy) as response:
            return response

    def _download_driver(self, url:str):

        """Downloads the driver from provided URL"""

        response = self._get_driver_response(url)

        if response.status_code != 200:
            print(f"Unable to download driver version {self.desired_version}")
            response.raise_for_status()

        self._folder_check()
        zipped_file = f"{self.driver_folder}\\chromedriver.zip"

        with open(zipped_file, 'wb') as file:
            for chunk in response.iter_content(chunk_size=128):
                file.write(chunk)

        with zipfile.ZipFile(zipped_file) as zf:
            zf.extractall(self.driver_folder)

        with open(self.version_file, 'w') as file:
            json.dump({'version': self.desired_version}, file)

    def _determine_desired_version(self):

        """Check if there is a driver for the current version of Chrome and uses the stable version if not"""

        current_driver = self.check_driver_version()
        chrome_version = self.get_chrome_version()
        stable_version = self.get_stable_chromedriver_version()

        url = self._url_builder(chrome_version)
        response = self._get_driver_response(url)
        if response.status_code == 200:
            self.desired_version = chrome_version
            return

        if current_driver == chrome_version or current_driver == stable_version:
            self.desired_version = current_driver

        if int(chrome_version.replace('.', '')) > int(stable_version.replace('.', '')):
            self.desired_version = stable_version
            return

        if int(chrome_version.replace('.', '')) < int(stable_version.replace('.', '')):
            raise Exception('Outdated Chrome', 'Chrome version should be updated', f'Chrome:{chrome_version}, Stable version:{stable_version}')

    def get_driver(self):

        """Get the desired driver"""

        if self.desired_version is None: self._determine_desired_version()

        if self.desired_version == self.check_driver_version(): return

        url = self._url_builder(self.desired_version)
        self._download_driver(url)
