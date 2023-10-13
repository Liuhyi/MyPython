import os
import re
import threading
import traceback
from requests_html import HTMLSession
import urllib3
import asyncio
from functools import partial


class Spider:
    def __init__(
        self, base_url, *, headers=None, proxy=None, cookies=None, proxy_url=None
    ):
        self.base_url = base_url
        self.session = HTMLSession()
        self.headers = headers
        self.initial_proxy = proxy
        self.proxy = proxy
        self.cookies = cookies
        self.proxy_url = proxy_url
        self.excel_saver = None
        self.counter = 0
        self.key = None
        self.numbers = None

    def fetch(self, url, max_outer_retries=10, params=None):
        """Makes an HTTP request and returns the HTML response."""

        def _make_request(proxies):
            """Helper function to make a request using provided proxy settings."""
            try:
                r = self.session.get(
                    url,
                    headers=self.headers,
                    proxies=proxies,
                    cookies=self.cookies,
                    params=params,
                    timeout=6,
                )
                r.raise_for_status()
                return r
            except Exception as e:
                print(
                    f"Error fetching {url} with params:{params} using {self.proxy}. Error: {e}"
                )
                traceback.print_exc()
                return None

        outer_retry_count = 0

        while outer_retry_count <= max_outer_retries:
            # Try fetching with the current proxy
            response = _make_request(self.proxy)
            if response:
                return response

            # If failed, try updating the proxy and fetching the response 5 times
            for _ in range(1):
                print(f"Error fetching {url} with params:{params} using {self.proxy}.")
                if self.proxy_url:
                    self._update_proxy()
                response = _make_request(self.proxy)
                if response:
                    return response

            # If all 5 tries failed, revert to the initial proxy
            self.proxy = self.initial_proxy
            print("Reverting to initial proxy")
            outer_retry_count += 1

        raise Exception(
            f"Reached maximum retry attempts for URL: {url} with params:{params}."
        )

    async def parse(self, response):
        """Parses the HTML response and saves the data to an Excel file."""
        try:
            data = response.json()
            tasks = []
            for item in data["data"][:-1]:
                loop = asyncio.get_event_loop()
                base_name = re.sub(r'[\\/:*?"<>|]', "",  item["fromPageTitleEnc"])
                if "setList" in item:
                    for sub_item in item["setList"]:
                        url = sub_item["objURL"]
                        task = asyncio.create_task(self.fetch_and_save(url, base_name))
                        tasks.append(task)
                else:
                    url = item["replaceUrl"][0]["ObjURL"].replace("\\", "")
                    urls = item["middleURL"]
                    task = asyncio.create_task(self.fetch_and_save(url, base_name, urls))
                    tasks.append(task)
            await asyncio.gather(*tasks)
        except Exception as e:
            print(f"Error parsing data. Error: {e}")
            traceback.print_exc()

    async def fetch_and_save(self, url, base_name, urls=None):
        loop = asyncio.get_event_loop()
        try:
            partial_func = partial(
                self.session.get,
                url,
                headers={"Referer": "https://image.baidu.com/"},
                timeout=6,
                verify=False
            )
            r = await loop.run_in_executor(None, partial_func)
            if r.headers["Content-Type"] == "image/gif":
                raise Exception("GIF")
            r.raise_for_status()
        except Exception as e:
            print(f"Error fetching {url}. Error: {e}")
            traceback.print_exc()
            url = urls
            partial_func = partial(
                self.session.get,
                url,
                headers={"Referer": "https://image.baidu.com/"},
                timeout=6,
                verify=False
            )
            r = await loop.run_in_executor(None, partial_func)
        partials = partial(self.save, r, f"{base_name}.jpg")
        await loop.run_in_executor(None, partials)

    def save(self, data, filename):
        with threading.Lock():
            self.counter += 1
            path = f"{self.key}/{self.counter}_{filename}"
            with open(path, "wb") as f:
                f.write(data.content)
            print(f"第{self.counter}张图片下载完成")
            if self.counter == self.numbers:
                print("下载完成")
                os._exit(0)

    def _update_proxy(self):
        count = 0
        while count <= 2:
            try:
                response = self.session.get(self.proxy_url).json()
                print(response)
                ip = response["RESULT"][0]["ip"]
                port = response["RESULT"][0]["port"]
            except Exception as e:
                print(f"Error fetching proxy. Error: {e}")
                print("." * 30 + "Attempting to update proxy" + "." * 30)
                traceback.print_exc()
                count += 1
            else:
                self.proxy = dict(
                    http=f"http://{ip}:{port}", https=f"http://{ip}:{port}"
                )
                print("=" * 30 + "Proxy updated" + "=" * 30)
                break

    def run(self):
        loop = asyncio.get_event_loop()
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.key = input("请输入图片关键字：")
        self.numbers = int(input("请输入图片数量："))
        if not os.path.exists(self.key):
            os.mkdir(self.key)
        each = 30
        page = 0
        while self.counter < self.numbers:
            query = dict(
                word=self.key,
                queryWord=self.key,
                pn=page * each,
                rn=each,
            )
            page += 1
            response = self.fetch(self.base_url, params=query)
            loop.run_until_complete(self.parse(response))


if __name__ == "__main__":
    spider = Spider("https://image.baidu.com/search/acjson?tn=resultjson_com&logid=10386623690106717899&ipn=rj&ct"
                    "=201326592&is=&fp=result&fr=ala&cl=2&lm=-1&ie=utf-8&oe=utf-8&adpicid=&st=&z=&ic=&hd=&latest"
                    "=&copyright=&s=&se=&tab=&width=&height"
                    "=&face=&istype=&qc=&nc=&expermode=&nojc=&isAsync=&gsm=1e&1696248884995=")
    spider.run()
