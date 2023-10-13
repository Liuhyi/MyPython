import traceback
from requests_html import HTMLSession
from saveexcel.saveitem import ExcelSaver
from concurrent.futures import ThreadPoolExecutor


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
        self.max_page = 20
        self.excel_headers = [
            "PostId",
            "RecruitPostId",
            "RecruitPostNam",
            "CountryName",
            "LocationName",
            "CategoryName",
            "Responsibility",
            "LastUpdateTime",
            "PostURL",
            "SourceID",
            "IsCollect",
            "IsValid",
        ]
        self.excel_saver = None

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

    def parse_first(self, response):
        max_page = int(
            response.html.xpath(
                "//div[@class='application-main']//a[@href='#100']/text()"
            )[0]
        )
        self.max_page = max_page
        data = self.parse(response)
        return data

    @staticmethod
    def parse(response):
        try:
            posts = response.json()["Data"]["Posts"]
            # 检查posts是否为null或空列表
            if not posts:
                return None  # 返回None作为特殊标记
            for job in posts:
                post = dict(
                    PostId=job["PostId"],
                    RecruitPostId=job["RecruitPostId"],
                    RecruitPostNam=job["RecruitPostName"],
                    CountryName=job["CountryName"],
                    LocationName=job["LocationName"],
                    CategoryName=job["CategoryName"],
                    Responsibility=job["Responsibility"],
                    LastUpdateTime=job["LastUpdateTime"],
                    PostURL=job["PostURL"],
                    SourceID=job["SourceID"],
                    IsCollect=job["IsCollect"],
                    IsValid=job["IsValid"],
                )
                yield post
        except TypeError:
            print("=" * 30 + "解析错误" + "=" * 30)

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

    def task(self, query):
        response = self.fetch(self.base_url, params=query)
        data = self.parse(response)
        return data

    def run(self):
        file_name = "腾讯岗位数据.xlsx"
        self.excel_saver = ExcelSaver(
            column_headers=self.excel_headers,
            output_filename=file_name,
            max_rows_per_sheet=5000,
        )
        futures = []
        try:
            with ThreadPoolExecutor(max_workers=10) as executor:
                for page in range(1, self.max_page + 1):
                    query = dict(pageSize=200, pageIndex=page)
                    future = executor.submit(self.task, query)
                    futures.append(future)
                for future in futures:
                    datas = future.result()
                    if datas:
                        for data in datas:
                            data = list(data.values())
                            self.excel_saver.save_data_item(data)
        except Exception as e:
            print(e)
            traceback.print_exc()
        finally:
            self.excel_saver.close()


if __name__ == "__main__":
    spider = Spider(
        "https://careers.tencent.com/tencentcareer/api/post/Query?&language=zh-cn"
    )
    spider.run()
