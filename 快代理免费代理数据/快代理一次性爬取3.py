import time
from saveexcel.saveitems import ExcelWriter
from requests_html import HTMLSession


class Spider:
    def __init__(self, base_url, *, headers=None, proxy=None, cookies=None, proxy_url=None):
        self.data = []
        self.max_page = None
        self.excel_headers = None
        self.initial_proxy = proxy
        self.base_url = base_url
        self.session = HTMLSession()
        self.headers = headers
        self.proxy = proxy
        self.cookies = cookies
        self.proxy_url = proxy_url

    def fetch(self, url, max_outer_retries=3):
        """Makes an HTTP request and returns the HTML response."""

        def _make_request(proxies):
            """Helper function to make a request using provided proxy settings."""
            try:
                r = self.session.get(
                    url,
                    headers=self.headers,
                    proxies=proxies,
                    cookies=self.cookies
                )
                r.raise_for_status()
                return r
            except Exception as e:
                print(f"Error fetching {url} using {self.proxy}. Error: {e}")
                return None

        outer_retry_count = 0

        while outer_retry_count <= max_outer_retries:
            # Try fetching with the current proxy
            response = _make_request(self.proxy)
            if response:
                return response

            # If failed, try updating the proxy and fetching the response 5 times
            for _ in range(5):
                self._update_proxy()
                response = _make_request(self.proxy)
                if response:
                    return response

            # If all 5 tries failed, revert to the initial proxy
            self.proxy = self.initial_proxy
            outer_retry_count += 1

        raise Exception(f"Reached maximum retry attempts for URL: {url}.")

    def parse_first(self, response):
        self.excel_headers = response.html.xpath("//table/thead/tr/th/text()")  # 获取表头
        self.max_page = int(response.html.xpath("//div[@id='list']/div[last()-1]/ul/li[last()-2]/a/text()")[0])  # 获取最大页数
        self.parse(response)

    def parse(self, response):
        for tr in response.html.xpath("//table/tbody/tr"):
            data = tr.xpath("//td//text()")
            self.data.append(data)

    def save(self, data):
        excel_writer = ExcelWriter(column_headers=self.excel_headers, output_filename="快代理爬取结果.xlsx")
        excel_writer.save_data(data)

    def _update_proxy(self):
        """Updates the proxy."""
        while True:
            try:
                response = self.session.get(self.proxy_url).json()
                print(response)
                ip = response['RESULT'][0]['ip']
                port = response['RESULT'][0]['port']
            except Exception as e:
                print(f"Error fetching proxy. Error: {e}")
                print("." * 30 + "Attempting to update proxy" + "." * 30)
                time.sleep(5)
            else:
                self.proxy = dict(http=f"http://{ip}:{port}", https=f"https://{ip}:{port}")
                print("=" * 30 + "Proxy updated" + "=" * 30)
                break

    def run(self):
        start_url = self.base_url.format(1)
        print(f"正在请求第{1}页数据")
        try:
            response = self.fetch(start_url)
            self.parse_first(response)
        except Exception as e:
            print(e)
            print("程序出错，正在重新运行")
            self.run()
        for page in range(2, self.max_page + 1):
            url = self.base_url.format(page)
            print(f"正在请求第{page}页数据")
            while True:
                try:
                    response = self.fetch(url)
                    self.parse(response)
                    break
                except Exception as e:
                    print(f"q请求第{page}页数据出错，正在重新请求: {url},Error: {e}")
        self.save(self.data)


if __name__ == "__main__":

    spider = Spider("http://www.kuaidaili.com/free/inha/{}/", proxy_url="http://api.xdaili.cn/xdaili-api//greatRecharge/getGreatIp?spiderId=afea3e76c783499aaca8aa3ddd61edf7&orderno=YZ20239150188SDxj02&returnType=2&count=1")
    spider.run()
