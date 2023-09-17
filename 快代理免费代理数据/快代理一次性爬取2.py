import time
from saveexcel.saveitems import ExcelWriter
from requests_html import HTMLSession


class Spider:
    def __init__(self, base_url, *, headers=None, proxy=None, cookies=None, proxy_url=None):
        self.data = []
        self.max_page = None
        self.excel_headers = None
        self.base_url = base_url
        self.session = HTMLSession()
        self.headers = headers
        self.proxy = proxy
        self.cookies = cookies
        self.proxy_url = proxy_url

    def fetch(self, url):
        """Makes an HTTP request and returns the HTML response."""
        while True:
            try:
                response = self.session.get(
                    url,
                    headers=self.headers,
                    proxies=self.proxy,
                    cookies=self.cookies
                )
                response.raise_for_status()
                return response
            except Exception as e:
                print("." * 30 + f"Error fetching {url}. Error: {e}" + "." * 30)
                if self.proxy_url:
                    self._update_proxy()

    def parse_first(self, response):
        self.excel_headers = response.html.xpath("//table/thead/tr/th/text()")  # 获取表头
        self.max_page = int(response.html.xpath("//div[@id='list']/div[last()-1]/ul/li[last()-2]/a/text()")[0])  # 获取最大页数
        self.parse(response)

    def parse(self, response):
        for tr in response.html.xpath("//table/tbody/tr"):
            data = tr.xpath("//td//text()")
            self.data.append(data)

    def save(self, data):
        excel_writer = ExcelWriter(column_headers=self.headers, output_filename="快代理爬取结果.xlsx")
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
                self.proxy = dict(http=f"http://{ip}:{port}", https=f"http://{ip}:{port}")
                print("=" * 30 + "Proxy updated" + "=" * 30)
                break

    def run(self):
        start_url = self.base_url.format(1)
        print("=" * 30 + f"正在请求第{1}页数据" + "=" * 30)
        response = self.fetch(start_url)
        self.parse_first(response)
        for page in range(2, self.max_page + 1):
            url = self.base_url.format(page)
            print("=" * 30 + f"正在请求第{page}页数据" + "=" * 30)
            response = self.fetch(url)
            self.parse(response)
        self.save(self.data)


if __name__ == "__main__":

    spider = Spider("http://www.kuaidaili.com/free/inha/{}/")
    spider.run()
