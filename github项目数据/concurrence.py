import time
import traceback

from requests_html import HTMLSession
from saveexcel.saveitem import ExcelSaver
from concurrent.futures import ThreadPoolExecutor


class Spider:
    def __init__(self, base_url, *, headers=None, proxy=None, cookies=None, proxy_url=None):
        self.base_url = base_url
        self.session = HTMLSession()
        self.headers = headers
        self.initial_proxy = proxy
        self.proxy = proxy
        self.cookies = cookies
        self.proxy_url = proxy_url
        self.max_page = 100
        self.excel_headers = ["项目名称", "项目地址", "项目描述", "项目标签", "项目语言", "项目stars", "项目更新时间"]
        self.excel_saver = None

    def fetch(self, url, max_outer_retries=3, params=None):
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
                    timeout=6
                )
                r.raise_for_status()
                return r
            except Exception as e:
                print(f"Error fetching {url} with params:{params} using {self.proxy}. Error: {e}")
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

        raise Exception(f"Reached maximum retry attempts for URL: {url} with params:{params}.")

    def parse_first(self, response):
        max_page = int(response.html.xpath('''//div[@class='application-main']//nav[@aria-label="Pagination"]//a[last()-1]/text()''')[0])
        self.max_page = max_page
        data = self.parse(response)
        return data

    @staticmethod
    def parse(response):
        data = []
        for div in response.html.xpath("//div[@data-testid='results-list']/div"):
            title = "".join(div.xpath(".//h3//a//text()"))
            site = "https://github.com" + div.xpath(".//h3//a/@href")[0]
            description = div.xpath('.//div[@class="Box-sc-g0xbh4-0 LjnbQ"]//text()')
            description = [d.strip() for d in description]
            description = "".join(description)
            topics = div.xpath(".//div[@class = 'Box-sc-g0xbh4-0 frRVAS']//text()")
            topics = [topic.strip() for topic in topics]
            topics = str(topics) if topics else "无"
            language = div.xpath(".//ul/li[last()-2]//text()")
            language = language[0] if language else "无"
            stars = div.xpath(".//ul/li[last()-1]//text()")[0]
            update = div.xpath(".//ul/li[last()]//text()")
            update = [u.strip() for u in update]
            update = ' '.join(update)
            data.append([title, site, description, topics, language, stars, update])
        return data

    def _update_proxy(self):
        count = 0
        while count <= 2:
            try:
                response = self.session.get(self.proxy_url).json()
                print(response)
                ip = response['RESULT'][0]['ip']
                port = response['RESULT'][0]['port']
            except Exception as e:
                print(f"Error fetching proxy. Error: {e}")
                print("." * 30 + "Attempting to update proxy" + "." * 30)
                time.sleep(5)
                count += 1
            else:
                self.proxy = dict(http=f"http://{ip}:{port}", https=f"http://{ip}:{port}")
                print("=" * 30 + "Proxy updated" + "=" * 30)
                break

    def task(self, query):
        response = self.fetch(self.base_url, params=query)
        data = self.parse(response)
        return data

    def run(self):
        key = input("请输入要搜索的项目关键字：")
        s = input("请输入排序方式:\n输入stars按照stars排序,输入forks按照forks排序，输入updated按照更新时间排序：")
        file_name = f"{key}项目.xlsx"
        self.excel_saver = ExcelSaver(column_headers=self.excel_headers, output_filename=file_name,
                                      max_rows_per_sheet=5000)
        query = dict(q=key, s=s, o="desc")
        response = self.fetch(self.base_url, params=query)
        data = self.parse_first(response)
        self.excel_saver.save_data_item(data)
        futures = []
        try:
            with ThreadPoolExecutor(max_workers=10) as executor:
                for page in range(2, self.max_page + 1):
                    query = dict(q=key, s=s, o="desc", p=page)
                    future = executor.submit(self.task, query)
                    futures.append(future)
                for future in futures:
                    data = future.result()
                    self.excel_saver.save_data_item(data)
        except Exception as e:
            print(e)
            print(traceback.print_exc())
        finally:
            self.excel_saver.close()


if __name__ == "__main__":
    spider = Spider("https://github.com/search?&type=repositories", headers={"Accept": "text/html",
                    "Referer": "https://github.com/search?q=python&type=repositories",
                    "cookie": "_octo=GH1.1.179135992.1695187704; "
                               "_device_id=481314df80e638afb99b62b237270705; "
                               "user_session=9X647yWVgTE8cnShAMaJYwZahOZhBfMayqEs_jGdKoMC-AL5; "
                               "__Host-user_session_same_site=9X647yWVgTE8cnShAMaJYwZahOZhBfMayqEs_jGdKoMC"
                               "-AL5; logged_in=yes; dotcom_user=Liuhyi; has_recent_activity=1; "
                               "color_mode=%7B%22color_mode%22%3A%22auto%22%2C%22light_theme%22%3A%7B%22name"
                               "%22%3A%22light%22%2C%22color_mode%22%3A%22light%22%7D%2C%22dark_theme%22%3A"
                               "%7B%22name%22%3A%22dark%22%2C%22color_mode%22%3A%22dark%22%7D%7D; "
                               "preferred_color_mode=dark; tz=Asia%2FShanghai; "
                               "_gh_sess=FdYG2iQ9ifMPx%2Bkf41giJMEBNyBUJFBg1DK"
                               "%2BfTgn4yvWvmVyzhU6Z27Andd4Dck6RscHR5SqhuG47%2FNs%2BKGfAnxJubrtB4B3LARjlk"
                               "%2BGvcjRZQKVTTa%2FUALorVn%2BMFEbmKdvFZxPVE%2FGrpnZJS7OktoTUDZMR%2B%2FRLwm4NE0"
                               "%2F2nPJYozXHB%2BYAvOZwkfZlI%2FHvIm%2F7xeMUCjiA6jigrrxmSdBaODhrWGe8QSG2"
                               "%2FWIewlzR0cyfCNQ5jlPBCo3gZsJUpDM39yBdGqSVtZ5Tll6QAEzj0TvDDJUbfaGZYfu%2Bxe"
                               "%2BiCiS4ITJdUu7G%2BKVFbmkJl1Tq2c460SypQEBG7Xsvp%2BZh0SsMsn%2F"
                               "--E7dO3wOr2OuiLVq6--5KJ%2B0xXY1u6vB%2Fban4spRA%3D%3D"},
                    proxy_url="http://api.xdaili.cn/xdaili-api//greatRecharge/getGreatIp?spiderId=afea3e76c783499aaca8aa3ddd61edf7&orderno=YZ20239150188SDxj02&returnType=2&count=1")
    spider.run()
