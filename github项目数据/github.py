import time
import re
import traceback
from requests import session
from saveexcel.saveitem import ExcelSaver


class Spider:
    def __init__(self, base_url, *, headers=None, proxy=None, cookies=None, proxy_url=None):
        self.base_url = base_url
        self.session = session()
        self.headers = headers
        self.initial_proxy = proxy
        self.proxy = proxy
        self.cookies = cookies
        self.proxy_url = proxy_url
        self.max_page = 100
        self.excel_headers = ["id", "项目名称", "项目地址", "项目描述", "项目标签", "项目语言", "项目stars"]

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
                print("ip失效，正在拉取代理ip")
                if self.proxy_url:
                    self._update_proxy()
                response = _make_request(self.proxy)
                if response:
                    return response

            # If all 5 tries failed, revert to the initial proxy
            self.proxy = self.initial_proxy
            print("Reverting to initial proxy")
            print("代理ip失效，回退到电脑本机ip")
            outer_retry_count += 1

        raise Exception(f"Reached maximum retry attempts for URL: {url} with params:{params}.")

    def parse_first(self, response):
        r = response.json()
        self.max_page = r["payload"]["page_count"]
        data = self.parse(response)
        return data

    @staticmethod
    def parse(response):
        data = []
        r = response.json()
        for item in r["payload"]["results"]:
            id = item["id"]
            name = re.sub(r'<.*?>|-', '', item["hl_name"])
            site = "https://github.com/" + name
            description = re.sub(r'<.*?>', '', item["hl_trunc_description"]) if item["hl_trunc_description"] else "无"
            topics = str(item["topics"]) if item["topics"] else "无"
            language = item["language"] if item["language"] else "无"
            stars = item["followers"]
            data.append([id, name, site, description, topics, language, stars])
        return data

    def _update_proxy(self):
        while True:
            try:
                response = self.session.get(self.proxy_url).json()
                print(response)
                ip = response['RESULT'][0]['ip']
                port = response['RESULT'][0]['port']
            except Exception as e:
                print(f"Error fetching proxy. Error: {e}")
                traceback.print_exc()
                print("代理ip获取失败，5秒后重试")
                time.sleep(5)
            else:
                self.proxy = dict(http=f"http://{ip}:{port}", https=f"http://{ip}:{port}")
                print("代理ip获取成功")
                break

    def run(self):
        key = input("请输入要搜索的项目关键字：")
        s = int(input("请输入排序方式:\n1.默认排序\n2.stars\n3.forks\n"))
        guides = ["stars", "forks"]
        file_name = f"{key}项目.xlsx"
        excel_saver = ExcelSaver(self.excel_headers, output_filename=file_name, max_rows_per_sheet=5000)
        if s == 1:
            query = dict(q=key)
        else:
            query = dict(q=key, s=guides[s - 2], o="desc")
        try:
            response = self.fetch(self.base_url, params=query)
            data = self.parse_first(response)
            excel_saver.save_data_item(data)
            for page in range(2, self.max_page + 1):
                query["p"] = page
                response = self.fetch(self.base_url, params=query)
                data = self.parse(response)
                excel_saver.save_data_item(data)
        except Exception as e:
            print(f"异常: {e}")
            traceback.print_exc()
            print("代理更换次数已达上限,下载中止")
        else:
            print("下载完成")
        finally:
            excel_saver.close()
            input("按任意键退出")


if __name__ == "__main__":
    spider = Spider("https://github.com/search?&type=repositories",
                    headers={"Accept": "application/json",
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
                                       "--E7dO3wOr2OuiLVq6--5KJ%2B0xXY1u6vB%2Fban4spRA%3D%3D",
                             "User-Agent": 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/603.3.8 ('
                                           'KHTML, like Gecko) Version/10.1.2 Safari/603.3.8'},
                    proxy_url="http://api.xdaili.cn/xdaili-api//greatRecharge/getGreatIp?spiderId"
                              "=afea3e76c783499aaca8aa3ddd61edf7&orderno=YZ20239150188SDxj02&returnType=2&count=1")
    spider.run()
