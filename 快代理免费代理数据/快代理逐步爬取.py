import requests_html
import time
import random
from saveexcel.saveitem import ExcelSaver


class Spider:
    def __init__(self):
        self.url = "http://www.kuaidaili.com/free/inha/1/"  # 快代理免费ip网址
        self.session = requests_html.HTMLSession()
        self.max_page = None  # 最大页数
        self.excel_headers = None  # 表头
        self.proxies = {}
        self.excel_writer = None

    def single_request(self):
        while True:
            try:
                response = self.session.get(self.url, proxies=self.proxies, timeout=5)
                if response.status_code == 200:  # 响应是否成功
                    break  # 成功则跳出循环
            except Exception as e:
                print(e)
                print("." * 30 + "IP已失效，正在重新获取代理IP" + "." * 30)
                self.get_proxies()  # 获取代理IP
        return response

    def get_response(self):
        response = self.single_request()
        self.parse_first_response(response)
        for page in range(2, int(self.max_page[0]) + 1):
            self.url = f"https://www.kuaidaili.com/free/inha/{page}/"
            time.sleep(random.random())
            response = self.single_request()
            self.parse_response(response)

    def get_proxies(self):
        url = ('http://api.xdaili.cn/xdaili-api//greatRecharge/getGreatIp?spiderId=afea3e76c783499aaca8aa3ddd61'
               'edf7&orderno=YZ20235100936AbAo3g&returnType=2&count=1')
        while True:
            try:
                response = self.session.get(url).json()
                print(response)
                ip = response['RESULT'][0]['ip']
                port = response['RESULT'][0]['port']
            except Exception as e:
                print(e)
                print("." * 30 + "代理IP获取失败，正在重新获取代理IP" + "." * 30)
                time.sleep(5)  # 休眠5秒,防止请求过快,导致失败
            else:
                self.proxies = {
                    "http": f"http://{ip}:{port}"
                }
                print("=" * 30 + "代理IP获取成功" + "=" * 30)
                break

    def parse_first_response(self, response):
        self.excel_headers = response.html.xpath("//table/thead/tr/th/text()")  # 获取表头
        self.max_page = response.html.xpath("//div[@id='list']/div[last()-1]/ul/li[last()-2]/a/text()")  # 获取最大页数
        self.parse_response(response)

    def parse_response(self, response):
        data = []
        for tr in response.html.xpath("//table/tbody/tr"):
            data.append(tr.xpath("//td//text()"))
        self.save_data(data)

    def save_data(self, data):
        self.excel_writer = ExcelSaver(column_headers=self.excel_headers, output_filename="快代理免费ip.xls",
                                       max_rows_per_sheet=5000
                                       )
        self.excel_writer.save_data_item(data)

    def scape(self):
        self.get_response()


if __name__ == '__main__':
    spider = Spider()
    spider.scape()
