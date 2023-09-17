import requests
import json


class TencentSpider:
    def __init__(self):
        self.urls = [
            ("https://careers.tencent.com/tencentcareer/api/post/Query?"
             "&pageIndex={}&pageSize=200&language=zh-cn").format(i) for i in range(1, 21)
        ]
        self.headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) "
                                      "AppleWebKit/603.3.8 (KHTML, like Gecko) Version/10.1.2 Safari/603.3.8"}
        self.file_name = "tencent.json"
        self.response = None

    def get_response(self, url):
        self.response = requests.get(url, headers=self.headers)

    def parse_response(self):
        try:
            posts = self.response.json()["Data"]["Posts"]
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

    def save_to_json(self, data):
        # 读取已有的JSON文件
        try:
            with open(self.file_name, "r", encoding="utf-8") as file:
                data_list = json.load(file)
        except (FileNotFoundError, json.JSONDecodeError):
            data_list = []
        # 将新数据添加到列表中
        data_list.append(data)
        # 将更新后的列表写回JSON文件
        with open(self.file_name, "w", encoding="utf-8") as file:
            json.dump(data_list, file, ensure_ascii=False, indent=2)

    def scrape(self):
        for url in self.urls:
            self.get_response(url)
            posts = list(self.parse_response())
            if not posts:  # 检查是否收到特殊标记
                print("=" * 30 + "所有数据已经爬取完毕!" + "=" * 30)
                break
            for post in posts:
                self.save_to_json(post)
                print(post["RecruitPostNam"] + "=" * 30 + "保存成功")


if __name__ == "__main__":
    spider = TencentSpider()
    spider.scrape()
