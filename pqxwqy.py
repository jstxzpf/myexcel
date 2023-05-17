import requests
from bs4 import BeautifulSoup
def get_titles_on_page(url):
    # 发送请求获取网页内容
    response = requests.get(url)
    # 使用BeautifulSoup解析网页
    soup = BeautifulSoup(response.text, 'html.parser')
    # 获取所有标题标签
    titles = soup.find_all('a', class_='f14')
    # 遍历所有标题，如果包含“小微企业”字样，就将其写入文件
    with open('XWQY.txt', 'a', encoding='utf-8') as f:
        for title in titles:
            if '小微企业' in title.get_text():
                f.write(title.get_text() + '\n')
def crawl_pages(url, depth=2):
    # 如果已经达到指定深度，则返回
    if depth == 0:
        return
    # 发送请求获取网页内容
    response = requests.get(url)
    # 使用BeautifulSoup解析网页
    soup = BeautifulSoup(response.text, 'html.parser')
    # 获取当前页面上所有带有“小微企业”字样的标题并保存到文件中
    get_titles_on_page(url)
    # 获取下一页的URL
    next_page = soup.find('a', class_='next')
    if next_page:
        next_url = 'http://www.stats.gov.cn' + next_page['href']
        # 递归调用该函数，深度减1
        crawl_pages(next_url, depth-1)
crawl_pages('http://www.stats.gov.cn/search/s?qt=%E5%B0%8F%E5%BE%AE%E4%BC%81%E4%B8%9A', depth=2)
