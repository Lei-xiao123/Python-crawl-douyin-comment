"""

[环境使用]:
    Python 3.10
    Pycharm

[模块使用]:
    DrissionPage (自动化)
    pandas
    jieba
    wordcloud
    pyecharts
    csv
---------------------------------------------------------------------------------------------------
win + R 输入cmd 输入安装命令 pip install 模块名 (如果你觉得安装速度比较慢, 你可以切换国内镜像源)

---------------------------------------------------------------------------------------------------


技术不断的更新迭代的
    - 自动化模块: selenium
        目前很多网站对于 selenium 有检测的 (淘宝 boss 前程无忧..)
    - 推荐: DrissionPage (以后可能也会被检测)


"""

# 导入自动化模块
from DrissionPage import ChromiumPage
# 导入时间模块
import datetime
# 导入csv模块
import csv
# 创建文件对象
f = open('data.csv', mode='w', encoding='utf-8-sig', newline='')
# 字典写入方法
csv_writer = csv.DictWriter(f, fieldnames=['昵称', '地区', '时间', '评论'])
# 写入表头
csv_writer.writeheader()
# 打开浏览器
driver = ChromiumPage()
# 监听数据包
driver.listen.start('aweme/v1/web/comment/list/')
# 访问网站
driver.get('https://www.douyin.com/?modal_id=7402920366357761316')
for page in range(30):
    print(f'正在采集第{page+1}页的数据内容')
    # 下滑页面到底部
    driver.scroll.to_bottom()
    # 等待数据包加载 -> 数据包没有加载 / 没有数据了 卡住
    resp = driver.listen.wait()
    # 直接获取数据包返回的响应数据
    json_data = resp.response.body
    # 解析数据, 提取评论数据所在列表 comments
    comments = json_data['comments']
    # for循环遍历, 提取列表里面元素
    for index in comments:
        # 键值对取值, 提取相关内容
        text = index['text'] # 评论内容
        nickname = index['user']['nickname'] # 昵称
        create_time = index['create_time'] # 评论时间戳
        # 把时间戳转成日期
        date = str(datetime.datetime.fromtimestamp(create_time))
        ip_label = index['ip_label'] # 地区
        # 把数据放到字典里面
        dit = {
            '昵称': nickname,
            '地区': ip_label,
            '时间': date,
            '评论': text,
        }
        # 写入数据
        csv_writer.writerow(dit)
        print(dit)

###################################################################
#################################################################
#
# # 导入数据处理
# import pandas as pd
# # 导入结巴分词
# import jieba
# # 导入词云模块
# import wordcloud
# import numpy as np
# from PIL import Image
# # 读取形状图片
# mask = np.array(Image.open('python.png'))
# # 读取文件
# df = pd.read_csv('data.csv')
# # 获取评论内容
# content = ' '.join([str(i).replace('\n', '') for i in df['评论']])
# # 结巴分词处理
# string = ' '.join(jieba.lcut(content))
# # 词云图配置
# wc = wordcloud.WordCloud(
#     font_path='msyh.ttc',  # 字体文件
#     width=1000,  # 宽
#     height=700, # 高
#     mask=mask, # 词云图形状
#     stopwords={'了', '的', '我', '你', '是', '都', '把', '能', '就', '这', '还'}
# )
# # 导入词汇内容
# wc.generate(string)
# # 导出词云图
# wc.to_file('cy2.png')
# print(string)


