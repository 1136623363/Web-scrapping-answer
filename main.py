import threading
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'}
params = {'batch': '1', }


def getPrompts(languages):
    url = 'https://flowgpt.com/api/trpc/prompt.getPrompts'
    data = []
    for language in languages:
        with requests.Session() as s:
            acount = 0
            while True:
                lang = '中文' if language == 'zh' else '英文'
                print(f'开始爬取{lang}Prompts第{acount + 1}页')
                params[
                    'input'] = f'{{"0":{{"json":{{"skip":{acount * 36},"tag":null,"sort":null,"q":null,"language":"{language}"}},' \
                               f'"meta":{{"values":{{"tag":["undefined"],"sort":["undefined"],"q":["undefined"]}}}}}}}}'

                try:
                    response = s.get(url=url, headers=headers, params=params, timeout=3)
                except Exception:
                    print('爬取失败\n请检查网络是否通畅或者连接代理后再尝试')

                json_ = response.json()[0]['result']['data']['json']

                if len(json_) and acount < 20:  # 限制20页
                    data.extend(json_)
                    acount += 1
                    # break
                else:
                    break
    print(f'总共爬取到{len(data)}条Prompt')

    return data


def getComment(i):
    url = 'https://flowgpt.com/api/trpc/comment.getComments'
    params['input'] = f'{{"0": {{"json": "{i["id"]}"}}}}'

    response = requests.get(url, headers=headers, params=params)
    comment = [i["body"] for i in response.json()[0]["result"]["data"]["json"]]
    # 将列表连接成字符串，并添加序号
    result = '\n'.join([f'{i + 1}. {item}' for i, item in enumerate(comment)])

    i.update({'comment': result})


def getComments_multi(data):
    print('\n开始爬取评论')
    threads = []

    for i in data:
        if i['comments']:
            t = threading.Thread(target=getComment, args=(i,))
            threads.append(t)
            t.start()

    # 等待所有线程完成
    for t in threads:
        t.join()


def save_to_xlsx(data):
    # 创建一个新的工作簿
    workbook = Workbook()

    # 获取默认的工作表
    sheet = workbook.active

    # 定义要保存的字段名称
    fieldnames = ['title', 'description', 'uses', 'initPrompt', 'Tag', 'comment']
    # 'id', 'createdAt', 'updatedAt', 'title', 'description', 'saves', 'userId', 'thumbnailURL', 'upvotes', 'initPrompt',
    # 'live', 'accessibility', 'visibility', 'popularity', 'views', 'conversationId', 'editedAt', 'language', 'uses',
    # 'ranking', 'model', 'createdMethod', 'systemMessage', 'type', 'adminWeight', 'welcomeMessage', 'temperature',
    # 'impressions', 'shares', 'comments', 'cup', 'fop', 'rankingForNew', 'tip', 'trendingScore', 'Tag', 'User'

    # 写入字段名称
    sheet.append(fieldnames)

    # 写入数据行
    for item in data:
        row = [item.get(field, '') for field in fieldnames]

        # 将列表转换为字符串
        for i in range(len(row)):
            if isinstance(row[i], list):
                row[i] = ', '.join([tag['name'] for tag in row[i]])

        sheet.append(row)

    # 设置列宽和居中
    for column in sheet.columns:
        column_letter = column[0].column_letter
        if column_letter == 'C':  # 第三列
            sheet.column_dimensions[column_letter].width = 10
            for cell in column:
                cell.alignment = Alignment(horizontal='center')  # 居中
        else:
            sheet.column_dimensions[column_letter].width = 40

    try:
        # 保存工作簿到文件
        filename = 'Prompts.xlsx'
        workbook.save(filename)
        print('\n数据已成功保存到Prompts.xlsx')
    except PermissionError:
        input('\n保存失败\n请尝试关闭正在浏览xlsx文件的窗口\n并按回车重新保存\n')
        save_to_xlsx(data)


if __name__ == '__main__':
    data = getPrompts(['zh', 'en']) #爬取 中文、英文 提示词
    getComments_multi(data)
    save_to_xlsx(data)
