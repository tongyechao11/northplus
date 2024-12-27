import os
import requests
from bs4 import BeautifulSoup
import openpyxl

# 获取当前脚本所在的目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 生成 Excel 文件的路径，保存到当前目录
excel_path = os.path.join(current_dir, '网页标题.xlsx')

# 尝试打开现有的 Excel 文件，如果文件不存在，则创建一个新的
if os.path.exists(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
else:
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "网页标题"
    sheet.append(["网址链接", "标题"])  # 添加表头

# 设置列宽度
sheet.column_dimensions['A'].width = 50  # 设置A列（网址链接）的宽度
sheet.column_dimensions['B'].width = 50  # 设置B列（标题）的宽度

# 设置起始值和终止值（这里只爬取一个页面）
start_number = 2410800
end_number = 2410900

# 循环遍历数字，构建网址并抓取标题
for number in range(start_number, end_number + 1):
    url = f'https://north-plus.net/read.php?tid-{number}.html'  # 动态生成网址
    try:
        # 发送请求获取网页内容（不使用 headers 和 cookies）
        response = requests.get(url)

        if response.status_code == 200:
            # 解析网页内容
            soup = BeautifulSoup(response.text, 'html.parser')

            # 查找 id="subject_tpc" 的元素
            subject_tpc_element = soup.find(id="subject_tpc")

            if subject_tpc_element:
                # 获取该元素的文本内容
                content = subject_tpc_element.get_text()

                # 判断内容是否包含 "[三次元R18相关]"
                if "[三次元R18相关]" in content:
                    # 提取网页标题
                    title = soup.title.string if soup.title else "无标题"

                    # 只取 '|' 前面的部分
                    if "|" in title:
                        title = title.split("|")[0].strip()  # 使用 split() 切分并去掉多余的空格

                    # 输出当前网址和标题
                    print(f"抓取的网址: {url}")
                    print(f"网页标题: {title}")
                    print('-' * 50)  # 分隔符，便于阅读

                    # 将网址和标题写入 Excel 文件
                    sheet.append([url, title])
                else:
                    print(f"该页面不包含 '[三次元R18相关]'，跳过网址: {url}")
            else:
                print(f"页面没有找到 id='subject_tpc' 元素，跳过网址: {url}")
        else:
            print(f"无法访问 {url}, 状态码: {response.status_code}")

    except Exception as e:
        print(f"抓取 {url} 时出错: {e}")

# 保存 Excel 文件
wb.save(excel_path)
print(f"抓取完成，数据已保存到 '{excel_path}'")

# 打开生成的 Excel 文件
os.startfile(excel_path)  # 在 Windows 系统中，使用默认应用程序打开文件
