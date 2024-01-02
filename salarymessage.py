import requests
import json
import pandas as pd
import re
import time
import configparser
from pypinyin import pinyin, Style

def get_pinyin(word):
    # 将汉字转换为拼音，默认使用带声调的拼音风格
    pinyin_result = pinyin(word, style=Style.NORMAL)

    # 将列表中的拼音连接成字符串
    pinyin_str = ''.join([item[0] for item in pinyin_result])

    return pinyin_str

def send_wechat_message(corpID, agentID, corpSecret, toUser, markdown_table):
    # 获取 Access Token
    token_url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpID}&corpsecret={corpSecret}"
    token_response = requests.get(token_url)
    access_token = token_response.json().get("access_token", "")

    # 发送文本消息的 API
    message_api_url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"

    

    message_content = {
        "touser": toUser,
        "msgtype": "markdown",
        "agentid": agentID,
        "markdown": {
            "content": markdown_table,
        }
    }

    # 发送消息
    response = requests.post(message_api_url, data=json.dumps(message_content, ensure_ascii=False).encode('utf-8'))
    result = response.json()

    # 输出发送结果
    print(result)

def processfile(excel_file,yearmonth):    
    # 读取 Excel 文件，指定两行表头
    # df = pd.read_excel('D:\\hc\\xsj\\program\\python\\salary\\深圳欣视界- 2023年11月工资单.xlsx',skiprows=1, header=[0, 1])
    df = pd.read_excel(excel_file,skiprows=1, header=None)

    # 获取前两行作为表头
     
    header_rows = df.iloc[:2]
    header_rows = header_rows.fillna(method='ffill', axis=0)
    header_rows = header_rows.fillna(method='ffill', axis=1)
    print(header_rows)
    df.iloc[:2]=header_rows
    df=df.fillna(0)
   


    markdown_messages = []
    all_names = []
    # 遍历每一行数据
    for rowindex, row in df.iterrows():
        # 遍历每一列数据
        markdown_table = f"## 您{yearmonth}工资明细如下："
        preivous_level1 = 'empty'
        if rowindex < 2 or row[1]==0 or row[1].strip()=='合计：':
            continue
        is_level2=0
        for colindex,value in enumerate(row):
            df.iloc[0,colindex] = df.iloc[0,colindex].replace("\n", "")
            df.iloc[1,colindex] = df.iloc[1,colindex].replace("\n", "")
            col_message=''
            if df.iloc[0,colindex]=='序号':
                continue
            if preivous_level1!=df.iloc[0,colindex]:
                preivous_level1 = df.iloc[0,colindex]
                if is_level2==1:
                    col_message = '\n'+col_message
                    is_level2=0
                col_message = col_message + f'\n> ### {df.iloc[0,colindex]}:'                
            if df.iloc[0,colindex]==df.iloc[1,colindex]:
                col_message = col_message + f"{value}"
                # col_message = col_message + f"<font color='warning'>{value}</font>"                
            else:
                col_message = col_message + f"\n>> {df.iloc[1,colindex]}:{value}"
                # col_message = col_message + f"\n>> {df.iloc[1,colindex]}:<font color='warning'>{value}</font>"
                is_level2=1
            if df.iloc[1,colindex]=='姓名':
                all_names.append(value)
            markdown_table += col_message

        # print(markdown_table)
        # print(rowindex,'*'*100)
        markdown_messages.append(markdown_table)
    return markdown_messages,all_names
        

def main():
    
    # 创建 ConfigParser 对象
    config = configparser.ConfigParser()

    # 读取配置文件
    config.read('settings.ini')

    # 获取配置信息
    corpID = config.get('corp', 'corpID')
    agentID = config.get('corp', 'agentID')
    corpSecret = config.get('corp', 'corpSecret')


    # 读取JSON文件
    with open('nameid.json', 'r',encoding='utf-8') as file:
        namelist = json.load(file)
    
    print(namelist.keys())

    toUser= input('请输入您的企业微信ID（通常是姓名的全拼）：')
    eachUser=toUser
    
    # 用户输入文件名
    filename = input('请输入需要发送通知的工资单文件名：')
    pattern = r"\d{4}年\d+月"
    yearmonth=re.findall(pattern, filename)[0]


    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    all_messages,all_names = processfile(filename,yearmonth)
    for message in all_messages:
        send_wechat_message(corpID, agentID, corpSecret, toUser, message)
        time.sleep(2.3)
    confirm_ok=input(f'所有信息均已发送到{toUser}的企业微信，请检查一下内容是否正确，如果正确，则输入"Y"：')
    if confirm_ok=='Y' or confirm_ok=='y':
        confirm_send=input(f'您已确认信息内容正确，下面将会把工资条信息单独发送给每个人，确认请输入"Y"：')
    else:
        print('请检查excel文件内容后再次尝试，谢谢')
        exit(0)
    if confirm_send=='Y' or confirm_send=='y':
        for message_i,message in enumerate(all_messages):
            if all_names[message_i] in namelist.keys():
                eachUser = namelist[all_names[message_i] ]
            else:
                eachUser = get_pinyin(all_names[message_i])            
            print(message_i,eachUser)
            send_wechat_message(corpID, agentID, corpSecret, toUser, message)
            time.sleep(2.3)
        print(f'{yearmonth}工资条发送完毕，一共{len(all_messages)}条，谢谢！')
    else:
        print(f'{yearmonth}工资条未发送，谢谢！')

if __name__ == "__main__":
    main()  # 在程序执行时调用主函数
