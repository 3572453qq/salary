import requests
import json

# 企业微信应用的信息
corpID = "wwc75be524bd50ea62"
agentID = "1000019"
corpSecret = "9AuHHCNnwf8in7ZE2UcxDy3hSIzwfRH-q5o-W-tKpjU"

# 个人用户的 UserID
# toUser = "ChenXidd"
toUser = "HuangChao"
# toUser = "pengxuerou"
toUser = "huangchao"
# 发送消息的内容
messageContent = "工资条测试信息, 您本月实发工资300万"

# 获取 Access Token
token_url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpID}&corpsecret={corpSecret}"
print('token_url:',token_url)
token_response = requests.get(token_url)
access_token = token_response.json().get("access_token", "")
print(access_token)

# 发送文本消息的 API
message_api_url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"

markdown_table = """
## 您本月工资明细如下：
 > 应发：<font color='info'>100000</font>
 >> 基本工资：50000
 >> 绩效工资：50000

 > 代扣：<font color='warning'>90000</font>
 >> 社保：80000
 >> 个人所得税：10000
 
 > 实发：<font color='info'>10000</font>
"""

# 构建消息内容
message_content = {
    "touser": toUser,
    "msgtype": "markdown",
    "agentid": agentID,
    "markdown": {
        "content": markdown_table,
    }
}

# 发送消息
# response = requests.post(message_api_url, data=json.dumps(message_content), headers={"Content-Type": "application/json"})
response = requests.post(message_api_url, data=json.dumps(message_content, ensure_ascii=False).encode('utf-8'))
result = response.json()




# 输出发送结果
print(result)