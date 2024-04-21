from data_extract import pub_data_dict
from data_extract import pub_valid_flag

import requests
import time
from xml.sax.saxutils import escape
import json

jss_url = str(pub_data_dict['JAMF'][0])
bearer_token = str(pub_data_dict['JAMF'][1])

def get_last_user_id():
        global max_id
        id_url = jss_url + f"/JSSResource/users"
        headers = {"accept": "application/json","Authorization": f"Bearer {bearer_token}",}
        response = requests.get(id_url, headers=headers)
        data = json.loads(response.content)
        ids = [user['id'] for user in data['users']]
        max_id = max(ids)

get_last_user_id()

last_user_id = max_id + 1

for i in range(len(pub_data_dict['CN_Name'])):
    time.sleep(0.5)
    row_data = {header: values[i] for header, values in pub_data_dict.items()}

    stu_loginID = row_data['Login_ID']
    # stu_relName = row_data['CN_Name']
    stu_email_email = row_data['StuEmail']
    stu_form = row_data['Form']

    # 构建请求体
    request_body = f"""
    <user>
        <name>{stu_loginID}</name>
        <full_name>{stu_loginID}</full_name>
        <email>{stu_email_email}</email>
        <email_address>{stu_email_email}</email_address>
        <position>{stu_form}</position>
        <sites>
            <site>
                <id>-1</id>
                <name>None</name>
            </site>
        </sites>
    </user>
    """

    # 构建请求头
    headers = {
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": f"Bearer {bearer_token}",
    }

    user_id = last_user_id + i

    # 发起请求
    url = jss_url + "/JSSResource/users/id/" + f"user_id"
    response = requests.post(url, data=request_body, headers=headers)

    # 处理响应
    if response.status_code == 201:
        print(f"用户 {stu_loginID} 创建成功!")
    else:
        print(f"用户 {stu_loginID} 创建失败，错误代码: {response.status_code}")
        # print(response.text)
