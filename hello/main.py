from flask import Flask, request, jsonify, render_template
import requests
import time
import openpyxl
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from apscheduler.schedulers.background import BackgroundScheduler
import uuid  # 导入uuid模块

app = Flask(__name__)
scheduler = BackgroundScheduler()
scheduler.start()

def load_config(config_file):
    config = {}
    with open(config_file, 'r') as file:
        lines = file.readlines()
        for line in lines:
            key, value = line.strip().split('=')
            config[key] = value
    return config

def search_places(api_key, query, region, fetched_results, limit):
    url = f"https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {
        'key': api_key,
        'query': query,
        'region': region
    }
    results = []
    while len(fetched_results) < limit:
        response = requests.get(url, params=params)
        data = response.json()
        if 'error_message' in data:
            print(f"Error: {data['error_message']}")
            break
        fetched_data = data.get('results', [])
        for result in fetched_data:
            result['country'] = region  # Add country information to each result
        fetched_results.extend(fetched_data)
        results.extend(fetched_data)
        if 'next_page_token' not in data or len(fetched_results) >= limit:
            break
        params['pagetoken'] = data['next_page_token']
        time.sleep(3)  # 增加延迟，确保API有时间准备下一页数据
    print(f"Total fetched results: {len(results)}")
    return results

def filter_places_with_phone(api_key, places, limit):
    filtered_results = []
    for place in places:
        if len(filtered_results) >= limit:
            break
        place_id = place['place_id']
        details = get_place_details(api_key, place_id)
        if 'formatted_phone_number' in details:
            details['country'] = place['country']  # Preserve country information
            filtered_results.append(details)
    print(f"Filtered results with phone numbers: {len(filtered_results)}")
    return filtered_results

def get_place_details(api_key, place_id):
    url = f"https://maps.googleapis.com/maps/api/place/details/json"
    fields = "name,formatted_address,formatted_phone_number,international_phone_number,website,rating,price_level,business_status,place_id,icon"
    params = {
        'key': api_key,
        'place_id': place_id,
        'fields': fields
    }
    response = requests.get(url, params=params)
    return response.json().get('result', {})

def save_to_excel(results, query, keyword):
    print("Original Results: ", results)  # 打印原始查询结果
    
    # 创建Excel文件
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"

    # 写入标题
    headers = [
        '关键词', '国家', 'business_status', 'formatted_address', 'formatted_phone_number', 'icon',
        'international_phone_number', 'name', 'place_id', 'rating', 'website'
    ]
    ws.append(headers)

    # 写入数据
    for result in results:
        row = [
            keyword,  # 关键词
            result.get('country', ''),  # 国家
            str(result.get('business_status', '')),
            str(result.get('formatted_address', '')),
            str(result.get('formatted_phone_number', '')),
            str(result.get('icon', '')),
            str(result.get('international_phone_number', '')),
            str(result.get('name', '')),
            str(result.get('place_id', '')),
            str(result.get('rating', '')),
            str(result.get('website', ''))
        ]
        ws.append(row)
    
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    unique_id = uuid.uuid4()  # 生成一个随机UUID
    if not os.path.exists('data'):
        os.makedirs('data')
    filename = f"data/{query}_{timestamp}_{unique_id}.xlsx"  # 将UUID加入到文件名中
    wb.save(filename)
    return filename

def send_email(smtp_server, smtp_port, smtp_user, smtp_password, to_email, subject, body, attachment):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    attachment_name = os.path.basename(attachment)
    with open(attachment, "rb") as attach_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attach_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', attachment_name))
        msg.attach(part)

    try:
        print(f"Connecting to SMTP server {smtp_server} on port {smtp_port} using SSL")
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())
        server.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

def process_query(keywords, countries, email, config):
    api_key = config['api_key']
    limit = int(config['limit'])
    smtp_server = config['smtp_server']
    smtp_port = int(config['smtp_port'])
    smtp_user = config['smtp_user']
    smtp_password = config['smtp_password']

    global_fetched_results = []
    global_filtered_results = []

    for keyword in keywords:
        for country in countries:
            if len(global_filtered_results) >= limit:
                break
            fetched_country_results = []
            while len(global_filtered_results) < limit:
                search_results = search_places(api_key, keyword.strip(), country.strip(), fetched_country_results, limit)
                if not search_results:
                    print(f"No more results found for keyword '{keyword}' in country '{country}'")
                    break
                filtered_results = filter_places_with_phone(api_key, search_results, limit - len(global_filtered_results))
                global_filtered_results.extend(filtered_results)
                if len(global_filtered_results) >= limit:
                    break

    filename = save_to_excel(global_filtered_results, "_".join(keywords), keyword)
    if filename:
        email_subject = f"{','.join(keywords)}的查询结果"
        send_email(smtp_server, smtp_port, smtp_user, smtp_password, email, 
                   email_subject, "请查看附件中的查询结果。", filename)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/query', methods=['POST'])
def query():
    data = request.json
    keywords = data['keywords'].split(',')
    countries = data['countries']
    email = data['email']

    config = load_config('config.txt')

    # 使用后台任务处理查询
    scheduler.add_job(func=process_query, args=[keywords, countries, email, config], trigger='date')

    return jsonify({'success': True})

if __name__ == '__main__':
    app.run(debug=True)
