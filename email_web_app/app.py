import os
import json
import smtplib
import imaplib
import email
import markdown
import bleach
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import decode_header
from email import encoders
from flask import Flask, render_template, request, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = 'a_very_secret_key_for_flash_messaging_12345'
CONFIG_FILE = 'config.json'

# --- 核心邏輯 ---
def load_config():
    if not os.path.exists(CONFIG_FILE): return None
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return None

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

def find_server_config(username, config):
    for server in config.get('smtp_servers', []):
        if server.get('username') == username:
            return server
    return None

def send_email(server_config, to_addr, subject, plain_text_body, html_body, attachments=None, sender_nickname=None):
    smtp_server = server_config['host']
    smtp_port = server_config['port']
    username = server_config['username']
    password = server_config['password']
    msg = MIMEMultipart('mixed')
    if sender_nickname and sender_nickname.strip():
        msg['From'] = f"{sender_nickname.strip()} <{username}>"
    else:
        msg['From'] = username
    msg['To'] = to_addr
    msg['Subject'] = subject
    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    part_plain = MIMEText(plain_text_body, 'plain', 'utf-8')
    part_html = MIMEText(html_body, 'html', 'utf-8')
    msg_alternative.attach(part_plain)
    msg_alternative.attach(part_html)
    if attachments:
        for f in attachments:
            if f and f.filename:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{f.filename}"')
                msg.attach(part)
    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=15)
        server.starttls()
        server.login(username, password)
        server.sendmail(username, to_addr.split(','), msg.as_string())
        server.quit()
        return True, f"從 {username} 發送成功！"
    except Exception as e:
        return False, f"從 {username} 發送失敗，錯誤: {e}"

# --- Flask 路由 ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send', methods=['GET', 'POST'])
def send_form():
    config = load_config()
    if request.method == 'POST':
        server_config = find_server_config(request.form['sender'], config)
        if not server_config:
            flash('未找到發件服務器配置。', 'error')
            return redirect(url_for('send_form'))

        raw_body = request.form['body']
        body_format = request.form['body_format']
        if body_format == 'plain':
            plain_text_body = raw_body
            html_body = f"<p>{raw_body.replace(chr(10), '<br>')}</p>"
        elif body_format == 'html':
            plain_text_body = "（此郵件為HTML格式，請在支持HTML的客戶端中查看）"
            html_body = raw_body
        else:
            plain_text_body = raw_body
            html_body = markdown.markdown(raw_body, extensions=['fenced_code', 'tables'])

        attachments = request.files.getlist('attachments')
        
        success, message = send_email(
            server_config=server_config,
            to_addr=request.form['to_addr'],
            subject=request.form['subject'],
            plain_text_body=plain_text_body,
            html_body=html_body,
            attachments=attachments,
            sender_nickname=request.form['nickname']
        )
        if success:
            flash(message, 'success')
        else:
            flash(message, 'error')
        return redirect(url_for('send_form'))
        
    return render_template('send_email.html', config=config)

@app.route('/fetch', methods=['GET', 'POST'])
def fetch_form():
    config = load_config()
    if request.method == 'POST':
        account_username = request.form['account']
        server_config = find_server_config(account_username, config)
        if not server_config:
            flash('未找到郵箱配置。', 'error')
            return redirect(url_for('fetch_form'))
        
        try:
            mail = imaplib.IMAP4_SSL(server_config['host'], 993)
            mail.login(server_config['username'], server_config['password'])
            mail.select('inbox')
            status, uids = mail.uid('search', None, 'UNSEEN')
            
            emails_list = []
            if uids[0]:
                email_uids = uids[0].split()
                for email_uid in reversed(email_uids[-20:]):
                    status, msg_data = mail.uid('fetch', email_uid, '(RFC822)')
                    msg = email.message_from_bytes(msg_data[0][1])
                    subject, encoding = decode_header(msg['Subject'])[0]
                    if isinstance(subject, bytes): subject = subject.decode(encoding if encoding else 'utf-8', 'ignore')
                    
                    emails_list.append({
                        'uid': email_uid.decode(),
                        'from': msg.get('From'),
                        'subject': subject,
                        'date': msg.get('Date')
                    })
            mail.logout()
            return render_template('fetch_emails.html', config=config, emails=emails_list, account_checked=account_username)
        except Exception as e:
            flash(f"獲取郵件失敗: {e}", 'error')
            return redirect(url_for('fetch_form'))

    return render_template('fetch_emails.html', config=config, emails=None)

@app.route('/view_email/<account_username>/<email_uid>')
def view_email(account_username, email_uid):
    config = load_config()
    server_config = find_server_config(account_username, config)
    if not server_config:
        flash('未找到郵箱配置。', 'error')
        return redirect(url_for('fetch_form'))

    email_data = {'from': '', 'to': '', 'subject': '', 'date': '', 'body': ''}
    try:
        mail = imaplib.IMAP4_SSL(server_config['host'], 993)
        mail.login(server_config['username'], server_config['password'])
        mail.select('inbox')
        
        status, msg_data = mail.uid('fetch', email_uid.encode(), '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])

        subject, encoding = decode_header(msg['Subject'])[0]
        if isinstance(subject, bytes): subject = subject.decode(encoding if encoding else 'utf-8', 'ignore')

        email_data['subject'] = subject
        email_data['from'] = msg.get('From')
        email_data['to'] = msg.get('To')
        email_data['date'] = msg.get('Date')

        body_html = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    body_html = part.get_payload(decode=True).decode(errors='ignore')
                    break
        elif msg.get_content_type() == 'text/html':
            body_html = msg.get_payload(decode=True).decode(errors='ignore')
        
        if not body_html: # Fallback to plain text if no HTML part
             for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    plain_body = part.get_payload(decode=True).decode(errors='ignore')
                    body_html = f"<pre>{plain_body}</pre>" # Wrap plain text in <pre>
                    break

        allowed_tags = bleach.sanitizer.ALLOWED_TAGS | {'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'br', 'hr', 'pre', 'code', 'img', 'table', 'thead', 'tbody', 'tr', 'td', 'th'}
        allowed_attrs = {
            '*': ['style', 'class'],
            'a': ['href', 'title'],
            'img': ['src', 'alt', 'style']
        }
        email_data['body'] = bleach.clean(body_html, tags=allowed_tags, attributes=allowed_attrs)
        mail.logout()
    except Exception as e:
        flash(f"讀取郵件失敗: {e}", 'error')

    return render_template('view_email.html', email=email_data)

@app.route('/manage', methods=['GET'])
def manage_configs():
    config = load_config()
    return render_template('manage_configs.html', config=config)

@app.route('/manage/add_server', methods=['POST'])
def add_server():
    config = load_config()
    new_server = {
        'host': request.form['host'],
        'port': int(request.form['port']),
        'username': request.form['username'],
        'password': request.form['password']
    }
    config['smtp_servers'].append(new_server)
    save_config(config)
    flash('發件服務器添加成功！', 'success')
    return redirect(url_for('manage_configs'))

@app.route('/manage/delete_server', methods=['POST'])
def delete_server():
    config = load_config()
    username_to_delete = request.form['username']
    config['smtp_servers'] = [s for s in config['smtp_servers'] if s['username'] != username_to_delete]
    save_config(config)
    flash('發件服務器刪除成功！', 'success')
    return redirect(url_for('manage_configs'))

@app.route('/manage/add_contact', methods=['POST'])
def add_contact():
    config = load_config()
    config['contacts'][request.form['name']] = request.form['email']
    save_config(config)
    flash('聯絡人添加成功！', 'success')
    return redirect(url_for('manage_configs'))

@app.route('/manage/delete_contact', methods=['POST'])
def delete_contact():
    config = load_config()
    name_to_delete = request.form['name']
    if name_to_delete in config['contacts']:
        del config['contacts'][name_to_delete]
    save_config(config)
    flash('聯絡人刪除成功！', 'success')
    return redirect(url_for('manage_configs'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
