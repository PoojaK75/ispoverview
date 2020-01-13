import os, csv
import win32com.client as win32
from win32com.client import Dispatch
from flask import Flask, render_template, request, redirect

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/project-single_SM.html')
def team_SM():
    return render_template('project-single_SM.html')

@app.route('/project-single_MIBO.html')
def team_MIBO():
    return render_template('project-single_MIBO.html')

@app.route('/project-single_Acc.html')
def team_ACC():
    return render_template('project-single_Acc.html')

@app.route('/project-single_Unix.html')
def team_UNIX():
    return render_template('project-single_Unix.html')

@app.route('/project-single_DBA.html')
def team_DBA():
    return render_template('project-single_DBA.html')

@app.route('/project-single_Ops.html')
def team_OPS():
    return render_template('project-single_Ops.html')

@app.route('/contact.html')
def contact():
    return render_template('contact.html')

@app.route('/<string:page_name>')
def html_page(page_name):
    return render_template(page_name)

def write_to_file(data):
    with open('DB.txt', mode='a') as database:
        name=data["name"]
        phone=data["phone"]
        email=data["email"]
        message=data["message"]
        file=database.write(f'\n{name},{phone},{email},{message}')
        
def write_to_csv(data):
    with open('DB2.csv', newline='', mode='a') as database2:
        name=data["name"]
        phone=data["phone"]
        email=data["email"]
        message=data["message"]
        csv_writer=csv.writer(database2, delimiter=',',quotechar='|',quoting=csv.QUOTE_MINIMAL)
        csv_writer.writerow([name,phone,email,message])
        
def send_mail():
    if request.method=='POST':
        EMailIDs=open('C:\\Users\\PO323206\\Downloads\\python scripts\\web server\\templates\\emails.txt').read()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)        
        mail.To = EMailIDs[0]
        mail.Cc = EMailIDs[1]
        mail.Subject = "New Message from ISP Web Page"
        mail.HTMLBody = """<html><left><b><i><font size=8 color="#cc00cc" face="Brush Script MT">Hi</i></b></font></left></html>"""
        mail.Send()
    
@app.route('/submit_form', methods=['POST', 'GET'])
def submit():
    if request.method=='POST':
        data=request.form.to_dict()
        write_to_csv(data)
        return redirect('/thankyou.html')  