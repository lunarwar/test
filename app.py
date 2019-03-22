from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from flask_cors import CORS
import PDTRobot_new
import SDTRobot_new
import owsBot_INCI
import PTRobot_new
import TT_UPDATE_D
import datetime
import WORKORDER
import cmdb_project
import selenium
from selenium import webdriver

app = Flask(__name__)
CORS(app)
app.secret_key = '!@#$%^^7'

# use this only if the directory does not exist
# os.makedirs(os.path.join(app.instance_path, 'htmlfi'), exist_ok=True)


class Dbhandles(object):
    def __init__(self, user, pwd, hst, db):
        self.user = user
        self.password = pwd
        self.host = hst
        self.database = db


def days_between(d1, d2):
    d1 = datetime.datetime.strptime(d1, "%Y-%m-%d")
    d2 = datetime.datetime.strptime(d2, "%Y-%m-%d")
    return abs((d2 - d1).days)

@app.route('/index')
def index():
    if session.get('login_status') is True:
        msgbody = []
        return render_template('owsobot.html', data=msgbody, lenght=len(msgbody), username=session.get('username'))
    else:
        return redirect(url_for('default_page'))


@app.route('/')
def default_page():
    return render_template('login.html')


@app.route('/default', methods=['POST', 'GET'])
def signin():
        session['username'] = request.form['username']
        session['password'] = request.form['password']
        session['login_status'] = True
        if session['password'] != '' and session['username'] != '':
            return redirect(url_for('index'))
        else:
            return redirect(url_for('default_page'))


@app.route('/pdt_data', methods=['POST', 'GET'])
def pdt_data():
    filtr = request.form['pdt_filter']
    startdate = request.form['pdt_end_date']
    enddate = request.form['pdt_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status =  PDTRobot_new.weboperations(pwd,usrname,startdate, enddate, filtr)
    if status == "operation successfull":
        return redirect(url_for('index'))


@app.route('/sdt_data', methods=['POST', 'GET'])
def sdt_data():
    filtr = request.form['sdt_filter']
    startdate = request.form['sdt_end_date']
    enddate = request.form['sdt_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status =  SDTRobot_new.weboperations(pwd, usrname, startdate, enddate, filtr,"")
    if status == "operation successfull":
        return redirect(url_for('index'))
    


@app.route('/pt_data', methods=['POST', 'GET'])
def pt_data():
    filtr = request.form['pt_filter']
    startdate = request.form['pt_end_date']
    enddate = request.form['pt_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status =  PTRobot_new.weboperations(pwd,usrname,startdate, enddate, filtr)
    if status == "operation successfull":
        return redirect(url_for('index'))


@app.route('/tt_data', methods=['POST', 'GET'])
def tt_data():
    filtr = request.form['tt_filter']
    startdate = request.form['tt_start_date']
    enddate = request.form['tt_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status =  TT_UPDATE_D.weboperations(pwd,usrname,startdate, enddate, filtr)
    if status == "operation successfull":
        return redirect(url_for('index'))

@app.route('/cmdb', methods=['POST', 'GET'])
def cmdb():
    #filtr = request.form['tt_filter']
    #startdate = request.form['tt_start_date']
    #enddate = request.form['tt_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status =  cmdb_project.weboperations(pwd,usrname)
    if status == "operation successfull":
        return redirect(url_for('index'))

@app.route('/workd', methods=['POST', 'GET'])
def workd():
    filtr = request.form['wo_filter']
    startdate = request.form['wo_start_date']
    enddate = request.form['wo_end_date']
    pwd = session.get('password')
    usrname = session.get('username')
    status = WORKORDER.weboperations(pwd,usrname,startdate, enddate, filtr)
    if status == "operation successfull":
        return redirect(url_for('index'))


@app.route('/logout')
def logout():
    session.pop('login_status', None)
    return redirect(url_for('default_page'))




if __name__ == '__main__':
    app.run(debug=True, port=45574)

