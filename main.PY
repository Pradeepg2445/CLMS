from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtWidgets import QApplication, QWidget,QShortcut,QFrame, QDesktopWidget, QTableWidget, QTableWidgetItem, QGridLayout,QTableWidget, QDialog, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtGui import QPixmap, QCloseEvent,QKeySequence,QBrush,QColor
import sys
import sqlite3
import xlsxwriter
import pandas as pd
import os
from datetime import date
import time
from datetime import timedelta
import datetime 
import re
import smtplib
import ssl
from email.mime.text import MIMEText
import pygame
import requests
from plyer import notification
import random
from cryptography.fernet import Fernet
import shutil
from PyQt5.QtCore import Qt,QDate
import itertools




def load():
        
        global app
        global sign_up_win
        global login_win
        global main_win
        global today
        global login_status
        global PublicKey
        global PrivateKey
        global key
        global fernet
        key=b'oc7b6PWUvDNEWcdSaauMYDo-aBq0aLbKX5_7Kqk-eTQ='
        fernet=Fernet(key)
        app=QtWidgets.QApplication(sys.argv)
        login_win = uic.loadUi("application_ui/ui/login.ui")
        sign_up_win = uic.loadUi("application_ui/ui/sign_up.ui")
        main_win = uic.loadUi("application_ui/ui/main.ui")

        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        list_of_tables=c.execute(""" SELECT name FROM sqlite_master WHERE type='table' AND name='login_details'; """).fetchall()
        if list_of_tables==[]:
                
                login_details_table = """CREATE TABLE login_details (
                user_id INTEGER(1) PRIMARY KEY,
                user_name VARCHAR(64) NOT NULL,
                email VARCHAR(64) NOT NULL,
                password_1 VARCHAR(32) NOT NULL,
                password_2 VARCHAR(32) NOT NULL,
                themes VARCHAR(1) NOT NULL,
                sound VARCHAR(1) NOT NULL,
                student_login VARCHAR(1) NOT NULL,
                backup VARCHAR(1) NOT NULL,
                login_status VARCHAR(1) NOT NULL
                );"""       
                c.execute(login_details_table)
                try:
                        book_details_table = """CREATE TABLE book_details (
                book_id INTEGER PRIMARY KEY AUTOINCREMENT,
                book_title VARCHAR(256) NOT NULL,
                author VARCHAR(256) NOT NULL,
                isbn VARCHAR(16) NOT NULL,
                available INTEGER(1024) NOT NULL,
                quantity INTEGER(1024) NOT NULL,
                publisher VARCHAR(256),
                price VARCHAR(64),
                publication_date VARCHAR(16),
                tags VARCHAR(256),
                description VARCHAR(2048),
                total_issued INTEGER(1024)
                );"""
                        c.execute(book_details_table)
                         
                        course_details_table = """CREATE TABLE course_details (
                course_id INTEGER PRIMARY KEY AUTOINCREMENT,
                course_name VARCHAR(64) NOT NULL,
                total_year INTEGER(1) NOT NULL,
                description VARCHAR(2048)
                );"""
                        c.execute(course_details_table)
                         
                        batch_details_table = """CREATE TABLE batch_details (
                batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_year VARCHAR(4) NOT NULL,
                end_year VARCHAR(4) NOT NULL,
                course_id INTEGER,
                description VARCHAR(2048),
                FOREIGN KEY (course_id) REFERENCES course_details(course_id)
                );"""
                        c.execute(batch_details_table)
                        student_details_table ="""CREATE TABLE student_details(
                roll_number VARCHAR PRIMARY KEY,
                register_number VARCHAR(1024),
                name VARCHAR(128),
                date_of_birth VARCHAR(16),
                father_name VARCHAR(128),
                mother_name VARCHAR(128),
                phone VARCHAR(10),
                email VARCHAR(128),
                address VARCHAR(256),
                aadhar_number VARCHAR(15),
                religion VARCHAR(32),
                community VARCHAR(32),
                caste VARCHAR(256),
                anual_income VARCHAR(16),
                blood_group VARCHAR(8),
                father_occupation VARCHAR(256),
                mother_occupation VARCHAR(256),
                area_type VARCHAR(256),
                sslc_mark VARCHAR(8),
                sslc_passing_year VARCHAR(32),
                sslc_school_name VARCHAR(256),
                sslc_medium VARCHAR(32),
                sslc_school_type VARCHAR(64),
                hsc_mark VARCHAR(8),
                hsc_passing_year VARCHAR(32),
                hsc_school_name VARCHAR(256),
                hsc_medium VARCHAR(32),
                hsc_school_type VARCHAR(64),
                admission_date VARCHAR(16),
                description VARCHAR(2048),
                batch_id INTEGER,
                course_id INTEGER,

                FOREIGN KEY (course_id) REFERENCES course_details(course_id),
                FOREIGN KEY (batch_id) REFERENCES batch_details(batch_id)
                );"""
                        c.execute(student_details_table)
                         

                        
                except Exception as e:
                        print("here")
                        print(e)
                try:
                        book_quantity_details_table = """CREATE TABLE book_quantity_details (
                book_quantity_id INTEGER(1024) PRIMARY KEY,
                book_id INTEGER(1024),
                inserted_on VARCHAR(16),
                description VARCHAR(2048),
                FOREIGN KEY (book_id) REFERENCES book_details(book_id)
                );"""
                        c.execute(book_quantity_details_table)
                         

                        issued_book_details_table = """CREATE TABLE issued_book_details (
                issued_book_id VARCHAR(1024)PRIMARY KEY,
                roll_number VARCHAR(1024),
                book_quantity_id VARCHAR(1024),
                issue_date VARCHAR(12),
                last_date VARCHAR(12),
                book_id VARCHAR(1024),
                course_id VARCHAR(1024),
                batch_id VARCHAR(1024),
                FOREIGN KEY (roll_number) REFERENCES student_details(roll_number),
                FOREIGN KEY (book_quantity_id) REFERENCES book_quantity_details(book_quantity_id),
                FOREIGN KEY (book_id) REFERENCES book_details(book_id)
                );"""
                        c.execute(issued_book_details_table)
                         
                        book_history_table = """CREATE TABLE book_history (
                issued_book_id VARCHAR(1024),
                roll_number VARCHAR(1024),
                book_quantity_id VARCHAR(1024),
                issue_date VARCHAR(12),
                last_date VARCHAR(12),
                given_date VARCHAR(12),
                book_id VARCHAR(1024),
                course_id VARCHAR(1024),
                batch_id VARCHAR(1024),
                FOREIGN KEY (roll_number) REFERENCES student_details(roll_number),
                FOREIGN KEY (book_quantity_id) REFERENCES book_quantity_details(book_quantity_id)
                );"""
                        c.execute(book_history_table)
                        course_delete_trigger='''
CREATE TRIGGER IF NOT EXISTS course_delete_trigger
BEFORE DELETE ON course_details
BEGIN
DELETE FROM batch_details WHERE course_id = old.course_id ;
DELETE FROM student_details WHERE course_id = old.course_id ;
DELETE FROM book_history WHERE course_id = old.course_id ;
DELETE FROM issued_book_details WHERE course_id = old.course_id ;
END;
'''
                        c.execute(course_delete_trigger)#$%^&*
                        batch_delete_trigger='''
CREATE TRIGGER IF NOT EXISTS batch_delete_trigger
BEFORE DELETE ON batch_details
BEGIN
DELETE FROM student_details WHERE batch_id = old.batch_id ;
DELETE FROM book_history WHERE batch_id = old.batch_id ;
DELETE FROM issued_book_details WHERE batch_id = old.batch_id ;
END;
'''
                        c.execute(batch_delete_trigger)#$%^&*
                        student_delete_trigger='''
CREATE TRIGGER IF NOT EXISTS student_delete_trigger
BEFORE DELETE ON student_details
BEGIN
DELETE FROM book_history WHERE roll_number = old.roll_number ;
DELETE FROM issued_book_details WHERE roll_number = old.roll_number ;
END;
'''
                        c.execute(student_delete_trigger)#$%^&*
                        book_delete_trigger='''
CREATE TRIGGER IF NOT EXISTS book_delete_trigger
BEFORE DELETE ON book_details
BEGIN
DELETE FROM book_quantity_details WHERE book_id = old.book_id ;
DELETE FROM issued_book_details WHERE book_id = old.book_id ;
DELETE FROM book_history WHERE book_id = old.book_id ;
END;
'''
                        c.execute(book_delete_trigger)#$%^&*
                        book_qty_delete_trigger='''
CREATE TRIGGER IF NOT EXISTS book_qty_delete_trigger
BEFORE DELETE ON book_details
BEGIN
DELETE FROM issued_book_details WHERE book_id = old.book_id ;
DELETE FROM book_history WHERE book_id = old.book_id ;
END;
'''
                        c.execute(book_qty_delete_trigger)#$%^&*
                        
                except Exception as e:
                        print("s")
                        print(e)
                conn.commit()
                conn.close()  
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT login_status FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()
        
        if a is None:
                sign_up_ui()
                exist_life()
        else:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()

                login_status="3"
                query = ('''UPDATE login_details SET login_status = "{}" WHERE user_id=1;'''.format(login_status))
                c.execute(query)
                conn.commit()
                conn.close()
                back_up(0)
                login_ui()
                exist_life()



                        
def refresh(login_status):
        try:
                if login_status=="1" or login_status=="2":
                        close_all_window()
                        main_ui()
        except Exception as e:
                print(e)

def batch_name_to_id(batch_name):
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                batch_name=batch_name.replace('-',' ')
                batch_name=batch_name.split(' ')
                start_year=batch_name[0]
                end_year=batch_name[1]
                query = ('''SELECT batch_id FROM batch_details WHERE start_year="{}" and end_year="{}";'''.format(start_year,end_year))
                c.execute(query)
                a = c.fetchone()
                conn.close()
                batch_id=a[0]
                return batch_id
        except Exception as e:
                print(e)
        

def batch_id_to_name(batch_id):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('''SELECT start_year,end_year FROM batch_details WHERE batch_id="{}";'''.format(batch_id))
        c.execute(query)
        a = c.fetchone()
        conn.close()        
        start_year=a[0]
        end_year=a[1]
        batch_name=str(start_year)+"-"+str(end_year)
        return batch_name

def course_name_to_id(course_name):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('''SELECT course_id FROM course_details WHERE course_name="{}";'''.format(course_name))
        c.execute(query)
        a = c.fetchone()
        conn.close()
        course_id=a[0]
        return course_id

def course_id_to_name(course_id):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('''SELECT course_name FROM course_details WHERE course_id="{}";'''.format(course_id))
        c.execute(query)
        a = c.fetchone()
        conn.close()
        course_name=a[0]
        return course_name
def book_quantity_id_to_book_id(qty_id):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('''SELECT book_id FROM book_quantity_details WHERE book_quantity_id="{}";'''.format(qty_id))
        c.execute(query)
        a = c.fetchone()
        conn.close()
        book_id=a[0]
        return book_id
        
        
def sound_1():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT sound FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()
               
        data=a[0]
        if data=='1':
                pygame.mixer.init()
                sound = pygame.mixer.Sound("application_ui/music/click.wav")
                sound.play()
def sound_2():
        
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT sound FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()       
        data=a[0]
        if data=='1':
                pygame.mixer.init()
                sound = pygame.mixer.Sound("application_ui/music/error.wav")
                sound.play()
def sound_3():
        
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT sound FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()        
        conn.close()       
        data=a[0]
        if data=='1':
                pygame.mixer.init()
                sound = pygame.mixer.Sound("application_ui/music/success.mp3")
                sound.play()
def sound_4():
        
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT sound FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()        
        data=a[0]
        if data=='1':
                pygame.mixer.init()
                sound = pygame.mixer.Sound("application_ui/music/click_2.wav")
                sound.play()
     
                
def sound_5():
        pygame.mixer.init()
        sound = pygame.mixer.Sound("application_ui/music/click_2.wav")
        sound.play()

def sound_6():
        pygame.mixer.init()
        sound = pygame.mixer.Sound("application_ui/music/error.wav")
        sound.play()
def sound_7():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT sound FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()       
        data=a[0]
        if data=='1':
                pygame.mixer.init()
                sound = pygame.mixer.Sound("application_ui/music/exit.wav")
                sound.play()
                time.sleep(0.28)
        
def sign_up_ui():
  #get user data and save it on ###ogin_details###
  #set login column true in ###staff details###
        try:
                sound_5()
                close_all_window()
                qtRectangle=sign_up_win.frameGeometry()
                centerPoint=QDesktopWidget().availableGeometry().center()
                qtRectangle.moveCenter(centerPoint)
                sign_up_win.move(qtRectangle.topLeft())
                sign_up_win.show()
                sign_up_win.next_btn.clicked.connect(sign_up_process)
        except Exception as e:
                print(e)
def sign_up_process():
        user_name=sign_up_win.user_name.text()
        email=sign_up_win.email.text()
        password_1=sign_up_win.password_1.text()
        password_2=sign_up_win.password_2.text()

        if user_name=='':
                sound_6()
                sign_up_win.error_message.setText("Please enter the user name !")
                return
        if email=='':
                sound_6()
                sign_up_win.error_message.setText("Please enter the email !")
                return
        else:
                regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
                if not re.fullmatch(regex, email):
                        sound_6()
                        sign_up_win.error_message.setText("Please enter the valid email !")
                        return
        if password_1=='':
                sound_6()
                sign_up_win.error_message.setText("Please enter the password 1 !")
                return
        else:
                if not len(password_1)>=3:
                        sound_6()
                        sign_up_win.error_message.setText("Please enter the password 1 atleast in 3 character !")
                        return
        if password_2=='':
                sound_6()
                sign_up_win.error_message.setText("Please enter the Password 2 !")
                return
        else:
                if not len(password_2)>=3:
                        sound_6()
                        sign_up_win.error_message.setText("Please enter the password 2 atleast in 3 character !")
                        return
        try:
                themes="1"
                sound="1"
                student_login="1"
                backup="1"
                login_status="1"
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
     
                c.execute("insert into login_details values(?,?,?,?,?,?,?,?,?,?)",(1,user_name,email,password_1,password_2,themes,sound,student_login,backup,login_status))
                conn.commit()
                conn.close()
                login_status="1"
                refresh(login_status)
        except Exception as e:
                print(e)
        


def login_ui():
  #display the ***login UI***
  #get user name and password
  #if user name found check password, if password match, set login column true in ###staff details###
  #else print ***user name or password was not matching***
        sound_4()
        close_all_window()
        qtRectangle=login_win.frameGeometry()
        centerPoint=QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        login_win.move(qtRectangle.topLeft())
        login_win.show()
        set_button_color_for_active_window()
        login_win.progressBar.hide()
        login_win.forgot_password.clicked.connect(forgot_password_ui)
        login_win.login_btn.clicked.connect(login_process)


def login_process():
        try:
                if login_win.user_name.text()=='' or login_win.password.text()=='':
                        sound_2()
                        reply=QMessageBox.critical(login_win,"No input given","Please input the login details",QMessageBox.Ok)    
                        return
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')#database connect
                c = conn.cursor()
                query = '''SELECT * FROM login_details WHERE user_id=1;'''
                c.execute(query)
                a = c.fetchone()
                conn.close()
                
                user_name=a[1]
                email=a[2]
                password_1=a[3]
                password_2=a[4]
                login_status=a[9]
                
                if login_win.password.text()=='student':
                        main_win.issue_book_btn.setEnabled(False)
                        main_win.return_book_btn.setEnabled(False)
                        main_win.issued_book_details_btn.setEnabled(False)
                        main_win.manage_students_btn.setEnabled(False)
                        main_win.manage_book_quantity_btn.setEnabled(False)
                        main_win.settings_btn.setEnabled(False)
                        main_win.batch_and_course_btn.setEnabled(False)
                        main_win.role_number_4.setText(login_win.user_name.text())
                        main_win.role_number_4.setEnabled(False)
                        
                        main_win.add_book_pop_up_btn.setEnabled(False)
                        main_win.delete_book_btn.setEnabled(False)
                        main_win.manage_books_save_btn.setEnabled(False)
                        login_status="2"
                else:
                        if login_win.user_name.text()==user_name:
                                if login_win.password.text()==password_1:
                                        login_status="1"
                                elif login_win.password.text()==password_2:
                                        login_status="2"
                                else:
                                        sound_2()
                                        reply=QMessageBox.critical(login_win,"Failure",("Your password or user name is invalid !"),QMessageBox.Ok)
                                        login_status="3"
                        else:
                                sound_2()
                                reply=QMessageBox.critical(login_win,"Failure",("Your password or user name is invalid !"),QMessageBox.Ok)
                                login_status="3"
                
                                
                login_win.user_name.setText('')
                login_win.password.setText('')
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''UPDATE login_details SET login_status = "{}" WHERE user_id=1;'''.format(login_status))
                c.execute(query)
                conn.commit()
                conn.close()
              

                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT login_status FROM login_details WHERE user_id=1;'''
                c.execute(query)
                a = c.fetchone()
                conn.close()
                login_status=a[0]
                refresh(login_status)

        except Exception as e:
                sound_2()
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)
    

def forgot_password_ui():
        context = ssl.create_default_context()
        url = "https://www.bhphotovideo.com"
        timeout = 8
        sound_1()
        try:
                request = requests.get(url, timeout=timeout)
        except (requests.ConnectionError, requests.Timeout) as exception:
                sound_2()
                reply=QMessageBox.critical(login_win,"Failure",("Please connecet the internet and try again ! {}".format(exception)),QMessageBox.Ok)
                return
        # Try to log in to server and send email
        try:
                login_win.progressBar.show()
                login_win.progressBar.setValue(1)
                server = smtplib.SMTP('smtp.gmail.com', 587)
                login_win.progressBar.setValue(17)
                server.ehlo()  # check connection
                login_win.progressBar.setValue(29)
                server.starttls(context=context)  # Secure the connection
                login_win.progressBar.setValue(43)
                server.ehlo()  # check connection
                login_win.progressBar.setValue(56)
                server.login('pradeepdowney64@gmail.com','dsce nrao hwdx txzo')
                login_win.progressBar.setValue(65)
                conn = sqlite3.connect('application_data/databases/clms_database.db')#database connect
                login_win.progressBar.setValue(67)
                c = conn.cursor()
                login_win.progressBar.setValue(70)
                query = '''SELECT * FROM login_details WHERE user_id=1;'''
                login_win.progressBar.setValue(73)
                c.execute(query)
                login_win.progressBar.setValue(75)
                a = c.fetchone()
                login_win.progressBar.setValue(80)
                user_name=a[1]
                login_win.progressBar.setValue(81)
                email=a[2]
                login_win.progressBar.setValue(82)
                password_1=a[3]
                login_win.progressBar.setValue(83)
                password_2=a[4]
                login_win.progressBar.setValue(85)
                conn.commit()
                login_win.progressBar.setValue(87)
                conn.close()
                login_win.progressBar.setValue(91)
                
                
                # Send email here
                message = MIMEText('\nHi {},\n\n User Name: {} \n password_1 (admin): {}\n password_2 (staff): {}\n\n This mail was generated by CLMS software, so please do not reply. Thank you for using CLMS!'.format(user_name,user_name,password_1,password_2,))
                message['Subject'] = 'CLMS OTP Verification'
                message['From'] = "pradeepdowney64@gmail.com"
                message['To'] = "pradeepdowney63@gmail.com"
                login_win.progressBar.setValue(94)
                server.sendmail('pradeepdowney64@gmail.com', email, message.as_string())
                login_win.progressBar.setValue(100)
                login_win.progressBar.hide()
                sound_3()
                reply=QMessageBox.information(login_win,"Success",("Your login details are successfully sent to your email {} !".format(email)),QMessageBox.Ok)
       
        except Exception as e:
                # Print any error messages
                sound_2()
                reply=QMessageBox.critical(login_win,"Error",str(e),QMessageBox.Ok)                
        finally:
                server.quit()
                
def main_ui():
#if cookie!=true:
#go to -->login_ui()
#else show ***main_ui***

  qtRectangle=main_win.frameGeometry()
  centerPoint=QDesktopWidget().availableGeometry().center()
  centerPoint.setY(centerPoint.y()-7)
  qtRectangle.moveCenter(centerPoint)
  main_win.move(qtRectangle.topLeft())
  main_win.show()

  main_win.home_btn.clicked.connect(home_ui)
  main_win.issue_book_btn.clicked.connect(issue_book_ui)
  main_win.return_book_btn.clicked.connect(return_book_ui)
  main_win.issued_book_details_btn.clicked.connect(issued_book_details_ui)
  main_win.defaulter_btn.clicked.connect(defaulter_ui)
  main_win.manage_students_btn.clicked.connect(manage_students_ui)
  main_win.manage_books_btn.clicked.connect(manage_books_ui)
  main_win.manage_book_quantity_btn.clicked.connect(manage_book_quantity_ui)
  main_win.settings_btn.clicked.connect(settings_ui)
  main_win.batch_and_course_btn.clicked.connect(manage_batch_ui)
  main_win.exit_btn.clicked.connect(on_sys_exit) #PROGRAM KILL sys.exit


  main_win.student_profile_btn_7.clicked.connect(student_profile_ui)
  main_win.student_profile_btn_10.clicked.connect(personal_details_ui)
  main_win.book_history_btn.clicked.connect(book_history_ui)


  main_win.export_manage_books_btn.clicked.connect(export_manage_books_process)


  main_win.roll_number.textChanged.connect(refresh_student_info)
  main_win.book_qty_id.textChanged.connect(refresh_book_info)


  main_win.issue_btn.clicked.connect(issue_query)
  main_win.issue_book_input_clear_btn.clicked.connect(issue_book_input_clear)

  main_win.manage_course_btn.clicked.connect(manage_course_ui)
  main_win.manage_batch_btn.clicked.connect(manage_batch_ui)

  main_win.change_user_name_btn.clicked.connect(change_user_name_ui)
  main_win.change_user_name_process_btn.clicked.connect(change_user_name_process)
  main_win.change_user_name_pop_up_close_btn.clicked.connect(change_user_name_pop_up_close)
  main_win.change_user_name_pop_up_close_btn_2.clicked.connect(change_user_name_pop_up_close)
  main_win.change_password_process_btn.clicked.connect(change_password_process)
  main_win.change_email_btn.clicked.connect(change_email_ui)
  main_win.change_email_process_btn.clicked.connect(change_email_process)
  main_win.change_email_pop_up_close_btn.clicked.connect(change_email_pop_up_close)
  main_win.change_email_pop_up_close_btn_2.clicked.connect(change_email_pop_up_close)
  main_win.change_password_pop_up_close_btn.clicked.connect(change_password_pop_up_close)
  main_win.change_password_pop_up_close_btn_2.clicked.connect(change_password_pop_up_close)
  main_win.change_password_btn.clicked.connect(change_password_ui)

  main_win.add_book_quantity_btn.clicked.connect(add_book_quantity_process)
  main_win.add_book_quantity_close_btn.clicked.connect(add_book_quantity_close)
  main_win.save_book_quantity_btn.clicked.connect(save_add_book_quntity_details_process)
  main_win.delete_book_quantity_btn.clicked.connect(delete_book_quantity_process)
  
  main_win.manage_student_pop_up_btn.clicked.connect(manage_student_pop_up_process)
  main_win.close_manage_student_btn.clicked.connect(close_manage_student_process)
  main_win.delete_manage_student_btn.clicked.connect(delete_student_process)

  main_win.next_1.clicked.connect(next_2_process)
  main_win.next_2.clicked.connect(next_3_process)
  main_win.next_3.clicked.connect(next_4_process)

  main_win.prev_2.clicked.connect(next_1_process)
  main_win.prev_3.clicked.connect(next_2_process)
  main_win.prev_4.clicked.connect(next_3_process)
  
  main_win.finish.clicked.connect(finish_process)

  main_win.add_course_btn.clicked.connect(add_course_process)
  main_win.close_add_course_btn.clicked.connect(close_add_course_process)
  main_win.save_add_course_btn.clicked.connect(save_add_course_process)
  main_win.delete_course_btn.clicked.connect(delete_course_process)

  main_win.add_batch_btn.clicked.connect(add_batch_pop_up)
  main_win.add_batch_pop_up_close_btn.clicked.connect(close_add_batch_process)
  main_win.save_add_batch_btn.clicked.connect(save_add_batch_process)
  main_win.delete_batch_btn.clicked.connect(delete_batch_process)
  
  main_win.filter_btn.clicked.connect(filter_pop_up_ui)
  main_win.close_filter_pop_up_btn.clicked.connect(close_filter_pop_up)

  main_win.reset_btn.clicked.connect(reset_process)
  main_win.restore_btn.clicked.connect(restore_process)
  main_win.delete_back_up_btn.clicked.connect(delete_back_up_process)
  main_win.back_up_btn.clicked.connect(lambda:back_up(1))

  main_win.add_book_pop_up_btn.clicked.connect(add_book_pop_up)
  main_win.add_book_close_btn.clicked.connect(add_book_close_process)
  main_win.save_add_book_details_btn.clicked.connect(save_add_book_details_process)
  main_win.delete_book_btn.clicked.connect(delete_book_process)
  main_win.delete_issued_book_btn.clicked.connect(delete_issued_book_row)


  main_win.clear_return_book_btn.clicked.connect(clear_return_book_process_2)
  
  main_win.clear_btn.clicked.connect(clear_book_history_process_2)
  
  main_win.return_btn.clicked.connect(return_query)

  main_win.manage_books_save_btn.clicked.connect(manage_books_table_changed)




  main_win.book_filter.textChanged.connect(refresh_book_details_table)
  main_win.student_filter.textChanged.connect(refresh_student_details_table)
  main_win.clear_filter.clicked.connect(clear_filter_process)
  main_win.save_student_btn.clicked.connect(manage_students_details_table_changed)
  main_win.course_7.currentTextChanged.connect(refresh_student_details_table)
  main_win.batch_6.currentTextChanged.connect(refresh_student_details_table)

  main_win.clear_filter_btn.clicked.connect(clear_filter_proces)
  main_win.export_student_details_list_btn.clicked.connect(export_student_details_list_process)
                                                           

  

  clear_main_ui()
  home_ui()
  
def home_ui():
        sound_4()
        clear_main_ui()
        
        pixmap1=QPixmap("application_ui/images/college_background_image.png")
        main_win.home_screen_saver.setPixmap(pixmap1)
        main_win.home_screen_saver.show()
        lists=random.sample(range(1,20),4)
                        
        pixmap2=QPixmap("application_ui/images/{}.png".format(lists[0]))
        main_win.sub_image_1.setPixmap(pixmap2)
        main_win.sub_image_1.show()
        
        pixmap3=QPixmap("application_ui/images/{}.png".format(lists[1]))
        main_win.sub_image_2.setPixmap(pixmap3)
        main_win.sub_image_2.show()
        
        pixmap4=QPixmap("application_ui/images/{}.png".format(lists[2]))
        main_win.sub_image_3.setPixmap(pixmap4)
        main_win.sub_image_3.show()
        
        pixmap5=QPixmap("application_ui/images/{}.png".format(lists[3]))
        main_win.sub_image_4.setPixmap(pixmap5)
        main_win.sub_image_4.show()
        
        main_win.home_frame.show()
        set_button_color_for_active_window()
        main_win.setWindowTitle("Home - College Library Management System")

def issue_book_ui():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT roll_number FROM student_details;'''
                c.execute(query)
                dx = c.fetchone()
                conn.close()
                if dx is None:
                        sound_2()
                        reply=QMessageBox.warning(main_win,"Warning","Please add student details.\nStudent table is empty! ",QMessageBox.Ok)
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_quantity_id FROM book_quantity_details;'''
                c.execute(query)
                dy = c.fetchone()
                conn.close()
                if dy is None:
                        sound_2()
                        reply=QMessageBox.warning(main_win,"Warning","Please add book quantity details.\nBook quantity table is empty! ",QMessageBox.Ok)
                        return
                sound_1()
                clear_main_ui()
                main_win.issue_book_frame.show()
                set_button_color_for_active_window()
                main_win.setWindowTitle("Issue Book - College Library Management System")
                main_win.today_date.setDateTime(QtCore.QDateTime.currentDateTime())
                main_win.today_date.setDisplayFormat("dd/MM/yyyy")
                ds=datetime.datetime.now().day
                ms=datetime.datetime.now().month
                ys=datetime.datetime.now().year
                ds=int(ds)
                ms=int(ms)
                ys=int(ys)
                d=QDate(ys,ms,ds).addMonths(5)
                main_win.last_date.setDate(d)
                main_win.last_date.setDisplayFormat("dd/MM/yyyy")
        except Exception as e:
                print("kkke")
                print(e)




  
def issue_query():
        try:
                roll_number=main_win.roll_number.text()
                book_qty_id=str(main_win.book_qty_id.text())
                if roll_number=='':
                        sound_2()
                        reply=QMessageBox.warning(main_win,"warning","Please enter roll number ! ",QMessageBox.Ok)
                        return
                if book_qty_id=='':
                        sound_2()
                        reply=QMessageBox.warning(main_win,"warning","Please enter the book quantity id ! ",QMessageBox.Ok)
                        return
                if main_win.phone_number.text()=='':
                        sound_2()
                        reply=QMessageBox.warning(main_win,"warning","Student roll number not found ! ",QMessageBox.Ok)                        
                        return
                if main_win.book_number.text()=='':
                        sound_2()
                        reply=QMessageBox.warning(main_win,"warning","Book quantity id not found ! ",QMessageBox.Ok)                        
                        return
                #!!!!!!!!!!!!!        
                        
               
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_quantity_id FROM issued_book_details;'''
                c.execute(query)
                sx = c.fetchall()
                conn.close()
                if sx!=None:
                        length_of_sx=len(sx)
                        for i in range(0,length_of_sx):
                                if sx[i][0]==book_qty_id:
                                        sound_2()
                                        reply=QMessageBox.warning(main_win,"warning","Book was already issued ! ",QMessageBox.Ok)
                                        return
                        
                issued_book_id=str(roll_number)+str("-")+str(book_qty_id)
                issue_date=main_win.today_date.date()
                issue_date=issue_date.toString("dd/MM/yyyy")
                last_date=main_win.last_date.date()
                last_date=last_date.toString("dd/MM/yyyy")

                ds=book_quantity_id_to_book_id(book_qty_id)
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''SELECT total_issued FROM book_details WHERE book_id="{}";'''.format(ds))
                c.execute(query)
                df = c.fetchone()
                conn.close()

                df=df[0]
                df=int(df)+1
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('UPDATE book_details SET total_issued="{}" where book_id="{}"'.format(df,ds))
                c.execute(query)
                conn.commit()
                conn.close()

                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''SELECT course_id,batch_id FROM student_details WHERE roll_number="{}";'''.format(roll_number))
                c.execute(query)
                sa = c.fetchone()
                conn.close()

                course_id=sa[0]
                batch_id=sa[1]
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                c.execute("INSERT INTO issued_book_details (issued_book_id,roll_number,book_quantity_id,issue_date,last_date,book_id,course_id,batch_id)VALUES(?,?,?,?,?,?,?,?)",(issued_book_id,roll_number,book_qty_id,issue_date,last_date,ds,course_id,batch_id))
                conn.commit()
                conn.close()

                sound_3()
                reply=QMessageBox.information(main_win,"Success","The book was issued! ",QMessageBox.Ok)
                main_win.roll_number.setText('')
                main_win.book_qty_id.setText('')
                refresh_student_info()
                refresh_book_info()
        #if data is found; save to the database
        #alert success
        #clear screan
        #set same student  id
        except Exception as e:
                print("uue")
                print(e)
def issue_book_input_clear():
        try:
                sound_4()
                main_win.roll_number.setText('')
                main_win.book_qty_id.setText('')
                refresh_student_info()
                refresh_book_info()
        except Exception as e:
                print("eghjhg")
                print(e)
        


def refresh_student_info():
        try:
                main_win.name.setText('')
                main_win.course.setText('')
                main_win.batch.setText('')
                main_win.phone_number.setText('')
                main_win.dob.setText('')
                main_win.address.setText('')
                main_win.student_description.setPlainText('')

                roll_number=main_win.roll_number.text()

                if roll_number == '':
                        main_win.name.setText('')
                        main_win.course.setText('')
                        main_win.batch.setText('')
                        main_win.phone_number.setText('')
                        main_win.dob.setText('')
                        main_win.address.setText('')
                        main_win.student_description.setPlainText('')
                else:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = '''SELECT roll_number,name,date_of_birth,phone,address,description,batch_id,course_id FROM student_details;'''
                        c.execute(query)
                        b = c.fetchall()
                        conn.close()
                        if b==None:
                                return
                        row=()
                        z=()
                        i=0
                        for i,z in enumerate(b):
                            if z[0]==roll_number:
                                break
                            else:
                                z=()
                        i=0
                        if z==():
                            return
                        print(z)
                        name=z[1]
                        date_of_birth=z[2]
                        phone=z[3]
                        address=z[4]
                        description=z[5]
                        batch_id=z[6]
                        course_id=z[7]

                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('''SELECT start_year,end_year FROM batch_details WHERE batch_id="{}" ;'''.format(batch_id))
                        c.execute(query)
                        a = c.fetchone()
                        conn.close()
                        print(a)
                        if a==None:
                                return
                        print("a")
                        
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('''SELECT course_name FROM course_details WHERE course_id="{}" ;'''.format(course_id))
                        c.execute(query)
                        d = c.fetchone()
                        conn.close()
                        if d==None:
                                return
                        print("b")

                        start_year=a[0]
                        end_year=a[1]
                        batch_name=str(start_year+"-"+end_year)
                        course_name=d[0]
                        main_win.name.setText(name)
                        main_win.course.setText(course_name)
                        main_win.batch.setText(batch_name)
                        main_win.phone_number.setText(phone)
                        main_win.dob.setText(date_of_birth)
                        main_win.address.setText(address)
                        main_win.student_description.setPlainText(description)
                #check roll number  is valid
                #if valid; get data
                #print data
                

        except Exception as e:
                print("jkjkje")
                print(e)
def refresh_book_info():
        try:
                main_win.book_number.setText('')
                main_win.title.setText('')
                main_win.author.setText('')
                main_win.isbn.setText('')
                main_win.publisher.setText('')
                main_win.release_date.setText('')
                main_win.book_description.setPlainText('')
                b=()
                book_qty_id=main_win.book_qty_id.text()
                if book_qty_id == '':
                        main_win.book_number.setText('')
                        main_win.title.setText('')
                        main_win.author.setText('')
                        main_win.isbn.setText('')
                        main_win.publisher.setText('')
                        main_win.release_date.setText('')
                        main_win.book_description.setPlainText('')
                else:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = '''SELECT book_quantity_id,book_id,description FROM book_quantity_details;'''
                        c.execute(query)
                        b = c.fetchall()
                        conn.close()
                if b==None:
                        return   
                row=()
                z=()
                i=0
                for i,z in enumerate(b):
                        if str(z[0])==book_qty_id:
                                break
                        else:
                                z=()
                i=0
                if z==():
                        return
                hh=z[1]
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''SELECT book_title,author,isbn,publisher,publication_date FROM book_details WHERE book_id="{}" ;'''.format(hh))
                c.execute(query)
                a = c.fetchone()
                conn.close()
                
                if a==None:
                        return

                book_id=z[1]
                title=a[0]
                author=a[1]
                isbn=a[2]
                publisher=a[3]
                publication_date=a[4]
                description=z[2]
                main_win.book_number.setText(str(book_id))                
                main_win.title.setText(title)
                main_win.author.setText(author)
                main_win.isbn.setText(isbn)
                main_win.publisher.setText(publisher)
                main_win.release_date.setText(publication_date)
                main_win.book_description.setPlainText(description)
                #check roll number  is valid
                #if valid; get data
                #print data
        except Exception as e:
                print("e")
                print(e)
        
def return_book_ui():
        sound_1()
        clear_main_ui()
        main_win.return_book_frame.show()
        set_button_color_for_active_window()
        main_win.setWindowTitle("Return Book - College Library Management System")
        main_win.roll_number_4.textChanged.connect(refresh_return_book)
def clear_return_book_process_2():
        main_win.roll_number_4.setText('')
        main_win.book_qty_id.setText('')
        clear_return_book_process()#@@@@@@@@@@@@@@@@@@@
        
def clear_return_book_process():
        try:
                roll_number=main_win.roll_number_4.text()
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM issued_book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()
                ad=()
                rowCount=main_win.return_book_listing_table.rowCount()
                
                
                for Current_row in range(0,rowCount):
                        main_win.return_book_listing_table.removeRow(0)
                main_win.name_5.setText('')
                main_win.course_5.setText('')
                main_win.batch_4.setText('')
                main_win.phone_4.setText('')
                main_win.dob_2.setText('')
                main_win.address_5.setText('')
                main_win.description_7.setPlainText('')
        except Exception as e:
                print("lllll")
                print(e)
def refresh_return_book():
        try:
                clear_return_book_process()
                main_win.name_5.setText('')
                main_win.course_5.setText('')
                main_win.batch_4.setText('')
                main_win.phone_4.setText('')
                main_win.dob_2.setText('')
                main_win.address_5.setText('')
                main_win.description_7.setPlainText('')
                roll_number=main_win.roll_number_4.text()

                if roll_number == '':
                        main_win.name_5.setText('')
                        main_win.course_5.setText('')
                        main_win.batch_4.setText('')
                        main_win.phone_4.setText('')
                        main_win.dob_2.setText('')
                        main_win.address_5.setText('')
                        main_win.description_7.setPlainText('')
                else:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = '''SELECT roll_number,name,date_of_birth,phone,address,description,batch_id,course_id FROM student_details;'''
                        c.execute(query)
                        b = c.fetchall()
                        conn.close()
                        if b==None:
                                return
                        row=()
                        z=()
                        i=0
                        for i,z in enumerate(b):
                            if z[0]==roll_number:
                                break
                            else:
                                z=()
                        i=0
                        if z==():
                            return
                        name=z[1]
                        date_of_birth=z[2]
                        phone=z[3]
                        address=z[4]
                        description=z[5]
                        batch_id=z[6]
                        course_id=z[7]

                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('''SELECT start_year,end_year FROM batch_details WHERE batch_id="{}" ;'''.format(batch_id))
                        c.execute(query)
                        a = c.fetchone()
                        conn.close()
                        if a==None:
                                return
                        
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('''SELECT course_name FROM course_details WHERE course_id="{}" ;'''.format(course_id))
                        c.execute(query)
                        d = c.fetchone()
                        conn.close()
                        if d==None:
                                return
                
                        start_year=a[0]
                        end_year=a[1]
                        batch_name=str(start_year+"-"+end_year)
                        course_name=d[0]

                        main_win.name_5.setText(name)
                        main_win.course_5.setText(course_name)
                        main_win.batch_4.setText(batch_name)
                        main_win.phone_4.setText(phone)
                        main_win.dob_2.setText(date_of_birth)
                        main_win.address_5.setText(address)
                        main_win.description_7.setPlainText(description)
                #check roll number  is valid
                #if valid; get data
                #print data

#************************
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = '''SELECT * FROM issued_book_details;'''
                        c.execute(query)
                        a=c.fetchall()
                        conn.close()


                        ad=()
                        for i in range(0,len(a)):
                                if str(a[i][1])==str(roll_number):
                                        ad+=(a[i],)
                                

                        rowCount=0
                        rowCount_2=0
                        rowPosition=0
                        rowCount=main_win.return_book_listing_table.rowCount()
                        rowCount_2=len(ad)
                        clear_defaulter_list()

                        if not rowCount == rowCount_2:
                                for Current_row in range(rowCount, rowCount_2):
                                        main_win.return_book_listing_table.insertRow(Current_row)

                        rowPosition=0
                        rowCount=main_win.return_book_listing_table.rowCount()
                        rowCount_2=len(a)
                        ds=datetime.datetime.now().day
                        ms=datetime.datetime.now().month
                        ys=datetime.datetime.now().year

                        ds=int(ds)
                        ms=int(ms)
                        ys=int(ys)
                        d=QDate(ys,ms,ds)
                        d=d.toPyDate()
                        out=0
                        main_win.issued_book_details_table.setStyleSheet("")
                        for rowPosition in range(0,rowCount):
                                for column,item in enumerate(ad[rowPosition]):
                                        if item==None:
                                                continue
                                        if column==None:
                                                continue
                                        
                                        main_win.return_book_listing_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
                                        
                                        if column==5:
                                                continue
                                        dd=datetime.datetime.strptime(ad[rowPosition][4],'%d/%m/%Y').date()
                                        out=dd-d
                                        out=int(out.days)
                                        if out<0:
                                                main_win.return_book_listing_table.item(rowPosition,column).setBackground(QBrush(QColor(244,164,96)))
                                        #@@@@@@@@@@@@@@@@
                                
        except Exception as e:
                print("Refresh self detect return book ")
                print(e)


def return_query():
        roll_number=main_win.roll_number_4.text()
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM issued_book_details;'''
        c.execute(query)
        qq=c.fetchall()
        conn.close()
        try:
                a=()
                for i in range(0,len(qq)):
                        if str(qq[i][1])==str(roll_number):
                                a+=(qq[i],)
                last_row=len(a)
                currentIndex=main_win.return_book_listing_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.return_book_listing_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to return the book?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to return the book?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        ds=datetime.datetime.now().day
                        ms=datetime.datetime.now().month
                        ys=datetime.datetime.now().year

                        ds=int(ds)
                        ms=int(ms)
                        ys=int(ys)
                        d=QDate(ys,ms,ds)
                        d=d.toPyDate()
                        print("1")
                        
                        query = 'SELECT * FROM issued_book_details where issued_book_id="{}"'.format(currentIndex_2)
                        c.execute(query)
                        b=c.fetchone()
                        roll_number=b[1]
                        book_quantity_id=b[2]
                        issue_date=b[3]
                        last_date=b[4]
                        book_id=b[5]
                        course_id=b[6]
                        batch_id=b[7]
                        given_date=d
                        print("2")

                        c.execute("INSERT INTO book_history(issued_book_id,roll_number,book_quantity_id,issue_date,last_date,given_date,book_id,course_id,batch_id)VALUES(?,?,?,?,?,?,?,?,?)",(currentIndex_2,roll_number,book_quantity_id,issue_date,last_date,given_date,book_id,batch_id,course_id))
                        
                        print("3")
                        
                        query3 = 'DELETE from issued_book_details where issued_book_id="{}"'.format(currentIndex_2)
                        c.execute(query3)
                        main_win.return_book_listing_table.removeRow(currentIndex)

                        print("4")



                        
                        print("6")
                        
                        conn.commit()
                        conn.close()

                        
                        print("8")
                        refresh_return_book()
                        
                        print("finish")
        except Exception as e:
                print("ejkjhghjhghjkjhgfghjhgf")
                print(e)


def issued_book_details_ui():
        try:
                sound_1()
                clear_main_ui()
                main_win.issued_book_details_frame.show()
                set_button_color_for_active_window()
                main_win.setWindowTitle("Issued Book Details - College Library Management System")
                refresh_issued_book_list()

        except Exception as e:
                print(e)
def clear_issued_book_list():
        
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM issued_book_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.issued_book_details_table.rowCount()
        rowCount_2=len(a)
        for i in range(0,rowCount):
                        main_win.issued_book_details_table.removeRow(0)   
       
def refresh_issued_book_list():
        
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM issued_book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.issued_book_details_table.rowCount()
                rowCount_2=len(a)
                clear_issued_book_list()
                for Current_row in range(0, rowCount_2):
                        main_win.issued_book_details_table.insertRow(Current_row)
                                
                rowPosition=0
                rowCount=main_win.issued_book_details_table.rowCount()
                
                ds=datetime.datetime.now().day
                ms=datetime.datetime.now().month
                ys=datetime.datetime.now().year
                ds=int(ds)
                ms=int(ms)
                ys=int(ys)
                d=QDate(ys,ms,ds)
                d=d.toPyDate()
                out=0
                rowCount_2=len(a)
                main_win.issued_book_details_table.setStyleSheet("")
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(a[rowPosition]):
                                if item==None:
                                        continue
                                main_win.issued_book_details_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
                                dd=datetime.datetime.strptime(a[rowPosition][4],'%d/%m/%Y').date()
                                out=dd-d
                                out=int(out.days)
                                if out<0:
                                        main_win.issued_book_details_table.item(rowPosition,column).setBackground(QBrush(QColor(244,164,96)))
                                      

                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                

                for i in range(0,rowCount):
                                roll_num = main_win.issued_book_details_table.item(i, 1)
                                roll_num = roll_num.text()
                                query = ('''SELECT name,batch_id,course_id FROM student_details WHERE roll_number="{}";'''.format(roll_num))
                                c.execute(query)
                                ex=c.fetchone()
                                batch=batch_id_to_name(ex[1])
                                course=course_id_to_name(ex[2])
                                main_win.issued_book_details_table.setItem(i,5,QTableWidgetItem(str(ex[0])))
                                main_win.issued_book_details_table.setItem(i,6,QTableWidgetItem(str(batch)))
                                main_win.issued_book_details_table.setItem(i,7,QTableWidgetItem(str(course)))

                conn.close()
                                          #@@@@@@@@@@@@@@@@
                                             
                                
        except Exception as e:
                print("Refresh self detect issued book detail")
                print(e)


def delete_issued_book_row():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM issued_book_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.issued_book_details_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.issued_book_details_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        
                        
                        query4='SELECT book_quantity_id FROM issued_book_details where issued_book_id="{}"'.format(currentIndex_2)
                        c.execute(query4)
                        eq=c.fetchone()
                        qw=eq[0]
                        er=book_quantity_id_to_book_id(qw)
                        query6='SELECT total_issued FROM book_details where book_id="{}"'.format(er)
                        c.execute(query6)
                        et=c.fetchone()
                        ett=et[0]
                        ett=int(ett)-1
                        #77088
                        query6='UPDATE book_details SET total_issued="{}" where book_id="{}"'.format(ett,er)
                        c.execute(query6)
                        et=c.fetchone()


                        query = ('DELETE from issued_book_details where issued_book_id="{}"'.format(currentIndex_2))
                        c.execute(query)

                        conn.commit()
                        conn.close()
                        main_win.issued_book_details_table.removeRow(currentIndex) 
                        refresh_issued_book_list()
        except Exception as e:
                print("e")
                print(e)

def defaulter_ui():
        sound_1()
        clear_main_ui()
        main_win.defaulter_frame.show()
        set_button_color_for_active_window()
        main_win.setWindowTitle("Defaulter List - College Library Management System")
        refresh_defaulter_table()
def clear_defaulter_list():
        #**************************
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM issued_book_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.defaulter_table.rowCount()
        rowCount_2=len(a)
        
        for i in range(0,rowCount):
                main_win.defaulter_table.removeRow(0)
                        
                
  
       
def refresh_defaulter_table():
        
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM issued_book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                
                ds=datetime.datetime.now().day
                ms=datetime.datetime.now().month
                ys=datetime.datetime.now().year
                ds=int(ds)
                ms=int(ms)
                ys=int(ys)
                d=QDate(ys,ms,ds)
                d=d.toPyDate()
                ad=()
                out=0
                for i in range(0,len(a)):
                        dd=datetime.datetime.strptime(a[i][4],'%d/%m/%Y').date()
                        
                        out=dd-d
                        out=int(out.days)
                        if out<0:
                                ad+=(a[i],)
                        out=0
                        

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.defaulter_table.rowCount()
                rowCount_2=len(ad)
                clear_defaulter_list()



                if not rowCount == rowCount_2:
                        for Current_row in range(rowCount, rowCount_2):
                                main_win.defaulter_table.insertRow(Current_row)
                                
                rowPosition=0
                rowCount=main_win.defaulter_table.rowCount()
                rowCount_2=len(a)
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(ad[rowPosition]):
                                if item==None:
                                        continue
                                main_win.defaulter_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))

                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()

                for i in range(0,rowCount):
                                roll_num = main_win.defaulter_table.item(i, 1)
                                roll_num = roll_num.text()
                                query = ('''SELECT name,batch_id,course_id FROM student_details WHERE roll_number="{}";'''.format(roll_num))
                                c.execute(query)
                                ex=c.fetchone()
                                batch=batch_id_to_name(ex[1])
                                course=course_id_to_name(ex[2])
                                main_win.defaulter_table.setItem(i,5,QTableWidgetItem(str(ex[0])))
                                main_win.defaulter_table.setItem(i,6,QTableWidgetItem(str(batch)))
                                main_win.defaulter_table.setItem(i,7,QTableWidgetItem(str(course)))

                conn.close()
                
                                
                                
        except Exception as e:
                print("Refresh self detect issued book detail")
                print(e)




def manage_students_ui():
    try:
        sound_1()
        clear_main_ui()
        main_win.manage_students_frame.show()
        set_button_color_for_active_window()
        main_win.setWindowTitle("Manage Students - College Library Management System")
        refresh_student_details_table()
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT course_name FROM course_details;'''
        c.execute(query)
        ass = c.fetchall()
        conn.close()
        print(ass)
        if ass == []:
                sound_2()
                reply=QMessageBox.warning(main_win,"Warning","Please add batch details.\nbatch table is empty! ",QMessageBox.Ok)
                return
                

        qq=('All',)
        for i,d in enumerate(ass):
                qq+=(d[0],)
                main_win.course_7.clear()
                main_win.course_7.addItems(qq)
                
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT start_year,end_year FROM batch_details;'''
        c.execute(query)
        ass = c.fetchall()
        conn.close()

        qq=('All',)
        for i,d in enumerate(ass):
                san=str(d[0])+"-"+str(d[1])
                qq+=(san,)
                main_win.batch_6.clear()
                main_win.batch_6.addItems(qq)
        


    except Exception as e:
        print("Manage Student Ui ERROR\n{}".format(e))

def manage_student_pop_up_process():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT start_year FROM batch_details;'''
                c.execute(query)
                a = c.fetchall()
                conn.close()
                if a is None:
                        sound_2()
                        reply=QMessageBox.warning(main_win,"Warning","Please add batch details.\nbatch table is empty! ",QMessageBox.Ok)
                        return

                clear_pop_up()
                main_win.manage_student_pop_up.show()
                next_1_process()#fd

                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT course_name FROM course_details;'''
                c.execute(query)
                ass = c.fetchall()
                conn.close()

                qq=()
                for i,d in enumerate(ass):
                        qq+=(d[0],)
                        main_win.course_3.clear()
                        main_win.course_3.addItems(qq)
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT start_year,end_year FROM batch_details;'''
                c.execute(query)
                ass = c.fetchall()
                conn.close()

                qq=()
                for i,d in enumerate(ass):
                        san=str(d[0])+"-"+str(d[1])
                        qq+=(san,)
                        main_win.batch_2.clear()
                        main_win.batch_2.addItems(qq)
        except Exception as e:
                print(e)
        
  
def next_1_process():
        sound_4()
        clear_manage_student_pop()
        main_win.get_details_1.show()
def next_2_process():
        if main_win.roll_number_2.text()=='':
                sound_2()
                reply=QMessageBox.warning(main_win,"Warning","Please enter the roll number ! ",QMessageBox.Ok)
                return
        if main_win.phone.text()=='':
                sound_2()
                reply=QMessageBox.warning(main_win,"Warning","Please enter the phone number ! ",QMessageBox.Ok)
                return

        sound_4()
        clear_manage_student_pop()
        main_win.get_details_2.show()
def next_3_process():
        sound_4()
        clear_manage_student_pop()
        main_win.get_details_3.show()
def next_4_process():
        sound_4()
        clear_manage_student_pop()
        main_win.get_details_4.show()

def finish_process():
        roll_number=main_win.roll_number_2.text()
        name=main_win.name_2.text()
        date_of_birth=main_win.date_of_birth.text()
        father_name=main_win.father_name.text()
        mother_name=main_win.mother_name.text()
        phone=main_win.phone.text()
        email=main_win.email_2.text()
        address=main_win.address_2.text()
        aadhar_number=main_win.aadhar_number.text()
        anual_income=main_win.anual_income.text()
        blood_group=main_win.blood_group.text()
        caste=main_win.caste.text()
        community=main_win.community.text()
        father_occupation=main_win.father_occupation.text()
        mother_occupation=main_win.mother_occupation.text()
        religion=main_win.religion.text()
        area_type=main_win.area_type.text()
        hsc_mark=main_win.hsc_mark.text()
        hsc_passing_year=main_win.hsc_passing_year.text()
        sslc_mark=main_win.sslc_mark.text()
        sslc_medium=main_win.sslc_medium.text()
        sslc_passing_year=main_win.sslc_passing_year.text()
        sslc_school_name=main_win.sslc_school_name.text()
        sslc_school_type=main_win.sslc_school_type.text()
        batch=main_win.batch_2.currentText()
        course=main_win.course_3.currentText()
        course_id=course_name_to_id(course)
        batch_id=batch_name_to_id(batch)
        hsc_medium=main_win.hsc_medium.text()
        hsc_school_name=main_win.hsc_school_name.text()
        hsc_school_type=main_win.hsc_school_type.text()
        register_number=main_win.register_number.text()
        admission_date=main_win.admission_date_2.text()
        description=main_win.description_5.toPlainText()
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                c.execute("INSERT INTO student_details(roll_number,register_number,course_id,batch_id,name,date_of_birth,father_name,mother_name,phone,email,address,aadhar_number,religion,community,caste,anual_income,blood_group,father_occupation,mother_occupation,area_type, sslc_mark,sslc_passing_year,sslc_school_name,sslc_medium,sslc_school_type,hsc_mark,hsc_passing_year,hsc_school_name, hsc_medium,hsc_school_type,admission_date,description)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(roll_number,register_number,course_id,batch_id,name,date_of_birth,father_name,mother_name,phone,email,address,aadhar_number,religion,community,caste,anual_income,blood_group,father_occupation,mother_occupation,area_type,sslc_mark,sslc_passing_year,sslc_school_name,sslc_medium,sslc_school_type,hsc_mark,hsc_passing_year,hsc_school_name,hsc_medium,hsc_school_type,admission_date,description))
                conn.commit()
                conn.close()
        except Exception as e:
                print(e)


        main_win.roll_number_2.setText('')
        main_win.name_2.setText('')
        main_win.date_of_birth.setText('')
        main_win.father_name.setText('')
        mother_name=main_win.mother_name.setText('')
        main_win.phone.setText('')
        main_win.email_2.setText('')
        main_win.address_2.setText('')
        main_win.aadhar_number.setText('')
        main_win.anual_income.setText('')
        main_win.blood_group.setText('')
        main_win.caste.setText('')
        main_win.community.setText('')
        main_win.father_occupation.setText('')
        main_win.mother_occupation.setText('')
        main_win.religion.setText('')
        main_win.area_type.setText('')
        main_win.hsc_mark.setText('')
        main_win.hsc_passing_year.setText('')
        main_win.sslc_mark.setText('')
        main_win.sslc_medium.setText('')
        main_win.sslc_passing_year.setText('')
        main_win.sslc_school_name.setText('')
        main_win.sslc_school_type.setText('')
        main_win.batch_2.currentText()
        main_win.course_3.currentText()
        main_win.hsc_medium.setText('')
        main_win.hsc_school_name.setText('')
        main_win.hsc_school_type.setText('')
        main_win.register_number.setText('')
        main_win.admission_date_2.setText('')
        main_win.description_5.clear()


        refresh_student_details_table()

        if main_win.manage_student_pop_up.isVisible():
                        sound_3()
                        main_win.manage_student_pop_up.hide()
                        reply=QMessageBox.information(main_win,"Success","Student details was successfully added ! ",QMessageBox.Ok)   
        
        
def clear_student_details_table():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM student_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.manage_students_details_table.rowCount()
        rowCount_2=len(a)
        if rowCount_2==None:
                return
        for i in range(0, rowCount_2):
                        main_win.manage_students_details_table.removeRow(0)



def refresh_student_details_table():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM student_details;'''
                c.execute(query)
                pp=c.fetchall()
                conn.close()

                course=main_win.course_7.currentText()
                batch=main_win.batch_6.currentText()
                rr=()
                if course!=None:
                        if course=='All':
                                rr=pp
                        else:
                                if course=='':
                                        course_id=''
                                else:
                                        course_id=course_name_to_id(course)
                                for i in range(0,len(pp)):
                                        if pp[i][31]==course_id:
                                                rr+=(pp[i],)
                        
                dd=()
                if batch!=None:
                        if batch=='All':
                                dd=rr
                        else:
                                print(batch)
                                if batch=='':
                                        batch_id=''
                                else:
                                        batch_id=batch_name_to_id(batch)
                                for i in range(0,len(rr)):
                                        if rr[i][30]==batch_id:
                                                dd+=(rr[i],)
                
                                        
    #fd
                e=()
                keyword1=main_win.student_filter.text()
                for i in range(0,len(dd)):
                        for j,item in enumerate(dd[i]):
                                item=str(item)
                                if re.search(keyword1, item, re.IGNORECASE):
                                        e+=(dd[i],)
                if keyword1=='':
                        e=dd
                if e=='':
                        clear_student_details_table()
                e=sorted(e)
                
                a=[e[i] for i in range(len(e)) if i==0 or e[i]!=e[i-1]]
                
#ed
                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.manage_students_details_table.rowCount()
                rowCount_2=len(a)
                if rowCount_2==None:
                        return
                clear_student_details_table()
                for Current_row in range(0, rowCount_2):
                        main_win.manage_students_details_table.insertRow(Current_row)
                rowPosition=0
                rowCount=main_win.manage_students_details_table.rowCount()
                rowCount_2=len(a)
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(a[rowPosition]):
                                if item==None:
                                        continue
                                main_win.manage_students_details_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))




        except Exception as e:
                print("Refresh self detect student detail")
                print(e)


def delete_student_process():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM student_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.manage_students_details_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.manage_students_details_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('DELETE from student_details where roll_number="{}"'.format(currentIndex_2))
                        c.execute(query)
                        a = c.fetchone()
                        conn.commit()
                        conn.close()
                        main_win.manage_students_details_table.removeRow(currentIndex) 
                        refresh_student_details_table()
        except Exception as e:
                print("e")
                print(e)



                        
def manage_students_details_table_changed():#er
    try:
        t_row=main_win.manage_students_details_table.rowCount()
        t_column=main_win.manage_students_details_table.columnCount()
        for i in range(0,t_row+1):
            row_name=main_win.manage_students_details_table.item(i, 0)
            phone_num=main_win.manage_students_details_table.item(i, 6)
            if row_name==None:
                break
            if phone_num==None:
                break
            row_name=row_name.text()
            phone_num=phone_num.text()
            for j in range(0, t_column):
                    
                data=main_win.manage_students_details_table.item(i,j).text()
                if data==None:
                    break
                
                column_name = main_win.manage_students_details_table.horizontalHeaderItem(j).text()
                if column_name==-1:
                    break
                column_name=column_name.lower()
                column_name=column_name.replace(' ','_')
            
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                if column_name=='roll_number':
                        continue
                else:

                    query = ('UPDATE student_details SET {}="{}" where roll_number="{}"'.format(column_name,data,row_name))

                c.execute(query)
                conn.commit()
                conn.close()                
            
                         
    except Exception as e:
        print("here")
        print(e)

    finally:
            refresh_student_details_table()
            sound_3()
            reply=QMessageBox.information(main_win,"Success",("The Book table was saved! "),QMessageBox.Ok)
                        
   
def clear_filter_proces():
        try:
                main_win.student_filter.setText('')
                main_win.course_7.setCurrentText('All')
                main_win.batch_6.setCurrentText('All')
                close_filter_pop_up()
        except Exception as e:
                print(e)
        
                        
def close_manage_student_process():
        if main_win.manage_student_pop_up.isVisible():
                sound_7()
                main_win.manage_student_pop_up.hide()


def filter_pop_up_ui():
        sound_4()
        clear_pop_up()
        main_win.filter_pop_up.show()
        
        
def close_filter_pop_up():
        if main_win.filter_pop_up.isVisible():
                sound_7()
                main_win.filter_pop_up.hide()
        
        
     
        
def export_student_details_list_process():
        sound_4()

        try:
                x=()
                a=()
                t_row=main_win.manage_students_details_table.rowCount()
                t_column=main_win.manage_students_details_table.columnCount()
                for i in range(0,t_row):
                        for j in range(0,t_column):
                                item=main_win.manage_students_details_table.item(i, j)
                                if item==None:
                                        continue
                                item=item.text()
                                x+=(item,)
                        a+=(x,)
                        x=()
                        
                
                data=pd.DataFrame(a)                
                data.columns=['Roll Number','Register Number','Name','Date Of Birth','Father Name','Mother Name','Phone','Email','Address','Aadhar Number','Religion','Community','Caste','Anual Income','Bloood Group','Father Occupation','Mother Occupation','Area Type','SSSLC Mark','SSLC Passing Year','SSLC School Name','SSLC Medium','SSLC School Type','HSC Mark','HSC Passing Year','HSC School Name','HSC Medium','HSC School Type','Admission','Description']
                              
        except Exception as e:
                print(e)



                filename=QFileDialog.getSaveFileName(main_win,"Export Student Details List", 'c:\\', "Excel Files(*.xlsx)")
                if filename[0] == '':
                        return
                if os.path.exists(filename[0]):
                        try:
                                os.rename(filename[0],filename[0])
                        except OSError as e:
                                sound_2()
                                reply=QMessageBox.critical(main_win,"Failure - File Not Saved",str(e),QMessageBox.Ok)
                                return
                writer = pd.ExcelWriter(filename[0], engine='xlsxwriter')
                data.to_excel(writer,index=False)
                writer.save()
                sound_3()
                reply=QMessageBox.information(main_win,"Success","Student Details list was successfully exported ! ",QMessageBox.Ok)
        except Exception as e:
                print(e)
                
def export_manage_books_process():
        sound_4()
        filename=QFileDialog.getSaveFileName(main_win,"Export Book Details List", 'c:\\', "Excel Files(*.xlsx)")
        if filename[0] == '':
                return
                
        
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM book_details;'''
                c.execute(query)
                dd=c.fetchall()
                conn.close()
                                #fd
                e=()
                keyword1=main_win.book_filter.text()
                for i in range(0,len(dd)):
                        for j,item in enumerate(dd[i]):
                                item=str(item)
                                if re.search(keyword1, item, re.IGNORECASE):
                                        e+=(dd[i],)
                if keyword1=='':
                        e=dd
                if e=='':
                        clear_book_details_table()
                e=sorted(e)
                a=[e[i] for i in range(len(e)) if i==0 or e[i]!=e[i-1]]
                

                rowCount=main_win.manage_books_table.rowCount()
                rowPosition=0
                rowCount_2=len(a)
                rows=()
                lists=()
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(a[rowPosition]):
                                if column==11:
                                        print(item)
                                        continue
                                rows=rows+(item,)
                        lists=lists+(rows,)
                        rows=()
                        
                data=pd.DataFrame(lists)
                data.columns=['Book ID','Book Title','Author','ISBN','Available','Quantity','Publisher','Price','Publication Date','tags','Description']
        except Exception as e:
                print(e)
        if os.path.exists(filename[0]):
                try:
                        os.rename(filename[0],filename[0])
                except OSError as e:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Failure - File Not Saved",str(e),QMessageBox.Ok)
                        return
        writer = pd.ExcelWriter(filename[0], engine='xlsxwriter')
        data.to_excel(writer,index=False)
        writer.save()
        sound_3()
        reply=QMessageBox.information(main_win,"Success","Book Details list was successfully exported ! ",QMessageBox.Ok)        
                
def manage_books_ui():
        
        sound_1()
        clear_main_ui()
        main_win.setWindowTitle("Manage Books - College Library Management System")
        main_win.manage_books_frame.show()
        set_button_color_for_active_window()
        try:
                refresh_book_details_table()

        except Exception as e:
                print("ii")
                print(e)
        
        
def add_book_pop_up():
        try:
                clear_pop_up()
                main_win.add_book_pop_up.show()
                sound_4()
        except Exception as e:
                print(e)
def add_book_close_process():
        refresh_book_details_table()
        main_win.book_title.setText('')
        main_win.author_2.setText('')
        main_win.isbn_2.setText('')
        main_win.publisher_2.setText('')
        main_win.publication_date.setText('')
        main_win.price.setText('')
        main_win.book_tags.setText('')
        main_win.description.clear()
        if main_win.add_book_pop_up.isVisible():
                sound_7()
                main_win.add_book_pop_up.hide()
def save_add_book_details_process():
        

        try:
                book_title=main_win.book_title.text()
                author_2=main_win.author_2.text()
                isbn_2=main_win.isbn_2.text()
                publisher_2=main_win.publisher_2.text()
                publication_date=main_win.publication_date.text()
                price=main_win.price.text()
                book_tags=main_win.book_tags.text()
                description=main_win.description.toPlainText()
                available=0
                quantity=1
                total_issued=0

                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                c.execute("INSERT INTO book_details(book_title,author,isbn,available,quantity,publisher,price,publication_date,tags,description,total_issued)VALUES(?,?,?,?,?,?,?,?,?,?,?)",(book_title,author_2,isbn_2,available,quantity,publisher_2,price,publication_date,book_tags,description,total_issued))
                conn.commit()
                conn.close()
                
                refresh_book_details_table()
                main_win.book_title.setText('')
                main_win.author_2.setText('')
                main_win.isbn_2.setText('')
                main_win.publisher_2.setText('')
                main_win.publication_date.setText('')
                main_win.price.setText('')
                main_win.book_tags.setText('')
                main_win.description.clear()
                
                if main_win.add_book_pop_up.isVisible():
                        sound_3()
                        main_win.add_book_pop_up.hide()
                        reply=QMessageBox.information(main_win,"Success","Book details was successfully added ! ",QMessageBox.Ok)   
        
        except Exception as e:
                print("ooo")
                print(e)


def clear_book_details_table():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM book_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.manage_books_table.rowCount()
        rowCount_2=len(a)

        for Current_row in range(0,rowCount):
                        main_win.manage_books_table.removeRow(0)                        

def refresh_book_details_table():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM book_details;'''
                c.execute(query)
                dd=c.fetchall()
                conn.close()
                
                #fd
                e=()
                keyword1=main_win.book_filter.text()
                for i in range(0,len(dd)):
                        for j,item in enumerate(dd[i]):
                                item=str(item)
                                if re.search(keyword1, item, re.IGNORECASE):
                                        e+=(dd[i],)
                if keyword1=='':
                        e=dd
                if e=='':
                        clear_book_details_table()
                e=sorted(e)
                a=[e[i] for i in range(len(e)) if i==0 or e[i]!=e[i-1]]
                
                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.manage_books_table.rowCount()
                rowCount_2=len(a)
                clear_book_details_table()
                for Current_row in range(0, rowCount_2):
                        main_win.manage_books_table.insertRow(Current_row)
                rowPosition=0
                rowCount=main_win.manage_books_table.rowCount()
                rowCount_2=len(a)
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(a[rowPosition]):
                                if column==4:

                                        item=int(a[rowPosition][5])-int(a[rowPosition][11])
                                        print("{}-{}={}".format(a[rowPosition][5],a[rowPosition][11],item))
                                        main_win.manage_books_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
                                else:
                                        if item==None:
                                                continue
                                        main_win.manage_books_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))

        except Exception as e:
                print("Refresh self detect manage book")
                print(e)

def delete_book_process():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM book_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.manage_books_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.manage_books_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('DELETE from book_details where book_id={}'.format(currentIndex_2))
                        c.execute(query)
                        a = c.fetchone()
                        conn.commit()
                        conn.close()
                        main_win.manage_books_table.removeRow(currentIndex) 
                        refresh_book_details_table()
        except Exception as e:
                print("e")
                print(e)

def clear_filter_process():
        main_win.book_filter.setText('')
        refresh_book_details_table()
        


def manage_books_table_changed():
    try:
        t_row=main_win.manage_books_table.rowCount()
        t_column=main_win.manage_books_table.columnCount()
        for i in range(0,t_row+1):
            row_name=main_win.manage_books_table.item(i, 0)
            isbn_num=main_win.manage_books_table.item(i, 3)
            if row_name==None:
                break
            if isbn_num==None:
                break
            row_name=row_name.text()
            isbn_num=isbn_num.text()
            for j in range(0, t_column):
                data=main_win.manage_books_table.item(i,j).text()
                if data==None:
                    break
                
                column_name = main_win.manage_books_table.horizontalHeaderItem(j).text()
                if column_name==-1:
                    break
                column_name=column_name.lower()
                column_name=column_name.replace(' ','_')
                if column_name=='available':
                        continue

                if column_name=='quantity':
                        continue
                
            
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                if column_name=='book_id':
                    if not data.isnumeric():
                            sound_2()
                            reply=QMessageBox.critical(main_win,"Warning",("Please enter the numeric value! "),QMessageBox.Ok)
                            return
                    query = ('UPDATE book_details SET book_id="{}" where isbn="{}"'.format(data,isbn_num))
                    
                else:
                    query = ('UPDATE book_details SET {}="{}" where book_id="{}"'.format(column_name,data,row_name))

                c.execute(query)
                conn.commit()
                conn.close()
        refresh_book_details_table()
        sound_3()
        reply=QMessageBox.information(main_win,"Success",("The Book table was saved! "),QMessageBox.Ok)
                         
    except Exception as e:
        print("here")
        print(e)
                        
        
        
def manage_book_quantity_ui():#$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        sound_1()
        clear_main_ui()
        main_win.setWindowTitle("Manage Book Quantity - College Library Management System")
        main_win.manage_book_quantity_frame.show()
        set_button_color_for_active_window()
        try:
                refresh_book_quantity_details_table()
                main_win.save_book_quantity_btn_2.clicked.connect(manage_books_quantity_table_changed)
        except Exception as e:
                print("ii")
                print(e)

def add_book_quantity_process():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_title FROM book_details;'''
                c.execute(query)
                a = c.fetchone()
                conn.close()
        
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_id,book_title FROM book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()
                print(a)
                f=()
                for i,item in enumerate(a):
                        st=str(a[i][0])+". "+str(a[i][1])
                        f+=(st,)
                main_win.book_title_list.clear()
                main_win.book_title_list.addItems(f)#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                if a is None:
                        reply=QMessageBox.warning(main_win,"Warning","Please add book details.\nBook table is empty! ",QMessageBox.Ok)
                        return
                clear_pop_up()
                main_win.add_book_quantity_pop_up.show()
                sound_4()
        except Exception as e:
                print(e)
def add_book_quantity_close():
        refresh_book_quantity_details_table()
        main_win.book_quantity_id.setText('')
        main_win.description_2.clear()
        if main_win.add_book_quantity_pop_up.isVisible():
                sound_7()
                main_win.add_book_quantity_pop_up.hide()



def save_add_book_quntity_details_process():
        try:
                book_id=main_win.book_title_list.currentText()
                b=book_id.split('. ')
                book_id=int(b[0])
                if book_id==None:
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_id FROM book_details;'''
                c.execute(query)
                a = c.fetchall()
                conn.close()
                for ii in range(0,len(a)):
                        f=a[ii][0]
                        if str(book_id)==str(f):
                                bo=1
                                break
                        else:
                                bo=''
                if bo!=1:
                        reply=QMessageBox.warning(main_win,"Warning","Book id was not found! ",QMessageBox.Ok)
                        return
               

               
                        
                book_quantity_id=main_win.book_quantity_id.text()
                if book_quantity_id==None:
                        return
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT book_quantity_id FROM book_quantity_details;'''
                c.execute(query)
                a = c.fetchall()
                conn.close()
                boo=''
                for ii in range(0,len(a)):
                        f=a[ii][0]
                        if str(book_quantity_id)==str(f):
                                boo=1
                                break
                        else:
                                boo=''
                if boo==1:
                        reply=QMessageBox.warning(main_win,"Warning","Book Quantity id already added! ",QMessageBox.Ok)
                        return
                
                description=main_win.description_2.toPlainText()

                ds=datetime.datetime.now().day
                ms=datetime.datetime.now().month
                ys=datetime.datetime.now().year
                ds=int(ds)
                ms=int(ms)
                ys=int(ys)
                d=QDate(ys,ms,ds)
                d=d.toPyDate()
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                c.execute("INSERT INTO book_quantity_details(book_quantity_id,inserted_on,book_id,description)VALUES(?,?,?,?)",(book_quantity_id,d,book_id,description))
                conn.commit()
                conn.close()
                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('SELECT count(book_id) FROM book_quantity_details where book_id="{}"'.format(book_id))         
                c.execute(query)
                at=c.fetchone()
                conn.commit()
                conn.close()
                qt=int(at[0])
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('UPDATE book_details SET quantity="{}" where book_id="{}"'.format(qt,book_id))         
                c.execute(query)
                conn.commit()
                conn.close()
                
                main_win.book_quantity_id.setText('')
                main_win.description_2.clear()
                refresh_book_quantity_details_table()
                if main_win.add_book_quantity_pop_up.isVisible():
                        sound_3()
                        main_win.add_book_quantity_pop_up.hide()
                        reply=QMessageBox.information(main_win,"Success","Book quantity details was successfully added ! ",QMessageBox.Ok)   
        
        except Exception as e:
                print("ooo")
                print(e)


def clear_book_quantity_details_table():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM book_quantity_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.manage_book_quantity_table.rowCount()
        for Current_row in range(0,rowCount):
                        main_win.manage_book_quantity_table.removeRow(0)                        



def refresh_book_quantity_details_table():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM book_quantity_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.manage_book_quantity_table.rowCount()
                rowCount_2=len(a)
                clear_book_quantity_details_table()
                
                for Current_row in range(0, rowCount_2):
                        main_win.manage_book_quantity_table.insertRow(Current_row)
                rowPosition=0
                rowCount=main_win.manage_book_quantity_table.rowCount()
                rowCount_2=len(a)



                for rowPosition in range(0,rowCount_2):
                        for column,item in enumerate(a[rowPosition]):

                                if item==None:
                                        continue
                                main_win.manage_book_quantity_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
                return

        except Exception as e:
                print("Quantity Refresh self detect")
                print(e)

def delete_book_quantity_process():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM book_quantity_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.manage_book_quantity_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.manage_book_quantity_table.item(currentIndex, 0)
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('DELETE from book_quantity_details where book_quantity_id="{}"'.format(currentIndex_2))
                        c.execute(query)
                        a = c.fetchone()
                        conn.commit()
                        conn.close()
                        main_win.manage_book_quantity_table.removeRow(currentIndex)
                        refresh_book_quantity_details_table()
        except Exception as e:
                print("e")
                print(e)

def manage_books_quantity_table_changed():
        try:
                t_row=main_win.manage_book_quantity_table.rowCount()
                t_column=main_win.manage_book_quantity_table.columnCount()
                for i in range(0,t_row+1):
                        row_name=main_win.manage_book_quantity_table.item(i, 0)
                        if row_name==None:
                                break
                        row_name=row_name.text()
                        for j in range(0, t_column):
                                data=main_win.manage_book_quantity_table.item(i,j).text()
                                if data==None:
                                        break
                                column_name = main_win.manage_book_quantity_table.horizontalHeaderItem(j).text()
                                if column_name==-1:
                                        break
                                column_name=column_name.lower()
                                column_name=column_name.replace(' ','_')
                                if column_name=='book_quantity_id':
                                        continue
                                conn = sqlite3.connect('application_data/databases/clms_database.db')
                                c = conn.cursor()
                                query = ('UPDATE book_quantity_details SET {}="{}" where book_quantity_id="{}"'.format(column_name,data,row_name))
                                c.execute(query)
                                conn.commit()
                                conn.close()
        except Exception as e:
                print("here")
                print(e)
        finally:
                refresh_book_quantity_details_table()
                sound_3()
                reply=QMessageBox.information(main_win,"Success",("The book quantity table was saved! "),QMessageBox.Ok)
                        
        




def change_user_name_ui():
        clear_pop_up()
        main_win.change_user_name_pop_up.show()
        sound_4()
def change_user_name_process():
        try:
                user_name=main_win.user_name.text()
                if user_name=='':
                        sound_2()
                        main_win.error_message_5.setText("Please enter the user name !")
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                login_status="3"
                query = ('UPDATE login_details SET user_name ="{}"WHERE user_id=1;'.format(user_name))
                c.execute(query)
                conn.commit()
                conn.close()
                sound_3()
                clear_pop_up()
                reply=QMessageBox.information(main_win,"Success","Your user name has been successfully updated ! ",QMessageBox.Ok)
        except Exception as e:
                sound_2()
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)
def change_user_name_pop_up_close():        
        if main_win.change_user_name_pop_up.isVisible():
                main_win.change_user_name_pop_up.hide()
                sound_7()
     
        
def change_email_ui():
        clear_pop_up()
        main_win.change_email_pop_up.show()
        sound_4()
def change_email_process():
        try:
                email=main_win.email.text()
                if email=='':
                        sound_2()
                        main_win.error_message_4.setText("Please enter the email !")
                        return
                regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
                if not re.fullmatch(regex, email):
                        sound_6()
                        main_win.error_message_4.setText("Please enter the valid email !")
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                login_status="3"
                query = ('UPDATE login_details SET email ="{}"WHERE user_id=1;'.format(email))
                c.execute(query)
                conn.commit()
                conn.close()
                sound_3()
                clear_pop_up()
                reply=QMessageBox.information(main_win,"Success","Your email has been successfully updated ! ",QMessageBox.Ok)
        except Exception as e:
                sound_2()
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)
def change_email_pop_up_close():        
        if main_win.change_email_pop_up.isVisible():
                main_win.change_email_pop_up.hide()
                sound_7()
     
        
def change_password_ui():
        clear_pop_up()
        main_win.change_password_pop_up.show()
        sound_4()
def change_password_process():
        try:
                pass_1=main_win.password_1.text()
                pass_2=main_win.password_2.text()
                if pass_1=='':
                        sound_2()
                        main_win.error_message_3.setText("Please enter the password_1 !")
                        return
                if pass_2=='':
                        sound_2()
                        main_win.error_message_3.setText("Please enter the password_2 !")
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                login_status="3"
                query = ('UPDATE login_details SET password_1 ="{}", password_2="{}" WHERE user_id=1;'.format(pass_1,pass_2))
                c.execute(query)
                conn.commit()
                conn.close()
                sound_3()
                clear_pop_up()
                reply=QMessageBox.information(main_win,"Success","Your Password has been successfully updated ! ",QMessageBox.Ok)
        except Exception as e:
                sound_2()
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)        
def change_password_pop_up_close():        
        if main_win.change_password_pop_up.isVisible():
                main_win.change_password_pop_up.hide()
                sound_7()
     


def settings_ui():
        try:
                refresh_settings()
                sound_1()
                clear_main_ui()
                main_win.setWindowTitle("Settings - College Library Management System")
                main_win.settings_frame.show()
                set_button_color_for_active_window()
                main_win.gray.toggled.connect(lambda:change_themes('1'))
                main_win.gold.toggled.connect(lambda:change_themes('2'))
                main_win.brown.toggled.connect(lambda:change_themes('3'))
                main_win.violet.toggled.connect(lambda:change_themes('4'))
                main_win.orange.toggled.connect(lambda:change_themes('5'))
                main_win.pink.toggled.connect(lambda:change_themes('6'))
                main_win.red.toggled.connect(lambda:change_themes('7'))
                main_win.blue.toggled.connect(lambda:change_themes('8'))
                main_win.green.toggled.connect(lambda:change_themes('9'))
                
                main_win.studentLoginAllow.toggled.connect(lambda:change_student_login('1'))
                main_win.studentLoginDeny.toggled.connect(lambda:change_student_login('2'))

                main_win.soundYes.toggled.connect(lambda:change_sound('1'))
                main_win.soundNo.toggled.connect(lambda:change_sound('2'))

                main_win.backupAllow.toggled.connect(lambda:change_backup('1'))
                main_win.backupDeny.toggled.connect(lambda:change_backup('2'))
        except Exception as e:
                reply=QMessageBox.error(main_win,"Error",str(e),QMessageBox.Ok)
                
def change_themes(x):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('UPDATE login_details SET themes ="{}" WHERE user_id=1;'.format(x))
        c.execute(query)
        conn.commit()
        conn.close()
        set_button_color_for_active_window()
        load_themes()

def change_student_login(x):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('UPDATE login_details SET student_login ="{}" WHERE user_id=1;'.format(x))
        c.execute(query)
        conn.commit()
        conn.close()

 
def change_sound(x):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('UPDATE login_details SET sound ="{}" WHERE user_id=1;'.format(x))
        c.execute(query)
        conn.commit()
        conn.close()

def change_backup(x):
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = ('UPDATE login_details SET backup ="{}" WHERE user_id=1;'.format(x))
        c.execute(query)
        conn.commit()
        conn.close()


        
                        
def refresh_settings():
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT login_status,sound,student_login,backup,themes FROM login_details WHERE user_id=1;'''

                c.execute(query)
                a = c.fetchone()
                conn.close()
                login_status=a[0]
                sound=a[1]
                student_login=a[2]
                backup=a[3]
                themes=a[4]
                if login_status=='2':
                        sound_2()
                        reply=QMessageBox.warning(main_win,"Warning","Access is denied!\nThis option only for Admin.",QMessageBox.Ok)
                        return
                if backup=='1':
                        main_win.backupAllow.setChecked(True)
                else:
                        main_win.backupDeny.setChecked(True)
                if student_login=='1':
                        main_win.studentLoginAllow.setChecked(True)
                else:
                        main_win.studentLoginDeny.setChecked(True)
                if sound=='1':
                        main_win.soundYes.setChecked(True)
                else:
                        main_win.soundNo.setChecked(True)

                match themes:
                        case '1':
                                main_win.gray.setChecked(True)
                        case '2':
                                main_win.gold.setChecked(True)
                        case '3':
                                main_win.brown.setChecked(True)
                        case '4':
                                main_win.violet.setChecked(True)
                        case '5':
                                main_win.orange.setChecked(True)
                        case '6':
                                main_win.pink.setChecked(True)
                        case '7':
                                main_win.red.setChecked(True)
                        case '8':
                                main_win.blue.setChecked(True)
                        case '9':
                                main_win.green.setChecked(True)
def student_profile_ui():
        try:
                sound_1()
                clear_main_ui()
                clear_student_profile()
                main_win.setWindowTitle("Student Profile - College Library Management System")
                main_win.student_profile_frame.show()
                main_win.personal_details.show()
                set_button_color_for_active_window()
                main_win.role_number_4.textChanged.connect(profile_show)
        except Exception as e:
                print("profile")
                print(e)
def profile_show():
        try:
                clear_book_history_table() 
                role_number=main_win.role_number_4.text()
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT roll_number,name,date_of_birth,phone,address,description,batch_id,course_id,father_name,mother_name,aadhar_number,religion,email,community,caste,anual_income,blood_group,father_occupation,mother_occupation,area_type,sslc_mark,sslc_passing_year,sslc_school_name,sslc_medium,sslc_school_type,hsc_mark,hsc_passing_year,hsc_school_name,hsc_medium,hsc_school_type,description FROM student_details;'''
                c.execute(query)
                b = c.fetchall()
                conn.close()
                if b==None:
                        return
                row=()
                z=()
                i=0
                
                for i,z in enumerate(b):
                    if z[0]==role_number:
                        break
                    else:
                        z=()
                i=0
                if z==():
                    return
                name=z[1]
                date_of_birth=z[2]
                phone=z[3]
                address=z[4]
                description=z[5]
                batch_id=z[6]
                course_id=z[7]
                
                father_name=z[8]
                mother_name=z[9]
                aadhar_number=z[10]
                religion=z[11]
                email=z[12]
                community=z[13]
                caste=z[14]
                anual_income=z[15]
                blood_group=z[16]
                father_occupation=z[17]
                mother_occupation=z[18]
                area_type=z[19]
                sslc_mark=z[20]
                sslc_passing_year=z[21]
                sslc_school_name=z[22]
                sslc_medium=z[23]
                sslc_school_type=z[24]
                hsc_mark=z[25]
                hsc_passing_year=z[26]
                hsc_school_name=z[27]
                hsc_medium=z[28]
                hsc_school_type=z[29]
                description=z[30]

                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''SELECT start_year,end_year FROM batch_details WHERE batch_id="{}" ;'''.format(batch_id))
                c.execute(query)
                a = c.fetchone()
                conn.close()
                if a==None:
                        return
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = ('''SELECT course_name FROM course_details WHERE course_id="{}" ;'''.format(course_id))
                c.execute(query)
                d = c.fetchone()
                conn.close()
                if d==None:
                        return
                start_year=a[0]
                end_year=a[1]
                batch_name=str(start_year+"-"+end_year)
                course_name=d[0]


                main_win.roll_number_5.setText(role_number)
                main_win.name_4.setText(name)
                main_win.course_4.setText(course_name)
                main_win.batch_3.setText(batch_name)
                main_win.date_of_birth_3.setText(date_of_birth)
                main_win.father_name_3.setText(father_name)
                main_win.mother_name_3.setText(mother_name)
                main_win.phone_3.setText(phone)
                main_win.email_4.setText(email)
                main_win.address_4.setText(address)
                main_win.aadhar_number_3.setText(aadhar_number)
                main_win.religion_3.setText(religion)
                main_win.community_3.setText(community)
                main_win.caste_3.setText(caste)
                main_win.anual_income_3.setText(anual_income)
                main_win.blood_group_2.setText(blood_group)
                main_win.father_occupation_3.setText(father_occupation)
                main_win.mother_occupation_3.setText(mother_occupation)
                main_win.area_type_3.setText(area_type)
                main_win.sslc_mark_3.setText(sslc_mark)
                main_win.sslc_passing_year_3.setText(sslc_passing_year)
                main_win.sslc_school_name_3.setText(sslc_school_name)
                main_win.sslc_medium_3.setText(sslc_medium)
                main_win.sslc_school_type_3.setText(sslc_school_type)
                main_win.hsc_mark_3.setText(hsc_mark)
                main_win.hsc_passing_year_3.setText(hsc_passing_year)
                main_win.hsc_school_name_3.setText(hsc_school_name)
                main_win.hsc_medium_3.setText(hsc_medium)
                main_win.hsc_school_type_3.setText(hsc_school_type)
                main_win.description_6.setText(description)
        except Exception as e:
                print(e)


def personal_details_ui():
        sound_4()
        clear_student_profile()
        main_win.student_profile_frame.show()
        main_win.personal_details.show()
        
def book_history_ui():
        if main_win.phone_3.text()=='':
                sound_2()
                reply=QMessageBox.warning(main_win,"Warning","Student details not found ! ",QMessageBox.Ok)
                return
        sound_4()
        roll_num=main_win.role_number_4.text()
        clear_student_profile()
        main_win.student_profile_frame.show()
        main_win.book_history.show()
        if roll_num=='':
                return
        else:
                refresh_book_history(roll_num)
def clear_book_history_process_2():
        main_win.role_number_4.setText('')
        clear_book_history_table()
def clear_book_history_table():
        try:
                
                roll_number=main_win.role_number_4.text()
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM issued_book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()
                ad=()
                rowCount=main_win.book__history_table.rowCount()
                
                
                for Current_row in range(0,rowCount):
                        main_win.book__history_table.removeRow(0)
                        
                main_win.roll_number_5.setText('')
                main_win.name_4.setText('')
                main_win.course_4.setText('')
                main_win.batch_3.setText('')
                main_win.date_of_birth_3.setText('')
                main_win.father_name_3.setText('')
                main_win.mother_name_3.setText('')
                main_win.phone_3.setText('')
                main_win.email_4.setText('')
                main_win.address_4.setText('')
                main_win.aadhar_number_3.setText('')
                main_win.religion_3.setText('')
                main_win.community_3.setText('')
                main_win.caste_3.setText('')
                main_win.anual_income_3.setText('')
                main_win.blood_group_2.setText('')
                main_win.father_occupation_3.setText('')
                main_win.mother_occupation_3.setText('')
                main_win.area_type_3.setText('')
                main_win.sslc_mark_3.setText('')
                main_win.sslc_passing_year_3.setText('')
                main_win.sslc_school_name_3.setText('')
                main_win.sslc_medium_3.setText('')
                main_win.sslc_school_type_3.setText('')
                main_win.hsc_mark_3.setText('')
                main_win.hsc_passing_year_3.setText('')
                main_win.hsc_school_name_3.setText('')
                main_win.hsc_medium_3.setText('')
                main_win.hsc_school_type_3.setText('')
                main_win.description_6.setText('')
        except Exception as e:
                print(e)
        
        
def refresh_book_history(roll_num):
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM issued_book_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                ad=()
                for i in range(0,len(a)):
                        if str(a[i][1])==str(roll_num):
                                ad+=(a[i],)
                                

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.book__history_table.rowCount()
                rowCount_2=len(ad)
                clear_defaulter_list()

                if not rowCount == rowCount_2:
                        for Current_row in range(rowCount, rowCount_2):
                                main_win.book__history_table.insertRow(Current_row)

                rowPosition=0
                rowCount=main_win.book__history_table.rowCount()
                rowCount_2=len(a)
                ds=datetime.datetime.now().day
                ms=datetime.datetime.now().month
                ys=datetime.datetime.now().year

                ds=int(ds)
                ms=int(ms)
                ys=int(ys)
                d=QDate(ys,ms,ds)
                d=d.toPyDate()
                out=0
                main_win.book__history_table.setStyleSheet("")
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(ad[rowPosition]):
                                if item==None:
                                        continue
                                main_win.book__history_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
                                        
                                dd=datetime.datetime.strptime(ad[rowPosition][4],'%d/%m/%Y').date()
                                out=dd-d
                                out=int(out.days)
                                if out<0:
                                        main_win.book__history_table.item(rowPosition,column).setBackground(QBrush(QColor(244,164,96)))
                                else:
                                        
                                        main_win.book__history_table.item(rowPosition,column).setBackground(QBrush(QColor(152,251,152)))
                                        
        #^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        except Exception as e:
                print(e)
        
def manage_course_ui():
        sound_4()
        clear_main_ui()
        main_win.setWindowTitle("Manage Course - College Library Management System")
        main_win.course_frame.show()
        set_button_color_for_active_window()
        try:
                refresh_course_details_table()
                main_win.save_course_btn.clicked.connect(manage_course_table_changed)
        except Exception as e:
                print("ii")
                print(e)


def add_course_process():
        clear_pop_up()
        main_win.add_course_pop_up.show()
        sound_4()
def close_add_course_process():
        
        main_win.course_name.setText('')
        main_win.total_year.setText('')
        main_win.description_3.clear()
        if main_win.add_course_pop_up.isVisible():
                main_win.add_course_pop_up.hide()
                sound_7()








def save_add_course_process():
        try:
                course_name=main_win.course_name.text()
                total_year=main_win.total_year.text()
                description_3=main_win.description_3.toPlainText()

                
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                c.execute("INSERT INTO course_details(course_name,total_year,description)VALUES(?,?,?)",(course_name,total_year,description_3))
                conn.commit()
                conn.close()
                
                refresh_course_details_table()
                main_win.course_name.setText('')
                main_win.total_year.setText('')
                main_win.description_3.clear()
				
                if main_win.add_course_pop_up.isVisible():
                        sound_3()
                        main_win.add_course_pop_up.hide()
                        reply=QMessageBox.information(main_win,"Success","Book details was successfully added ! ",QMessageBox.Ok)   
        
        except Exception as e:
                print("ooo")
                print(e)




def clear_course_details_table():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM course_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.manage_course_table.rowCount()
        rowCount_2=len(a)
        if not rowCount == rowCount_2:
                for Current_row in range(rowCount, rowCount_2):
                        main_win.manage_course_table.removeRow(Current_row)                        






def refresh_course_details_table():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM course_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.manage_course_table.rowCount()
                rowCount_2=len(a)
                clear_course_details_table()

                if not rowCount == rowCount_2:
                        for Current_row in range(rowCount, rowCount_2):
                                main_win.manage_course_table.insertRow(Current_row)
                rowPosition=0
                rowCount=main_win.manage_course_table.rowCount()
                rowCount_2=len(a)
                for rowPosition in range(0,rowCount_2):
                        for column,item in enumerate(a[rowPosition]):
                                if item==None:
                                        return
                                main_win.manage_course_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))

        except Exception as e:
                print("Refresh self detect")
                print(e)





def delete_course_process():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM course_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.manage_course_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.manage_course_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        #$%^

                        query = ('SELECT batch_id FROM batch_details WHERE course_id="{}"'.format(currentIndex_2))
                        c.execute(query)
                        aws = c.fetchall()
                        
                        query = ('DELETE from course_details where course_id={}'.format(currentIndex_2))
                        c.execute(query)
                        a = c.fetchone()
                        conn.commit()
                        conn.close()
                        main_win.manage_course_table.removeRow(currentIndex) 
                        refresh_course_details_table()
        except Exception as e:
                print("e")
                print(e)



def manage_course_table_changed():
        
        try:
                t_row=main_win.manage_course_table.rowCount()
                t_column=main_win.manage_course_table.columnCount()
                for i in range(0,t_row+1):
                        row_name=main_win.manage_course_table.item(i, 0)
                        if row_name==None:
                                break
                        row_name=row_name.text()
                        for j in range(0, t_column):
                                data=main_win.manage_course_table.item(i,j).text()
                                if data==None:
                                        break
                                column_name = main_win.manage_course_table.horizontalHeaderItem(j).text()
                                if column_name==-1:
                                        break
                                column_name=column_name.lower()
                                column_name=column_name.replace(' ','_')
                                if column_name=='course_id':
                                        continue
                                conn = sqlite3.connect('application_data/databases/clms_database.db')
                                c = conn.cursor()
                                query = ('UPDATE course_details SET {}="{}" where course_id="{}"'.format(column_name,data,row_name))
                                c.execute(query)
                                conn.commit()
                                conn.close()
        except Exception as e:
                print("here")
                print(e)
        finally:
                refresh_course_details_table()
                sound_3()
                reply=QMessageBox.information(main_win,"Success",("The course table was saved! "),QMessageBox.Ok)
                        
        











def manage_batch_ui():
        sound_1()
        clear_main_ui()
        main_win.setWindowTitle("Manage Batch - College Library Management System")
        main_win.batch_frame.show()
        set_button_color_for_active_window()
        try:
                refresh_batch_details_table()
                main_win.save_batch_btn.clicked.connect(manage_batch_table_changed)
        except Exception as e:
                print("ii")
                print(e)
        
def add_batch_pop_up():
        
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT course_name FROM course_details;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()
        if a is None:
                reply=QMessageBox.warning(main_win,"Warning","Please add course details.\nCourse table is empty! ",QMessageBox.Ok)
                return
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT course_name FROM course_details;'''
        c.execute(query)
        ass = c.fetchall()
        conn.close()
        clear_pop_up()
        try:
                qq=()
                for i,d in enumerate(ass):
                        print(d[0])
                        qq+=(d[0],)
                        main_win.course_2.clear()
                        main_win.course_2.addItems(qq)
        except Exception as e:
                print(e)
        
        main_win.add_batch_pop_up.show()
        sound_4()

def close_add_batch_process():
        main_win.start_year.setText('')
        main_win.end_year.setText('')
        main_win.course_2.setCurrentText("Select Course")
        main_win.description_4.clear()
        if main_win.add_batch_pop_up.isVisible():
                main_win.add_batch_pop_up.hide()
                sound_7()
def save_add_batch_process():
        try:

                
                start_year=main_win.start_year.text()
                end_year=main_win.end_year.text()
                course_2=main_win.course_2.currentText()
                course_id=course_name_to_id(course_2)
                #$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                #course_id get from text
                description_4=main_win.description_4.toPlainText()


                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                
                
                c.execute("INSERT INTO batch_details(start_year,end_year,course_id,description)VALUES(?,?,?,?)",(start_year,end_year,course_id,description_4))

                conn.commit()
                conn.close()
                
                refresh_course_details_table()
                main_win.start_year.setText('')
                main_win.end_year.setText('')
                main_win.course_2.setCurrentText("Select Course")
                main_win.description_4.clear()

                refresh_batch_details_table()
				
                if main_win.add_batch_pop_up.isVisible():
                        sound_3()
                        main_win.add_batch_pop_up.hide()
                reply=QMessageBox.information(main_win,"Success","Batch details was successfully added ! ",QMessageBox.Ok)   
        
        except Exception as e:
                print("ooo")
                print(e)






def clear_batch_details_table():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM batch_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        
        rowCount=0
        rowCount_2=0
        rowPosition=0
        rowCount=main_win.manage_batch_table.rowCount()
        rowCount_2=len(a)
        for i in range(0,rowCount):
                main_win.manage_batch_table.removeRow(0)
        





def refresh_batch_details_table():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                query = '''SELECT * FROM batch_details;'''
                c.execute(query)
                a=c.fetchall()
                conn.close()

                rowCount=0
                rowCount_2=0
                rowPosition=0
                rowCount=main_win.manage_batch_table.rowCount()
                rowCount_2=len(a)
                clear_batch_details_table()
                for Current_row in range(0, rowCount_2):
                        main_win.manage_batch_table.insertRow(Current_row)
                rowPosition=0
                rowCount=main_win.manage_batch_table.rowCount()
                rowCount_2=len(a)
                for rowPosition in range(0,rowCount):
                        for column,item in enumerate(a[rowPosition]):
                                if item==None:
                                        continue                                
                                main_win.manage_batch_table.setItem(rowPosition,column,QTableWidgetItem(str(item)))
               

        except Exception as e:
                print("Refresh self detect")
                print(e)




def delete_batch_process():
        conn = sqlite3.connect('application_data/databases/clms_database.db')
        c = conn.cursor()
        query = '''SELECT * FROM batch_details;'''
        c.execute(query)
        a=c.fetchall()
        conn.close()
        try:
                last_row=len(a)
                currentIndex=main_win.manage_batch_table.currentRow()
                if currentIndex==-1:
                        sound_2()
                        reply=QMessageBox.critical(main_win,"Error","Please select the row!",QMessageBox.Ok)
                        return
                
                firstColumnInRow = main_win.manage_batch_table.item(currentIndex, 0) # returns QTableWidgetItem
                currentIndex_2 = firstColumnInRow.text()

             
                sound_2()
                if currentIndex==last_row-1:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete last row?",QMessageBox.Yes | QMessageBox.Cancel)
                else:
                        reply=QMessageBox.warning(main_win,"Warning","Do you want to delete the row?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        conn = sqlite3.connect('application_data/databases/clms_database.db')
                        c = conn.cursor()
                        query = ('DELETE from batch_details where batch_id={}'.format(currentIndex_2))
                        c.execute(query)
                        a = c.fetchone()
                        conn.commit()
                        conn.close()
                        main_win.manage_batch_table.removeRow(currentIndex) 
                        refresh_batch_details_table()
        except Exception as e:
                print("e")
                print(e)


def manage_batch_table_changed():
        try:
                
                rowCount=main_win.manage_batch_table.rowCount()
                if rowCount==0:
                        sound_2()
                        return
                t_row=main_win.manage_batch_table.rowCount()
                t_column=main_win.manage_batch_table.columnCount()
                for i in range(0,t_row+1):
                        row_name=main_win.manage_batch_table.item(i, 0)
                        if row_name==None:
                                break
                        row_name=row_name.text()
                        for j in range(0, t_column-1):
                                data=main_win.manage_batch_table.item(i,j).text()
                                if data==None:
                                        break
                                column_name = main_win.manage_batch_table.horizontalHeaderItem(j).text()
                                if column_name==-1:
                                        break
                                column_name=column_name.lower()
                                column_name=column_name.replace(' ','_')
                                if column_name=='batch_id':
                                        continue
                                if column_name=='total_student':
                                        continue
                                conn = sqlite3.connect('application_data/databases/clms_database.db')
                                c = conn.cursor()
                                query = ('UPDATE batch_details SET {}="{}" where batch_id="{}"'.format(column_name,data,row_name))
                                c.execute(query)
                                conn.commit()
                                conn.close()
                sound_3()
                reply=QMessageBox.information(main_win,"Success",("The batch table was saved! "),QMessageBox.Ok)
                                
        except Exception as e:
                print("here")
                print(e)      
        





def delete_back_up_process():
        try:
                sound_2()
                if not os.path.exists("application_data/databases/clms_back_up_database.db"):
                         reply=QMessageBox.information(main_win,"Information","The backup data  was already deleted!",QMessageBox.Ok)
                         return
                reply=QMessageBox.warning(main_win,"Warning","Note:You cannot recover the backup data, the entire backup database is deleted!\nDo you want to delete all the data?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        if os.path.exists("application_data/databases/clms_back_up_database.db"):
                                os.remove("application_data/databases/clms_back_up_database.db")
                                sound_3()
                                reply=QMessageBox.information(main_win,"Success","The backup data was deleted!",QMessageBox.Ok)
        except Exception as e:
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)
        






def restore_process():
        shutil.copyfile("application_data/databases/clms_back_up_database.db","application_data/databases/clms_database.db")
        sound_3()
        reply=QMessageBox.information(main_win,"Success","Database restore successfully finished!",QMessageBox.Ok)

                



                
                
def reset_process():
        try:
                sound_2()
                reply=QMessageBox.warning(main_win,"Warning","Note:You cannot recover the data, the entire database is deleted!\nDo you want to delete all the data?",QMessageBox.Yes | QMessageBox.Cancel)
                if reply==16384:
                        if os.path.exists("application_data/databases/clms_database.db"):
                                if os.path.exists("application_data/databases/clms_back_up_database.db"):
                                        os.remove("application_data/databases/clms_back_up_database.db")
                                os.remove("application_data/databases/clms_database.db")
                                reply=QMessageBox.information(main_win,"Success","The CLMS has been reseted!\nThe CLMS want to exit!\nYou want to start CLMS again!",QMessageBox.Ok)
                                close_all_window()
                                sys.exit()
        except Exception as e:
                reply=QMessageBox.critical(main_win,"Error",str(e),QMessageBox.Ok)

def back_up(sound):
        conn = sqlite3.connect('application_data/databases/clms_database.db')#database connect
        c = conn.cursor()
        query = '''SELECT backup FROM login_details WHERE user_id=1;'''
        c.execute(query)
        a = c.fetchone()
        conn.close()
        back_up=a[0]
        if back_up=='1':
                shutil.copyfile("application_data/databases/clms_database.db","application_data/databases/clms_back_up_database.db")
                if sound==1:
                        sound_3()
                        reply=QMessageBox.information(main_win,"Success","Database backup successfully finished!",QMessageBox.Ok)

                
def close_all_window():
        if login_win.isActiveWindow():
                login_win.destroy()
        if main_win.isActiveWindow():
                main_win.destroy()
        if sign_up_win.isActiveWindow():
                sign_up_win.destroy()
        
        
        
def clear_main_ui():
        clear_pop_up()
        if main_win.home_frame.isVisible():
                main_win.home_frame.hide()
        if main_win.issue_book_frame.isVisible():
                main_win.issue_book_frame.hide()
        if main_win.return_book_frame.isVisible():
                main_win.return_book_frame.hide()
        if main_win.issued_book_details_frame.isVisible():
                main_win.issued_book_details_frame.hide()
        if main_win.defaulter_frame.isVisible():
                main_win.defaulter_frame.hide()
        if main_win.manage_students_frame.isVisible():
                main_win.manage_students_frame.hide()
        if main_win.manage_books_frame.isVisible():
                main_win.manage_books_frame.hide()
        if main_win.manage_book_quantity_frame.isVisible():
                main_win.manage_book_quantity_frame.hide()
        if main_win.settings_frame.isVisible():
                main_win.settings_frame.hide()
        if main_win.batch_frame.isVisible():
                main_win.batch_frame.hide()

                
        if main_win.student_profile_frame.isVisible():
                main_win.student_profile_frame.hide()
                
        if main_win.course_frame.isVisible():
                main_win.course_frame.hide()
        if main_win.batch_frame.isVisible():
                main_win.batch_frame.hide()


def clear_student_profile():
        if main_win.book_history.isVisible():
                main_win.book_history.hide()
        if main_win.personal_details.isVisible():
                main_win.personal_details.hide()
                
def clear_pop_up():
        if main_win.add_book_quantity_pop_up.isVisible():
                main_win.add_book_quantity_pop_up.hide()
        if main_win.change_user_name_pop_up.isVisible():
                main_win.change_user_name_pop_up.hide()
        if main_win.change_password_pop_up.isVisible():
                main_win.change_password_pop_up.hide()
        if main_win.change_email_pop_up.isVisible():
                main_win.change_email_pop_up.hide()
        if main_win.manage_student_pop_up.isVisible():
                main_win.manage_student_pop_up.hide()
        if main_win.filter_pop_up.isVisible():
                main_win.filter_pop_up.hide()
        if main_win.add_course_pop_up.isVisible():
                main_win.add_course_pop_up.hide()
        if main_win.add_batch_pop_up.isVisible():
                main_win.add_batch_pop_up.hide()
        if main_win.add_book_pop_up.isVisible():
                main_win.add_book_pop_up.hide()
                
        
                
def clear_manage_student_pop():
        if main_win.get_details_1.isVisible():
                main_win.get_details_1.hide()
        if main_win.get_details_2.isVisible():
                main_win.get_details_2.hide()
        if main_win.get_details_3.isVisible():
                main_win.get_details_3.hide()
        if main_win.get_details_4.isVisible():
                main_win.get_details_4.hide()
                
                
def  load_themes():
        conn = sqlite3.connect('application_data/databases/clms_database.db')#database connect
        c = conn.cursor()
        query = '''SELECT themes FROM login_details WHERE user_id=1;'''
        c.execute(query)
        b = c.fetchone()
        conn.close()
        
        data=b[0]
        match data:
                case '1':
                        bg_color="lightgray"
                        c="application_ui/css/gray_btn_non_active_color.css"
                case '2':
                        bg_color="#ffdf37"#gold
                        c="application_ui/css/gold_btn_non_active_color.css"
                case '3':
                        bg_color="#bc9e82"#brown
                        c="application_ui/css/brown_btn_non_active_color.css"
                case '4':
                        bg_color="violet"#violet
                        c="application_ui/css/violet_btn_non_active_color.css"
                case '5':
                        bg_color="#ffd580"#orange
                        c="application_ui/css/orange_btn_non_active_color.css"
                case '6':
                        bg_color="#ffb6c1"#pink
                        c="application_ui/css/pink_btn_non_active_color.css"
                case '7':
                        bg_color="#fa8072"#red
                        c="application_ui/css/red_btn_non_active_color.css"
                case '8':
                        bg_color="#55aaff"#blue
                        c="application_ui/css/blue_btn_non_active_color.css"
                case '9':
                        bg_color="#50c878"#green
                        c="application_ui/css/green_btn_non_active_color.css"
                case default:
                        bg_color="lightgray"#gray
                        c="application_ui/css/gray_btn_non_active_color.css"

        if main_win.isActiveWindow():
                main_win.setStyleSheet(str("QMainWindow {background-color:"+bg_color+'}'))
        if login_win.isActiveWindow():
                login_win.setStyleSheet(str("QWidget {background-color:"+bg_color+'}'))

        return c
                
def set_button_color_for_active_window():
        c=load_themes()
        set_button_color_for_non_active_window(c)
        active_btn_color="application_ui/css/active_btn_color.css"
        if main_win.isActiveWindow():
                if main_win.home_frame.isVisible():
                        main_win.home_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.issue_book_frame.isVisible():
                        main_win.issue_book_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.return_book_frame.isVisible():
                        main_win.return_book_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.issued_book_details_frame.isVisible():
                        main_win.issued_book_details_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.defaulter_frame.isVisible():
                        main_win.defaulter_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.manage_students_frame.isVisible():
                        main_win.manage_students_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.manage_books_frame.isVisible():
                        main_win.manage_books_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.manage_book_quantity_frame.isVisible():
                        main_win.manage_book_quantity_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.settings_frame.isVisible():
                        main_win.settings_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.batch_frame.isVisible():
                        main_win.batch_and_course_btn.setStyleSheet(open(active_btn_color).read())
                if main_win.student_profile_frame.isVisible():
                        main_win.student_profile_btn_7.setStyleSheet(open(active_btn_color).read())

 


def set_button_color_for_non_active_window(c):
        if main_win.isActiveWindow():
                main_win.home_btn.setStyleSheet(open(c).read())
                main_win.issue_book_btn.setStyleSheet(open(c).read())
                main_win.return_book_btn.setStyleSheet(open(c).read())
                main_win.issued_book_details_btn.setStyleSheet(open(c).read())
                main_win.defaulter_btn.setStyleSheet(open(c).read())
                main_win.manage_students_btn.setStyleSheet(open(c).read())
                main_win.manage_books_btn.setStyleSheet(open(c).read())
                main_win.manage_book_quantity_btn.setStyleSheet(open(c).read())
                main_win.settings_btn.setStyleSheet(open(c).read())
                main_win.batch_and_course_btn.setStyleSheet(open(c).read())
                main_win.student_profile_btn_7.setStyleSheet(open(c).read())
                main_win.exit_btn.setStyleSheet(open(c).read())
                
        if login_win.isActiveWindow():
                login_win.forgot_password.setStyleSheet(open(c).read())
                login_win.login_btn.setStyleSheet(open(c).read())
                


      
def exist_life():
  sys.exit(app.exec_())
  



def on_sys_exit():
        try:
                conn = sqlite3.connect('application_data/databases/clms_database.db')
                c = conn.cursor()
                login_status="3"
                query = ('''UPDATE login_details SET login_status = "{}" WHERE user_id=1;'''.format(login_status))
                c.execute(query)
                conn.commit()
                conn.close()
                sound_7()
                back_up(0)
                notification.notify(title='Logout Success !', message='Thanks For Using CLMS !',app_icon="application_ui/images/icon.ico",timeout=20)
                close_all_window()
                sys.exit()
        except Exception as e:
                print(e)
        
if __name__ == '__main__':
        try:
                load()
        except Exception as e:
            if main_win.isActiveWindow():
                    sound_2()
                    reply=QMessageBox.critical(main_win,"Exception",str(e),QMessageBox.Ok)
