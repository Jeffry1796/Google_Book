from flask import Flask, redirect, url_for, session
from authlib.integrations.flask_client import OAuth
import requests
import pandas as pd
import time

app = Flask(__name__)
app.secret_key = 'random secret'

## Line 11-23 Google Authentification Parameter
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id='674920413266-cf3ta4172t2g6epapjc60co5179j12gg.apps.googleusercontent.com',
    client_secret='_aEzVHICmhsvYFN9F9LOo4L1',
    access_token_url='https://accounts.google.com/o/oauth2/token',
    access_token_params=None,
    authorize_url='https://accounts.google.com/o/oauth2/auth',
    authorize_params=None,
    api_base_url='https://www.googleapis.com/oauth2/v1/',
    userinfo_endpoint='https://www.googleapis.com/auth/books',  # This is only needed if using openId to fetch user info
    client_kwargs={'scope': 'https://www.googleapis.com/auth/books'},
)

@app.route('/')
def hello_world():
    return 'Type /login in address bar after your local IP e.g "http://127.0.0.1:5000/login"'

## Line 30-34 Google Authentification using your email address that you have configure to access Google Book API
@app.route('/login')
def login():
    google = oauth.create_client('google')  # create the google oauth client
    redirect_uri = url_for('authorize', _external=True)
    return google.authorize_redirect(redirect_uri)

## Line 37-133 Access Google Book API and Export data to Excel
@app.route('/authorize')
def authorize():
    google = oauth.create_client('google')  # create the google oauth client
    token = google.authorize_access_token()  # Access token from google (needed to get user info)        
    return auth_complete(google)

@app.route('/auth_complete')
def auth_complete(google):
    url_api = 'https://www.googleapis.com/books/v1/mylibrary/bookshelves?key=AIzaSyCI-44UQjQ1ppM_fOM4m-rJkeqbCbwDS0w'
    res_data = google.get(url = url_api)
    data = res_data.json()

    volum_count,idx,dfx = [],[],[]

    ##Input bookshelf name that you plan to export the data
    tit_bookshelf = ['My Google eBooks','To read','KUNTUL']
    
    for tit_book in range (len(data['items'])):
        for x in tit_bookshelf:
            if data['items'][tit_book]['title'].upper() == x.upper():
                dfx.append(x)
                volum_count.append(data['items'][tit_book]['volumeCount'])
                idx.append(data['items'][tit_book]['id'])
                
    max_dat = 40
    writer = pd.ExcelWriter('Google Books Data.xlsx',engine='xlsxwriter')
   
    for cnt in range(len(volum_count)):
        list_author,list_title,list_date,list_page,list_link = [],[],[],[],[] 
        start = 0
        
        loop = int(volum_count[cnt]) / 40
        max_data_xls = volum_count[cnt]
        
        for d in range (int(loop)+1):
            check_url = 'https://www.googleapis.com/books/v1/mylibrary/bookshelves/'+str(idx[cnt])+'/volumes?key=AIzaSyCI-44UQjQ1ppM_fOM4m-rJkeqbCbwDS0w&startIndex='+str(start)+'&maxResults=40'
            check_r = google.get(url = check_url)
            check_data = check_r.json()

            if max_data_xls >= 40:
                max_dat = 40
            else:
                max_dat = max_data_xls

            for x in range (max_dat):
                try:
                    author = check_data['items'][x]['volumeInfo']['authors']
                except:
                    author = ['NO DATA']

                try:
                    title = check_data['items'][x]['volumeInfo']['title']
                except:
                    title = 'NO DATA'

                try:
                    date = check_data['items'][x]['volumeInfo']['publishedDate']
                except:
                    date = 'NO DATA'

                try:
                    page = check_data['items'][x]['volumeInfo']['pageCount']
                except:
                    page = 'NO DATA'

                try:        
                    weblink = check_data['items'][x]['volumeInfo']['previewLink']
                except:
                    weblink = 'NO DATA'

                if author == '':
                    author = ['NO DATA']
                if title == '':
                    title = 'NO DATA'
                if date == '':
                    date = 'NO DATA'
                if page == '':
                    page == 'NO DATA'
                if weblink == '':
                    weblink == 'NO DATA'
                    
                list_author.append(','.join(author))
                list_title.append(title)
                list_date.append(date)
                list_page.append(page)
                list_link.append(weblink)

                print('Title: '+title)
##                print('Author: '+','.join(author))
##                print('Date: ' + date)
##                print('Page: ' +str(page))
                print('\n')

            start = start+40
            max_data_xls = max_data_xls - 40

            if int(max_data_xls) < 0:
                break

        ##Always close your excel file if you export the new excel file with the same name
        try:
            data_excel = {'Title': list_title, 'Author':list_author, 'Date':list_date, 'Page':list_page, 'Link':list_link}
            dfx[cnt] = pd.DataFrame(data_excel, columns=['Title','Author','Date','Page','Link'])
            dfx[cnt].to_excel (writer, sheet_name=tit_bookshelf[cnt] ,index = False, header=True)
            time.sleep(2)
        except Exception as e:
            print(e)
            return redirect ('/error')

    writer.save()
        
    return 'EXPORT DATA COMPLETE'        

@app.route('/error')
def error():
    return 'EXPORT DATA FAILED'           

if __name__ == "__main__":
    app.run()
    
##import requests
##import pandas as pd
##
####https://www.googleapis.com/books/v1/users/106317554425208272757/bookshelves/1001/volumes?key=AIzaSyCI-44UQjQ1ppM_fOM4m-rJkeqbCbwDS0w
##list_author,list_title,list_date,list_page = [],[],[],[]
##
##URL = 'https://www.googleapis.com/books/v1/users/106317554425208272757/bookshelves/1001?key=AIzaSyCI-44UQjQ1ppM_fOM4m-rJkeqbCbwDS0w'
##r = requests.get(url = URL)
##data = r.json()
##
##loop = data['volumeCount']/40
##
##start = 0
##max_api =  40
##max_dat = 40 #(max=40)
##
##for d in range (int(loop)+1):
##    check_url = 'https://www.googleapis.com/books/v1/users/106317554425208272757/bookshelves/1001/volumes?key=AIzaSyCI-44UQjQ1ppM_fOM4m-rJkeqbCbwDS0w&startIndex='+str(start)+'&maxResults='+str(max_dat)
##    check_r = requests.get(url = check_url)
##    check_data = check_r.json()
##
##    if data['volumeCount'] > max_dat:
##        max_dat = max_dat
##    else:
##        max_dat = data['volumeCount']
##        
##    for x in range (max_dat):
##        try:
##            author = check_data['items'][x]['volumeInfo']['authors']
##        except:
##            author = ['NO DATA']
##
##        try:
##            title = check_data['items'][x]['volumeInfo']['title']
##        except:
##            title = 'NO DATA'
##
##        try:
##            date = check_data['items'][x]['volumeInfo']['publishedDate']
##        except:
##            date = 'NO DATA'
##
##        try:
##            page = check_data['items'][x]['volumeInfo']['pageCount']
##        except:
##            page = 'NO DATA'
##
##        if author == '':
##            author = ['NO DATA']
##        if title == '':
##            title = 'NO DATA'
##        if date == '':
##            date = 'NO DATA'
##        if page == '':
##            page == 'NO DATA'
##
##        list_author.append(','.join(author))
##        list_title.append(title)
##        list_date.append(date)
##        list_page.append(page)
##
##        print('Title: '+title)
##        print('Author: '+','.join(author))
##        print('Date: ' + date)
##        print('Page: ' +str(page))
##        print('\n')
##
##    start = start+40
##    data['volumeCount'] = data['volumeCount'] - 40
##
##data_excel = {'Title': list_title, 'Author':list_author, 'Date':list_date, 'Page':list_page}
##df = pd.DataFrame(data_excel, columns=['Title','Author','Date','Page'])
##df.to_excel ('export_dataframe.xlsx', index = False, header=True)
