# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 20:45:30 2020

@author: dell
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 20:37:20 2020

@author: dell
"""
#import pygame
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser as wb
import os
import smtplib
import requests
#from pprint import pprint
from selenium import webdriver
from threading import Thread
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By 
import time, os 
from newscatcher import Newscatcher
import vlc
import struct
import ctypes
import subprocess
from time import sleep
import wolframalpha
#import connect_database
from sound import Sound
import pandas as pd
import requests
import pandas as pd
import pandas_ta as ta

from plotly.offline import plot
import plotly.graph_objs as go
from mpl_finance import candlestick_ohlc
import datetime as dt
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import json

walf_app_id='your api key'
from twilio.rest import Client

import win32com.client as win32

import cv2 

#pygame.init()
#
#display_width = 800
#display_height = 600
#
#black = (0,0,0)
#alpha = (0,88,255)
#white = (255,255,255)
#red = (200,0,0)
#green = (0,200,0)
#bright_red = (255,0,0)
#bright_green = (0,255,0)
#
#gameDisplay = pygame.display.set_mode((display_width,display_height))
#pygame.display.set_caption('GUI Speech Recognition')
#
#
#
#gameDisplay.fill(white)
#carImg = pygame.image.load('img.jpg')
#gameDisplay.blit(carImg,(0,0))


wb.register('firefox',
	None,
	wb.BackgroundBrowser("C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe"))




SPI_SETDESKWALLPAPER = 20
WALLPAPER_PATH = 'C:\\Users\\dell\\Pictures\\road.png'

chromedir='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe %s'
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=C:/Users/dell/AppData/Local/Google/Chrome/User Data")

#########Voice intialization ##########################
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[1].id)
engine.setProperty('voice',voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def welcome():
    speak("Hello There!")
    hour = int(datetime.datetime.now().hour)
    #print(hour)
    year = int(datetime.datetime.now().year)
    month = int(datetime.datetime.now().month)
    date = int(datetime.datetime.now().day)
    Time = datetime.datetime.now().strftime("%I:%M:%S") 
   
    if hour>=6 and hour<12:
        speak("Good Morning! the current time is")
        speak(Time)

    elif hour>=12 and hour<16:
        speak("Good Afternoon!the current time is")
        speak(Time)
    elif hour>=16 and hour<24:
        speak("Good Evening!the current time is")
        speak(Time)
    else:
        speak("Good Night!")

    speak("Wendy at your Service! how May I help You? ")
#welcome()
def takeCommand():

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold =1
        r.adjust_for_ambient_noise(source,duration=0.51)
        audio = r.listen(source,phrase_time_limit=4)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"You Said:{query}\n")
        
    
    except Exception as e:
        print(e)
        print("Say that again Please...")
        #speak("Say that again Please...")
        return "None"
    return query   

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('', '')
    server.sendmail('', to, content)
    server.close()

#def facebook():
#    driver = webdriver.Chrome('chromedriver.exe',chrome_options=options)#add the location of the chrome Drivers
#    driver.get("https://www.facebook.com/")#Add the webhost name
#    elem1 = driver.find_element_by_class("RNmpXc")
#    elem1.click()

def ticker(text):    
    ticker=pd.read_csv("tickers.csv")
    try:
        return ticker.loc[ticker['name'].str.contains(text), 'ticker'].iloc[0]
        
    except IndexError:
        pass

def youtube():
    driver = webdriver.Chrome('chromedriver.exe',chrome_options=options)
    driver.get("https://www.youtube.com/")#Add the webhost name
    time.sleep(10)
    try:
        speak("what would you like to search")
        name = takeCommand().title()
        elem1 = driver.find_element_by_name('search_query')
        elem1.click
        elem1.send_keys(name)
        elem2= driver.find_element_by_xpath('//*[@id="search-icon-legacy"]')
        elem2.click()
    except Exception:
        speak("sorry! I couldn't do that")
        
def headlines():
    website=Newscatcher('washingtonpost.com')
#   speak("you want me to read the headlines?")
#    msg = takeCommand().lower() 
#    if msg=="Yes" or "Yeah":
#        try:   
    results=website.headlines
    speak(results) 
#            quit()        
#        except Exception as e:
#            print(e)
    
def whatsapp():
    driver = webdriver.Chrome('chromedriver.exe',chrome_options=options)
    driver.get("https://web.whatsapp.com/")#Add the webhost name
    wait = WebDriverWait(driver, 60)
    time.sleep(10)
    try:
        speak("mention the name")
        name = takeCommand().title()
                #name=takeCommand().lower() 
        arg= '//*[contains(@title,"{}")]'.format(name)
        
            
        user = driver.find_element_by_xpath(arg)
        user.click()
        if name:
            speak("what to say")
                  
            msg = takeCommand().lower() 
            msg_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
            
            for i in range(1):
                    msg_box.send_keys(msg)
                    button = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]')
                    button.click()
    except Exception as e:
        print(e)
        #print("Say that again Please...")
        
        return None
  



def is_64_windows():
    """Find out how many bits is OS. """
    return struct.calcsize('P') * 8 == 64


def get_sys_parameters_info():
    """Based on if this is 32bit or 64bit returns correct version of SystemParametersInfo function. """
    return ctypes.windll.user32.SystemParametersInfoW if is_64_windows() \
        else ctypes.windll.user32.SystemParametersInfoA


def change_wallpaper():
    sys_parameters_info = get_sys_parameters_info()
    r = sys_parameters_info(SPI_SETDESKWALLPAPER, 0, WALLPAPER_PATH, 3)

    # When the SPI_SETDESKWALLPAPER flag is used,
    # SystemParametersInfo returns TRUE
    # unless there is an error (like when the specified file doesn't exist).
    if not r:
        print(ctypes.WinError())


def text_objects(text, font):
    textSurface = font.render(text, True, alpha)
    return textSurface, textSurface.get_rect()





def button(msg,x,y,w,h,ic,ac,action=None):
    mouse = pygame.mouse.get_pos()
    click = pygame.mouse.get_pressed()
    if x+w > mouse[0] > x and y+h > mouse[1] > y:
        pygame.draw.rect(gameDisplay, ac,(x,y,w,h))

        if click[0] == 1 and action != None:
           pass        
        else:
            pygame.draw.rect(gameDisplay, ic,(x,y,w,h))

        smallText = pygame.font.SysFont("comicsansms",20)
        textSurf, textRect = text_objects(msg, smallText)
        textRect.center = ( (x+(w/2)), (y+(h/2)) )
        gameDisplay.blit(textSurf, textRect)



def message_display(text):
    largeText = pygame.font.Font('freesansbold.ttf',30)
    TextSurf, TextRect = text_objects(text, largeText)
    TextRect.center = ((display_width/2),(display_height/2))
    gameDisplay.blit(TextSurf, TextRect)

    pygame.display.update()


def play_music():
#    while True:
#        playlist=['C:\\Users\\dell\\enya.mp3']
#        for song in playlist:
#            media_player = vlc.MediaPlayer(song)          
#            media_player.play()


    music_dir = 'C:\\Users\\dell\\AI_Sudoku\\music'
    songs = os.listdir(music_dir)
    print(songs)    
    os.startfile(os.path.join(music_dir, songs[0]))

def pkill (process_name):
    try:
        os.system("taskkill /f /im "+ process_name+".exe")
       # killed = os.system('tskill ' + process_name+'.exe')
    except (Exception) as e:
        killed = 0
    #return killed


def real_time_price(Name):
    url=('https://financialmodelingprep.com/api/v3/stock/real-time-price/'+Name+'')
    res = requests.get(url)
    
    data = res.json()
    return(data['price'])


def company_profile(Name):
    url =('https://financialmodelingprep.com/api/v3/company/profile/'+Name+'')
    res = requests.get(url)
    data = res.json()
    speak("The profile of the symbol "+Name+" is as follows")
    speak("Company name :{}".format(data["profile"]["companyName" ]))
    speak("industry :{}".format(data["profile"]["industry" ]))
    speak("description :{}".format(data["profile"]["description" ]))
    speak("Present share Price in market:{} dollars".format(data["profile"]["price"]))
    speak("Market Capitalization: {} Billion dollars".format(data["profile"]["mktCap" ]))
    speak("CEO :{}".format(data["profile"]["ceo" ]))
    speak("Beta :{}".format(data["profile"]["beta" ]))
    speak("Volume Average :{}Million".format(data["profile"]["volAvg" ]))


def financial_summary(Name):
    speak("Please hold while I open the financial summary of the page for you")
    url=('https://financialmodelingprep.com/financial-summary/'+Name+'')
    wb.get('firefox').open(url)
    speak("You can find every detail of the company in this page, Please feel free to navigate")


def phonecall():
    speak("what to say")
                  
    msg = takeCommand().lower() 
    account_sid = 'get your sid'
    auth_token = 'get your token id'
    client = Client(account_sid, auth_token)
    
    call = client.calls.create(
                            twiml='<Response><Say>'+msg+'</Say></Response>',
                            to=' the dialer',
                            from_='+ twilio number'
                        )

    print(call.sid)
    




def historical_price(Name):
        speak("Please give me one more second while I try to fetch a few technical indicators for the company taking on the last 120 days")
        url=('https://financialmodelingprep.com/api/v3/historical-price-full/'+Name+'')
        res = requests.get(url)
        
        data = res.json()
        data=data['historical'][-120:]
        bs = pd.DataFrame.from_dict(data)
        bs.set_index('date', inplace=True)
       # print(bs.keys)
        
        #bs['close'].plot()
    
        
        def cscheme(colors):
            aliases = {
                'BkBu': ['black', 'blue'],
                'gr': ['green', 'red'],
                'grays': ['silver', 'gray'],
                'mas': ['black', 'green', 'orange', 'red'],
            }
            aliases['default'] = aliases['gr']
            return aliases[colors]
        
        last_ = bs.shape[0]
        price_size=(8, 6) 
        
        def machart(kind, fast, medium, slow, append=True, last=last_, figsize=price_size, colors=cscheme('mas')):
            
            title = kind
            ma1 = bs.ta(kind=kind, length=fast, append=append)
            ma2 = bs.ta(kind=kind, length=medium, append=append)
            ma3 = bs.ta(kind=kind, length=slow, append=append)
            
            #figure1 = plt.Figure(figsize=(6,5), dpi=100)
            #ax1 = figure1.add_subplot(111)
            #bar1 = FigureCanvasTkAgg(figure1, tab3)
            #bar1.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
            madf = pd.concat([bs['close'], bs[[ma1.name, ma2.name, ma3.name]]], axis=1, sort=False).tail(last)
            madf.plot(figsize=figsize, title=title, color=colors, grid=True) 
            #ax1.set_title('MACD')
    
        def rsi_plot(bs):
    
            window_length=14
            close = bs['close']
            # Get the difference in price from previous step
            delta = close.diff()
            # Get rid of the first row, which is NaN since it did not have a previous 
            # row to calculate the differences
            delta = delta[1:] 
            
            # Make the positive gains (up) and negative gains (down) Series
            up, down = delta.copy(), delta.copy()
            up[up < 0] = 0
            down[down > 0] = 0
            
            # Calculate the EWMA
            roll_up1 = up.ewm(span=window_length).mean()
            roll_down1 = down.abs().ewm(span=window_length).mean()
            
            # Calculate the RSI based on EWMA
    #        RS1 = roll_up1 / roll_down1
    #        RSI1 = 100.0 - (100.0 / (1.0 + RS1))
            
            # Calculate the SMA
            roll_up2 = up.rolling(window_length).mean()
            roll_down2 = down.abs().rolling(window_length).mean()
            
            # Calculate the RSI based on SMA
            RS2 = roll_up2 / roll_down2
            RSI2 = 100.0 - (100.0 / (1.0 + RS2))
            
            # Compare graphically
            plt.figure(figsize=(8, 6))
#            figure2 = plt.Figure(figsize=(6,5), dpi=100)
#            ax2 = figure2.add_subplot(111)
#            bar2 = FigureCanvasTkAgg(figure2, tab3)
#            bar2.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
            #RSI1.plot()
            
            RSI2.plot()
            plt.legend(['RSI via SMA'])
            plt.axhline(y=30,     color='red',   linestyle='-')
            plt.axhline(y=70,     color='blue',  linestyle='-')
            plt.show()
            #ax2.set_title('RSI')
        def macd(bs):        
            recent=120
            ind_size = (8, 6)
            macddf = bs.ta.macd(fast=8, slow=21, signal=9, min_periods=None, append=True)
           # print(macddf)
            
            macddf[[macddf.columns[0], macddf.columns[2]]].tail(recent).plot(figsize=(16, 2), color=cscheme('BkBu'), linewidth=1.3)
           
            macddf[macddf.columns[1]].tail(recent).plot.area(figsize=ind_size, stacked=False, color=['silver'], linewidth=1, title="macd", grid=True).axhline(y=0, color="black", lw=1.1)
    
        def aroon(bs):
            arn=ta.aroon(bs['close'],length=None,offset=None)
            
            arn.plot()
        
        def bol_band(bs):
            b=ta.bbands(bs['close'], length=None, std=None, mamode=None, offset=None)
            bs['close'].plot()
            b=pd.concat([bs['close'],b['BBL_20'],b['BBM_20'],b['BBU_20']], axis=1, sort=False)
            
            b.plot()
            
        def stochos(bs):
            b=ta.stoch(bs['high'],bs['low'],bs['close'],fast_k=None, slow_k=None, slow_d=None, offset=None)
            
            b.plot()
    
        def chmf(bs):
            b=ta.cmf(bs['high'],bs['low'],bs['close'],bs['volume'],bs['open'],length=None, offset=None)
            
            b.plot()
     
        def cdlp(bs):
            bs = bs[['open', 'high', 'low', 'close']]
            bs.reset_index(level=0, inplace=True) 
            #print(bs)
            bs['date'] = bs['date'].map(mdates.datestr2num)
           # df['Date'] = df['Date'].map(mdates.date2num)
            
            ax = plt.subplot()
            candlestick_ohlc(ax,bs.values, width=5, colorup='g', colordown='r')
            ax.xaxis_date()
            ax.grid(True)
            
            plt.show()
        #plt.subplots(3,3) 
        
        rsi_plot(bs)
        machart('ema', 8, 21, 50, last=120)
        machart('sma',8, 21, 50, last=120)
        macd(bs)
        #aroon(bs)
        bol_band(bs)   
        #stochos(bs)
        #chmf(bs)
       # cdlp(bs)    



if __name__ == "__main__":
    welcome()
    while True:
        query = takeCommand().lower()
        
#        pygame.display.update()
#        button("Speak!",150,450,100,50,green,bright_green,s2t)

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
            
        elif 'whatsapp for me' and 'whatsapp' in query:
            speak("Please wait while the whatsapp application opens")
            whatsapp()
        
        elif 'youtube' in query:
            speak("Please wait while the youtube opens")
            speak("Here you go")
            youtube()
            
        elif 'search for website' in query:
            speak("what should i search?")
            chrome_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe %s' #Add the Location of the chrome browser

            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("Listening...")
                r.pause_threshold =1
                r.adjust_for_ambient_noise(source,duration=0.51)
                audio = r.listen(source,phrase_time_limit=4)

            try:
                text = r.recognize_google(audio)
                print('wendy thinks you said:\n' +text +'.com')
                wb.open(text+'.com')
            except Exception as e:
                print(e)
        
        elif 'how is the weather' and 'weather' in query:
            speak("Hold on, fetching the weather report of Visakhapatnam!");
            url = 'https://api.openweathermap.org/data/2.5/weather?id=1253102&appid=5d6d85d0f2dec9a4fc2149002d4dadc8'#Open api link here

            res = requests.get(url)

            data = res.json()

            weather = data['weather'] [0] ['main'] 
            
            temp = data['main']['temp']
            temp = float(temp)-273.15
            wind_speed = data['wind']['speed']

            latitude = data['coord']['lat']
            longitude = data['coord']['lon']

            description = data['weather'][0]['description']
            speak('Temperature : {} degree celcius'.format(temp))
            print('Wind Speed : {} m/s'.format(wind_speed))
            print('Latitude : {}'.format(latitude))
            print('Longitude : {}'.format(longitude))
            print('Description : {}'.format(description))
            print('weather is: {} '.format(weather))
            speak('weather is: {} '.format(weather))


        elif 'what time is it' and 'the time' in query:
            strTime = datetime.datetime.now().strftime("%I:%M:%S")    
            speak(f"Sam, the time is {strTime}")
        
        elif 'the date' in query:
            year = int(datetime.datetime.now().year)
            month = int(datetime.datetime.now().month)
            date = int(datetime.datetime.now().day)
            speak("the current Date is")
            speak(date)
            speak(month)
            speak(year)


        elif 'email to pinky' and 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = ""    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry my friend . I am not able to send this email")      

        elif 'how are you' in query:
            speak("I am absolutely fantastic!")

            #    codePath = "C:\\Users\\user account\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"#ADD THE PATH OF THE PROGEM HERE
         #   os.startfile(codePath)


        elif 'open' in query:
            os.system('explorer C://{}'.format(query.replace('Open','')))

        
#        elif 'turn on lights' in query:
#            speak("OK, turning on the Lights")
#            lighton()
#            speak("Lights are on")
#        
#        elif 'turn off lights' in query:
#            speak("OK, turning off the Lights")
#            lightoff()
#            speak("Lights are off")
#
        elif 'news headlines' in query:
            speak("OK,sam here are the top headLines from huffingtonpost")
            headlines()
            cmnd=takeCommand().lower()
            if cmnd=="stop":
                break
            speak("That's the update from the top headlines")

        elif 'go offline' in query:
            speak("ok sam! shutting down the system")
            break
            quit()
        
        elif 'change my wallpaper' and 'wallpaper' in query:
            speak("Ok, Changing your wallpaper")
            change_wallpaper()
            speak("Wallpaper changed")
            
        elif 'play music' in query:
            speak("here you go")
            play_music()
            
        
        elif 'stop' in query:
            query = query.replace("stop","")
            print(query)
            pkill(query)
    
        elif 'price of' in query:
             query = query.replace("price of","")
             query=query.title().strip()
             print(query)
             tck=ticker(query)
             print(tck)
             result=real_time_price(tck)
             speak('The current price of '+query+'is{} dollars'.format(result))
#             except TypeError:
#                 speak("i didn't get that")
#                 
        elif "calculate" in query: 
              
            # write your wolframalpha app_id here 
            app_id = "api key here" 
            client = wolframalpha.Client(app_id) 
  
            indx = query.split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text 
            speak("The answer is {}".format(answer))    
        
        elif 'document with subject' in query:
            subject = query.replace("document with subject", "")
            
            speak("what do you want me to save ?")
            
            
            
            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("Listening...")
                r.pause_threshold =1
                r.adjust_for_ambient_noise(source,duration=1)
                audio = r.listen(source,phrase_time_limit=15)

            try:
                text = r.recognize_google(audio)
                print('google think you said:\n' +text)
                if text: 
                    word = win32.Dispatch('Word.Application')
                    doc = word.Documents.Add()
                    word.Visible = True
                    rng = doc.Range(0,0)
                    rng.InsertAfter(subject)
                    r = sr.Recognizer()
                    eng=doc.Range(3,8)
                    rng.InsertAfter(text)
                    sleep(2)
                    doc.SaveAs(subject+'.docx')
                        #doc.Close(False)
                        #word.Quit()
                    speak('document saved')    
            except Exception as e:
                print(e)
        


        elif "hey wendy" in query: 
            speak("Yes sir! How may I help you")
            
        elif "mute" in query:
            speak("muting")
            Sound.mute()
            
        elif "volume up" in query:
            Sound.volume_up()
        
        elif "company analysis" in query:
            query = query.replace("company analysis of","")
            query=query.title().strip()
            print(query)
            tck=ticker(query)
            print(tck)
            company_profile(tck)
            financial_summary(tck)
            historical_price(tck)
            
        elif "phone call" in query:
            speak("dialing in")
            phonecall()
            speak("the call has been dialed")
            
        elif 'shut down' in query:
            speak("ok sam! shutting down the system")
            break
            quit()    
            
        elif 'introduce' in query:
            speak("Hi! I am Wendy! A virtual assistant to help you ease off your tasks through a simple voice command. I can help you with many tasks like new headlines, browse youtube, whatsapp your friend, search for a website, get the price of a stock, make company analysis , play music, do math, report weather and many more, just say a command and I would do it!")
         
        elif 'break' in query:
            pkill(vlc)
            
          
        elif 'start camera' and 'camera' in query:
             speak("opening camera")
             cap=cv2.VideoCapture(0)
             cap.set(cv2.CAP_PROP_FPS, int(300))
             while True:
                _,im=cap.read()
                cv2.imshow('display',im)
                if cv2.waitKey(1) == 27: 
                    break  # esc to quit
             cv2.destroyAllWindows()