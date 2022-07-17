import datetime
import json
import os
import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox

import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

# Init Dialog
root = tk.Tk()
root.withdraw()

# Setting basePath
base_path = os.getcwd() + '/settings/'

# youtube data api
YOUTUBE_API_SERVICE_NAME = 'youtube'
YOUTUBE_API_VERSION = 'v3'

# youtube analitics api
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
YOUTUBE_ANALYTICS_API_SERVICE_NAME = 'youtubeAnalytics'
YOUTUBE_ANALYTICS_API_VERSION = 'v2'
SCOPES_YOUTUBE = [
    'https://www.googleapis.com/auth/youtube.readonly',
    'https://www.googleapis.com/auth/yt-analytics-monetary.readonly',
    'https://www.googleapis.com/auth/yt-analytics.readonly'
    ]

# spreadsheet
SCOPES_SPREAD = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
    ]

#---------------------------------------
# 設定ファイルの読み込み
#---------------------------------------
class Settings:
    def __init__(self):
        try:
            self.resultflg = True

            # 設定ファイル読み込み
            print("設定ファイル読み込み")
            json_open = open(base_path + 'settings.json', 'r')
            json_load = json.load(json_open)

            # スプレッドシート情報取得
            print("スプレッドシート情報取得")
            self.spread_key = json_load['spreadsheet']['key']
            self.spread_secreatfile = base_path + json_load['spreadsheet']['secret_file']

            # YouTube情報取得
            print("YouTube情報取得")
            self.youtube_key = json_load['youtube']['key']
            self.channels = []
            for channel in json_load['youtube']['channels']:
                if('name' in channel and 'id' in channel and 'secret_file' in channel):
                    self.channels.append(channel)
                else:
                    self.resultflg = False
                    self.resultmsg = '設定ファイルの「youtube_channels」の形式が不正です。'
                    break
        
        except:
            self.resultflg = False
            self.resultmsg = '設定ファイルの読み込みに失敗しました。'
setting_info = Settings()

#---------------------------------------
# YOUTUBE DATA API実行用クラス
#---------------------------------------
class VIDEO_INFO:
    def __init__(self, key, id, secret_file):
        try:
            os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
            self.results = []
            self.id = id
            self.secret_file = base_path + secret_file

            self.youtube_url = 'https://www.youtube.com/watch?v='
            self.youtube = build(
                YOUTUBE_API_SERVICE_NAME,
                YOUTUBE_API_VERSION,
                developerKey=key
            )

            if not secret_file == '':

                flow = InstalledAppFlow.from_client_secrets_file(self.secret_file, SCOPES_YOUTUBE)
                flow.run_local_server(port=8080, prompt="consent")
                credentials = flow.credentials
                self.youtube_analitics = build(
                    YOUTUBE_ANALYTICS_API_SERVICE_NAME,
                    YOUTUBE_ANALYTICS_API_VERSION,
                    credentials = credentials
                )
            self.afterBeforeList = self.GetAfterBefore()

            self.initflg = True
        except Exception as e:
            self.initflg  = False
            self.errormsg = "YouTubeAPIを使用するための認証に失敗しました。"
            print(e)

    # 再生時間を「XX:XX」に変換する
    def durationToSeconds(self, duration):
        """
        duration - ISO 8601 time format
        examples :
            'P1W2DT6H21M32S' - 1 week, 2 days, 6 hours, 21 mins, 32 secs,
            'PT7M15S' - 7 mins, 15 secs
        """
        try:
            split   = duration.split('T')
            period  = split[0]
            time    = split[1]
            timeD   = {}

            # days & weeks
            if len(period) > 1:
                timeD['days']  = int(period[-2:-1])
            if len(period) > 3:
                timeD['weeks'] = int(period[:-3].replace('P', ''))

            # hours, minutes & seconds
            if len(time.split('H')) > 1:
                timeD['hours'] = time.split('H')[0].zfill(2)
                time = time.split('H')[1]
            if len(time.split('M')) > 1:
                timeD['minutes'] = time.split('M')[0].zfill(2)
                time = time.split('M')[1]    
            if len(time.split('S')) > 1:
                timeD['seconds'] = time.split('S')[0].zfill(2)

            if 'hours' in timeD:
                if not 'minutes' in timeD:
                    timeD['minutes'] = '00'
                if not 'seconds' in timeD:
                    timeD['seconds'] = '00'
                strtime = timeD['hours'] + ':' + timeD['minutes'] + ':' + timeD['seconds']
            elif 'minutes' in timeD:
                if not 'seconds' in timeD:
                    timeD['seconds'] = '00'
                strtime = timeD['minutes'] + ':' + timeD['seconds']
            elif 'seconds' in timeD:
                strtime = '00:' + timeD['seconds']
        except:
            strtime = '00:00'

        return strtime        

    # 過去2年の日付配列を生成
    def GetAfterBefore(self):
        now = datetime.datetime.now()
        year = now.year

        periodList = []
        periodList.append(year)

        afterBeforeList = []
        for counter in range(year):
            year = year -1
            periodList.append(year)

        count = 0
        for period in periodList:
            afterBeforeList.append(['{}-01-01T00:00:00Z'.format(period), '{}-12-31T00:00:00Z'.format(period)])
            if count == 1:
                break
            else:
                count += 1

        return afterBeforeList

    # 動画情報を取得
    def Get_videos_Info(self):
        
        for afterBefore in self.afterBeforeList:

            nextPageToken = ''
            nextflg = True
            while nextflg:

                # チャンネル動画の取得
                response = self.youtube.search().list(
                    part = "snippet",
                    channelId = self.id,
                    maxResults = 50,
                    order = "date",
                    publishedAfter = afterBefore[0],
                    publishedBefore = afterBefore[1],
                    pageToken=nextPageToken
                ).execute()

                if 'nextPageToken' in response:
                    nextPageToken = response['nextPageToken']
                    nextflg = True
                else:
                    nextPageToken = ''
                    nextflg = False

                # チャンネルの動画情報を1件ずつ取得
                for item in response['items']:
                    # 取得したjsonから動画情報だけを取得
                    if item['id']['kind'] == 'youtube#video':
                        d = datetime.datetime.strptime(item['snippet']['publishedAt'], '%Y-%m-%dT%H:%M:%SZ')
                        # 情報出力
                        result = {
                            "id"          : item['id']['videoId'],
                            "title"       : item['snippet']['title'],
                            "url"         : self.youtube_url + item['id']['videoId'],
                            "publishedAt" : d.strftime("%Y/%m/%d %H:%M:%S"),
                            "thumbnail"   : item['snippet']['thumbnails']['default']['url'],
                            "width"       : item['snippet']['thumbnails']['default']['width'],
                            "height"      : item['snippet']['thumbnails']['default']['height'],
                            "duration"    : ""
                        }

                        # 結果配列に格納
                        self.results.append(result)
            
    # 動画情報（詳細情報）を取得　※動画長さなど
    def Get_video_Detail(self):
        for result in self.results:
            video_responce = self.youtube.videos().list(
                part = 'contentDetails',
                id   = result["id"]
            ).execute()
          
            dur = self.durationToSeconds(video_responce["items"][0]["contentDetails"]["duration"])
            result["duration"] = dur

#---------------------------------------
# スプレッドシートへの出力
#---------------------------------------
# 出力用のリストを作成
def Create_celllist(ws, data, start_cell, end_cell, col_num):
    cell_list = ws.range(start_cell[0], start_cell[1], end_cell[0], end_cell[1])
    data_num = len(data)
    all_num = col_num * len(data)

    # データの格納
    for i, item in enumerate(data):
        start_num = i * col_num
        cell_list[start_num].value   = item['id']
        cell_list[start_num+1].value = item['title']
        cell_list[start_num+2].value = item['url']
        cell_list[start_num+3].value = item['publishedAt']
        cell_list[start_num+4].value = '=IMAGE("' + item['thumbnail'] + '",4,' + str(item['height']) + ',' + str(item['width']) + ')'


    return cell_list

# スプレッドシートへの出力
def spreadsheet_export(results):
    #認証情報設定
    file_path = base_path + setting_info.spread_secreatfile
    credentials = ServiceAccountCredentials.from_json_keyfile_name(setting_info.spread_secreatfile, SCOPES_SPREAD)
    gc = gspread.authorize(credentials)
    wb = gc.open_by_key(setting_info.spread_key)

    # ワークシートの取得
    worksheets = wb.worksheets()
    bsheet = False
    for sheet in worksheets:
        if 'チャンネル動画データ' == sheet.title:
            bsheet = True
            break
    
    # 出力用のワークシートがなければ作成する
    if bsheet == False:
        ws = wb.add_worksheet(title = 'チャンネル動画データ', rows=1000, cols=50)
    else:
        ws = wb.worksheet('チャンネル動画データ')

    # 取得結果情報を出力する
    col_num = 5
    start_cell = [3, 1]
    end_cell   = [len(results) + 2, col_num]
    cell_list = Create_celllist(ws, results, start_cell, end_cell, col_num)

    ws.update_cells(cell_list,value_input_option='USER_ENTERED')

#---------------------------------------
# メイン関数
#---------------------------------------
def run_button(select):

    # 選択されたチャンネル情報を取得
    print ("設定ファイル情報を取得中")
    for channel in setting_info.channels:
        if channel['name'] == select:
            id          = channel['id']
            secret_file = channel['secret_file']
            break

    # YOUTUBE情報を取得
    print ("YouTube情報を取得中")
    videoInfo = VIDEO_INFO(setting_info.youtube_key, id, secret_file)
    if videoInfo.initflg == True:
        videoInfo.Get_videos_Info()
        #videoInfo.Get_video_Detail()

        # Analytics情報を取得
        #videoInfo.Get_Analytics()

        # スプレッドシートへの結果出力
        print ("スプレッドシートへ出力中")
        spreadsheet_export(videoInfo.results)

        # 処理終了
        print ("処理終了")
        messagebox.showinfo(title='処理完了', message="処理が正常に終了しました。")
    else:
        messagebox.showerror(title='エラー', message=videoInfo.errormsg)

def close():
    sys.exit()

#---------------------------------------
# 初期実行
#---------------------------------------
if __name__ == '__main__':
    # 設定ファイル読み込み
    if setting_info.resultflg == False:
        res = messagebox.showerror(title='エラー', message=setting_info.resultmsg)
    else:
        # メインウィンドウ
        root = tk.Tk()
        root.title(u"YouTube Analytics")
        root.geometry("440x120")
        root.resizable(width=False, height=False)

        # メインフレーム
        frame = tk.Frame(root)
        frame.grid(column=0, row=0, padx=5, pady=10)

        # ラベル
        Label1 = tk.Label(frame, text=u'Select Youtube Channel')

        # コンボボックス
        list_values = []
        for channel in setting_info.channels:
            list_values.append(channel['name'])
        Combobox = ttk.Combobox(frame, values=list_values, state='readonly', width=50)
        Combobox.set(list_values[0])

        # 実行ボタン
        Button1 = ttk.Button(frame, text="実行",command=lambda:run_button(Combobox.get()))

        # 閉じるボタン
        Button2 = ttk.Button(frame, text="閉じる",command=lambda:close())

        # 配置
        Label1 .grid(columnspan=4, column=0, row=0, sticky=tk.W)
        Combobox.grid(columnspan=4, column=0, row=1, pady=10)
        Button1.grid(column=2, row=2, sticky=tk.W + tk.E)
        Button2.grid(column=1, row=2, sticky=tk.W + tk.E)

        # 画面表示
        root.mainloop()
