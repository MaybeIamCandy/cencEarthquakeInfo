import requests, time, json, ctypes, sys
from win10toast_click import ToastNotifier
import win32com.client

prevData = None
speaker = win32com.client.Dispatch('SAPI.SpVoice')
toaster = ToastNotifier()
speaker.volume = 70 #语音合成音量
firstRun = True
debugMode = True if sys.gettrace() else False
headers = {'User-Agent': 'Mozilla/5.0 (compatible; cencEarthquakeInfoApp/0.1.4)'}
response = requests.Session()
response.trust_env = False

if debugMode:
    print('debug模式运行中')
else:
    print('未处于debug模式')
    console = ctypes.windll.kernel32.GetConsoleWindow()
    if console != 0: #隐藏窗口
        ctypes.windll.user32.ShowWindow(console, 0)
        ctypes.windll.kernel32.CloseHandle(console)

def getData(url):
    try:
        global res
        res = response.get(url, headers=headers)
        if res.status_code == 200:
            return res.json()
        else:
            print('Failed to fetch data from URL:', url)
    except Exception as e:
        print('Error:', e)
    return None

def sendNotification(title, message):
    toaster.show_toast(title=u'{}'.format(title), msg=u'{}'.format(message), icon_path=r'.\\ico\\cenc.ico', duration=None, threaded=True)

def timestampConvert(timestamp):
    global timeConverted
    #毫秒级时间戳转换为中文格式日期，如：2023年9月8日22时04分16秒
    timestamp = float(timestamp/1000)
    timeArray = time.localtime(timestamp)
    timeConverted = time.strftime('%m月%d日%H时%M分%S秒', timeArray)
    print(timeConverted)
    return timeConverted

def getContent():
    text = res.text
    text = json.loads(text)
    global epicenter, epicenterLat, epicenterLon, depth, magnitude, timestamp
    epicenter = text['No0']['epicenter']
    epicenterLat = text['No0']['epicenterLat']
    epicenterLon = text['No0']['epicenterLon']
    depth = text['No0']['depth']
    magnitude = text['No0']['magnitude']
    timestamp = int(text['No0']['time'])
    reportNum = text['No0']['reportNum']

    #报文状态判断，state仅为“自动”或“正式”，stateText为附加文案
    global state, stateText
    if reportNum == '1': #0=自动报，1=正式报
        state = '正式'
        stateText = '。'
        epicenterText = ''
        magAutoText = ''
    else: #自动报，需要附加文案“最终结果以正式速报为准。”
        state = '自动'
        stateText = '，最终结果以正式速报为准。'
        epicenterText = '附近'
        magAutoText = '左右'
    #经纬度判断
    global latText, lonText
    if '-' in epicenterLat:
        latText = '南纬'
        epicenterLat = epicenterLat.lstrip('-')
    else:
        latText = '北纬'

    if '-' in epicenterLon:
        lonText = '西经'
        epicenterLon = epicenterLon.lstrip('-')
    else:
        lonText = '东经'

    timestampConvert(timestamp=timestamp)

    reportStateText = '中国地震台网{0}测定'.format(state)
    descText = '{0}在{1}{9}（{2}{3}度，{4}{5}度）发生{6}级{10}地震，震源深度{7}千米{8}'.format(timeConverted, epicenter, latText, epicenterLat, lonText, epicenterLon, magnitude, depth, stateText, epicenterText, magAutoText)
    sendNotification(title=reportStateText, message=descText)
    finalText = reportStateText+'：'+descText
    speaker.Speak(u'{}'.format(finalText), 1) #语音合成
    print(finalText)
    with open('telop.txt', 'w', encoding='utf-8') as f:
        f.write('')
        f.write(f'{finalText}')

def checkUrl(url):
    global prevData, firstRun
    while True:
        _ = str(int(time.time()))
        currentData = getData(url+'?_='+_)
        
        if currentData is not None:
            if prevData is None:
                prevData = currentData
                print('返回数据：', currentData)
            elif currentData != prevData:
                print('内容已更新')
                getContent()
                prevData = currentData
            elif firstRun:
                firstRun = False
                print('内容相同')
                getContent()
            else:
                print('内容相同')
        else:
            print('请求返回数据为空')
        time.sleep(5)

if __name__ == '__main__':
    url = 'https://api.sweetcandy233.top/cenc/phrasedTelegram'
    checkUrl(url)
