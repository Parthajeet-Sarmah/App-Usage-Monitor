#import pywinctl as pwc
import win32gui, win32process
import random
from datetime import datetime
import psutil
import time
import math
import plotly.graph_objects as go
import ctypes

#problems - program stops on 4 tracks due to the TRACK_LIMIT but i also need the functionality of the TRACK_LIMIT

TRACK_THRESHOLD = 0.01
SPI_SETDESKWALLPAPER = 20
OTHER_TRACK_COUNT_THRESHOLD = 4

firstRun = False
takeProcessNames = False
filtered = ["Visual Studio Code", "Google Chrome", "Brave", "VALORANT"]
ignoredProcesses = ["explorer.exe"]

colors = ['#003f5c',
'#2f4b7c',
'#665191',
'#a05195',
'#d45087',
'#f95d6a',
'#ff7c43',
'#ffa600']

random.shuffle(colors)

def filter_names(title):
    for i in filtered:
        try:
            if(title.find(i) != -1):
                return i
        except AttributeError:
            continue
    
    return title
    
def get_active_window_pid():
    processId = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[1]
    return processId

def get_active_window_process_name():
    try:
        current_process = psutil.Process(get_active_window_pid())
        return current_process.name()
    except ValueError:
        pass
    except psutil.NoSuchProcess:
        pass

tracks = {}
other_tracks = {}

def calculate_total_time():
    t=0
    for i in tracks.values():
        t += i
    return t

other_track_count = 0
other_track_sum = 0

def transfer_tracks():

    #need to find a limiting condition for accumulation to other tracks
        #count >= 5
        #threshold = 1%
        #Other value < 5%

    #calculate threshold
    threshold = int(TRACK_THRESHOLD * calculate_total_time())

    #calculate track count less than threshold in tracks{}
    global other_track_count
    global other_track_sum

    other_track_count = 0
    other_track_sum=0

    for i in tracks:
        if tracks[i] < threshold and i != "Other":
            other_track_count += 1
            other_track_sum += tracks[i]

    print("count ", other_track_count)
    print("threshold ", threshold)
    #send all processes less than threshold to other tracks
    for i in list(tracks.keys()):
        if other_track_count < OTHER_TRACK_COUNT_THRESHOLD or other_track_sum >= threshold * 5:
            break
        if tracks[i] < threshold and i != "Other":
            #send to other tracks or add to it if already exists
            if i in other_tracks:
                other_tracks[i] += tracks[i]
            else:
                other_tracks[i] = tracks[i]
                del tracks[i]
    #send all other processes greater than threshold to tracks
    for i in list(other_tracks.keys()):
        if other_tracks[i] >= threshold:
            #send to other tracks or add to it if already exists
            if i in tracks:
                tracks[i] += other_tracks[i]
            else:
                tracks[i] = other_tracks[i]
                del other_tracks[i]
    totalOtherTime=0

    for i in other_tracks.values():
        totalOtherTime += i
    print("total other time: ", totalOtherTime)
    tracks["Other"] = totalOtherTime

def send_to_tracks_or_others(process, time):
    global other_track_count
    threshold = TRACK_THRESHOLD * calculate_total_time()
    if time < threshold and other_track_count >= 5:
        other_tracks[process] = int(time)
    else:
        tracks[process] = int(time)

def get_previous_track_status():
    currentDateAndTime = datetime.now()

    print(currentDateAndTime)

def track_active_window_time():
    prev_window = None
    start_time = time.time()

    while True:
        #current_window = get_active_window_process_name() if takeProcessNames else filter_names(pwc.getActiveWindowTitle())
        current_window = filter_names(get_active_window_process_name())
        if current_window != prev_window:
            if prev_window and prev_window not in ignoredProcesses:
                end_time = time.time()
                active_time = end_time - start_time
                if math.floor(active_time) <= 0:
                    continue
                elif prev_window in tracks:
                    tracks[prev_window] += math.floor(active_time)
                elif prev_window in other_tracks:
                    other_tracks[prev_window] += math.floor(active_time)
                else:
                    send_to_tracks_or_others(prev_window, active_time)
            start_time = time.time()
            prev_window = current_window
        transfer_tracks()
        update_wallpaper(tracks)

def generate_pie_chart(data):
    labels = list(data.keys())
    values = list(data.values())

    imgpath = "C:\\wallpaper\\app_usage_pie.png"
    fig = go.Figure(data=go.Pie(labels=labels, values=values, hole=0.3, texttemplate="%{value}s"))
    fig.update_layout(template="plotly_dark", width=1920, height=1080, legend_font_size=20)
    fig.update_traces(textfont_color='white', textfont_size=20, marker=dict(colors=colors, line=dict(color='#ffffff', width=2)))
    fig.write_image(imgpath)

    return imgpath

def update_wallpaper(data):
    imgpath = generate_pie_chart(data)
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, imgpath, 3)


update_wallpaper(tracks)
track_active_window_time()

        
    

