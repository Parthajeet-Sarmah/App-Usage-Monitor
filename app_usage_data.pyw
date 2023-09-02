#import pywinctl as pwc
import win32gui, win32process
import random
import psutil
import atexit
import time
import datetime
import math
import plotly.graph_objects as go
import ctypes
import os

TRACK_THRESHOLD = 0.01
SPI_SETDESKWALLPAPER = 20
OTHER_TRACK_COUNT_THRESHOLD = 4

class AppUsageData():
    def __init__(self):    
        self.filter_search_terms = ["VALORANT"] #process names with these terms will be filtered...
        self.filter_change_terms = ["VALORANT.exe"] #....to these terms
        self.ignoredProcesses = ["explorer.exe", "ApplicationFrameHost.exe", "LockApp.exe", "ApplicationFrameHost.exe", "SearchApp.exe", "ShellExperienceHost.exe"] #processes that will be ignored and not counted for the app usage
        self.pie_chart_colors = ['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087', '#f95d6a', '#ff7c43', '#ffa600']
        
        #self.tracks - our main dictionary where the information of the processes with their time taken
        if not os.path.exists("C:\\wallpaper\\save_data.txt"): 
            self.tracks = {}
        else:
            with open("C:\\wallpaper\\save_data.txt", 'r') as file:
                self.tracks = file.read() 
        
        self.other_tracks = {} #an assimilation of very small tracks in the Other section, are kept here
        self.other_track_count = 0
        self.other_track_sum = 0
        self.other_burst = False #if any track has present in the Other track, other_burst set to True, else False
        self.hms_values = [] #the HH:MM:SS values for the run times of the tracks

        random.shuffle(self.pie_chart_colors) #shuffle all the pie chart colors so that we get different order of colours

    def filter_names(self, title):
        for i in range(len(self.filter_search_terms)):
            try:
                if(title.find(self.filter_search_terms[i]) != -1):
                    return self.filter_change_terms[i]
            except AttributeError:
                continue
        
        return title
    
    def get_active_window_pid(self):
        processId = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[1]
        return processId

    def get_active_window_process_name(self):
        try:
            current_process = psutil.Process(self.get_active_window_pid())
            return current_process.name()
        except ValueError:
            pass
        except psutil.NoSuchProcess:
            pass

    def calculate_total_time(self):
        t=0
        for i in self.tracks.values():
            t += i
        return t

    def transfer_tracks(self):
        #calculate threshold that will be boundary to what's a main track and a other track
        threshold = int(TRACK_THRESHOLD * self.calculate_total_time())

        for i in self.tracks:
            if self.tracks[i] < threshold and i != "Other":
                self.other_track_count += 1
                self.other_track_sum += self.tracks[i]

        #send all processes less than threshold to other_tracks
        for i in list(self.tracks.keys()):
            if self.other_burst==False and (self.other_track_count < OTHER_TRACK_COUNT_THRESHOLD or self.other_track_sum >= threshold * 5):
                break
            if self.tracks[i] < threshold and i != "Other":
                self.other_burst = True
                #send to other_tracks or add to it if already exists
                if i in self.other_tracks:
                    self.other_tracks[i] += self.tracks[i]
                else:
                    self.other_tracks[i] = self.tracks[i]
                    del self.tracks[i]
        #send all other processes greater than threshold to self.tracks
        for i in list(self.other_tracks.keys()):
            if self.other_tracks[i] >= threshold:
                #send to tracks or add to it if already exists
                if i in self.tracks:
                    self.tracks[i] += self.other_tracks[i]
                else:
                    self.tracks[i] = self.other_tracks[i]
                    del self.other_tracks[i]
                    if len(self.other_tracks) <= 0:
                        self.other_burst = False
        
        totalOtherTime=0

        for i in self.other_tracks.values():
            totalOtherTime += i
        self.tracks["Other"] = totalOtherTime

    def send_to_tracks_or_others(self, process, time):
        
        threshold = TRACK_THRESHOLD * self.calculate_total_time()
        if time < threshold and self.other_track_count >= 5:
            self.other_tracks[process] = int(time)
        else:
            self.tracks[process] = int(time)

    def track_active_window_time(self):
        prev_window = None
        start_time = time.time()

        while True:
            current_window = self.filter_names(self.get_active_window_process_name())
            if current_window != prev_window:
                if prev_window and prev_window not in self.ignoredProcesses:
                    end_time = time.time()
                    active_time = end_time - start_time
                    if math.floor(active_time) <= 0:
                        pass
                    elif prev_window in self.tracks:
                        self.tracks[prev_window] += math.floor(active_time)
                    elif prev_window in self.other_tracks:
                        self.other_tracks[prev_window] += math.floor(active_time)
                    else:
                        self.send_to_tracks_or_others(prev_window, active_time)
                start_time = time.time()
                prev_window = current_window
            self.transfer_tracks()
            self.update_wallpaper(self.tracks)

    def convert_to_hms_value(self, value):
        mintues, seconds = divmod(value, 60)
        hours, mintues = divmod(mintues, 60)

        h = "" if hours == 0 else str(hours) + "h"
        m = "" if mintues == 0 else str(mintues) + "m"
        s = "" if seconds == 0 or hours != 0 else str(seconds) + "s"

        return h+m+s

    def generate_pie_chart(self, data):
        labels = list(data.keys())
        values = list(data.values())

        for i in range(len(values)):
            t = self.convert_to_hms_value(values[i])
            if i < len(self.hms_values):
                self.hms_values[i] = t
            else:
                self.hms_values.append(t)    

        tt = "Total Time: \n" + self.convert_to_hms_value(self.calculate_total_time()) if len(self.tracks) > 1 else ""
        imgpath = "C:\\wallpaper\\app_usage_pie.png"
        fig = go.Figure(data=go.Pie(labels=labels, values=values, hole=0.5, text=self.hms_values, textinfo="label+text"))
        fig.update_layout(template="plotly_dark", width=1920, height=1080, showlegend=False,
                        annotations=[dict(text=tt, x=0.5, y=0.5, font_size=30, showarrow=False)])
        fig.update_traces(textfont_color='white', textfont_size=20, marker=dict(colors=self.pie_chart_colors, line=dict(color='#ffffff', width=2)))
        fig.write_image(imgpath)

        return imgpath
    
    def update_wallpaper(self, data):
        imgpath = self.generate_pie_chart(data)
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, imgpath, 3)

    def shutdown_save(self):
        file = open("C:\\wallpaper\\save_data.txt", 'r')
        file.close()

try:
    app = AppUsageData()
    app.update_wallpaper(app.tracks)
    app.track_active_window_time()
except:
    atexit.register(app.shutdown_save)