import pkg_resources
import serial
import serial.tools.list_ports
import time
import win32gui
import win32process
import atexit
import threading
from infi.systray import SysTrayIcon
import GUI_manager
import data
import psutil


#----- Variables -----#
connected = 0
old_song = ""
old_time = ""
first_data_transfer = 1
spotify_open = 0
time_prefix = "sgxdfgchjkl"
exit_by_user = 0
spotify_closed_key = "kyawesgtlu"
closed_by_setting = 0
paused_by_setting = 0
paused = "Paused"
show_song = 1
char = "-"
result = ""
spotify_pid = None
arduino = None

spotify_check_interval = 3  # Check every 5 seconds
last_spotify_check = 0
cached_song = spotify_closed_key
cached_artist = ""

# Exit Handler
#-----------------------------------
def exit_handler():
    global arduino
    try:
        arduino.write("ServerExitedByUser".encode())
    except:
        pass

#-----------------------------------
# SysTray
#-----------------------------------
def on_quit_callback(systray):
    global exit_by_user
    exit_by_user = 1

def settings_callout(systray):
    if data.gui_counter != 0:
        data.gui_counter = 0
        GUI_manager.initiate_GUI()

menu_options = (("Settings", None, settings_callout),)
systray = SysTrayIcon("icon.ico", "Arduino Spotify Checker", menu_options, on_quit=on_quit_callback)
systray.start()

# Get Spotify Window Info
#-----------------------------------
def get_info_windows():
    global result, spotify_pid
    result = ""
    pids = []
    try:
        for proc in psutil.process_iter(['name', 'pid']):
            if proc.name().lower() == 'spotify.exe':
                pids.append(proc.info["pid"])

        def callback(hwnd, pid):
            global result, spotify_pid
            pid_list = win32process.GetWindowThreadProcessId(hwnd)
            if pid == pid_list[1]:
                window_title = win32gui.GetWindowText(hwnd)
                if char in window_title:
                    result = window_title
                    spotify_pid = pid
                elif window_title == "Spotify Free" or window_title == "Spotify Premium":
                    result = "Paused - Paused"

        for pid in pids:
            win32gui.EnumWindows(callback, pid)

    except:
        print("Error while getting song info")

    parts = result.split(char)
    if len(parts) == 2:
        return parts[1].strip(), parts[0].strip()  # song, artist
    return spotify_closed_key, ""

#-----------------------------------
def connect():
    global arduino, connected
    ports = list(serial.tools.list_ports.comports())
    for p in ports:
        if "USB-SERIAL CH340" in p.description:
            port = p.name
            try:
                arduino = serial.Serial(port, 9600)
                print("Connected")
                connected = 1
                return
            except:
                connected = 0
    connected = 0

#-----------------------------------
def TimeNDate():
    global time_prefix, arduino, old_time
    current_time = time.strftime("%H:%M", time.localtime())
    if current_time != old_time:
        try:
            old_time = current_time
            arduino.write((time_prefix + old_time).encode())
            print("Time sent:", time_prefix + old_time)
        except:
            print("Board Disconnected")
            connect()

#-----------------------------------
def send_song_to_arduino(song):
    global arduino, old_song
    if song != old_song:
        try:
            arduino.write(song.encode())
            print("Song sent:", song)
            old_song = song
        except:
            print("Board Disconnected")
            connect()

#-----------------------------------
def update_spotify_cache():
    global cached_song, cached_artist, last_spotify_check
    if time.time() - last_spotify_check >= spotify_check_interval:
        cached_song, cached_artist = get_info_windows()
        last_spotify_check = time.time()

#-----------------------------------
def main_loop():
    global exit_by_user, connected, first_data_transfer
    global closed_by_setting, paused_by_setting, paused, old_time
    global show_song, cached_song

    delay = 1
    time_single_signal = 0

    while not exit_by_user:
        time.sleep(delay)

        if data.on == 0:
            try:
                arduino.write("ServerExitedByUser".encode())
            except:
                pass
            closed_by_setting = 1
            continue

        if connected == 0:
            connect()
            if connected:
                first_data_transfer = 1
                time.sleep(4)

        if closed_by_setting == 1:
            TimeNDate()
            update_spotify_cache()
            send_song_to_arduino(cached_song)
            closed_by_setting = 0
            continue

        if data.paused == 0 and not paused_by_setting:
            paused = spotify_closed_key
            paused_by_setting = 1
            old_time = "invalid"
        elif data.paused == 1 and paused_by_setting:
            paused = "Paused"
            TimeNDate()
            update_spotify_cache()
            send_song_to_arduino(cached_song)
            paused_by_setting = 0
            closed_by_setting = 1
            continue

        if data.show_song == 0:
            show_song = 0
        elif data.show_song == 1:
            show_song = 1

        update_spotify_cache()

        if show_song:
            TimeNDate()
            send_song_to_arduino(cached_song)
        else:
            send_song_to_arduino(spotify_closed_key)

#----------------------------------------------------------------------
atexit.register(exit_handler)
main_thread = threading.Thread(target=main_loop)
main_thread.start()
