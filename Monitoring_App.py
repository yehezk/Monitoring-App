from pynput.keyboard import Listener
import json
import os
import datetime
import threading
import win32gui
import win32process
import psutil
import pytz
import time

# Mendapatkan path ke direktori %localappdata%
local_app_data = os.getenv('LOCALAPPDATA')

# Membuat direktori MonitoringApp di %localappdata%
directory_path_localappdata = os.path.join(local_app_data, 'MonitoringApp')

log_file_apps = "log_apps.json"  # Nama file log aplikasi
# Menggabungkan path file log aplikasi
log_file_path_apps = os.path.join(directory_path_localappdata, log_file_apps)

stop_event = threading.Event()  # Membuat objek threading.Event()
stop_event.initial_processes = []  # Menyimpan proses awal
initial_window = None
click_count = 0
program_start_time = time.time()  # Waktu mulai program
is_typing = False
active_window_start_time = None
last_activity_time = time.time()
last_window = None
last_window_duration = 0

# structure json
log_structure = {
    "buckets": {
        "aw-watcher-window": {
            "events": []
        },
        "aw-watcher-input": {
            "events": []
        },
        "aw-watcher-afk": {
            "events": []
        }
    }
}


def main():
    """
    Fungsi utama yang menjalankan program.
    """
    create_log_files()  # Membuat file log jika tidak ada
    # Membuat thread untuk mendengarkan keyboard
    listener_thread_key = threading.Thread(target=listen_keyboard)
    listener_thread_key.start()  # Memulai thread mendengarkan keyboard
    # Membuat thread untuk mendengarkan aplikasi
    listener_thread_apps = threading.Thread(target=listen_apps)
    listener_thread_apps.start()  # Memulai thread mendengarkan aplikasi
    # Thread untuk memantau aktivitas keyboard
    activity_thread = threading.Thread(target=check_and_log_activity)
    activity_thread.start()


def create_log_files():
    """
    Membuat file log keyboard dan log aplikasi jika tidak ada.
    """
    # Mengecek apakah direktori tidak ada
    if not os.path.exists(directory_path_localappdata):
        # Membuat direktori jika tidak ada
        os.makedirs(directory_path_localappdata)
        # print("directory belum ada")
    if os.path.exists(log_file_path_apps):
        # File sudah ada, lakukan penanganan yang sesuai
        # Misalnya, tampilkan pesan error atau keluar dari program.
        pass
        # print("File log sudah ada.")
        # Cek keberadaan file log aplikasi
    if not os.path.exists(log_file_path_apps):
        # Jika file log aplikasi tidak ada, buat file baru
        # Menulis struktur awal ke file log
        # print("File log tidak ada.")
        with open(log_file_path_apps, "w", encoding="utf-8") as file:
            json.dump(log_structure, file, indent=4)


def on_press(key):
    """
    Fungsi yang dipanggil saat tombol ditekan.
    Mencatat waktu saat mulai mengetik dan memulai timer untuk berhenti mengetik.
    """
    global is_typing, click_count, last_activity_time
    last_activity_time = time.time()  # Memperbarui waktu aktivitas terakhir
    if not is_typing:  # Jika tidak sedang mengetik
        is_typing = True  # Set status sedang mengetik menjadi True
        click_count = 0  # Reset jumlah klik keyboard
    click_count += 1  # Menambah jumlah klik keyboard


def listen_keyboard():
    """
    Fungsi untuk mendengarkan keyboard menggunakan pynput.
    """
    global is_typing, click_count,  last_activity_time

    with Listener(on_press=on_press):

        while True:
            stop_event.wait(1)  # Menunggu interval 1 detik
            if is_typing:  # Jika sedang mengetik
                # print("keyboard")
                timestamp = datetime.datetime.now(pytz.utc).isoformat()
                try:
                    with open(log_file_path_apps, "r", encoding="utf-8") as file:
                        log_data = json.load(file)
                except (FileNotFoundError, json.decoder.JSONDecodeError):
                    log_data = {}

                log_entry = {
                    "timestamp": timestamp,
                    "duration": 1,
                    "data": {
                        "presses": click_count,
                    }
                }
                log_data["buckets"]["aw-watcher-input"]["events"].append(
                    log_entry)

                with open(log_file_path_apps, "w", encoding="utf-8") as file:
                    json.dump(log_data, file, indent=4)

            click_count = 0  # Reset jumlah klik keyboard
            is_typing = False  # Set status sedang mengetik menjadi False


def get_active_window():
    """
    Mendapatkan judul jendela aktif.
    """
    active_window = None
    try:
        # Mendapatkan handle jendela aktif
        active_window_handle = win32gui.GetForegroundWindow()
        # Mendapatkan ID proses jendela aktif
        active_window_pid = win32process.GetWindowThreadProcessId(
            active_window_handle)
        if active_window_pid[-1] <= 0:  # Memeriksa ID proses yang valid
            return active_window

        # Mendapatkan objek proses menggunakan ID proses
        active_window_process = psutil.Process(active_window_pid[-1])
        # Mendapatkan nama eksekusi proses
        active_window_exe = active_window_process.exe()
        # Mendapatkan judul jendela menggunakan handle jendela
        active_window_title = win32gui.GetWindowText(active_window_handle)

        active_window = {
            "exe": active_window_exe,
            "title": active_window_title
        }
    except (psutil.AccessDenied, psutil.NoSuchProcess, psutil.ZombieProcess):
        pass

    return active_window


def listen_apps():
    """
    Mencatat perubahan dalam daftar jendela aplikasi yang sedang berjalan.
    """
    global initial_window, last_activity_time  # variabel global

    initial_window = get_active_window()

    active_window_start_time = None

    while not stop_event.is_set():
        current_window = get_active_window()
        current_timestamp = datetime.datetime.now()
        if current_window and current_window != initial_window and current_window.get("exe"):
            last_activity_time = time.time()
            if not current_window == 'C:\\Windows\\explorer.exe' and initial_window.get("title") != '':
                if not current_window == 'C:\\Windows\\explorer.exe' and initial_window.get("title") != 'Task Switching':
                    if active_window_start_time is not None:
                        duration = (current_timestamp -
                                    active_window_start_time).total_seconds()
                        log_application("active_window", duration,
                                        initial_window, active_window_start_time)
            initial_window = current_window
            active_window_start_time = current_timestamp


def log_application(event, duration, window, active_window_start_time):
    """
    Mencatat peristiwa pembukaan, penutupan aplikasi, atau perubahan jendela aktif ke file log apps.
    """

    active_window_start_time
    if event == "active_window":  # Jika event adalah perubahan jendela aktif
        # print("window")
        timestamp = datetime.datetime.now(pytz.utc).isoformat()
        log_entry = {
            "timestamp": timestamp,
            "duration": duration,  # Durasi jendela sebelumnya
            "data": {
                "app": os.path.basename(window["exe"]),
                "title": window["title"]
            }
        }
        try:
            with open(log_file_path_apps, "r", encoding="utf-8") as file:
                log_data = json.load(file)
        except (FileNotFoundError, json.decoder.JSONDecodeError):
            log_data = {}

        log_data["buckets"]["aw-watcher-window"]["events"].append(log_entry)

        with open(log_file_path_apps, "w", encoding="utf-8") as file:
            json.dump(log_data, file, indent=4)


def log_afk(duration):

    timestamp = datetime.datetime.now(pytz.utc).isoformat()

    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        log_data = {}

    afk_entry = {
        "timestamp": timestamp,
        "duration": duration,
        "data": {
            "status": "afk"
        }
    }

    log_data["buckets"]["aw-watcher-afk"]["events"].append(afk_entry)

    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        json.dump(log_data, file, indent=4)


def log_not_afk(duration):

    timestamp = datetime.datetime.now(pytz.utc).isoformat()

    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        log_data = {}

    not_afk_entry = {
        "timestamp": timestamp,
        "duration": duration,
        "data": {
            "status": "not-afk"
        }
    }

    log_data["buckets"]["aw-watcher-afk"]["events"].append(not_afk_entry)

    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        json.dump(log_data, file, indent=4)


def check_and_log_activity():
    global last_activity_time, program_start_time

    afk_timeout = 900  # Timeout AFK 15 menit
    afk_start_time = None
    statusnow = "not-afk"  # Status awal "not-afk"

    while True:
        current_time = time.time()
        if statusnow == "not-afk":
            if current_time - last_activity_time >= afk_timeout:
                afk_start_time = time.time()
                duration = afk_start_time - program_start_time
                # Durasi awal sebelum status diubah menjadi "afk"
                statusnow = "afk"
                log_not_afk(duration)
                # print("status afk, durasi dari not afk sampai afk:",
                #       duration, "Detik")
        elif statusnow == "afk":
            if current_time - last_activity_time < afk_timeout:
                duration = current_time - afk_start_time
                statusnow = "not-afk"
                log_afk(duration)
                # print("Status NOT AFK, durasi dari afk sampai not afk:",
                #       duration, "Detik")
        time.sleep(1)


if __name__ == '__main__':
    main()
