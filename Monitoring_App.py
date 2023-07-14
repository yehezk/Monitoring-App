import keyboard
import json
import os
import datetime
import threading
import win32gui
import win32process
import psutil
import pytz
import time
import traceback
import socket
import paramiko
import schedule
from scp import SCPClient

# Mendapatkan path ke direktori %localappdata%
local_app_data = os.getenv('LOCALAPPDATA')

# get device name
hostname = socket.gethostname()

# Membuat direktori MonitoringApp di %localappdata%
directory_path_localappdata = os.path.join(local_app_data, 'MonitoringApp')

log_file_apps = "log_apps_" + hostname.lower() + ".json"  # Nama file log aplikasi
# Menggabungkan path file log aplikasi
log_file_path_apps = os.path.join(directory_path_localappdata, log_file_apps)

# variabel stop event
stop_event = threading.Event()  # Membuat objek threading.Event()

# variabel global untuk fungsi
initial_window = None
click_count = 0  # Menghitung jumlah klik
program_start_time = time.time()  # Waktu mulai program
is_typing = False  # Menyimpan status False sebagai inisiasi awal
last_activity_time = time.time()  # aktifitas time terakhir

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
    listen_thread_afk = threading.Thread(target=listen_afk)
    listen_thread_apps = threading.Thread(target=listen_apps)
    listen_thread_key = threading.Thread(target=listen_keyboard)
    listen_thread_apps.start()
    listen_thread_key.start()
    keyboard.on_press(on_press)
    listen_thread_afk.start()

    schedule.every(2).minutes.do(run_upload)  # upload setiap 1 menit

    # Looping untuk menjalankan jadwal tugas
    while True:
        schedule.run_pending()


def run_upload():
    """
    Fungsi yang menjalankan upload file.
    """
    # Pengaturan koneksi SCP
    hostname = "185.207.9.74"
    port = 22
    username = "root"
    password = "43RGtfYy2k"
    # Path file lokal
    local_file_path = log_file_path_apps
    # Path file remote
    remote_file_path = "/var/www/html/data/" + log_file_apps

    send_file_scp(hostname, port, username, password,
                  local_file_path, remote_file_path)


def send_file_scp(hostname, port, username, password, local_file_path, remote_file_path):
    try:
        # Membuat objek SSHClient
        client = paramiko.SSHClient()
        client.load_system_host_keys()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        # Menghubungkan ke server SSH menggunakan informasi yang diberikan
        client.connect(hostname, port=port,
                       username=username, password=password)

        # Membuat objek SCPClient dengan menggunakan koneksi SSH
        scp_client = SCPClient(client.get_transport())
        # Menyalin file lokal ke server menggunakan SCP
        scp_client.put(local_file_path, remote_file_path)
        scp_client.close()  # Menutup koneksi SCP

        client.close()  # Menutup koneksi SSH
        print("File transferred successfully!")
    except paramiko.AuthenticationException:
        traceback.print_exc()
    except paramiko.SSHException:
        traceback.print_exc()
    except Exception as e:
        traceback.print_exc()


def create_log_files():
    """
    Membuat file log keyboard dan log aplikasi jika tidak ada.
    """
    # Mengecek apakah direktori tidak ada
    if not os.path.exists(directory_path_localappdata):
        # Membuat direktori jika tidak ada
        os.makedirs(directory_path_localappdata)
    if os.path.exists(log_file_path_apps):
        pass
        # Cek keberadaan file log aplikasi
    if not os.path.exists(log_file_path_apps):
        # Jika file log aplikasi tidak ada, buat file baru
        with open(log_file_path_apps, "w", encoding="utf-8") as file:
            # Menulis struktur awal ke file log
            json.dump(log_structure, file, indent=4)


def on_press(event):
    """
    Fungsi yang dipanggil saat tombol ditekan.
    Mencatat waktu saat mulai mengetik dan memulai timer untuk berhenti mengetik.
    """
    global is_typing, click_count, last_activity_time
    last_activity_time = time.time()  # Memperbarui waktu aktivitas terakhir
    if not is_typing:  # Jika tidak sedang mengetik
        is_typing = True  # Set status sedang mengetik menjadi True
    click_count += 1  # click count +1


def listen_keyboard():
    """
    Fungsi untuk mendengarkan keyboard.
    """
    global is_typing, click_count

    while True:
        stop_event.wait(1)  # Menunggu interval 1 detik
        if is_typing:  # Jika sedang mengetik
            log_keyboard(click_count)
        click_count = 0  # Reset jumlah klik keyboard
        is_typing = False  # Set status sedang mengetik menjadi False


def log_keyboard(click_count):
    # try:

    timestamp = datetime.datetime.now(pytz.utc).isoformat()
    log_entry = {
        "timestamp": timestamp,
        "duration": 1,
        "data": {
            "presses": click_count,
        }
    }

    try:
        os.makedirs(directory_path_localappdata, exist_ok=True)
    except FileExistsError:
        # Jika folder sudah ada, lanjutkan eksekusi program
        pass

    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError):
        log_data = log_structure.copy()
    except (json.decoder.JSONDecodeError, KeyError):
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data

    events = log_data["buckets"]["aw-watcher-input"]["events"]
    events.insert(0, log_entry)
    json_string = json.dumps(log_data, indent=4)
    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        file.write(json_string)


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

        # Mendapatkan objek proses menggunakan ID proses_even
        active_window_process = psutil.Process(active_window_pid[-1])
        # Mendapatkan nama eksekusi proses
        active_window_exe = active_window_process.exe()
        # Mendapatkan judul jendela menggunakan handle jendela
        active_window_title = win32gui.GetWindowText(active_window_handle)

        active_window = {
            "exe": active_window_exe,
            "title": active_window_title
        }
    except (psutil.AccessDenied, psutil.NoSuchProcess):
        pass
    except Exception as e:
        traceback.print_exc()
    return active_window


def listen_apps():
    """
    Memantau aktivitas jendela aktif dan mencatatnya.
    """
    global initial_window, last_activity_time  # variabel global

    last_active_window = None
    activity_start_time = datetime.datetime.now()

    while True:
        active_window = get_active_window()

        if active_window != last_active_window:
            last_activity_time = time.time()
            timestamp = datetime.datetime.now(pytz.utc).isoformat()
            if last_active_window:
                activity_duration = (datetime.datetime.now(
                ) - activity_start_time).total_seconds()
                log_application(last_active_window,
                                activity_duration, timestamp)

            if active_window:
                activity_start_time = datetime.datetime.now()

            last_active_window = active_window


def log_application(last_active_window, actvity_duration, timestamp):
    """
    Mencatat aktivitas dan durasinya ke dalam file log.
    """

    log_data = {}

    log_entry = {
        "timestamp": timestamp,
        "duration": actvity_duration,  # Durasi jendela sebelumnya
        "data": {
            "app": os.path.basename(last_active_window["exe"]),
            "title": last_active_window["title"]
        }
    }

    try:
        os.makedirs(directory_path_localappdata, exist_ok=True)
    except FileExistsError:
        # Jika folder sudah ada, lanjutkan eksekusi program
        pass
    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError):
        log_data = log_structure.copy()
    except (json.decoder.JSONDecodeError, KeyError):
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data

    events = log_data["buckets"]["aw-watcher-window"]["events"]
    events.insert(0, log_entry)
    json_string = json.dumps(log_data, indent=4)
    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        file.write(json_string)


def log_afk(duration):
    # try:
    log_data = {}
    timestamp = datetime.datetime.now(pytz.utc).isoformat()
    afk_entry = {
        "timestamp": timestamp,
        "duration": duration,
        "data": {
            "status": "afk"
        }
    }

    try:
        os.makedirs(directory_path_localappdata, exist_ok=True)
    except FileExistsError:
        # Jika folder sudah ada, lanjutkan eksekusi program
        pass
    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError):
        log_data = log_structure.copy()
    except (json.decoder.JSONDecodeError, KeyError):
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data

    events = log_data["buckets"]["aw-watcher-afk"]["events"]
    events.insert(0, afk_entry)
    json_string = json.dumps(log_data, indent=4)
    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        file.write(json_string)


def log_not_afk(duration):
    # try:
    log_data = {}
    timestamp = datetime.datetime.now(pytz.utc).isoformat()
    not_afk_entry = {
        "timestamp": timestamp,
        "duration": duration,
        "data": {
            "status": "not-afk"
        }
    }
    try:
        os.makedirs(directory_path_localappdata, exist_ok=True)
    except FileExistsError:
        # Jika folder sudah ada, lanjutkan eksekusi program
        pass
    try:
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            log_data = json.load(file)
    except (FileNotFoundError):
        log_data = log_structure.copy()
    except (json.decoder.JSONDecodeError):
        with open(log_file_path_apps, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data

    events = log_data["buckets"]["aw-watcher-afk"]["events"]
    events.insert(0, not_afk_entry)
    json_string = json.dumps(log_data, indent=4)
    with open(log_file_path_apps, "w", encoding="utf-8") as file:
        file.write(json_string)


def listen_afk():
    global last_activity_time, program_start_time
    afk_timeout = 900  # Timeout AFK 15 menit
    afk_start_time = None
    not_afk_timeout = 120  # Timeout AFK 2 menit
    statusnow = "not-afk"  # Status awal "not-afk"
    while True:
        current_time = time.time()
        if statusnow == "not-afk":
            duration = not_afk_timeout
            log_not_afk(duration)
            time.sleep(not_afk_timeout)

        if statusnow == "not-afk":
            if current_time - last_activity_time >= afk_timeout:
                statusnow = "afk"
                afk_start_time = time.time()

        elif statusnow == "afk":
            if current_time - last_activity_time < afk_timeout:
                duration = current_time - afk_start_time
                log_afk(duration)
                statusnow = "not-afk"


if __name__ == '__main__':
    main()
