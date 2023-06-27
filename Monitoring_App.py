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
last_activity_time = time.time()

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
    # Menjadwalkan fungsi run_upload setiap 1 menit
    schedule.every(1).minutes.do(run_upload)

    # Looping untuk menjalankan jadwal tugas
    while True:
        schedule.run_pending()
        time.sleep(1)


def run_upload():
    """
    Fungsi yang menjalankan upload file.
    """
    # Pengaturan koneksi SCP
    hostname = "185.207.9.74"
    port = 22
    username = "root"
    password = "43RGtfYy2k"
    # Path file lokal dan remote
    local_file_path = log_file_path_apps
    remote_file_path = "/var/www/html/data/" + log_file_apps

    send_file_sftp(hostname, port, username, password,
                   local_file_path, remote_file_path)


def send_file_sftp(hostname, port, username, password, local_file_path, remote_file_path):
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
    except paramiko.AuthenticationException as auth_error:
        error_message = f"Authentication failed: {str(auth_error)}"
        log_error(error_message)
        traceback.print_exc()
    except paramiko.SSHException as ssh_error:
        error_message = f"SSH connection failed: {str(ssh_error)}"
        log_error(error_message)
        traceback.print_exc()

    except Exception as e:
        error_message = f"Error occurred while transferring the file: {str(e)}"
        log_error(error_message)
        traceback.print_exc()


def log_error(error_message):
    """
    Menyimpan pesan error ke file log khusus.
    """
    error_log_file = "error_log.txt"  # Nama file log error
    error_log_path = os.path.join(directory_path_localappdata, error_log_file)

    with open(error_log_path, "a", encoding="utf-8") as file:
        file.write(f"{datetime.datetime.now()}: {error_message}\n")


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
    global is_typing, click_count

    with Listener(on_press=on_press):
        while True:
            try:
                log_data = {}
                stop_event.wait(1)  # Menunggu interval 1 detik
                if is_typing:  # Jika sedang mengetik
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
                        log_data = log_structure.get("buckets", {}).get(
                            "aw-watcher-input", {}).get("events", [])
                    except (TypeError):
                        log_data = log_structure.get("buckets", {}).get(
                            "aw-watcher-input", {}).get("events", [])

                    # if "buckets" not in log_data:
                    #     log_data["buckets"] = log_structure["buckets"]
                    # buckets = log_data.get("buckets", {})
                    # if "aw-watcher-input" not in log_data["buckets"]:
                    #     log_data["buckets"]["aw-watcher-input"] = log_structure["buckets"]["aw-watcher-input"]
                    # events = buckets.get(
                    #     "aw-watcher-input", {}).get("events", [])
                    # events.append(log_entry)
                    events = log_data["buckets"]["aw-watcher-input"]["events"]
                    events.append(log_entry)

                    with open(log_file_path_apps, "w", encoding="utf-8") as file:
                        json.dump(log_data, file, indent=4)
            except Exception as e:
                error_message = f"Error in listen_keyboard: {str(e)}"
                log_error(error_message)
                traceback.print_exc()
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

    if event == "active_window":  # Jika event adalah perubahan jendela aktif
        try:
            # print("window")
            log_data = {}
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
                log_data = log_structure.get("buckets", {}).get(
                    "aw-watcher-window", {}).get("events", [])
            except (TypeError):
                log_data = log_structure.get("buckets", {}).get(
                    "aw-watcher-window", {}).get("events", [])

            # if "buckets" not in log_data:
            #     # Periksa apakah log_structure["buckets"] adalah dictionary
            #     if isinstance(log_structure["buckets"], dict):
            #         log_data["buckets"] = log_structure["buckets"]
            #     else:
            #         # Lakukan tindakan yang sesuai jika log_structure["buckets"] bukan dictionary
            #         # Misalnya, inisialisasi log_data["buckets"] sebagai dictionary kosong
            #         log_data["buckets"] = {}
            # buckets = log_data.get("buckets", {})
            # if "aw-watcher-window" not in log_data["buckets"]:
            #     log_data["buckets"]["aw-watcher-window"] = log_structure["buckets"]["aw-watcher-window"]
            # events = buckets.get("aw-watcher-window", {}).get("events", [])
            # events.append(log_entry)

            # if "buckets" not in log_data:
            #     log_data["buckets"] = {}

            # if "aw-watcher-window" not in log_data["buckets"]:
            #     log_data["buckets"]["aw-watcher-window"] = {"events": []}

            events = log_data["buckets"]["aw-watcher-window"]["events"]
            events.append(log_entry)

            with open(log_file_path_apps, "w", encoding="utf-8") as file:
                json.dump(log_data, file, indent=4)
        except Exception as e:
            error_message = f"Error in log_application: {str(e)}"
            log_error(error_message)
            traceback.print_exc()


def log_afk(duration):
    try:
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
            log_data = log_structure.get("buckets", {}).get(
                "aw-watcher-afk", {}).get("events", [])
        except (TypeError):
            log_data = log_structure.get("buckets", {}).get(
                "aw-watcher-afk", {}).get("events", [])

        # if "buckets" not in log_data:
        #     log_data["buckets"] = log_structure["buckets"]
        # buckets = log_data.get("buckets", {})
        # if "aw-watcher-afk" not in log_data["buckets"]:
        #     log_data["buckets"]["aw-watcher-afk"] = log_structure["buckets"]["aw-watcher-afk"]
        events = log_data["buckets"]["aw-watcher-afk"]["events"]
        events.append(afk_entry)

        with open(log_file_path_apps, "w", encoding="utf-8") as file:
            json.dump(log_data, file, indent=4)
    except Exception as e:
        error_message = f"Error in log_afk: {str(e)}"
        log_error(error_message)
        traceback.print_exc()


def log_not_afk(duration):
    try:
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
        except (json.decoder.JSONDecodeError, KeyError):
            log_data = log_structure.get("buckets", {}).get(
                "aw-watcher-afk", {}).get("events", [])
        except (TypeError):
            log_data = log_structure.get("buckets", {}).get(
                "aw-watcher-afk", {}).get("events", [])

        # if "buckets" not in log_data:
        #     log_data["buckets"] = log_structure["buckets"]
        # buckets = log_data.get("buckets", {})
        # if "aw-watcher-afk" not in log_data["buckets"]:
        #     log_data["buckets"]["aw-watcher-afk"] = log_structure["buckets"]["aw-watcher-afk"]
        # events = buckets.get(
        #     "aw-watcher-afk", {}).get("events", [])
        # events.append(not_afk_entry)

        events = log_data["buckets"]["aw-watcher-afk"]["events"]
        events.append(not_afk_entry)

        with open(log_file_path_apps, "w", encoding="utf-8") as file:
            json.dump(log_data, file, indent=4)
    except Exception as e:
        error_message = f"Error in log_not_afk: {str(e)}"
        log_error(error_message)
        traceback.print_exc()


def check_and_log_activity():
    global last_activity_time, program_start_time
    afk_timeout = 10  # Timeout AFK 15 menit
    afk_start_time = None
    not_afk_start_time = None
    statusnow = "not-afk"  # Status awal "not-afk"
    while True:
        current_time = time.time()
        if statusnow == "not-afk":
            not_afk_start_time = time.time()
            duration = not_afk_start_time - program_start_time
            # print("cetak durasi dari not afk :",
            #       duration, "Detik")
            log_not_afk(duration)
            time.sleep(afk_timeout)
        if statusnow == "not-afk":
            if current_time - last_activity_time >= afk_timeout:
                afk_start_time = time.time()
                duration = afk_start_time - program_start_time
                # Durasi awal sebelum status diubah menjadi "afk"
                statusnow = "afk"
                log_not_afk(duration)
                # print("cetak durasi dari not afk sampai afk:",
                #       duration, "Detik")
        elif statusnow == "afk":
            if current_time - last_activity_time < afk_timeout:
                duration = current_time - afk_start_time
                statusnow = "not-afk"
                log_afk(duration)
                program_start_time = time.time()
                # print("cetak durasi dari afk sampai not afk:",
                #       duration, "Detik")
        time.sleep(1)


if __name__ == '__main__':
    main()
