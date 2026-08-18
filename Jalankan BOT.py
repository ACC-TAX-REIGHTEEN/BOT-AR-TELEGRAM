import os
import subprocess
import sys


def jalankan_otomatisasi():
    folder_dapur = "Dapur"

    file_syarat = [
        "0_DownloaderData.py",
        "1_CopyData.py",
        "2_AdjDateFormat.py",
        "3_ARBotTelegram.py",
        "config.conf",
    ]

    print("--> Memeriksa struktur folder dan file...")

    if not os.path.exists(folder_dapur) or not os.path.isdir(folder_dapur):
        print(f"--> [ERROR] Folder '{folder_dapur}' tidak ditemukan!")
        input("--> Tekan Enter untuk keluar.")
        return

    for file in file_syarat:
        jalur_file = os.path.join(folder_dapur, file)
        if not os.path.isfile(jalur_file):
            print(
                f"--> [ERROR] File '{file}' tidak ditemukan di dalam folder '{folder_dapur}'."
            )
            input("--> Tekan Enter untuk keluar.")
            return

    print("--> Semua file syarat ditemukan. Memulai proses...\n")

    scripts_persiapan = ["0_DownloaderData.py", "1_CopyData.py", "2_AdjDateFormat.py"]

    try:
        for script in scripts_persiapan:
            print(f"--> Memulai eksekusi: {script}")

            proses = subprocess.run([sys.executable, script], cwd=folder_dapur)

            if proses.returncode != 0:
                print(
                    f"--> [ERROR] {script} berhenti dengan kesalahan (Code {proses.returncode})."
                )
                input("--> Tekan Enter untuk keluar.")
                return

            print(f"--> Selesai: {script}\n")

        print("--> ==================================================")
        print("--> Memulai eksekusi 3_ARBotTelegram.py")
        print("--> Bot Telegram aktif dan berjalan secara terus-menerus.")
        print("--> Tekan Ctrl+C kapan saja untuk menghentikan Bot.")
        print("--> ==================================================\n")

        subprocess.run([sys.executable, "3_ARBotTelegram.py"], cwd=folder_dapur)

    except KeyboardInterrupt:
        print(
            "--> [INFO] Bot Telegram berhasil dihentikan oleh pengguna (Ctrl+C)."
        )
    except Exception as e:
        print(f"--> [ERROR] Terjadi kesalahan yang tidak terduga: {e}")

    print("--> Seluruh alur kerja telah selesai.")
    input("--> Tekan Enter untuk keluar.")


if __name__ == "__main__":
    jalankan_otomatisasi()