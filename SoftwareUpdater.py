import customtkinter as ctk
from PIL import Image
import ctypes
import requests
import locale
import sys
import os
import platform
import subprocess
import webbrowser
import threading

VERSION = '1.1'

CHECK_NEW_VERSION_URL = 'https://api.github.com/repos/bartekl1/SoftwareUpdater/releases/latest'


polish_text = [
    'Przygotowywanie...',
    'Wszystko w porządku.',
    'Ostrzeżenia',
    'Błędy',
    'Program działa tylko na systemie Windows.',
    'Brak połączenia z internetem.',
    'Dostępna nowa wersja.',
    'Nie można sprawdzić dostępności nowej wersji.',
    'Nie uruchomiono programu jako administrator.',
    'Winget nie jest dostępny',
    'Pobierz',
    'Uruchom jako administrator',
    'Dalej',
    'Sprawdzanie dostępnych aktualizacji...',
    'Dostępne aktualizacje',
    'Instalowanie aktualizacji...',
    'Zakończono',
    'Sukces:',
    'Wymagane ponowne uruchomienie komputera:',
    'Błąd:',
    'Zakończ',
    'Uruchom ponownie'
]

english_text = [
    'Preparing...',
    'Everything is fine.',
    'Warnings',
    'Errors',
    'Program works only on Windows.',
    'No internet connection.',
    'New version available.',
    'Can\'t check for new version.',
    'Not running as administrator.',
    'Winget is not available.',
    'Download',
    'Run as administrator',
    'Next',
    'Checking for updates...',
    'Available updates',
    'Installing updates...',
    'Completed',
    'Success:',
    'Required PC reboot:',
    'Error:',
    'Exit',
    'Reboot'
]

possible_issues = [
    'not_windows',
    'no_internet_connection',
    'new_version_available',
    'cannot_check_new_version',
    'not_admin',
    'winget_not_available'
]


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def get_text(text_number):
    if language == 'pl_PL':
        return polish_text[text_number]
    else:
        return english_text[text_number]


def download_new_version():
    webbrowser.open(
        'https://github.com/bartekl1/SoftwareUpdater/releases/latest',
        new=2)


def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1)
    sys.exit(0)


def check_system():
    global preparing_frame, preparing_progressbar

    home_frame.forget()

    preparing_frame = ctk.CTkFrame(window,
                                   fg_color='transparent')

    preparing_label = ctk.CTkLabel(preparing_frame,
                                   text=get_text(0),
                                   font=('Calibri', 22))
    preparing_label.pack()

    preparing_progressbar = ctk.CTkProgressBar(preparing_frame,
                                               width=300)
    preparing_progressbar.pack(pady=4)
    preparing_progressbar.configure(mode="indeterminnate")
    preparing_progressbar.start()

    preparing_frame.pack(padx=4, pady=4)

    t = threading.Thread(target=check_system2)
    t.start()


def check_system2():
    global check_system_frame

    issues = []

    # 1. Check OS (Windows only)
    if platform.system() != 'Windows':
        issues.append({'type': 'critical',
                       'message': 'not_windows'})

    # 2. Check internet connection
    try:
        requests.get('https://google.com/')
    except Exception:
        issues.append({'type': 'critical',
                       'message': 'no_internet_connection'})

    # 3. Check for new version
    try:
        r = requests.get(CHECK_NEW_VERSION_URL)
        latest_version = r.json()['name']
        if latest_version != VERSION:
            issues.append({
                'type': 'warning',
                'message': 'new_version_available',
                'version': latest_version
            })
    except Exception:
        issues.append({
            'type': 'warning',
            'message': 'cannot_check_new_version'
        })

    # 4. Check for admin privileges
    if not ctypes.windll.shell32.IsUserAnAdmin():
        issues.append({
            'type': 'warning',
            'message': 'not_admin'
        })

    # 5. Check for winget
    try:
        subprocess.run('winget', stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    except FileNotFoundError:
        issues.append({
            'type': 'critical',
            'message': 'winget_not_available'
        })

    preparing_progressbar.stop()
    preparing_frame.forget()

    check_system_frame = ctk.CTkFrame(window,
                                      fg_color='transparent')
    check_system_frame.pack(padx=4, pady=4)

    if len(issues) == 0:
        check_system_label = ctk.CTkLabel(check_system_frame,
                                          text=get_text(1),
                                          font=('Calibri', 22),
                                          text_color='green')
    elif 'critical' not in [issue['type'] for issue in issues]:
        check_system_label = ctk.CTkLabel(check_system_frame,
                                          text=get_text(2),
                                          font=('Calibri', 22),
                                          text_color='orange')
    else:
        check_system_label = ctk.CTkLabel(check_system_frame,
                                          text=get_text(3),
                                          font=('Calibri', 22),
                                          text_color='red')
    check_system_label.pack()

    issues_frame = ctk.CTkFrame(check_system_frame,
                                fg_color='transparent')
    issues_frame.pack()

    added_rows = 0

    for i, issue in enumerate(issues):
        if issue['type'] == 'warning':
            issue_icon_image = ctk.CTkImage(
                Image.open(resource_path('img/warning.png')),
                size=(36, 36))
        else:
            issue_icon_image = ctk.CTkImage(
                Image.open(resource_path('img/error.png')),
                size=(36, 36))
        issue_icon_label = ctk.CTkLabel(issues_frame,
                                        image=issue_icon_image,
                                        text='')
        issue_icon_label.grid(column=0, row=i + added_rows,
                              padx=2, pady=2)

        if issue['message'] != 'new_version_available':
            issue_description_label = ctk.CTkLabel(issues_frame,
                                                   text=get_text(possible_issues.index(issue['message']) + 4))
        else:
            issue_description_label = ctk.CTkLabel(issues_frame,
                                                   text=f'{get_text(6)} ({VERSION} -> {issue["version"]})')
        issue_description_label.grid(column=1, row=i + added_rows,
                                     padx=10, pady=2, sticky='w')

        if issue['message'] == 'new_version_available':
            download_image = ctk.CTkImage(
                Image.open(resource_path('img/download.png')),
                size=(28, 28))
            download_new_version_button = ctk.CTkButton(issues_frame,
                                                        text=get_text(10),
                                                        image=download_image,
                                                        anchor='w',
                                                        command=download_new_version)
            download_new_version_button.grid(column=1, row=i + added_rows + 1,
                                             padx=10, pady=2, sticky='w')
            added_rows += 1
        elif issue['message'] == 'not_admin':
            run_as_admin_image = ctk.CTkImage(
                Image.open(resource_path('img/run_as_admin.png')),
                size=(28, 28))
            run_as_admin_button = ctk.CTkButton(issues_frame,
                                                text=get_text(11),
                                                image=run_as_admin_image,
                                                anchor='w',
                                                command=run_as_admin)
            run_as_admin_button.grid(column=1, row=i + added_rows + 1,
                                     padx=10, pady=2, sticky='w')
            added_rows += 1

    if 'critical' not in [issue['type'] for issue in issues]:
        next_button = ctk.CTkButton(check_system_frame,
                                    text=get_text(12),
                                    command=check_for_winget_packages_updates)
        next_button.pack(anchor='se', pady=12)


def check_for_winget_packages_updates():
    global checking_updates_frame

    check_system_frame.forget()

    checking_updates_frame = ctk.CTkFrame(window,
                                          fg_color='transparent')

    checking_label = ctk.CTkLabel(checking_updates_frame,
                                  text=get_text(13),
                                  font=('Calibri', 22))
    checking_label.pack()

    checking_progressbar = ctk.CTkProgressBar(checking_updates_frame,
                                              width=300)
    checking_progressbar.pack(pady=4)
    checking_progressbar.configure(mode="indeterminnate")
    checking_progressbar.start()

    checking_updates_frame.pack(padx=4, pady=4)

    t = threading.Thread(target=check_updates)
    t.start()


def check_updates():
    global available_updates_frame, \
        available_updates

    s = subprocess.run('winget upgrade',
                       stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    output = s.stdout.decode().split('\r\n')
    title_row = output[0].split('\r')[-1]
    lengths = {
        'name': len(title_row.split('Id')[0]),
        'id': len(title_row.split('Version')[0].split('Id')[1]) + len('Id'),
        'version': len(title_row.split('Available')[0].split('Version')[1]) + len('Version'),
        'available': len(title_row.split('Source')[0].split('Available')[1]) + len('Available')
    }
    rows = output[2:-2]
    available_updates = []
    for row in rows:
        if ' upgrades available.' in row:
            break
        available_updates.append({
            'name': row[:lengths['name']],
            'id': row[lengths['name']:lengths['name']+lengths['id']],
            'version': row[lengths['name']+lengths['id']:lengths['name']+lengths['id']+lengths['version']],
            'available': row[lengths['name']+lengths['id']+lengths['version']:lengths['name']+lengths['id']+lengths['version']+lengths['available']]
        })
    for i, row in enumerate(available_updates):
        new_row = {}
        for key, value in row.items():
            splitted_value = value.split(' ')
            while splitted_value[-1] == '':
                splitted_value.pop()
            new_row[key] = ' '.join(splitted_value)
        available_updates[i] = new_row

    checking_updates_frame.forget()

    available_updates_frame = ctk.CTkFrame(window,
                                           fg_color='transparent')
    available_updates_frame.pack(padx=4, pady=4)

    available_updates_label = ctk.CTkLabel(available_updates_frame,
                                           text=get_text(14),
                                           font=('Calibri', 22))
    available_updates_label.pack()

    updates_frame = ctk.CTkScrollableFrame(available_updates_frame,
                                           width=350)
    updates_frame.pack(pady=4)

    for i, update in enumerate(available_updates):
        row = update
        row['variable'] = ctk.IntVar()
        row['checkbox'] = ctk.CTkCheckBox(updates_frame,
                                          text=f'{update["name"]} ({update["version"]} -> {update["available"]})',
                                          variable=row['variable'],
                                          offvalue=0, onvalue=1,
                                          font=('Calibri', 14))
        row['checkbox'].pack(anchor='w', pady=2)

    next_button = ctk.CTkButton(available_updates_frame,
                                text=get_text(12),
                                command=instal_winget_packages_updates)
    next_button.pack(anchor='se', pady=12)


def instal_winget_packages_updates():
    global installing_frame, number_of_updates, installed_updates_var, \
        currently_installing

    available_updates_frame.forget()

    number_of_updates = [update['variable'].get()
                         for update in available_updates].count(1)

    installing_frame = ctk.CTkFrame(window,
                                    fg_color='transparent')
    installing_frame.pack(padx=4, pady=4)

    installing_label = ctk.CTkLabel(installing_frame,
                                    text=get_text(15),
                                    font=('Calibri', 22))
    installing_label.pack()

    installed_updates_var = ctk.DoubleVar(value=0)

    installing_progressbar = ctk.CTkProgressBar(installing_frame,
                                                width=300,
                                                variable=installed_updates_var)
    installing_progressbar.pack(pady=4)

    currently_installing = ctk.StringVar()

    name_label = ctk.CTkLabel(installing_frame,
                              textvariable=currently_installing)
    name_label.pack()

    t = threading.Thread(target=install_updates)
    t.start()


def install_updates():
    installed_updates = 0

    for i, update in enumerate(available_updates):
        if update['variable'].get() == 1:
            currently_installing.set(
                f'{update["name"]} ({update["version"]} -> {update["available"]})')

            s = subprocess.run(f'winget upgrade --id {update["id"]}',
                               stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            output = s.stdout.decode()

            if 'Successfully installed' in output:
                available_updates[i]['status'] = 'success'
            elif 'Installer failed with exit code: 3010' in output:
                available_updates[i]['status'] = 'reboot_required'
            else:
                available_updates[i]['status'] = 'error'

            installed_updates += 1
            installed_updates_var.set(installed_updates / number_of_updates)
        else:
            available_updates[i]['status'] = 'skipped'

    installing_frame.forget()

    end_frame = ctk.CTkFrame(window,
                             fg_color='transparent')
    end_frame.pack(padx=4, pady=4)

    end_label = ctk.CTkLabel(end_frame,
                             text=get_text(16),
                             font=('Calibri', 22))
    end_label.pack()

    installed_frame = ctk.CTkScrollableFrame(end_frame,
                                             width=350)
    installed_frame.pack(padx=4, pady=4)

    rows = 0

    if 'success' in [u['status'] for u in available_updates]:
        icon_image = ctk.CTkImage(
            Image.open(resource_path('img/success.png')),
            size=(36, 36))
        icon_label = ctk.CTkLabel(installed_frame,
                                  image=icon_image,
                                  text='')
        icon_label.grid(column=0, row=rows,
                        padx=2, pady=2)

        label = ctk.CTkLabel(installed_frame,
                             text=get_text(17),
                             text_color='green')
        label.grid(column=1, row=rows,
                   padx=10, pady=2, sticky='w')

        rows += 1

        for update in available_updates:
            if update['status'] == 'success':
                update_label = ctk.CTkLabel(installed_frame,
                                            text=f'{update["name"]} ({update["version"]} -> {update["available"]})')
                update_label.grid(column=1, row=rows,
                                  padx=10, pady=2, sticky='w')

                rows += 1

    if 'reboot_required' in [u['status'] for u in available_updates]:
        icon_image = ctk.CTkImage(
            Image.open(resource_path('img/warning.png')),
            size=(36, 36))
        icon_label = ctk.CTkLabel(installed_frame,
                                  image=icon_image,
                                  text='')
        icon_label.grid(column=0, row=rows,
                        padx=2, pady=2)

        label = ctk.CTkLabel(installed_frame,
                             text=get_text(18),
                             text_color='orange')
        label.grid(column=1, row=rows,
                   padx=10, pady=2, sticky='w')

        rows += 1

        for update in available_updates:
            if update['status'] == 'reboot_required':
                update_label = ctk.CTkLabel(installed_frame,
                                            text=f'{update["name"]} ({update["version"]} -> {update["available"]})')
                update_label.grid(column=1, row=rows,
                                  padx=10, pady=2, sticky='w')

                rows += 1

    if 'error' in [u['status'] for u in available_updates]:
        icon_image = ctk.CTkImage(
            Image.open(resource_path('img/error.png')),
            size=(36, 36))
        icon_label = ctk.CTkLabel(installed_frame,
                                  image=icon_image,
                                  text='')
        icon_label.grid(column=0, row=rows,
                        padx=2, pady=2)

        label = ctk.CTkLabel(installed_frame,
                             text=get_text(19),
                             text_color='red')
        label.grid(column=1, row=rows,
                   padx=10, pady=2, sticky='w')

        rows += 1

        for update in available_updates:
            if update['status'] == 'error':
                update_label = ctk.CTkLabel(installed_frame,
                                            text=f'{update["name"]} ({update["version"]} -> {update["available"]})')
                update_label.grid(column=1, row=rows,
                                  padx=10, pady=2, sticky='w')

                rows += 1

    if 'reboot_required' in [u['status'] for u in available_updates]:
        reboot_button = ctk.CTkButton(end_frame,
                                      text=get_text(21),
                                      font=('Calibri', 24),
                                      command=reboot)
        reboot_button.pack(anchor='se', pady=14)

    exit_button = ctk.CTkButton(end_frame,
                                text=get_text(20),
                                font=('Calibri', 24),
                                command=window.destroy)
    exit_button.pack(anchor='se', pady=14)


def reboot():
    os.system('shutdown -r -t 0')


def main():
    global language, \
        window, home_frame

    windll = ctypes.windll.kernel32
    language = locale.windows_locale[windll.GetUserDefaultUILanguage()]

    window = ctk.CTk()

    window.minsize(375, 375)

    window.title('Software Updater')
    window.iconbitmap(resource_path('img/elephant.ico'))

    home_frame = ctk.CTkFrame(window,
                              fg_color='transparent')
    home_frame.pack(padx=4, pady=4)

    title_label = ctk.CTkLabel(home_frame,
                               text=f'Software Updater {VERSION}',
                               font=('Calibri', 30))
    title_label.pack()

    start_button = ctk.CTkButton(home_frame,
                                 text='Start',
                                 font=('Calibri', 24),
                                 command=check_system)
    start_button.pack(anchor='se', pady=14)

    window.mainloop()


if __name__ == '__main__':
    main()
