import sys
import pythoncom
from os import path, unlink, listdir, mkdir, rename, chmod, environ
from stat import S_IWRITE
from wmi import WMI
from winreg import HKEY_CURRENT_USER, HKEY_USERS, KEY_ALL_ACCESS, DeleteKey, DeleteValue, QueryValueEx, REG_DWORD
from winreg import OpenKey, QueryInfoKey, EnumKey, ConnectRegistry, HKEY_LOCAL_MACHINE, KEY_SET_VALUE, SetValueEx
from shutil import rmtree
from subprocess import run
from time import sleep, time
from getpass import getuser
from win32net import NetShareEnum
from random import random
from datetime import datetime, timedelta
from json import load
from logging import getLogger, basicConfig, exception
from threading import Thread
from concurrent.futures import TimeoutError
from functools import wraps
from pynput import mouse
from concurrent.futures import ThreadPoolExecutor
from psutil import disk_usage
from pyad import adgroup, adquery, aduser
import tkinter.messagebox
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, INSERT, messagebox, END, ttk, CENTER, SEL


def redirect(output: str) -> None:
    console.configure(state="normal")
    console.insert(END, output)
    console.see(END)
    console.configure(state="disabled")


def print_error(obj: Text, output: str = "", additional: str = "", clear_: bool = False, see: bool = False,
                newline: bool = False) -> None:
    if clear_:
        clear_obj(obj)
    obj.configure(state="normal")
    obj.tag_configure('red', foreground='red')
    if additional:
        obj.insert(INSERT, additional)
    obj.insert(END, output, "red")
    if newline:
        obj.insert(END, "\n")
    if see:
        obj.see(END)
    obj.configure(state="disabled")


def print_success(obj: Text, output: str = "", additional: str = "", clear_: bool = False, see: bool = False,
                  newline: bool = False) -> None:
    if clear_:
        clear_obj(obj)
    obj.configure(state="normal")
    obj.tag_configure('green', foreground='green')
    if additional:
        obj.insert(INSERT, additional)
    obj.insert(END, output, "green")
    if newline:
        obj.insert(END, "\n")
    if see:
        obj.see(END)
    obj.configure(state="disabled")


def update(obj: Text, statement: str) -> None:
    obj.configure(state="normal")
    obj.delete("1.0", END)
    obj.insert(INSERT, statement)
    obj.configure(state="disabled")


def update_error(obj: Text, initial: str, statement: str) -> None:
    obj.configure(state="normal")
    obj.tag_configure('red', foreground='red')
    obj.delete("1.0", END)
    obj.insert(INSERT, initial)
    obj.insert(INSERT, statement, "red")
    obj.configure(state="disabled")


def clear_obj(obj: Text) -> None:
    obj.configure(state="normal")
    obj.delete("1.0", END)
    obj.configure(state="disabled")


def tsleep(secs: float) -> None:
    t = Thread(target=lambda: t_actual_sleep(secs), daemon=True)
    t.start()
    while t.is_alive():
        refresh()
        sleep(0.05)


def t_actual_sleep(secs: float) -> None:
    sleep(secs)


def clear_all() -> None:
    for obj in [[display_pc, "Current computer: "], [computer_status, "Computer status: "], [console, ""],
                [display_user, "Current user: "], [uptime, "Uptime: "], [space_c, "Space in C disk: "],
                [space_d, "Space in D disk: "], [ram, "Total RAM: "], [ie_fixed, "Internet explorer: "],
                [cpt_fixed, "Cockpit printer: "], [user_active, "User status: "]]:
        obj[0].configure(state="normal")
        obj[0].delete("1.0", END)
        obj[0].insert(INSERT, obj[1])
        obj[0].configure(state="disabled")
        config.current_computer = None
        config.current_user = None


def disable(disable_submit: bool = False) -> None:
    computer.unbind("<Return>")
    for obj in (reset_spool, fix_cpt, fix_ie, clear_space, get_printers, delete_ost, delete_users, fix_3_lang,
                copy_but):
        obj.configure(state="disabled", cursor="arrow")
    if disable_submit:
        submit.configure(state="disabled")
    if config.first_time:
        config.first_time = 0
        computer.bind("<Return>", lambda _: on_submit())
        submit.configure(cursor="hand2")


def enable() -> None:
    for obj in (reset_spool, fix_cpt, fix_ie, clear_space, get_printers, delete_ost, delete_users, fix_3_lang, submit,
                copy_but):
        obj.configure(state="normal", cursor="hand2")
    computer.bind("<Return>", lambda _: on_submit())
    if not config.current_user:
        delete_ost.configure(state="disabled", cursor="arrow")
        fix_cpt.configure(state="disabled", cursor="arrow")


def show_text(_) -> None:
    if computer.get() == "Computer or User":
        computer.delete(0, END)
        computer.config(justify="center")


def hide_text(_) -> None:
    if computer.get() == "":
        computer.insert(0, "Computer or User")
        computer.config(justify="center")


def enable_paste(event) -> None:
    ctrl = (event.state & 0x4) != 0

    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")

    if event.keycode == 65 and ctrl and event.keysym.lower() != "a":
        event.widget.event_generate("<<SelectAll>>")


def copy_clip(to_copy: str) -> None:
    cp = Tk()
    cp.withdraw()
    cp.clipboard_clear()
    cp.clipboard_append(to_copy)
    cp.update()
    cp.destroy()


def on_button_press(_, __, button, ___):
    if button == mouse.Button.middle:
        if window.wm_state() == 'iconic':
            window.deiconify()
            window.lift()
            window.focus_set()
        else:
            window.iconify()


def disable_middle_click(event):
    if event.num == 2:
        return "break"


def ignore_selection(obj, _):
    obj.tag_remove(SEL, "1.0", END)
    return "break"


def asset(filename: str) -> str:
    return fr"{config.assets}\{filename}"


def create_selection_window(options):
    config.yes_no = False
    config.wll_delete = []

    def on_done():
        newline = "\n"
        selected_options = [check.get() for check in option_vars if check.get()]
        canvas_.unbind_all("<MouseWheel>")
        selection_window.destroy()
        window.focus_set()
        if not selected_options:
            print("No users were chosen to be deleted")
            config.yes_no = False
            return
        yes_no = messagebox.askyesno("Warning", f"Are you sure you want to delete the following users?\n"
                                                f"{newline.join(selected_options).replace('{', '').replace('}', ' -')}")
        if yes_no:
            config.wll_delete = [selected.split()[-1] for selected in selected_options if selected]
            config.yes_no = True
            return
        print("Canceled users deletion")
        config.yes_no = False
        return

    def disable_main_window(selection_window_):
        def on_window_close():
            window.grab_release()
            selection_window_.destroy()
            canvas_.unbind_all("<MouseWheel>")
            print("Canceled users deletion")

        window.grab_set()
        selection_window.protocol("WM_DELETE_WINDOW", on_window_close)
        window.wait_window(selection_window)

    def on_mousewheel(event):
        canvas_.yview_scroll(int(-1 * (event.delta / 120)), "units")

    selection_window = tkinter.Toplevel(window)
    selection_window.title("Select users")

    canvas_ = tkinter.Canvas(selection_window, height=200)
    scrollbar = tkinter.Scrollbar(selection_window, orient="vertical", command=canvas_.yview)
    scrollable_frame = tkinter.Frame(canvas_)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas_.configure(
            scrollregion=canvas_.bbox("all")
        )
    )

    canvas_.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas_.configure(yscrollcommand=scrollbar.set)
    canvas_.bind("<Configure>", lambda event: canvas_.configure(scrollregion=canvas_.bbox("all")))
    canvas_.bind_all("<MouseWheel>", on_mousewheel)

    option_vars = []
    for option in options:
        var = tkinter.StringVar()
        option_vars.append(var)
        checkbox = tkinter.Checkbutton(scrollable_frame, text=option[0], variable=var, onvalue=option, offvalue="")
        checkbox.pack(anchor=tkinter.W)

    canvas_.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    done_button = tkinter.Button(selection_window, text="Done", command=on_done)
    done_button.pack(pady=10)

    selection_window.protocol("WM_DELETE_WINDOW", disable_main_window)
    disable_main_window(selection_window)


class ProgressBar:
    def __init__(self, total_items, title_, end_statement):
        self.total_items = total_items
        self.title = title_
        self.current_item = 0

        self.root = window
        self.end_statement = end_statement

        self.label = tkinter.Label(self.root, text=self.title, background="#545664",

                                   fg='#FFFFFF', font=("Arial", 13, "bold"), justify=CENTER, )
        self.label.place(x=594.0, y=420.0, anchor=CENTER)
        ttk.Style().theme_use("xpnative")
        ttk.Style().layout('text.Horizontal.TProgressbar',
                           [('Horizontal.Progressbar.trough',
                             {'children': [('Horizontal.Progressbar.pbar',
                                            {'side': 'left', 'sticky': 'ns'})],
                              'sticky': 'nswe'}),
                            ('Horizontal.Progressbar.label', {'sticky': 'nswe'})])
        ttk.Style().configure('text.Horizontal.TProgressbar', text='0 %', anchor='center',
                              foreground='black')
        self.progressbar = ttk.Progressbar(window, length=297, mode="determinate", style='text.Horizontal.TProgressbar')
        self.progressbar.place(x=450.0, y=440.0)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        print(self.end_statement)
        self.progressbar.destroy()
        self.label.destroy()

    def __call__(self):
        self.current_item += 1
        self.root.update()
        if self.current_item <= self.total_items:
            progress = (self.current_item / self.total_items) * 100
            ttk.Style().configure('text.Horizontal.TProgressbar', text=f"{int(progress)} %")
            self.progressbar["value"] = progress
            self.root.update_idletasks()


def refresh() -> None:
    window.update_idletasks()
    window.update()


def fix_ie_func() -> None:
    pc = config.current_computer
    if not reg_connect():
        print_error(console, output="Could not fix internet explorer", newline=True)
    refresh()
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:
        for key_name in (
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C}",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C"):
            refresh()
            try:
                with OpenKey(reg, key_name, 0, KEY_ALL_ACCESS) as key:
                    DeleteKey(key, "")
            except FileNotFoundError:
                pass
            except:
                print_error(console, output="Unable to fix internet explorer", newline=True)
                log()
                return

    key = r"Software\Microsoft\Internet Explorer\BrowserEmulation"
    try:
        with ConnectRegistry(pc, HKEY_CURRENT_USER) as reg:
            with OpenKey(reg, key, 0, KEY_SET_VALUE) as key:
                # Set the value for IntranetCompatibilityMode
                SetValueEx(key, "IntranetCompatibilityMode", 0, REG_DWORD, 0)
                SetValueEx(key, "MSCompatibilityMode", 0, REG_DWORD, 0)
    except FileNotFoundError:
        pass
    except:
        log()

    update(ie_fixed, "Internet explorer: Fixed")
    print_success(console, output=f"Fixed internet explorer", newline=True)


def fix_cpt_func() -> None:
    if not reg_connect():
        print_error(console, output="ERROR, could not connect to remote registry", newline=True)
        return
    refresh()
    with ConnectRegistry(config.current_computer, HKEY_CURRENT_USER) as reg:
        try:
            with OpenKey(reg, r"SOFTWARE\Jetro Platforms\JDsClient\PrintPlugIn", 0, KEY_ALL_ACCESS) as key:
                DeleteValue(key, "PrintClientPath")
        except FileNotFoundError:
            pass
        except:
            log()
            print_error(console, output="Failed to fix cpt printer", newline=True)
            return
    print_success(console, output="Fixed cpt printer", newline=True)
    update(cpt_fixed, "Cockpit printer: Fixed")


def fix_3_languages() -> None:
    if not reg_connect():
        print_error(console, output="ERROR, could not connect to remote registry", newline=True)
        return
    refresh()
    with ConnectRegistry(config.current_computer, HKEY_USERS) as reg:
        refresh()
        try:
            with OpenKey(reg, r".DEFAULT\Keyboard Layout\Preload", 0, KEY_ALL_ACCESS) as key:
                DeleteKey(key, "")
        except FileNotFoundError:
            pass
        except:
            print_error(console, output="Failed to fix 3 languages bug", newline=True)
            log()
            return
    print_success(console, output="Fixed 3 languages bug", newline=True)


def reset_spooler() -> None:
    try:
        # noinspection PyUnboundLocalVariable
        refresh()
        connection = WMI(computer=config.current_computer)
        service = connection.Win32_Service(name="Spooler")
        refresh()
        service[0].StopService()
        tsleep(1)
        refresh()
        service[0].StartService()
        print_success(console, output=f"Successfully restarted the spooler", newline=True)
    except:
        print_error(console, output=f"Failed to restart the spooler", newline=True)
        log()


def delete_the_ost() -> None:
    user_ = config.current_user
    pc = config.current_computer
    if not tkinter.messagebox.askyesno(title="OST deletion",
                                       message=f"Are you sure you want to delete "
                                               f"the ost of {user_name_trasnslation(user_)}?"):
        print_error(console, output="Canceled OST deletion", newline=True)
        return
    try:
        host = WMI(computer=pc)
        for procs in ("lync.exe", "outlook.exe", "UcMapi.exe"):
            for proc in host.Win32_Process(name=procs):
                refresh()
                if proc:
                    try:
                        proc.Terminate()
                    except:
                        log()
    except:
        print_error(console, output="Could not connect to the computer", newline=True)
        log()
        return
    if path.exists(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook"):
        ost = listdir(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook")
        for file in ost:
            if file.endswith("ost"):
                ost = fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook\{file}"
                try:
                    tsleep(1)
                    rename(ost, f"{ost}{random():.3f}.old")
                    print_success(console, output=f"Successfully removed the ost file", newline=True)
                except FileExistsError:
                    try:
                        rename(ost, f"{ost}{random():.3f}.old")
                        print_success(console, output=f"Successfully removed the ost file", newline=True)
                    except:
                        log()
                        print_error(console, f"Could not Delete the OST file", newline=True)
                except:
                    print_error(console, f"Could not Delete the OST file", newline=True)
                    log()
                return
        else:
            print_error(console, f"Could not find an OST file", newline=True)
            return
    else:
        print_error(console, f"Could not find an OST file", newline=True)


def my_rm(file_: str, bar_: callable) -> None:
    try:
        if path.isfile(file_) or path.islink(file_):
            unlink(file_)
        elif path.isdir(file_):
            rmtree(file_, ignore_errors=True)
    except (PermissionError, FileNotFoundError):
        pass
    except:
        log()
    bar_()


def rmtree_recreate(dir_: str, bar_: callable = None) -> None:
    try:
        rmtree(dir_, ignore_errors=True)
        mkdir(dir_)
    except (FileExistsError, PermissionError):
        pass
    except:
        log()
    if bar_:
        bar_()


def clear_space_func() -> None:
    refresh()
    pc = config.current_computer
    users_dirs = listdir(fr"\\{pc}\c$\users")
    refresh()

    # get initial space
    space_init = get_space(pc)
    flag = False
    refresh()

    # clears search edb
    edb_file = fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
    if path.exists(edb_file) and config.delete_edb:
        try:
            connection = WMI(computer=pc)
            service = connection.Win32_Service(name="WSearch")
            refresh()
            service[0].StopService()
            tsleep(0.6)
            refresh()
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except PermissionError:
            pass
        except:
            log()

    # Deletes None user-specific files
    if config.c_paths_with_msg:
        for path_msg in config.c_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            refresh()
            if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                refresh()
                files = [fr"\\{pc}\c$\{path_msg[0]}\{file}" for file in listdir(fr"\\{pc}\c$\{path_msg[0]}")]
                refresh()
                with ProgressBar(len(files), path_msg[1], path_msg[-1]) as bar:
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        jobs = [executor.submit(my_rm, file, bar) for file in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)
                            window.update()

    # deletes temps of each user
    if config.delete_user_temp:
        with ProgressBar(len(users_dirs), f"Deleting temps of {len(users_dirs)} users",
                         f"Deleted temps of {len(users_dirs)} users") as bar:
            dirs = [fr"\\{pc}\c$\users\{dir_}\AppData\Local\Temp" for dir_ in users_dirs if
                    (dir_.lower().strip() != config.user.lower().strip() and config.current_computer.lower()
                     != config.host.lower().strip())]
            with ThreadPoolExecutor(max_workers=8) as executor:
                jobs = [executor.submit(my_rm, dir_, bar) for dir_ in dirs]
                while not all([result.done() for result in jobs]):
                    sleep(0.1)
                    window.update()

    if config.u_paths_with_msg:
        for path_msg in config.u_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            msg_ = path_msg[1].replace("users_amount", str(len(users_dirs)))
            with ProgressBar(len(users_dirs), msg_, path_msg[-1].replace(str(len(users_dirs)))) as bar:
                for user in users_dirs:
                    refresh()
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        refresh()
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        refresh()
                        with ThreadPoolExecutor(max_workers=8) as executor:
                            jobs = [executor.submit(my_rm, file) for file in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                                window.update()
                    bar()

    if config.u_paths_without_msg or config.c_paths_without_msg:
        with ProgressBar(len(config.u_paths_without_msg) + len(config.c_paths_without_msg), "Deleting additional files",
                         "Deleted additional files") as bar:
            for path_msg in config.c_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                refresh()
                if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                    refresh()
                    files = listdir(fr"\\{pc}\c$\{path_msg[0]}")
                    refresh()
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        jobs = [executor.submit(my_rm, file) for file in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)
                            window.update()
                bar()

            for path_msg in config.u_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                for user in users_dirs:
                    refresh()
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        refresh()
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        refresh()
                        with ThreadPoolExecutor(max_workers=8) as executor:
                            jobs = [executor.submit(my_rm, file) for file in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                                window.update()
                bar()

    if not flag and config.delete_edb and path.exists(edb_file):
        try:
            connection = WMI(computer=pc)
            service = connection.Win32_Service(name="WSearch")
            refresh()
            service[0].StopService()
            tsleep(0.8)
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except (PermissionError, FileNotFoundError):
            pass
        except:
            log()
    if flag and config.delete_edb:
        print(fr"Deleted the search.edb file")
    else:
        if config.delete_edb:
            print_error(console, output="Failed to remove search.edb file", newline=True)
    space_final = get_space(pc)
    print_success(console, output=f"Cleared {abs((space_final - space_init)):.1f} GB from the disk", newline=True)
    try:
        space = get_space(pc)
        if space <= 5:
            update_error(space_c, "Space in C disk: ", f"{space:.1f}GB free out of "
                                                       f"{get_total_space(pc):.1f}GB")
        else:
            update(space_c, f"Space in C disk: {space:.1f}GB free out of {get_total_space(pc):.1f}GB")
    except:
        log()
        update_error(space_c, "Space in C disk: ", "ERROR")


def my_rmtree(dir_, bar_) -> None:
    if path.isdir(dir_):
        rmtree(dir_, onerror=on_rm_error)
    bar_()


def del_users() -> None:
    config.wll_delete = []
    config.yes_no = False
    pc = config.current_computer
    users_to_choose_delete = []
    for dir_ in listdir(fr"\\{pc}\c$\Users"):
        if dir_.lower() == config.current_user.lower() or dir_.lower() in config.exclude or \
                any([dir_.lower().startswith(exc_lude) for exc_lude in config.startwith_exclude]) \
                or not path.isdir(fr"\\{pc}\c$\users\{dir_}"):
            continue
        users_to_choose_delete.append([user_name_trasnslation(dir_), dir_])
    if not users_to_choose_delete:
        print("No users were found to delete")
        return
    create_selection_window(users_to_choose_delete)
    if config.yes_no:
        refresh()
        space_init = get_space(pc)
        with ProgressBar(len(config.wll_delete), f"Deleting {len(config.wll_delete)} folders",
                         f"Deleted {len(config.wll_delete)} users") as bar:
            config.wll_delete = [fr"\\{pc}\c$\users\{dir_}" for dir_ in config.wll_delete]
            with ThreadPoolExecutor(max_workers=5) as executor:
                jobs = [executor.submit(my_rmtree, dir_, bar) for dir_ in config.wll_delete]
                while not all([result.done() for result in jobs]):
                    sleep(0.1)
                    window.update()
        space_final = get_space(pc)
        print(f"Cleared {abs((space_final - space_init)):.1f} GB from the disk")
        try:
            space = get_space(pc)
            if space <= 5:
                update_error(space_c, "Space in C disk: ", f"{space:.1f}GB free out of "
                                                           f"{get_total_space(pc):.1f}GB")
            else:
                update(space_c, f"Space in C disk: {space:.1f}GB free out of {get_total_space(pc):.1f}GB")
        except:
            log()
            update_error(space_c, "Space in C disk: ", "ERROR")


def get_printers_func() -> None:
    found_any = False
    pc = config.current_computer
    if not reg_connect():
        return
    refresh()
    with ConnectRegistry(pc, HKEY_USERS) as reg:
        users_dict = {}
        sid_list = []
        with OpenKey(reg, "") as users:
            users_len = QueryInfoKey(users)[0]
            for i in range(users_len):
                try:
                    sid_list.append(EnumKey(users, i))
                except FileNotFoundError:
                    pass
                except:
                    log()
        refresh()

        with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as users_path:
            for sid in set(sid_list):
                refresh()
                try:
                    with OpenKey(users_path,
                                 fr"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{sid}") as profiles:
                        username = QueryValueEx(profiles, "ProfileImagePath")
                        if username[0].startswith("C:\\"):
                            username = username[0].split("\\")[-1]
                            users_dict[sid] = user_name_trasnslation(username)

                except FileNotFoundError:
                    pass
                except:
                    log()
        refresh()

        flag = False
        for sid in sid_list:
            try:
                with OpenKey(reg, fr"{sid}\Printers\Connections") as printer_path:
                    printers_len = QueryInfoKey(printer_path)[0]
                    for i in range(printers_len):
                        try:
                            printer = EnumKey(printer_path, i).replace(",", "\\").strip()
                            if not flag:
                                print("\n", "-" * 54, "Network printers", "-" * 53)
                                flag = True
                            print(f"{printer} was found on user {users_dict[sid]}")
                            refresh()
                            found_any = True
                        except:
                            log()

            except FileNotFoundError:
                pass
            except:
                log()
    flag = False
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:

        # TCP/IP printers with translation
        with OpenKey(reg, r'SYSTEM\CurrentControlSet\Control\Print\Printers') as printers:
            found = []
            printers_len = QueryInfoKey(printers)[0]
            for i in range(printers_len):
                refresh()
                with OpenKey(printers, EnumKey(printers, i)) as printer:
                    try:
                        prnt = QueryValueEx(printer, "Port")[0]
                        if "_" in prnt:
                            prnt = prnt.split("_")[0].strip()
                        if prnt in found:
                            continue
                        found.append(prnt)
                        int(prnt.split(".")[0])
                        if not flag:
                            print("\n", "-" * 54, " TCP/IP printers ", "-" * 53)
                            flag = True
                        print(f"TCP/IP Printer with an IP of {prnt} is located at {config.ip_printers[prnt.strip()]}" if
                              prnt in config.ip_printers else
                              f"Printer with an IP of {prnt} is not on any of the servers")
                        refresh()
                        found_any = True
                    except (FileNotFoundError, ValueError):
                        pass
                    except:
                        log()
        flag = False
        with OpenKey(reg, r"SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM") as printers:
            found = []
            printers_len = QueryInfoKey(printers)[0]
            for i in range(printers_len):
                refresh()
                with OpenKey(printers, EnumKey(printers, i)) as printer:
                    try:
                        prnt = QueryValueEx(printer, "LocationInformation")[0].split("/")[2].split(":")[0]
                        if prnt in found:
                            continue
                        if "_" in prnt:
                            prnt = prnt.split("_")[0]
                        found.append(prnt)
                        if not flag:
                            print("\n", "-" * 55, " WSD printers ", "-" * 55)
                            flag = True
                        print(
                            f"WSD printer with an IP of {prnt.strip()} is located at "
                            f"{config.ip_printers[prnt.strip()]}" if prnt.strip() in config.ip_printers else
                            f"WSD printer with an IP of {prnt} is not on any of the servers")
                        refresh()
                        found_any = True
                    except (FileNotFoundError, ValueError, IndexError):
                        pass
                    except:
                        log()
    if not found_any:
        print_error(console, output=f"No printers were found", newline=True)


def run_func(func: callable) -> None:
    disable(disable_submit=True)
    window.after("idle", lambda: run_it(func))


def run_it(func: callable) -> None:
    refresh()
    if not reg_connect():
        on_submit()
        return
    refresh()
    if not wmi_connectable():
        on_submit()
        return
    refresh()
    try:
        func()
        enable()
    except:
        log()
        on_submit(pc=config.current_computer)


def update_user(user_: str) -> None:
    try:
        user_s = query_user(user_)
        if user_s == 0:
            print_success(user_active, additional="User status: ", output="Active", clear_=True)
        elif user_s == 1:
            update_error(user_active, "User status: ", "Disabled")
        elif user_s == 2:
            update_error(user_active, "User status: ", "Locked")
        elif user_s == 3:
            update_error(user_active, "User status: ", "Expired")
        elif user_s == 4:
            update_error(user_active, "User status: ", "Password expired")
    except:
        update_error(user_active, "User status: ", "ERROR")


# noinspection PyUnboundLocalVariable
def on_submit(pc: bool = None, passed_user: str = None) -> None:
    refresh()
    clear_all()
    if not pc:
        pc = computer.get().strip()
    if not pc:
        disable()
        computer.bind("<Return>", lambda _: on_submit())
        return
    disable(disable_submit=True)
    refresh()
    if pc_in_domain(pc):
        copy_but.configure(state="normal")
        copy_clip(pc)
        config.current_computer = pc
        update(display_pc, f"Current computer: {pc}")
        refresh()
        if check_pc_active(pc):
            if not wmi_connectable():
                print_error(console, output="Could not connect to computer's WMI", newline=True)
                submit.configure(state="normal")
                computer.bind("<Return>", lambda _: on_submit())
                return
            if not reg_connect():
                print_error(console, output="Could not connect to computer's registry", newline=True)
                submit.configure(state="normal")
                computer.bind("<Return>", lambda _: on_submit())
                return
            print_success(computer_status, "ONLINE", "Computer status: ", clear_=True)
            refresh()
            user_ = get_username(pc)
            if user_:
                config.current_user = user_
                if not passed_user or passed_user.lower() == user_.lower():
                    update(display_user, f"Current user: {user_name_trasnslation(user_)}")
                else:
                    update_error(display_user, "Current user: ", user_)
            else:
                update_error(display_user, "Current user: ", "No user")
            refresh()
            try:
                r_pc = WMI(pc)
                for k in r_pc.Win32_OperatingSystem():
                    last_boot_time = datetime.strptime(k.LastBootUpTime.split('.')[0], '%Y%m%d%H%M%S')
                    current_time = datetime.strptime(k.LocalDateTime.split('.')[0], '%Y%m%d%H%M%S')
                    uptime_ = current_time - last_boot_time
                    if uptime_ > timedelta(days=7):
                        update_error(uptime, "Uptime: ", uptime_)
                    else:
                        update(uptime, f"Uptime: {uptime_}")
                    break
            except:
                update_error(uptime, "Uptime: ", "ERROR")
                log()
            refresh()

            try:
                space = get_space(pc)
                if space <= 5:
                    update_error(space_c, "Space in C disk: ", f"{space:.1f}GB free out of "
                                                               f"{get_total_space(pc):.1f}GB")
                else:
                    update(space_c, f"Space in C disk: {space:.1f}GB free out of {get_total_space(pc):.1f}GB")
            except:
                log()
                update_error(space_c, "Space in C disk: ", "ERROR")
            refresh()

            if path.exists(fr"\\{pc}\d$"):
                try:
                    space = get_space(pc, disk="d")
                    if space <= 5:
                        update_error(space_d, "Space in D disk: ", f"{space:.1f}GB free out of "
                                                                   f"{get_total_space(pc, disk='d'):.1f}GB")
                    else:
                        update(space_d, f"Space in D disk: {space:.1f}GB free out of "
                                        f"{get_total_space(pc, disk='d'):.1f}GB")
                except:
                    log()
                    update_error(space_d, "Space in D disk: ", "ERROR")
            else:
                update_error(space_d, "Space in D disk: ", "Does not exist")
            refresh()

            try:
                try:
                    r_pc
                except NameError:
                    r_pc = WMI(pc)
                for ram_ in r_pc.Win32_ComputerSystem():
                    total_ram = int(ram_.TotalPhysicalMemory) / (1024 ** 3)
                    if total_ram < 7:
                        update_error(ram, "Total RAM: ", f"{round(total_ram)}GB")
                    else:
                        update(ram, f"Total RAM: {round(total_ram)}GB")
            except:
                update_error(ram, "Total RAM: ", "ERROR")
                log()
            refresh()

            if is_ie_fixed(pc):
                update(ie_fixed, "Internet explorer: Fixed")
            else:
                update_error(ie_fixed, "Internet explorer: ", "Not fixed")
            refresh()

            if is_cpt_fixed(pc):
                update(cpt_fixed, "Cockpit printer: Fixed")
            else:
                update_error(cpt_fixed, "Cockpit printer", "Not fixed")
            refresh()

            if user_ or passed_user:
                if passed_user:
                    user_ = passed_user
                update_user(user_)
            else:
                update_error(user_active, "User status: ", "No user")
            refresh()
            enable()
        else:
            print_error(computer_status, "OFFLINE", "Computer status: ", clear_=True)
            submit.configure(state="normal")
            computer.bind("<Return>", lambda _: on_submit())
            if passed_user:
                update_user(passed_user)
    else:
        refresh()
        try:
            with open(f"{config.users_txt}\\{pc}.txt") as pc_file:
                refresh()
                user_ = pc
                pc = pc_file.read().strip()
                on_submit(pc=pc, passed_user=user_)
        except FileNotFoundError:
            refresh()
            if user_exists(pc):
                refresh()
                print_error(console, f"Could not locate the current or last computer {pc} has logged on to")
                update_user(pc)
            else:
                refresh()
                if any([pc.lower() in config.ip_printers, pc.lower() in config.svr_printers]):
                    pc = pc.lower()
                    if pc in config.ip_printers:
                        print(f"Printer with an IP of {pc} is at {config.ip_printers[pc]}")
                        pc = config.ip_printers[pc]
                    elif pc in config.svr_printers:
                        print(f"Printer {pc} has an ip of {config.svr_printers[pc]}")
                        pc = config.svr_printers[pc]
                    copy_clip(pc)
                else:
                    if r"\\" in pc:
                        print_error(console, f"Could not locate printer {pc}")
                    elif pc.count(".") > 2:
                        print_error(console, f"Could not locate TCP/IP printer with ip of {pc}")
                    else:
                        print_error(console, f"No such user or computer in the domain {pc}")
            submit.configure(state="normal", cursor="hand2")
            computer.bind("<Return>", lambda _: on_submit())
            return


sys.stdout.write = redirect


class SetConfig:
    def __init__(self, json_config_file: dict) -> None:
        self.config_file = json_config_file
        self.ip_printers = {}
        self.svr_printers = {}
        for svr in self.config_file["print_servers"]:
            try:
                ip_svr = {prnt_name[2].strip().lower(): fr"{svr.lower()}\{prnt_name[0].strip()}" for prnt_name in
                          NetShareEnum(svr)}
                prnt_svr = {fr"{svr.lower()}\{prnt_name[0].strip().lower()}": prnt_name[2].strip() for prnt_name in
                            NetShareEnum(svr)}
                self.ip_printers.update(ip_svr)
                self.svr_printers.update(prnt_svr)
            except:
                pass
        self.log = self.config_file["log"]
        self.host = environ["COMPUTERNAME"]
        self.user = getuser()
        self.domain = self.config_file["domain"]
        self.delete_edb = self.config_file["delete_edb"]
        self.delete_user_temp = self.config_file["delete_user_temp"]
        self.c_paths_with_msg = [path_with_msg for path_with_msg in self.config_file["to_delete"] if
                                 len(path_with_msg) > 1]
        self.c_paths_without_msg = [path_without_msg for path_without_msg in self.config_file["to_delete"] if
                                    len(path_without_msg) == 1]
        self.u_paths_with_msg = [path_with_msg for path_with_msg in self.config_file["user_specific_delete"] if
                                 len(path_with_msg) > 1]
        self.u_paths_without_msg = [path_without_msg for path_without_msg in self.config_file["user_specific_delete"] if
                                    len(path_without_msg) == 1]
        self.users_txt = self.config_file["users txt"]
        self.current_computer = None
        self.current_user = None
        self.exclude = self.config_file["do_not_delete"]
        self.startwith_exclude = self.config_file["start_with_exclude"]
        self.wll_delete = []
        self.yes_no = False
        self.users_equal = False
        self.wmi_connectable = False
        self.reg_connectable = False
        self.assets = self.config_file["assets"].replace("/", "\\")


try:
    with open("\\".join(__file__.split("\\")[:-1]) + "\\config.json", encoding="utf8") as config_file:
        config = SetConfig(load(config_file))
except FileNotFoundError:
    messagebox.showerror("config file error", "could not find the config file")
    sys.exit(0)

if config.log and not path.isfile(config.log):
    try:
        with open(config.log, "w") as _:
            pass
    except:
        config.log = False

if config.log:
    basicConfig(filename=config.log, filemode="a", format="%(message)s")
    logger = getLogger("logfile")


def log() -> None:
    if not config.log:
        return
    err_log = f"""{'_' * 145}\nat {datetime.now().strftime('%Y-%m-%d %H:%M')} an error occurred on {config.host}\
 - {config.user}\n"""
    exception(err_log)


if not config.log:
    basicConfig(filename="FATAL_errors.log", filemode="w", format="%(message)s")
    logger = getLogger("fatal exceptions")


def log_unexpected(exc_ption, exc_value, __) -> None:
    if exc_ption != KeyboardInterrupt:
        exception(f"""{'_' * 145}\nat {datetime.now().strftime('%Y-%m-%d %H:%M')} a FATAL error occurred on 
{config.host} - {config.user}\n""", exc_info=exc_value)
    sys.exit(1)


class TimeoutException(Exception):
    pass


def Timeout(timeout: int) -> bool:
    def deco(func: callable):
        @wraps(func)
        def wrapper(*args, **kwargs):
            res = [TimeoutException('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, timeout))]

            def newFunc():
                try:
                    res[0] = func(*args, **kwargs)
                except Exception as e:
                    res[0] = e

            t = Thread(target=newFunc, daemon=True)
            try:
                t.start()
                start = time()
                while time() - start < 4.5 and t.is_alive():
                    refresh()
                    sleep(0.1)
                t.join(timeout)
            except TimeoutError:
                raise TimeoutException('Timeout occurred!')
            except Exception as je:
                log()
                raise je
            ret = res[0]
            if isinstance(ret, BaseException):
                raise ret
            return ret

        return wrapper

    return deco


# noinspection PyCallingNonCallable
def wmi_connectable() -> bool:
    x = Timeout(timeout=1.5)(WMI_connectable_wrap)
    try:
        y = x()
    except TimeoutException:
        config.wmi_connectable = False
        return False
    except:
        log()
        config.wmi_connectable = False
        return False
    if y:
        config.reg_connectable = True
    else:
        config.reg_connectable = False
    return y


def WMI_connectable_wrap() -> bool:
    pc = config.current_computer
    try:
        pythoncom.CoInitialize()
        WMI(computer=pc)
        return True
    except:
        log()
        return False


def get_space(pc: str, disk="c") -> float:
    return disk_usage(fr"\\{pc}\{disk}$").free / (1024 ** 3)


def get_total_space(pc: str, disk="c") -> float:
    return disk_usage(fr"\\{pc}\{disk}$").total / (1024 ** 3)


def pc_in_domain(pc: str) -> str:
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["name"],
        where_clause=f"name = '{pc}'",
        base_dn=config.domain
    )
    result = ad.get_results()
    is_pc = None
    for p in result:
        is_pc = p["name"]
    return is_pc


def user_exists(username_: str) -> bool:
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["sAMAccountName"],
        where_clause=f"sAMAccountName='{username_}'",
        base_dn=config.domain
    )
    result = ad.get_results()
    for _ in result:
        return True
    return False


def query_user(username_: str) -> str:
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["sAMAccountName", "userAccountControl"],
        where_clause=f"sAMAccountName='{username_}'",
        base_dn=config.domain
    )
    result = list(ad.get_results())

    if result is not None:
        user_account_control = result[0]["userAccountControl"]
        if user_account_control & 2 == 2:
            return 1

        refresh()
        dom_info = run(f'net user {username_} /domain | findstr /i "Account expires"', shell=True,
                       capture_output=True, text=True).stdout.strip().split()
        exp = " ".join(dom_info)
        if "Account expires Never" not in exp:
            expires = exp.split("Account expires")[1].strip()
            expires = f"{expires[:10]} {expires[11:19]}"
            if date_is_older(expires):
                return 3

        refresh()

        if "Locked" in run(f'net user {username_} /domain | findstr /i "Account active"', shell=True,
                           capture_output=True, text=True).stdout.strip():
            return 2

        refresh()

        pass_expired = exp.split("Password expires")[1].strip()
        pass_expired = f"{pass_expired[:10]} {pass_expired[11:19]}"
        if date_is_older(pass_expired):
            return 4

    return 0


# check if user is currently on computer
def check_pc_active_wrap(pc: str) -> bool:
    return path.exists(fr"\\{pc}\c$")


def check_pc_active(pc=None):
    # noinspection PyCallingNonCallable
    x = Timeout(timeout=3)(check_pc_active_wrap)
    try:
        y = x(pc=pc)
    except TimeoutException:
        return False
    except:
        log()
        return False

    return y


# get username from remote PC
def get_username(pc: str) -> str:
    try:
        con = WMI(computer=pc)
        rec = con.query("SELECT * FROM Win32_ComputerSystem")
        for user_ in rec:
            try:
                user_ = user_.UserName.split("\\")[1]
                return user_
            except AttributeError:
                pass
            except:
                log()
        try:
            processes = con.query("SELECT * FROM Win32_Process WHERE Name='explorer.exe'")
            for process in processes:
                _, _, user_ = process.GetOwner()
                return user_
        except:
            log()
    except:
        log()


# AD translation username to displayname
def user_name_trasnslation(username_: str) -> str:
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["displayName"],
        where_clause=f"sAMAccountName='{username_}'",
        base_dn=config.domain
    )
    result = ad.get_results()
    for u in result:
        return u["displayName"]
    return username_


def date_is_older(date_string):
    provided_date = datetime.strptime(date_string, "%d/%m/%Y %H:%M:%S")
    return provided_date < datetime.now()


# Add member to AD group
def add_member(username_: str, groupname: str) -> None:
    try:
        group_name = adgroup.ADGroup.from_cn(groupname)
        user_cn = None
        ad = adquery.ADQuery()
        ad.execute_query(
            attributes=["cn"],
            where_clause=f"sAMAccountName='{username_}'",
            base_dn=config.domain
        )

        result = ad.get_results()
        for u in result:
            user_cn = u["cn"]
        user_cn = aduser.ADUser.from_cn(user_cn)
        if user_cn:
            group_name.add_members([user_cn])
    except:
        log()


def on_rm_error(_: callable, path_: path, error: tuple) -> None:
    try:
        if error[0] != PermissionError:
            return
        chmod(path_, S_IWRITE)
        if path.isfile(path_):
            unlink(path_)
        elif path.isdir(path_):
            rmtree(path_, ignore_errors=True)
    except PermissionError:
        pass
    except:
        log()


# noinspection PyCallingNonCallable
def reg_connect() -> bool:
    x = Timeout(timeout=1.5)(is_reg)
    try:
        y = x()
    except TimeoutException:
        config.reg_connectable = False
        return False
    except:
        config.reg_connectable = False
        log()
        return False
    if y:
        config.reg_connectable = True
    else:
        config.reg_connectable = False
    return y


def is_reg(pc: str = None) -> bool:
    if not pc:
        pc = config.current_computer
    try:
        with ConnectRegistry(pc, HKEY_USERS) as _:
            return True
    except FileNotFoundError:
        pass
    except:
        log()
    return False


def is_ie_fixed(pc: str) -> bool:
    try:
        with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg_:
            try:
                with OpenKey(reg_, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper "
                                   r"Objects\{1FD49718-1D00-4B19-AF5F-070AF6D5D54C}", 0, KEY_ALL_ACCESS) as _:
                    return False
            except FileNotFoundError:
                pass
            except:
                log()
            return True
    except FileNotFoundError:
        pass
    except:
        log()
    return False


def is_cpt_fixed(pc: str) -> bool:
    try:
        with ConnectRegistry(pc, HKEY_CURRENT_USER) as reg_:
            try:
                with OpenKey(reg_, r"SOFTWARE\Jetro Platforms\JDsClient\PrintPlugIn", 0, KEY_ALL_ACCESS) as key_:
                    QueryValueEx(key_, "PrintClientPath")
                    return False
            except FileNotFoundError:
                return True
            except:
                log()
    except FileNotFoundError:
        pass
    except:
        log()
    return False


config.first_time = 1

window = Tk()

window.geometry("758x758")
window.configure(bg="#545664")
window.title("your title here")
window.iconbitmap(asset("icon.ico"))

canvas = Canvas(
    window,
    bg="#545664",
    height=758,
    width=758,
    bd=0,
    highlightthickness=0,
    relief="ridge"
)

canvas.place(x=0, y=0)

generic_image = PhotoImage(file=asset("generic_text.png"))

user_active_bg = canvas.create_image(
    236.0,
    422.5,
    image=generic_image
)
user_active = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
user_active.bind("<<Selection>>", lambda event_: ignore_selection(user_active, event_))

user_active.place(
    x=31.0,
    y=412.6,
    width=410.0,
    height=21.2
)

cpt_fixed_bg = canvas.create_image(
    236.0,
    392.0,
    image=generic_image
)
cpt_fixed = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
cpt_fixed.place(
    x=31.0,
    y=382.2,
    width=410.0,
    height=21.2
)
cpt_fixed.bind("<<Selection>>", lambda event_: ignore_selection(cpt_fixed, event_))

ie_fixed_bg = canvas.create_image(
    236.0,
    361.5,
    image=generic_image
)
ie_fixed = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
ie_fixed.bind("<<Selection>>", lambda event_: ignore_selection(ie_fixed, event_))

ie_fixed.place(
    x=31.0,
    y=351.7,
    width=410.0,
    height=21.2
)

ram_bg = canvas.create_image(
    236.0,
    330.5,
    image=generic_image
)
ram = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
ram.bind("<<Selection>>", lambda event_: ignore_selection(ram, event_))

ram.place(
    x=31.0,
    y=320.7,
    width=410.0,
    height=21.4
)

space_d_bg = canvas.create_image(
    236.0,
    299.5,
    image=generic_image
)
space_d = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
space_d.bind("<<Selection>>", lambda event_: ignore_selection(space_d, event_))

space_d.place(
    x=31.0,
    y=290.0,
    width=410.0,
    height=21.2
)

space_c_bg = canvas.create_image(
    236.0,
    268.5,
    image=generic_image
)
space_c = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
space_c.bind("<<Selection>>", lambda event_: ignore_selection(space_c, event_))

space_c.place(
    x=31.0,
    y=259.0,
    width=410.0,
    height=21.2
)

uptime_bg = canvas.create_image(
    236.0,
    237.5,
    image=generic_image
)
uptime = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
uptime.bind("<<Selection>>", lambda event_: ignore_selection(uptime, event_))

uptime.place(
    x=31.0,
    y=228.0,
    width=410.0,
    height=21.2
)

display_user_bg = canvas.create_image(
    236.0,
    206.5,
    image=generic_image
)
display_user = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
display_user.bind("<<Selection>>", lambda event_: ignore_selection(display_user, event_))

display_user.place(
    x=31.0,
    y=197.0,
    width=410.0,
    height=21.2
)

entry_bg_7 = canvas.create_image(
    236.0,
    175.5,
    image=generic_image
)
computer_status = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
computer_status.bind("<<Selection>>", lambda event_: ignore_selection(computer_status, event_))

computer_status.place(
    x=31.0,
    y=166.0,
    width=410.0,
    height=21.2
)

rest_spool_image = PhotoImage(
    file=asset("button_1.png"))
reset_spool = Button(
    image=rest_spool_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(reset_spooler),
    relief="flat",
    background="#545664"
)
reset_spool.place(
    x=599.0,
    y=131.0,
    width=128.0,
    height=55.0
)

delete_ost_image = PhotoImage(
    file=asset("button_2.png"))
delete_ost = Button(
    image=delete_ost_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(delete_the_ost),
    relief="flat",
    background="#545664"
)
delete_ost.place(
    x=461.0,
    y=353.0,
    width=128.0,
    height=55.0
)

delete_users_image = PhotoImage(
    file=asset("button_3.png"))
delete_users = Button(
    image=delete_users_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(del_users),
    relief="flat",
    background="#545664"
)
delete_users.place(
    x=461.0,
    y=279.0,
    width=128.0,
    height=55.0
)

get_printers_image = PhotoImage(
    file=asset("button_4.png"))
get_printers = Button(
    image=get_printers_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(get_printers_func),
    relief="flat",
    background="#545664"
)
get_printers.place(
    x=461.0,
    y=205.0,
    width=128.0,
    height=55.0
)

fix_3_lang_image = PhotoImage(
    file=asset("button_5.png"))
fix_3_lang = Button(
    image=fix_3_lang_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(fix_3_languages),
    relief="flat",
    background="#545664"
)
fix_3_lang.place(
    x=461.0,
    y=131.0,
    width=128.0,
    height=55.0
)

fix_cpt_image = PhotoImage(
    file=asset("button_6.png"))
fix_cpt = Button(
    image=fix_cpt_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(fix_cpt_func),
    relief="flat",
    background="#545664"
)
fix_cpt.place(
    x=599.0,
    y=353.0,
    width=128.0,
    height=55.0
)

fix_ie_image = PhotoImage(
    file=asset("button_7.png"))
fix_ie = Button(
    image=fix_ie_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(fix_ie_func),
    relief="flat",
    background="#545664"
)
fix_ie.place(
    x=599.0,
    y=280.0,
    width=128.0,
    height=55.0
)

clear_space_image = PhotoImage(
    file=asset("button_8.png"))
clear_space = Button(
    image=clear_space_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: run_func(clear_space_func),
    relief="flat",
    background="#545664"
)
clear_space.place(
    x=599.0,
    y=205.0,
    width=128.0,
    height=55.0
)

computer_image = PhotoImage(
    file=asset("entry_8.png"))
computer_background = canvas.create_image(
    377.5,
    27.0,
    image=computer_image
)
computer = Entry(
    font=("Times", 12, "bold"),
    bd=0,
    bg="#C2EFF9",
    fg="#000716",
    justify="center",
    highlightthickness=0
)
computer.bind("<FocusIn>", show_text)
computer.bind("<FocusOut>", hide_text)
computer.bind_all("<Key>", enable_paste, "+")
computer.insert(INSERT, "Computer or User")
computer.bind("<Return>", lambda _: on_submit())
computer.place(
    x=216.0,
    y=6.0,
    width=323.0,
    height=40.0
)

submit_image = PhotoImage(
    file=asset("button_9.png"))
submit = Button(
    image=submit_image,
    activebackground="#545664",
    background="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=on_submit,
    relief="flat"
)
submit.place(
    x=306.0,
    y=55.0,
    width=146.0,
    height=45.0
)

display_pc_image = PhotoImage(
    file=asset("display_pc.png"))

display_pc_bg = canvas.create_image(
    195.0,
    142.5,
    image=display_pc_image
)
display_pc = Text(
    bd=0,
    bg="#D9D9D9",
    fg="#000716",
    highlightthickness=0,
    font=('Arial', 12, 'bold'),
    cursor="arrow"
)
display_pc.bind("<<Selection>>", lambda event_: ignore_selection(display_pc, event_))

display_pc.place(
    x=31.0,
    y=133.0,
    width=331.0,
    height=21.4
)

copy_but_image = PhotoImage(
    file=asset("copy.png"))

copy_but = Button(
    image=copy_but_image,
    activebackground="#545664",
    borderwidth=0,
    highlightthickness=0,
    command=lambda: copy_clip(config.current_computer[1:-4]),
    relief="flat",
    background="#545664"
)

copy_but.place(
    x=376.0,
    y=130.0,
    width=69.5,
    height=32.5
)

console_image = PhotoImage(
    file=asset("entry_10.png"))
entry_bg_10 = canvas.create_image(
    379.0,
    605.0,
    image=console_image
)
console = Text(
    bd=0,
    bg="#FFFFFF",
    fg="#000716",
    font=('Arial', 12, 'bold'),
    highlightthickness=0,
    state="disabled"
)

console.place(
    x=30.0,
    y=471.5,
    width=680.0,
    height=269.0
)

mouse_listener = mouse.Listener(on_click=on_button_press)
mouse_listener.start()
window.resizable(False, False)

clear_all()
disable()
window.mainloop()
