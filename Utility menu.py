from __future__ import annotations
import sys
import pythoncom
from os import path, unlink, listdir, mkdir, rename, chmod, environ
from stat import S_IWRITE
from wmi import WMI
from winreg import HKEY_CURRENT_USER, HKEY_USERS, KEY_ALL_ACCESS, DeleteKey, DeleteValue, QueryValueEx, REG_DWORD
from winreg import OpenKey, QueryInfoKey, EnumKey, ConnectRegistry, HKEY_LOCAL_MACHINE, KEY_SET_VALUE, SetValueEx
from shutil import rmtree
from subprocess import run
from time import sleep
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
from pyad import adquery
import tkinter
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, INSERT, messagebox, END, ttk, CENTER, SEL, Event


def redirect(*output: str) -> None:
    """redirects all output to the console Text object"""
    gui.console.configure(state="normal")
    gui.console.insert(END, str(*output).replace(".!text11", ""))
    gui.console.see(END)
    gui.console.configure(state="disabled")


def print_error(obj: Text, output: str = "", additional: str = "", clear_: bool = False, see: bool = False,
                newline: bool = False) -> None:
    """prints an error to an object, will be red colored"""
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
    """print green text to an object"""
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
    """updates an object with new text"""
    obj.configure(state="normal")
    obj.delete("1.0", END)
    obj.insert(INSERT, statement)
    obj.configure(state="disabled")


def update_error(obj: Text, initial: str, statement: str | timedelta) -> None:
    """"updates red text into an object"""
    obj.configure(state="normal")
    obj.tag_configure('red', foreground='red')
    obj.delete("1.0", END)
    obj.insert(INSERT, initial)
    obj.insert(INSERT, statement, "red")
    obj.configure(state="disabled")


def clear_obj(obj: Text) -> None:
    """"clears an object from text"""
    obj.configure(state="normal")
    obj.delete("1.0", END)
    obj.configure(state="disabled")


def clear_all() -> None:
    """"clears all objects in the gui window, and sets their default text"""
    for obj in [[gui.display_pc, "Current computer: "], [gui.computer_status, "Computer status: "], [gui.console, ""],
                [gui.display_user, "Current user: "], [gui.uptime, "Uptime: "], [gui.space_c, "Space in C disk: "],
                [gui.space_d, "Space in D disk: "], [gui.ram, "Total RAM: "], [gui.ie_fixed, "Internet explorer: "],
                [gui.cpt_fixed, "Cockpit printer: "], [gui.user_active, "User status: "]]:
        obj[0].configure(state="normal")
        obj[0].delete("1.0", END)
        obj[0].insert(INSERT, obj[1])
        obj[0].configure(state="disabled")
        config.current_computer = None
        config.current_user = None


def disable(disable_submit: bool = False) -> None:
    """"disables all the buttons, so they aren't clickable while a function is still executing, also disables submitting
    by pressing the enter key"""
    gui.computer.unbind("<Return>")
    for obj in (gui.reset_spool, gui.fix_cpt, gui.fix_ie, gui.clear_space, gui.get_printers, gui.delete_ost,
                gui.delete_users, gui.fix_3_lang, gui.copy_but):
        obj.configure(state="disabled", cursor="arrow")
    if disable_submit:
        gui.submit.configure(state="disabled", cursor="arrow")
    else:
        gui.submit.configure(state="normal", cursor="hand2")
        gui.computer.bind("<Return>", lambda _: run_func(on_submit))
    if config.current_computer:
        gui.copy_but.configure(state="normal", cursor="hand2")
    if config.first_time:
        config.first_time = 0
        gui.computer.bind("<Return>", lambda _: run_func(on_submit))
        gui.submit.configure(cursor="hand2")


def enable() -> None:
    """"enables the buttons back, also makes submitted by pressing enter enabled again"""
    for obj in (gui.reset_spool, gui.fix_cpt, gui.fix_ie, gui.clear_space, gui.get_printers, gui.delete_ost,
                gui.delete_users, gui.fix_3_lang, gui.submit, gui.copy_but):
        obj.configure(state="normal", cursor="hand2")
    gui.computer.bind("<Return>", lambda _: run_func(on_submit))
    if not config.current_user:
        gui.delete_ost.configure(state="disabled", cursor="arrow")
        gui.fix_cpt.configure(state="disabled", cursor="arrow")


def show_text(_: Event) -> None:
    """"puts the default text in the computer entry box if its empty and isn't focused"""
    if gui.computer.get() == "Computer or User":
        gui.computer.delete(0, END)
        gui.computer.config(justify="center")


def hide_text(_: Event) -> None:
    """hides the default text once the user starts interacting with the computer entry box"""
    if gui.computer.get() == "":
        gui.computer.insert(0, "Computer or User")
        gui.computer.config(justify="center")


def enable_paste(event: Event) -> None:
    """"enable pasting to the computer entry while the keyboard language isn't in english"""
    ctrl = (event.state & 0x4) != 0

    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    elif event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")

    elif event.keycode == 65 and ctrl and event.keysym.lower() != "a":
        event.widget.event_generate("<<SelectAll>>")


def copy_clip(to_copy: str) -> None:
    """"copies the computer name to the user's clipboard"""
    cp = Tk()
    cp.withdraw()
    cp.clipboard_clear()
    cp.clipboard_append(to_copy)
    cp.update()
    cp.destroy()


def on_button_press(_: float, __: float, button: Button, ___: bool):
    """bring the app to the front when the middle mouse button is pressed"""
    if button == mouse.Button.middle:
        if gui.root.wm_state() == 'iconic':
            gui.root.deiconify()
            gui.root.lift()
            gui.root.focus_set()
        else:
            gui.root.iconify()


def disable_middle_click(event: Event):
    """"disables middle mouse presses"""
    if event.num == 2:
        return "break"


def ignore_selection(obj: Text, _: Event):
    """prevents the user from being able to select and mark text on the Text boxes which display info"""
    obj.tag_remove(SEL, "1.0", END)
    return "break"


def asset(filename: str) -> str:
    """return the full path to images - assets that the script uses"""
    return fr"{config.assets}\{filename}"


def create_selection_window(options: list) -> None:
    """creates a selection window with checkboxes to choose which users to delete"""
    config.yes_no = False
    config.wll_delete = []

    def on_done() -> None:
        """cleans up after the main function and give last warning before deleting the users folders"""
        newline = "\n"
        selected_options = [check.get() for check in option_vars if check.get()]
        canvas_.unbind_all("<MouseWheel>")
        selection_window.destroy()
        config.tasks.append(lambda: gui.root.focus_set())
        if not selected_options:
            config.tasks.append(lambda: print("No users were chosen to be deleted"))
            config.yes_no = False
            return
        yes_no = messagebox.askyesno("Warning", f"Are you sure you want to delete the following users?\n"
                                                f"{newline.join(selected_options).replace('{', '').replace('}', ' -')}")
        if yes_no:
            config.wll_delete = [selected.split()[-1] for selected in selected_options if selected]
            config.yes_no = True
            return
        config.tasks.append(lambda: print("Canceled users deletion"))
        config.yes_no = False
        return

    def disable_main_window(selection_window_: tkinter.Toplevel) -> None:
        """disables the main window and brings the selection box to the front"""

        def on_window_close():
            """"deletes the checkbox window and unbind middle mouse wheel from scrolling in the checkbox window"""
            config.tasks.append(lambda: gui.root.grab_release())
            selection_window_.destroy()
            canvas_.unbind_all("<MouseWheel>")
            config.tasks.append(lambda: print("Canceled users deletion"))

        gui.root.grab_set()
        selection_window.protocol("WM_DELETE_WINDOW", on_window_close)
        gui.root.wait_window(selection_window)

    def on_mousewheel(event: Event) -> None:
        """"enables scrolling in the checkbox window via middle mouse wheel, in case there are over 10 users"""
        canvas_.yview_scroll(int(-1 * (event.delta / 120)), "units")

    selection_window = tkinter.Toplevel(gui.root)
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
    """"an easily deployable progressbar which can be called via a with statement"""

    def __init__(self, total_items: int, title_: str, end_statement: str = ""):
        """"initial configuration of the progressbar"""
        self.total_items = total_items
        self.title = title_
        self.current_item = 0

        self.root = gui.root
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
        self.progressbar = ttk.Progressbar(self.root, length=297, mode="determinate",
                                           style='text.Horizontal.TProgressbar')
        self.progressbar.place(x=450.0, y=440.0)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """cleans up after the progress bar finishes"""
        config.tasks.append(lambda: print(self.end_statement))
        self.progressbar.destroy()
        self.label.destroy()

    def __call__(self):
        """updates the progressbar when its being called"""
        self.current_item += 1
        self.root.update()
        if self.current_item <= self.total_items:
            progress = (self.current_item / self.total_items) * 100
            ttk.Style().configure('text.Horizontal.TProgressbar', text=f"{int(progress)} %")
            self.progressbar["value"] = progress


def refresh() -> None:
    """refreshes the main window"""
    gui.root.update_idletasks()
    gui.root.update()


def fix_ie_func() -> None:
    """fixes the Internet Explorer application via deleting appropriate registry keys, as well as disabling
    compatibility mode"""
    pc = config.current_computer
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:
        for key_name in (
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C}",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C"):
            try:
                with OpenKey(reg, key_name, 0, KEY_ALL_ACCESS) as key:
                    DeleteKey(key, "")
            except FileNotFoundError:
                pass
            except:
                config.tasks.append(lambda: print_error(gui.console, output="Unable to fix internet explorer",
                                                        newline=True))
                log()
                return

    key = r"Software\Microsoft\Internet Explorer\BrowserEmulation"
    try:
        with ConnectRegistry(pc, HKEY_CURRENT_USER) as reg:
            with OpenKey(reg, key, 0, KEY_SET_VALUE) as key:
                SetValueEx(key, "IntranetCompatibilityMode", 0, REG_DWORD, 0)
                SetValueEx(key, "MSCompatibilityMode", 0, REG_DWORD, 0)
    except FileNotFoundError:
        pass
    except:
        log()

    config.tasks.append(lambda: update(gui.ie_fixed, "Internet explorer: Fixed"))
    config.tasks.append(lambda: print_success(gui.console, output=f"Fixed internet explorer", newline=True))


def fix_cpt_func() -> None:
    """"fixes cockpit printer via deleting the appropriate registry keys"""
    with ConnectRegistry(config.current_computer, HKEY_CURRENT_USER) as reg:
        try:
            with OpenKey(reg, r"SOFTWARE\Jetro Platforms\JDsClient\PrintPlugIn", 0, KEY_ALL_ACCESS) as key:
                DeleteValue(key, "PrintClientPath")
        except FileNotFoundError:
            pass
        except:
            log()
            config.tasks.append(lambda: print_error(gui.console, output="Failed to fix cpt printer", newline=True))
            return
    config.tasks.append(lambda: print_success(gui.console, output="Fixed cpt printer", newline=True))
    config.tasks.append(lambda: update(gui.cpt_fixed, "Cockpit printer: Fixed"))


def fix_3_languages() -> None:
    """fixes 3 languages bug via deleting the appropriate registry keys"""
    with ConnectRegistry(config.current_computer, HKEY_USERS) as reg:
        try:
            with OpenKey(reg, r".DEFAULT\Keyboard Layout\Preload", 0, KEY_ALL_ACCESS) as key:
                DeleteKey(key, "")
        except FileNotFoundError:
            pass
        except:
            config.tasks.append(print_error(lambda: gui.console, output="Failed to fix 3 languages bug", newline=True))
            log()
            return
    config.tasks.append(lambda: print_success(gui.console, output="Fixed 3 languages bug", newline=True))


def reset_spooler() -> None:
    """resets the print spooler via WMI"""
    try:
        # noinspection PyUnboundLocalVariable
        pythoncom.CoInitialize()
        connection = WMI(computer=config.current_computer)
        service = connection.Win32_Service(name="Spooler")
        service[0].StopService()
        sleep(1)
        service[0].StartService()
        config.tasks.append(lambda: print_success(gui.console, output=f"Successfully restarted the spooler",
                                                  newline=True))
    except:
        config.tasks.append(lambda: print_error(gui.console, output=f"Failed to restart the spooler", newline=True))
        log()


def delete_the_ost() -> None:
    """renames the ost file to .old with random digits in order to avoid conflict with duplicate ost filenames
    handles shutting down outlook and skype on the remote computer so the OST could be renamed"""
    user_ = config.current_user
    pc = config.current_computer
    pythoncom.CoInitialize()
    if not tkinter.messagebox.askyesno(title="OST deletion",
                                       message=f"Are you sure you want to delete "
                                               f"the ost of {user_name_translation(user_)}?"):
        config.tasks.append(lambda: print_error(gui.console, output="Canceled OST deletion", newline=True))
        return
    host = WMI(computer=pc)
    for procs in ("lync.exe", "outlook.exe", "UcMapi.exe"):
        for proc in host.Win32_Process(name=procs):
            if proc:
                try:
                    proc.Terminate()
                except:
                    log()
    if not path.exists(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook"):
        config.tasks.append(lambda: print_error(gui.console, f"Could not find an OST file", newline=True))
        return

    ost = listdir(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook")
    for file in ost:
        if file.endswith("ost"):
            ost = fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook\{file}"
            try:
                sleep(1)
                rename(ost, f"{ost}{random():.3f}.old")
                config.tasks.append(lambda: print_success(gui.console,
                                                          output=f"Successfully removed the ost file", newline=True))
                return
            except FileExistsError:
                try:
                    rename(ost, f"{ost}{random():.3f}.old")
                    config.tasks.append(lambda: print_success(gui.console,
                                                              output=f"Successfully removed the ost file",
                                                              newline=True))
                    return
                except:
                    log()
                    config.tasks.append(
                        lambda: print_error(gui.console, f"Could not Delete the OST file", newline=True))
                    return
            except:
                config.tasks.append(lambda: print_error(gui.console, f"Could not Delete the OST file", newline=True))
                log()
                return
    else:
        config.tasks.append(lambda: print_error(gui.console, f"Could not find an OST file", newline=True))


def my_rm(file_: str, bar_: callable) -> None:
    """removes readonly files via changing the file permissions"""
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
    """removes the entire dir with its contents and then recreate it"""
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
    """clears spaces from the remote computer, paths, and other configurations as for which files to delete
    can be configured via the config file. using multithreading to delete the files faster"""
    pc = config.current_computer
    users_dirs = listdir(fr"\\{pc}\c$\users")
    pythoncom.CoInitialize()

    space_init = get_space(pc)
    flag = False

    edb_file = fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
    if path.exists(edb_file) and config.delete_edb:
        try:
            connection = WMI(computer=pc)
            service = connection.Win32_Service(name="WSearch")
            service[0].StopService()
            sleep(0.6)
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except PermissionError:
            pass
        except:
            log()

    if config.c_paths_with_msg:
        for path_msg in config.c_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                files = [fr"\\{pc}\c$\{path_msg[0]}\{file}" for file in listdir(fr"\\{pc}\c$\{path_msg[0]}")]
                with ProgressBar(len(files), path_msg[1], path_msg[-1]) as bar:
                    with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                        jobs = [executor.submit(my_rm, file, bar) for file in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)

    if config.delete_user_temp:
        with ProgressBar(len(users_dirs), f"Deleting temps of {len(users_dirs)} users",
                         f"Deleted temps of {len(users_dirs)} users") as bar:
            dirs = [fr"\\{pc}\c$\users\{dir_}\AppData\Local\Temp" for dir_ in users_dirs if
                    (dir_.lower().strip() != config.user.lower().strip() and config.current_computer.lower()
                     != config.host.lower().strip())]
            with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                jobs = [executor.submit(my_rm, dir_, bar) for dir_ in dirs]
                while not all([result.done() for result in jobs]):
                    sleep(0.1)

    if config.u_paths_with_msg:
        for path_msg in config.u_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            msg_ = path_msg[1].replace("users_amount", str(len(users_dirs)))
            with ProgressBar(len(users_dirs), msg_, path_msg[-1].replace(str(len(users_dirs)))) as bar:
                for user in users_dirs:
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                            jobs = [executor.submit(my_rm, file) for file in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                    bar()

    if config.u_paths_without_msg or config.c_paths_without_msg:
        with ProgressBar(len(config.u_paths_without_msg) + len(config.c_paths_without_msg), "Deleting additional files",
                         "Deleted additional files") as bar:
            for path_msg in config.c_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                    files = listdir(fr"\\{pc}\c$\{path_msg[0]}")
                    with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                        jobs = [executor.submit(my_rm, file) for file in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)
                bar()

            for path_msg in config.u_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                for user in users_dirs:
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                            jobs = [executor.submit(my_rm, file) for file in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                bar()

    if not flag and config.delete_edb and path.exists(edb_file):
        try:
            connection = WMI(computer=pc)
            service = connection.Win32_Service(name="WSearch")
            service[0].StopService()
            sleep(0.8)
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except (PermissionError, FileNotFoundError):
            pass
        except:
            log()
    if flag and config.delete_edb:
        config.tasks.append(lambda: print(fr"Deleted the search.edb file"))
    else:
        if config.delete_edb:
            print_error(gui.console, output="Failed to remove search.edb file", newline=True)
    space_final = get_space(pc)
    config.tasks.append(lambda: print_success(gui.console,
                                              output=f"Cleared {abs((space_final - space_init)):.1f} GB from the disk",
                                              newline=True))
    try:
        space = get_space(pc)
        if space <= 5:
            config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ",
                                                     f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB"))
        else:
            config.tasks.append(lambda: update(gui.space_c,
                                               f"Space in C disk: {space:.1f}GB free out of "
                                               f"{get_total_space(pc):.1f}GB"))
    except:
        log()
        config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ", "ERROR"))


def my_rmtree(dir_: str, bar_: callable) -> None:
    """delete folders and their contents, then calls the bar object"""
    if path.isdir(dir_):
        rmtree(dir_, onerror=on_rm_error)
    bar_()


def del_users() -> None:
    """gives you the option to choose which users folders to delete as well as multithreading deletion of folders
    will exclude the current user of the remote pc if found one.
    users to exclude could be configured in the config file"""
    pythoncom.CoInitialize()
    config.wll_delete = []
    config.yes_no = False
    pc = config.current_computer
    users_to_choose_delete = []
    for dir_ in listdir(fr"\\{pc}\c$\Users"):
        if dir_.lower() == str(config.current_user).lower() or dir_.lower() in config.exclude or \
                any([dir_.lower().startswith(exc_lude) for exc_lude in config.startwith_exclude]) \
                or not path.isdir(fr"\\{pc}\c$\users\{dir_}"):
            continue
        users_to_choose_delete.append([user_name_translation(dir_), dir_])
    if not users_to_choose_delete:
        config.tasks.append(lambda: print("No users were found to delete"))
        return
    create_selection_window(users_to_choose_delete)
    if not config.yes_no:
        return
    space_init = get_space(pc)
    with ProgressBar(len(config.wll_delete), f"Deleting {len(config.wll_delete)} folders",
                     f"Deleted {len(config.wll_delete)} users") as bar:
        config.wll_delete = [fr"\\{pc}\c$\users\{dir_}" for dir_ in config.wll_delete]
        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
            jobs = [executor.submit(my_rmtree, dir_, bar) for dir_ in config.wll_delete]
            while not all([result.done() for result in jobs]):
                sleep(0.1)
    space_final = get_space(pc)
    config.tasks.append(lambda: print(f"Cleared {abs((space_final - space_init)):.1f} GB from the disk"))
    try:
        space = get_space(pc)
        if space <= 5:
            config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ",
                                                     f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB"))
        else:
            config.tasks.append(lambda: update(gui.space_c, f"Space in C disk: "
                                                            f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB"))
    except:
        log()
        config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ", "ERROR"))


def get_printers_func() -> None:
    """retrieves all network printers installed on the remote computer
     achieves that via querying the appropriate registry keys"""
    found_any = False
    pc = config.current_computer
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

        with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as users_path:
            pythoncom.CoInitialize()
            for sid in set(sid_list):
                try:
                    with OpenKey(users_path,
                                 fr"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{sid}") as profiles:
                        username = QueryValueEx(profiles, "ProfileImagePath")
                        if username[0].startswith("C:\\"):
                            username = username[0].split("\\")[-1]
                            users_dict[sid] = user_name_translation(username)
                except FileNotFoundError:
                    pass
                except:
                    log()

        flag = False
        for sid in sid_list:
            try:
                with OpenKey(reg, fr"{sid}\Printers\Connections") as printer_path:
                    printers_len = QueryInfoKey(printer_path)[0]
                    for i in range(printers_len):
                        try:
                            printer = EnumKey(printer_path, i).replace(",", "\\").strip()
                            p = f"{printer} was found on user {users_dict[sid]}"
                            if not flag:
                                config.tasks.append(lambda: print("\n", "-" * 54, "Network printers", "-" * 53))
                                flag = True
                            config.tasks.append(lambda: print(p))
                            found_any = True
                        except:
                            log()
            except FileNotFoundError:
                pass
            except:
                log()
    flag = False
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:
        with OpenKey(reg, r'SYSTEM\CurrentControlSet\Control\Print\Printers') as printers:
            found = []
            printers_len = QueryInfoKey(printers)[0]
            for i in range(printers_len):
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
                            config.tasks.append(lambda: print("\n", "-" * 54, " TCP/IP printers ", "-" * 53))
                            flag = True
                        config.tasks.append(lambda: print(
                            f"TCP/IP Printer with an IP of {prnt} is located at {config.ip_printers[prnt.strip()]}" if
                            prnt in config.ip_printers else
                            f"Printer with an IP of {prnt} is not on any of the servers"))
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
                with OpenKey(printers, EnumKey(printers, i)) as printer:
                    try:
                        prnt = QueryValueEx(printer, "LocationInformation")[0].split("/")[2].split(":")[0]
                        if prnt in found:
                            continue
                        if "_" in prnt:
                            prnt = prnt.split("_")[0]
                        found.append(prnt)
                        if not flag:
                            config.tasks.append(lambda: print("\n", "-" * 55, " WSD printers ", "-" * 55))
                            flag = True
                        config.tasks.append(lambda: print(
                            f"WSD printer with an IP of {prnt.strip()} is located at "
                            f"{config.ip_printers[prnt.strip()]}" if prnt.strip() in config.ip_printers else
                            f"WSD printer with an IP of {prnt} is not on any of the servers"))
                        found_any = True
                    except (FileNotFoundError, ValueError, IndexError):
                        pass
                    except:
                        log()
    if not found_any:
        config.tasks.append(lambda: print_error(gui.console, output=f"No printers were found", newline=True))


def run_it(func: callable, tries: int = 0) -> None:
    """passes the function to run_it after the main window is idle, and gives the button time to be unpressed
    as well as disabling the buttons"""
    try:
        pythoncom.CoInitialize()
        func()
    except:
        log()
        if not tries:
            config.tasks.append(
                lambda: gui.root.after(60, lambda: run_func(lambda: on_submit(pc=config.current_computer), tries=1))
            )
            config.tasks.append(
                lambda: gui.root.after(
                    100, lambda: print_error(gui.console, "An error occurred during the execution of the last function")
                )
            )


def run_func(func: callable, tries: int = 0) -> None:
    """runs the function itself, catch any exception and logs it, checks if the issue is a network issue via running
    on_submit when an exception is caught"""
    disable(disable_submit=True)
    config.disable = False
    refresh()
    if not func == on_submit:
        if not reg_connect():
            run_func(on_submit, tries)
            return
        refresh()
        if not wmi_connectable():
            run_func(on_submit, tries)
            return
    refresh()
    t = Thread(target=lambda: run_it(func, tries), daemon=True)
    t.start()
    while t.is_alive() or config.tasks:
        refresh()
        if config.tasks:
            config.tasks.pop(0)()
        sleep(0.03)
    disable() if config.disable else enable()


def update_user(user_: str) -> None:
    """updates the user status to the user_active Text box"""
    try:
        user_s = query_user(user_)
        if user_s == 0:
            print_success(gui.user_active, additional="User status: ", output="Active", clear_=True)
        elif user_s == 1:
            update_error(gui.user_active, "User status: ", "Disabled")
        elif user_s == 2:
            update_error(gui.user_active, "User status: ", "Locked")
        elif user_s == 3:
            update_error(gui.user_active, "User status: ", "Expired")
        elif user_s == 4:
            update_error(gui.user_active, "User status: ", "Password expired")
    except:
        update_error(gui.user_active, "User status: ", "ERROR")


# noinspection PyUnboundLocalVariable
def on_submit(pc: str = None, passed_user: str = None) -> None:
    """checks if the passed string is a computer in the domain
    if it is, it checks if its online, if it is online it then proceed to display information on the computer
    if the passed string is not a computer in the domain it looks for a file with the same name and txt extension
    in the preconfigured path via the config file, if it finds any it treats the contents of the file as the computer
    name and rerun on_submit with the computer as the arg
    if the string is neither a username nor a computer name it checks if it's a printer - TCP/IP or installed via
    print server"""
    clear_all()
    if not pc:
        pc = gui.computer.get().strip()
    if not pc:
        config.tasks.append(lambda: gui.computer.bind("<Return>", lambda _: run_func(on_submit)))
        config.disable = True
        return
    config.tasks.append(lambda: disable(disable_submit=True))
    pythoncom.CoInitialize()
    if pc_in_domain(pc):
        config.tasks.append(lambda: gui.copy_but.configure(state="normal", cursor="hand2"))
        config.tasks.append(lambda: copy_clip(pc))
        config.current_computer = pc
        config.tasks.append(lambda: update(gui.display_pc, f"Current computer: {pc}"))
        if not check_pc_active(pc):
            config.tasks.append(lambda: print_error(gui.computer_status, "OFFLINE", "Computer status: ", clear_=True))
            config.tasks.append(lambda: gui.submit.configure(state="normal"))
            config.tasks.append(lambda: gui.computer.bind("<Return>", lambda _: run_func(on_submit)))
            config.disable = True
            if passed_user:
                config.tasks.append(lambda: update_user(passed_user))
            return
        if not wmi_connectable():
            config.tasks.append(lambda: print_error(gui.console,
                                                    output="Could not connect to computer's WMI", newline=True))
            config.tasks.append(lambda: gui.submit.configure(state="normal"))
            config.tasks.append(lambda: gui.computer.bind("<Return>", lambda _: run_func(on_submit)))
            config.disable = True
            return
        if not reg_connect():
            config.tasks.append(lambda: print_error(gui.console, output="Could not connect to computer's registry",
                                                    newline=True))
            config.tasks.append(lambda: gui.submit.configure(state="normal"))
            config.tasks.append(lambda: gui.computer.bind("<Return>", lambda _: run_func(on_submit)))
            config.disable = True
            return

        config.tasks.append(lambda: print_success(gui.computer_status, "ONLINE", "Computer status: ", clear_=True))
        user_ = get_username(pc)
        if user_:
            config.current_user = user_
            if not passed_user or passed_user.lower() == user_.lower():
                config.tasks.append(lambda: update(gui.display_user, f"Current user: {user_name_translation(user_)}"))
            else:
                config.tasks.append(lambda: update_error(gui.display_user, "Current user: ", user_))
        else:
            config.tasks.append(lambda: update_error(gui.display_user, "Current user: ", "No user"))
        try:
            r_pc = WMI(pc)
            for k in r_pc.Win32_OperatingSystem():
                last_boot_time = datetime.strptime(k.LastBootUpTime.split('.')[0], '%Y%m%d%H%M%S')
                current_time = datetime.strptime(k.LocalDateTime.split('.')[0], '%Y%m%d%H%M%S')
                uptime_ = current_time - last_boot_time
                if uptime_ > timedelta(days=7):
                    config.tasks.append(lambda: update_error(gui.uptime, "Uptime: ", uptime_))
                else:
                    config.tasks.append(lambda: update(gui.uptime, f"Uptime: {uptime_}"))
                break
        except:
            config.tasks.append(lambda: update_error(gui.uptime, "Uptime: ", "ERROR"))
            log()
        sleep(0.1)

        try:
            space = get_space(pc)
            if space <= 5:
                config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ",
                                                         f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB"))
            else:
                config.tasks.append(lambda: update(gui.space_c, f"Space in C disk: {space:.1f}GB free"
                                                                f" out of {get_total_space(pc):.1f}GB"))
        except:
            log()
            config.tasks.append(lambda: update_error(gui.space_c, "Space in C disk: ", "ERROR"))

        if path.exists(fr"\\{pc}\d$"):
            try:
                space = get_space(pc, disk="d")
                if space <= 5:
                    config.tasks.append(lambda: update_error(
                        gui.space_d, "Space in D disk: ", f"{space:.1f}GB free out of "
                                                          f"{get_total_space(pc, disk='d'):.1f}GB"))
                else:
                    config.tasks.append(lambda: update(gui.space_d, f"Space in D disk: {space:.1f}GB free out of "
                                                                    f"{get_total_space(pc, disk='d'):.1f}GB"))
            except:
                log()
                config.tasks.append(lambda: update_error(gui.space_d, "Space in D disk: ", "ERROR"))
        else:
            config.tasks.append(lambda: update_error(gui.space_d, "Space in D disk: ", "Does not exist"))

        try:
            try:
                r_pc
            except NameError:
                r_pc = WMI(pc)
            for ram_ in r_pc.Win32_ComputerSystem():
                total_ram = int(ram_.TotalPhysicalMemory) / (1024 ** 3)
                if total_ram < 7:
                    config.tasks.append(lambda: update_error(gui.ram, "Total RAM: ", f"{round(total_ram)}GB"))
                else:
                    config.tasks.append(lambda: update(gui.ram, f"Total RAM: {round(total_ram)}GB"))
        except:
            config.tasks.append(lambda: update_error(gui.ram, "Total RAM: ", "ERROR"))
            log()

        if is_ie_fixed(pc):
            config.tasks.append(lambda: update(gui.ie_fixed, "Internet explorer: Fixed"))
        else:
            config.tasks.append(lambda: update_error(gui.ie_fixed, "Internet explorer: ", "Not fixed"))

        if is_cpt_fixed(pc):
            config.tasks.append(lambda: update(gui.cpt_fixed, "Cockpit printer: Fixed"))
        else:
            config.tasks.append(lambda: update_error(gui.cpt_fixed, "Cockpit printer", "Not fixed"))

        if user_ or passed_user:
            if passed_user:
                user_ = passed_user
            config.tasks.append(lambda: update_user(user_))
        else:
            config.tasks.append(lambda: update_error(gui.user_active, "User status: ", "No user"))
        config.tasks.append(lambda: enable())

    else:
        try:
            with open(f"{config.users_txt}\\{pc}.txt") as pc_file:
                user_ = pc
                pc = pc_file.read().strip()
                on_submit(pc=pc, passed_user=user_)
        except FileNotFoundError:
            config.disable = True
            if user_exists(pc):
                config.tasks.append(lambda: print_error(gui.console, f"Could not locate the current or last "
                                                                     f"computer {pc} has logged on to"))
                config.tasks.append(lambda: update_user(pc))
            else:
                if any([pc.lower() in config.ip_printers, pc.lower() in config.svr_printers]):
                    pc = pc.lower()
                    if pc in config.ip_printers:
                        config.tasks.append(lambda: print(f"Printer with an IP of {pc} is at {config.ip_printers[pc]}"))
                        pc = config.ip_printers[pc]
                    elif pc in config.svr_printers:
                        config.tasks.append(lambda: print(f"Printer {pc} has an ip of {config.svr_printers[pc]}"))
                        pc = config.svr_printers[pc]
                    config.tasks.append(lambda: copy_clip(pc))
                else:
                    if r"\\" in pc:
                        config.tasks.append(lambda: print_error(gui.console, f"Could not locate printer {pc}"))
                    elif pc.count(".") > 2:
                        config.tasks.append(lambda: print_error(gui.console,
                                                                f"Could not locate TCP/IP printer with ip of {pc}"))
                    else:
                        config.tasks.append(lambda:
                                            print_error(gui.console, f"No such user or computer in the domain {pc}"))
            config.tasks.append(lambda: gui.submit.configure(state="normal", cursor="hand2"))
            config.tasks.append(lambda: gui.computer.bind("<Return>", lambda _: run_func(on_submit)))
            return


sys.stdout.write = redirect


class SetConfig:
    """"sets the basic config for the script to run on, these configs are needed in order for the script to run"""

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
        self.mx_w = self.config_file["max_workers"]
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
        self.title = self.config_file["title"]
        self.tasks = []
        self.disable = False
        self.first_time = 1


try:
    with open("config.json", encoding="utf8") as config_file:
        config = SetConfig(load(config_file))
except FileNotFoundError:
    messagebox.showerror("config file error", "could not find the config.json file")
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
    """logs the exceptions as well as date, host, and username to the logfile"""
    if not config.log:
        return
    err_log = f"""{'_' * 145}\nat {datetime.now().strftime('%Y-%m-%d %H:%M')} an error occurred on {config.host}\
 - {config.user}\n"""
    exception(err_log)


if not config.log:
    basicConfig(filename="FATAL_errors.log", filemode="w", format="%(message)s")
    logger = getLogger("fatal exceptions")


class TimeoutException(Exception):
    pass


def Timeout(timeout: int | float) -> bool | TimeoutException | callable:
    """"run a function with timeout limit via threading"""

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
    """timed out test to check that the computer is connectable via WMI"""
    x = Timeout(timeout=1.5)(WMI_connectable_actual)
    try:
        y = x()
    except TimeoutException:
        config.wmi_connectable = False
        return False
    except:
        log()
        config.wmi_connectable = False
        return False
    config.reg_connectable = True if y else False
    return y


def WMI_connectable_actual() -> bool:
    """"the actual WMI_connectable check"""
    pc = config.current_computer
    try:
        pythoncom.CoInitialize()
        WMI(computer=pc)
        return True
    except:
        log()
        return False


def get_space(pc: str, disk: str = "c") -> float:
    """"returns the free space in the disk in GB"""
    return disk_usage(fr"\\{pc}\{disk}$").free / (1024 ** 3)


def get_total_space(pc: str, disk: str = "c") -> float:
    """"returns the total space in the disk in GB"""
    return disk_usage(fr"\\{pc}\{disk}$").total / (1024 ** 3)


def pc_in_domain(pc: str) -> str | None:
    """"query if the computer is in the domain"""
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
    """checks if a user exists in the domain"""
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


def query_user(username_: str) -> int:
    """"checks the user status in the domain, such as locked, disabled, expired, password expired"""
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

        dom_info = run(f'net user {username_} /domain | findstr /i "Account expires"', shell=True,
                       capture_output=True, text=True).stdout.strip().split()
        exp = " ".join(dom_info)
        if "Account expires Never" not in exp:
            expires = exp.split("Account expires")[1].strip()
            expires = f"{expires[:10]} {expires[11:19]}"
            if date_is_older(expires):
                return 3

        if "Locked" in run(f'net user {username_} /domain | findstr /i "Account active"', shell=True,
                           capture_output=True, text=True).stdout.strip():
            return 2

        pass_expired = exp.split("Password expires")[1].strip()
        pass_expired = f"{pass_expired[:10]} {pass_expired[11:19]}"
        if date_is_older(pass_expired):
            return 4

    return 0


def check_pc_active_actual(pc: str) -> bool:
    """checks if the computer is reachable via UNC pathing"""
    return path.exists(fr"\\{pc}\c$")


def check_pc_active(pc: str = None) -> bool:
    """"timed out check if the computer is online and reachable"""
    # noinspection PyCallingNonCallable
    x = Timeout(timeout=3)(check_pc_active_actual)
    try:
        y = x(pc=pc)
    except TimeoutException:
        return False
    except:
        log()
        return False

    return y


def get_username(pc: str) -> str | None:
    """retrieves the active user on the remote computer"""
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


def user_name_translation(username_: str) -> str | None:
    """returns the display name of a user in the domain"""
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


def date_is_older(date_string: datetime.strptime) -> bool:
    """checks if the initial date has passed"""
    provided_date = datetime.strptime(date_string, "%d/%m/%Y %H:%M:%S")
    return provided_date < datetime.now()


def on_rm_error(_: callable, path_: str, error: tuple) -> None:
    """deletes readonly files via changing permission to the file or directory"""
    if error[0] != PermissionError:
        return
    try:
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
    """timed check if the registry is connectable"""
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
    config.reg_connectable = True if y else False
    return y


def is_reg(pc: str = None) -> bool:
    """checks if the remote registry is connectable"""
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
    """checks if Internet Explorer is fixed via querying the registry keys"""
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
    """checks if cockpit printer is fixed via querying the registry keys"""
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


class GUI:
    def __init__(self):
        self.root = Tk()
        self.root.geometry("758x758")
        self.root.configure(bg="#545664")
        self.root.title(config.title)
        self.root.iconbitmap(asset("icon.ico"))
        self.canvas = Canvas(
            self.root,
            bg="#545664",
            height=758,
            width=758,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )
        self.canvas.place(x=0, y=0)
        self.generic_image = PhotoImage(file=asset("generic_text.png"))
        self.canvas.create_image(
            236.0,
            422.5,
            image=self.generic_image
        )
        self.user_active = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.user_active.bind("<<Selection>>", lambda event_: ignore_selection(self.user_active, event_))
        self.user_active.place(
            x=31.0,
            y=412.6,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            392.0,
            image=self.generic_image
        )
        self.cpt_fixed = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.cpt_fixed.place(
            x=31.0,
            y=382.2,
            width=410.0,
            height=21.2
        )
        self.cpt_fixed.bind("<<Selection>>", lambda event_: ignore_selection(self.cpt_fixed, event_))
        self.canvas.create_image(
            236.0,
            361.5,
            image=self.generic_image
        )
        self.ie_fixed = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.ie_fixed.bind("<<Selection>>", lambda event_: ignore_selection(self.ie_fixed, event_))
        self.ie_fixed.place(
            x=31.0,
            y=351.7,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            330.5,
            image=self.generic_image
        )
        self.ram = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.ram.bind("<<Selection>>", lambda event_: ignore_selection(self.ram, event_))
        self.ram.place(
            x=31.0,
            y=320.7,
            width=410.0,
            height=21.4
        )
        self.canvas.create_image(
            236.0,
            299.5,
            image=self.generic_image
        )
        self.space_d = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.space_d.bind("<<Selection>>", lambda event_: ignore_selection(self.space_d, event_))
        self.space_d.place(
            x=31.0,
            y=290.0,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            268.5,
            image=self.generic_image
        )
        self.space_c = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.space_c.bind("<<Selection>>", lambda event_: ignore_selection(self.space_c, event_))
        self.space_c.place(
            x=31.0,
            y=259.0,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            237.5,
            image=self.generic_image
        )
        self.canvas.create_image(
            236.0,
            237.5,
            image=self.generic_image
        )
        self.uptime = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.uptime.bind("<<Selection>>", lambda event_: ignore_selection(self.uptime, event_))
        self.uptime.place(
            x=31.0,
            y=228.0,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            206.5,
            image=self.generic_image
        )
        self.display_user = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.display_user.bind("<<Selection>>", lambda event_: ignore_selection(self.display_user, event_))
        self.display_user.place(
            x=31.0,
            y=197.0,
            width=410.0,
            height=21.2
        )
        self.canvas.create_image(
            236.0,
            175.5,
            image=self.generic_image
        )
        self.computer_status = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.computer_status.bind("<<Selection>>", lambda event_: ignore_selection(self.computer_status, event_))
        self.computer_status.place(
            x=31.0,
            y=166.0,
            width=410.0,
            height=21.2
        )
        self.rest_spool_image = PhotoImage(
            file=asset("button_1.png"))
        self.reset_spool = Button(
            image=self.rest_spool_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(reset_spooler),
            relief="flat",
            background="#545664"
        )
        self.reset_spool.place(
            x=599.0,
            y=131.0,
            width=128.0,
            height=55.0
        )
        self.delete_ost_image = PhotoImage(
            file=asset("button_2.png"))
        self.delete_ost = Button(
            image=self.delete_ost_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(delete_the_ost),
            relief="flat",
            background="#545664"
        )
        self.delete_ost.place(
            x=461.0,
            y=353.0,
            width=128.0,
            height=55.0
        )
        self.delete_users_image = PhotoImage(
            file=asset("button_3.png"))
        self.delete_users = Button(
            image=self.delete_users_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(del_users),
            relief="flat",
            background="#545664"
        )
        self.delete_users.place(
            x=461.0,
            y=279.0,
            width=128.0,
            height=55.0
        )
        self.get_printers_image = PhotoImage(
            file=asset("button_4.png"))
        self.get_printers = Button(
            image=self.get_printers_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(get_printers_func),
            relief="flat",
            background="#545664"
        )
        self.get_printers.place(
            x=461.0,
            y=205.0,
            width=128.0,
            height=55.0
        )
        self.fix_3_lang_image = PhotoImage(
            file=asset("button_5.png"))
        self.fix_3_lang = Button(
            image=self.fix_3_lang_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(fix_3_languages),
            relief="flat",
            background="#545664"
        )
        self.fix_3_lang.place(
            x=461.0,
            y=131.0,
            width=128.0,
            height=55.0
        )
        self.fix_cpt_image = PhotoImage(
            file=asset("button_6.png"))
        self.fix_cpt = Button(
            image=self.fix_cpt_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(fix_cpt_func),
            relief="flat",
            background="#545664"
        )
        self.fix_cpt.place(
            x=599.0,
            y=353.0,
            width=128.0,
            height=55.0
        )
        self.fix_ie_image = PhotoImage(
            file=asset("button_7.png"))
        self.fix_ie = Button(
            image=self.fix_ie_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(fix_ie_func),
            relief="flat",
            background="#545664"
        )
        self.fix_ie.place(
            x=599.0,
            y=280.0,
            width=128.0,
            height=55.0
        )
        self.clear_space_image = PhotoImage(
            file=asset("button_8.png"))
        self.clear_space = Button(
            image=self.clear_space_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(clear_space_func),
            relief="flat",
            background="#545664"
        )
        self.clear_space.place(
            x=599.0,
            y=205.0,
            width=128.0,
            height=55.0
        )
        self.computer_image = PhotoImage(
            file=asset("entry_8.png"))
        self.canvas.create_image(
            377.5,
            27.0,
            image=self.computer_image
        )
        self.computer = Entry(
            font=("Times", 12, "bold"),
            bd=0,
            bg="#C2EFF9",
            fg="#000716",
            justify="center",
            highlightthickness=0
        )
        self.computer.bind("<FocusIn>", show_text)
        self.computer.bind("<FocusOut>", hide_text)
        self.computer.bind_all("<Key>", enable_paste, "+")
        self.computer.insert(INSERT, "Computer or User")
        self.computer.bind("<Return>", lambda _: run_func(on_submit))
        self.computer.place(
            x=216.0,
            y=6.0,
            width=323.0,
            height=40.0
        )
        self.submit_image = PhotoImage(
            file=asset("button_9.png"))
        self.submit = Button(
            image=self.submit_image,
            activebackground="#545664",
            background="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: run_func(on_submit),
            relief="flat"
        )
        self.submit.place(
            x=306.0,
            y=55.0,
            width=146.0,
            height=45.0
        )
        self.display_pc_image = PhotoImage(
            file=asset("display_pc.png"))
        self.canvas.create_image(
            195.0,
            142.5,
            image=self.display_pc_image
        )
        self.display_pc = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('Arial', 12, 'bold'),
            cursor="arrow"
        )
        self.display_pc.bind("<<Selection>>", lambda event_: ignore_selection(self.display_pc, event_))
        self.display_pc.place(
            x=31.0,
            y=133.0,
            width=331.0,
            height=21.4
        )
        self.copy_but_image = PhotoImage(
            file=asset("copy.png"))
        self.copy_but = Button(
            image=self.copy_but_image,
            activebackground="#545664",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: copy_clip(config.current_computer[1:-4]),
            relief="flat",
            background="#545664"
        )
        self.copy_but.place(
            x=376.0,
            y=130.0,
            width=69.5,
            height=32.5
        )
        self.console_image = PhotoImage(
            file=asset("entry_10.png"))
        self.canvas.create_image(
            379.0,
            605.0,
            image=self.console_image
        )
        self.console = Text(
            bd=0,
            bg="#FFFFFF",
            fg="#000716",
            font=('Arial', 12, 'bold'),
            highlightthickness=0,
            state="disabled"
        )
        self.console.place(
            x=30.0,
            y=471.5,
            width=680.0,
            height=269.0
        )
        self.root.resizable(False, False)


mouse.Listener(on_click=on_button_press).start()
gui = GUI()
clear_all()
disable()
gui.root.mainloop()
