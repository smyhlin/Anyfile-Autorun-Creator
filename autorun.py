import getpass
import os
import platform
import subprocess
import tkinter
import tkinter.messagebox
from tkinter import ttk, filedialog, messagebox
import customtkinter

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    WIDTH = 780
    HEIGHT = 390
    USER_NAME = getpass.getuser()
    ps1 = """function jumpReg ($registryPath)
                        {
                            New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit" `
                                            -Name "LastKey" `
                                            -Value $registryPath `
                                            -PropertyType String `
                                            -Force

                            regedit
                        }"""

    def __init__(self):
        super().__init__()

        self.title("Anyfile Autorun Creator")
        self.centerWindow()
        self.resizable(False, False)
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed

        #! ============ create two frames ============

        #! configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self,
                                                 width=180,
                                                 corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe")

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        #! ============ frame_left ============

        #! configure grid layout (1x11)
        self.frame_left.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(8, weight=1)  # empty row as spacing
        self.frame_left.grid_rowconfigure(8, minsize=20)    # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(11, minsize=10)  # empty row with minsize as spacing

        self.label_1 = customtkinter.CTkLabel(master=self.frame_left,
                                              text="Autorun habitats",
                                              text_font=("Roboto Medium", -16))  # font name and size in px
        self.label_1.grid(row=1, column=0, pady=10, padx=10)

        self.button_1 = customtkinter.CTkButton(master=self.frame_left,
                                                text="Autorun folder",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_windows_autorun_folder)
        self.button_1.grid(row=2, column=0, pady=5, padx=20, sticky="we")

        self.button_2 = customtkinter.CTkButton(master=self.frame_left,
                                                text="Task Scheduler",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_task_sheduler)
        self.button_2.grid(row=3, column=0, pady=5, padx=20, sticky="we")

        self.button_3 = customtkinter.CTkButton(master=self.frame_left,
                                                text="HKEY_LOCAL_MACHINE\...\Run",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_registry_run_local_machine)
        self.button_3.grid(row=4, column=0, pady=5, padx=20, sticky="we")

        self.button_4 = customtkinter.CTkButton(master=self.frame_left,
                                                text="HKEY_LOCAL_MACHINE\...\RunOnce",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_registry_run_once_local_machine)
        self.button_4.grid(row=5, column=0, pady=5, padx=20, sticky="we")

        self.button_5 = customtkinter.CTkButton(master=self.frame_left,
                                                text="HKEY_CURRENT_USER\...\Run",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_registry_run_current_user)
        self.button_5.grid(row=6, column=0, pady=5, padx=20, sticky="we")
        
        self.button_6 = customtkinter.CTkButton(master=self.frame_left,
                                                text="HKEY_CURRENT_USER\...\RunOnce",
                                                fg_color=("gray75", "gray30"),  # <- custom tuple-color
                                                command=self.open_registry_run_once_current_user)
        self.button_6.grid(row=7, column=0, pady=5, padx=20, sticky="we")
        
        self.switch_1 = customtkinter.CTkSwitch(master=self.frame_left,
                                                text="Auto-Replace spaces with '_'")
        self.switch_1.grid(row=10, column=0, pady=10, padx=20, sticky="w")
        
        self.switch_2 = customtkinter.CTkSwitch(master=self.frame_left,
                                                text="Dark Mode",
                                                command=self.change_mode)
        self.switch_2.grid(row=11, column=0, pady=10, padx=20, sticky="w")

        #! ============ frame_right ============

        #! configure grid layout (3x7)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=20)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.frame_right.grid_rowconfigure(3, weight=100)  # empty row as spacing

        
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=0, column=0, columnspan=2, rowspan=4, pady=20, padx=20, sticky="nsew")

        #! ============ frame_info ============

        #! configure grid layout (1x1)
        self.frame_info.rowconfigure(0, weight=1)
        self.frame_info.columnconfigure(0, weight=1)
        self.frame_info.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_info.grid_rowconfigure(8, weight=1)  # empty row as spacing
        self.frame_info.grid_rowconfigure(8, minsize=20)    # empty row with minsize as spacing
        self.frame_info.grid_rowconfigure(11, minsize=10)  # empty row with minsize as spacing

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                   text="Print custom autorun file name:",
                                                   height=70,
                                                   fg_color=("white", "gray38"),  # <- custom tuple-color
                                                   text_font=("Roboto Medium", -14),
                                                   justify=tkinter.LEFT)
        self.label_info_1.grid(row=0, column=0, sticky="wesn", padx=15, pady=15)


        self.entry = customtkinter.CTkEntry(master=self.frame_info,
                                            width=120,
                                            placeholder_text="Your custom file name")
        self.entry.grid(row=10, column=0, pady=3, padx=20, sticky="we")

        self.button_7 = customtkinter.CTkButton(master=self.frame_info,
                                                text="Select File",
                                                command=self.select_file_to_autorun)
        self.button_7.grid(row=11, column=0, pady=3, padx=20, sticky="we")
        separator = ttk.Separator(self.frame_info, orient='horizontal')
        separator.grid(row=12, column=0, pady=20, padx=20, sticky="wesn")
        #! ============ frame_right ============

        self.radio_var = tkinter.IntVar(value=0)

        self.label_radio_group = customtkinter.CTkLabel(master=self.frame_right,
                                                        text="Select autorun type:",
                                                        text_font=("Roboto Medium", -14))
        self.label_radio_group.grid(row=0, column=2, columnspan=1, pady=20, padx=10, sticky="nw")

        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           value=0,
                                                           text="Autorun folder")
        self.radio_button_1.grid(row=1, column=2, pady=3, padx=20, sticky="nw")

        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           value=1,
                                                           text="Task Scheduler")
        self.radio_button_2.grid(row=2, column=2, pady=3, padx=20, sticky="nw")

        self.radio_button_3 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           value=2,
                                                           text="Registry")
        self.radio_button_3.grid(row=3, column=2, pady=3, padx=20, sticky="nw")

        # set default values
        self.radio_button_1.select()
        self.switch_1.select()
        self.switch_2.select()

    def centerWindow(self):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()

        x = (sw - App.WIDTH) / 2
        y = (sh - App.HEIGHT) / 2
        self.geometry('%dx%d+%d+%d' % (App.WIDTH, App.HEIGHT, x, y))
        
    def button_event(self):
        print("Button pressed")

    def change_mode(self):
        if self.switch_2.get() == 1:
            customtkinter.set_appearance_mode("dark")
        else:
            customtkinter.set_appearance_mode("light")

    def remove_spaces_from_filename(self):
        if self.switch_1.get() == 1:
            os.rename(self.file_path, self.file_path.replace(' ', '_'))
            self.file_path = self.file_path.replace(' ', '_')
            return True
        else:
            messagebox.showinfo('To Do:', 
                                'Enable Auto-Rename switch\n'
                                'OR Chose another file!'
                                '\n\n CMD won`t work with spaces in file name')
            return False

    def select_file_to_autorun(self):
        ftypes = [('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        self.file_path = dlg.show()
        self.add_to_startup()

    def open_task_sheduler(self):
        os.system('taskschd.msc')

    def open_registry_run_current_user(self):
        hkey_run = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', App.ps1 + hkey])

    def open_registry_run_once_current_user(self):
        hkey_run = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', App.ps1 + hkey])

    def open_registry_run_local_machine(self):
        hkey_run = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', App.ps1 + hkey])

    def open_registry_run_once_local_machine(self):
        hkey_run = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', App.ps1 + hkey])

    def chose_file_to_autorun(self):
        ftypes = [('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        self.file_path = dlg.show()
        self.add_to_startup()

    def open_windows_autorun_folder(self):
        if platform.system() == "Windows":
            autorun_folder = rf'C:\Users\{self.USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\.'

            # Open file or directory by path
            path = os.path.normpath(autorun_folder)
            FILEBROWSER_PATH = os.path.join(
                os.getenv('WINDIR'), 'explorer.exe')

            if os.path.isdir(path):
                subprocess.run([FILEBROWSER_PATH, path])
            elif os.path.isfile(path):
                subprocess.run([FILEBROWSER_PATH, '/select,', path])

    def add_to_startup(self):
        # Function where we get valid file name|path and chose type of "autoruner" from radiobutton
        if platform.system() == "Windows":
            if self.file_path:
                user_answer = True
                if ' ' in self.file_path:  # Check for spaces in file name and replace with '_'
                    user_answer = self.remove_spaces_from_filename()

                if user_answer:
                    self.autorun_file_directory = "/".join(self.file_path.split("/")[:-1])
                    custom_autorun_filename = self.entry.get()
                    if ' ' in custom_autorun_filename:
                        custom_autorun_filename = custom_autorun_filename.replace('', '_')

                    if custom_autorun_filename.endswith(('Your custom file name', '')):  # When user not set custom_autorun_filename we took filename from system 
                        custom_autorun_filename = self.file_path.split('/')[-1]

                    match self.radio_var.get():
                        case 0:  # Autorun folder
                            self.add_to_autorun_folder(custom_autorun_filename)
                        case 1:  # Autorun sheduler
                            self.add_to_task_sheduler(custom_autorun_filename)
                        case 2:  # Autorun registry
                            self.add_to_autorun_registry(custom_autorun_filename)
        else:
            print('Not realized')

    def add_to_autorun_folder(self, autorun_filename):
        bat_path = fr'C:\Users\{self.USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'

        new_file = bat_path + '\\' + f"{autorun_filename}_autorun.bat"
        user_answer = 'yes'
        if os.path.exists(new_file):
            user_answer = messagebox.askquestion('Autorun File exists!', 
                                                 'Autorun File exists!\n\nDo you want to rewrite it?')

        if user_answer == 'yes':
            with open(new_file, "w+", encoding='cp866') as bat_file:
                if '.py' in self.file_path:
                    bat_file.write(fr'start pythonw "{self.file_path}"')
                else:
                    bat_file.write(fr'start /d "{self.autorun_file_directory}/" {autorun_filename}  && Exit')
            messagebox.showinfo('Done!', 'Done!')
        else:
            messagebox.showinfo('To Do:', 'Then enter custom name!')

    def add_to_task_sheduler(self, autorun_filename):
        import win32com.client
        
        # define constants
        scheduler = win32com.client.Dispatch('Schedule.Service')
        scheduler.Connect()
        root_folder = scheduler.GetFolder('\\')
        task_def = scheduler.NewTask(0)

        # Create trigger
        TASK_TRIGGER_LOGON = 9
        trigger = task_def.Triggers.Create(TASK_TRIGGER_LOGON)
        trigger.Id = "LogonTriggerId"
        # trigger.UserId = os.environ.get('USERNAME') # Run for current user only

        # Create action
        TASK_ACTION_EXEC = 0
        action = task_def.Actions.Create(TASK_ACTION_EXEC)
        action.ID = 'Anyfile-Autorun-Creator'
        action.Path = 'cmd.exe'
        action.Arguments = f'/k start /d "{self.autorun_file_directory}/" {autorun_filename} && Exit'

        # Set parameters
        task_def.RegistrationInfo.Description = autorun_filename
        task_def.RegistrationInfo.Author = "Anyfile-Autorun-Creator"
        task_def.Settings.Enabled = True
        task_def.Settings.StopIfGoingOnBatteries = False

        # Register task
        # If task already exists, it will be updated
        TASK_CREATE_OR_UPDATE = 6
        TASK_ON_LOGON = 3
        root_folder.RegisterTaskDefinition(
            autorun_filename,  # Task name
            task_def,
            TASK_CREATE_OR_UPDATE,
            '',  # No user
            '',  # No password
            TASK_ON_LOGON)
        messagebox.showinfo('Done!', 'Done!')

    def add_to_autorun_registry(self, autorun_filename):
        import winreg
        # key we want to change is HKEY_CURRENT_USER
        # key value is Software\Microsoft\Windows\CurrentVersion\Run
        key = winreg.HKEY_CURRENT_USER
        key_value = "Software\Microsoft\Windows\CurrentVersion\Run"

        # open the key to make changes to
        open = winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS)
        # modify the opened key
        winreg.SetValueEx(open, autorun_filename, 0,
                          winreg.REG_SZ, f'cmd.exe /k start /d "{self.autorun_file_directory}/" {autorun_filename} && Exit')

        # now close the opened key
        winreg.CloseKey(open)
        messagebox.showinfo('Done!', 'Done!')
        
    def on_closing(self, event=0):
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()