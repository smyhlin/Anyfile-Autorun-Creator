import getpass
import os
import platform
import subprocess
import winreg as reg
from tkinter import Frame, IntVar, Tk, END, NORMAL
from tkinter import filedialog, messagebox, Radiobutton, Button, Label, Entry, Menu


class App(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()
        self.centerWindow()
        self.USER_NAME = getpass.getuser()
        self.autorun_type_variable = 1
        self.ps1 = """function jumpReg ($registryPath)
                        {
                            New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit" `
                                            -Name "LastKey" `
                                            -Value $registryPath `
                                            -PropertyType String `
                                            -Force

                            regedit
                        }"""

    def centerWindow(self):
        w = 441
        h = 177

        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()

        x = (sw - w) / 2
        y = (sh - h) / 2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def initUI(self):
        self.parent.configure(background='#0e1621')
        self.parent.resizable(False, False)
        self.master.title("Anyfile Autorun Creator")
        self.columnconfigure(0)
        self.rowconfigure(0)
        self.initMenuBar()
        self.initTextElements()
        self.initButtons()

    def initMenuBar(self):
        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        file_menu = Menu(menubar, tearoff=False, background='#0e1621',
                         foreground='#6d7883', activebackground='white', activeforeground='black')

        menubar.add_cascade(label="Open",
                            menu=file_menu
                            )

        file_menu.add_command(label='Open autorun folder',
                              command=self.open_autorun_folder)
        file_menu.add_separator()

        file_menu.add_command(label='Open Task Scheduler',
                              command=self.open_task_sheduler)
        file_menu.add_separator()

        file_menu.add_command(label='HKEY_LOCAL_MACHINE\...\Run',
                              command=self.registry_run_local_machine_open)
        file_menu.add_separator()

        file_menu.add_command(label='HKEY_CURRENT_USER\...\Run',
                              command=self.registry_run_local_user_open)
        file_menu.add_separator()

        file_menu.add_command(label='HKEY_LOCAL_MACHINE\...\RunOnce',
                              command=self.registry_run_once_local_machine_open)
        file_menu.add_separator()

        file_menu.add_command(label='HKEY_CURRENT_USER\...\RunOnce',
                              command=self.registry_run_once_local_user_open)

    def initTextElements(self):
        lbl = Label(self.parent,
                    text="Print custom autorun file name:",
                    bg='#0e1621', fg='#6d7883',
                    font=('Helvetica 13 bold'))
        lbl.grid(row=0, column=1)

        lb2 = Label(self.parent,
                    text="Select autorun type:",
                    bg='#0e1621', fg='#6d7883',
                    font=('Helvetica 13 bold'))
        lb2.grid(row=3, column=1, sticky='W')

        self.entry = Entry(self.parent,
                           relief='groove',
                           bg='#17212b',
                           fg='#6d7883',
                           font=('Helvetica 13'))
        self.entry.insert(END, 'Your custom file name')
        self.on_entry_clicked = self.entry.bind(
            '<Button-1>', self.on_entry_click)
        self.entry.grid(row=0, column=2, pady=3, sticky='NSEW')

    def initButtons(self):
        btn1 = Button(self.parent,
                      text="Chose file",
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.chose_file_to_autorun,
                      font=('Helvetica 13')
                      )
        btn1.grid(row=1, column=2, pady=2, sticky='NSEW')

        self.autorun_type = IntVar()
        b_sticky = 'W'
        btn2 = Radiobutton(self.parent,
                           text="Autorun folder",
                           foreground='#6d7883',
                           background='#0e1621',
                           relief='flat',
                           variable=self.autorun_type,
                           value=1,
                           font=('Helvetica 13 bold'),
                           command=self.select_radio_button)
        btn2.grid(row=4, column=1, sticky=b_sticky)
        btn2.invoke()
        
        btn3 = Radiobutton(self.parent,
                           text="Task Scheduler",
                           foreground='#6d7883',
                           background='#0e1621',
                           relief='flat',
                           variable=self.autorun_type,
                           value=2,
                           font=('Helvetica 13 bold'),
                           command=self.select_radio_button)
        btn3.grid(row=5, column=1, sticky=b_sticky)

        btn4 = Radiobutton(self.parent,
                           text="Registry",
                           foreground='#6d7883',
                           background='#0e1621',
                           relief='flat',
                           variable=self.autorun_type,
                           value=3,
                           font=('Helvetica 13 bold'),
                           command=self.select_radio_button)
        btn4.grid(row=6, column=1, sticky=b_sticky)

    def select_radio_button(self):
        self.autorun_type_variable = self.autorun_type.get()

    def open_task_sheduler(self):
        os.system('taskschd.msc')

    def registry_run_local_user_open(self):
        hkey_run = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', self.ps1 + hkey])

    def registry_run_once_local_user_open(self):
        hkey_run = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', self.ps1 + hkey])

    def registry_run_local_machine_open(self):
        hkey_run = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', self.ps1 + hkey])

    def registry_run_once_local_machine_open(self):
        hkey_run = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
        hkey = f'jumpReg ("{hkey_run}")'
        subprocess.Popen(['powershell.exe', self.ps1 + hkey])

    def on_entry_click(self, event):
        self.entry.configure(state=NORMAL)
        self.entry.delete(0, END)
        self.entry.unbind('<Button-1>', self.on_entry_clicked)

    def chose_file_to_autorun(self):
        ftypes = [('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        self.file_path = dlg.show()
        self.add_to_startup()

    def open_autorun_folder(self):
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
        if platform.system() == "Windows":
            if self.file_path:
                autorun_filename = self.entry.get()
                if autorun_filename.endswith(('Your custom file name', '')):
                    autorun_filename = self.file_path.split('/')[-1]

                match self.autorun_type_variable:
                    case 1:  # Autorun folder
                        self.add_to_autorun_folder(autorun_filename)
                    case 2:  # Autorun sheduler
                        self.add_to_task_sheduler(autorun_filename)
                    case 3:  # Autorun registry
                        self.add_to_autorun_registry(autorun_filename)

    def add_to_autorun_folder(self, autorun_filename):
        bat_path = fr'C:\Users\{self.USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'

        new_file = bat_path + '\\' + f"{autorun_filename}_autorun.bat"
        user_answer = 'yes'
        if os.path.exists(new_file):
            user_answer = messagebox.askquestion(
                'File exists!', 'Do you want to rewrite it?')

        if user_answer == 'yes':
            with open(new_file, "w+") as bat_file:
                if '.py' in self.file_path:
                    bat_file.write(fr'start pythonw "{self.file_path}"')
                else:
                    bat_file.write(fr'start "{self.file_path}"')
            messagebox.showinfo('Done!', 'Done!')
        else:
            messagebox.showinfo(
                'To Do:', 'Then enter custom name!')
    
    def add_to_task_sheduler(self):
        pass
    
    def add_to_autorun_registry(self, key_name):
        # key we want to change is HKEY_CURRENT_USER
        # key value is Software\Microsoft\Windows\CurrentVersion\Run
        key = reg.HKEY_CURRENT_USER
        key_value = "Software\Microsoft\Windows\CurrentVersion\Run"

        # open the key to make changes to
        open = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
        # modify the opened key
        reg.SetValueEx(open, key_name, 0, reg.REG_SZ, self.file_path)

        # now close the opened key
        reg.CloseKey(open)
        messagebox.showinfo('Done!', 'Done!')

def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
