import getpass
import os
import platform
import subprocess
from tkinter import Frame, Tk, Button, Label, Entry, END, NORMAL, Menu
from tkinter import filedialog, messagebox


class App(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()
        self.centerWindow()
        self.USER_NAME = getpass.getuser()
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
        w = 381
        h = 236

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
        file_menu = Menu(menubar, tearoff=False,background='#0e1621', foreground='#6d7883', activebackground='white', activeforeground='black')

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

        self.entry = Entry(self.parent,
                           relief='groove',
                           bg='#17212b',
                           fg='#6d7883')
        self.entry.insert(END, 'Your custom file name')
        self.clicked = self.entry.bind('<Button-1>', self.on_click)
        self.entry.grid(row=0, column=2, pady=3, sticky='NSEW')

    def initButtons(self):
        btn1 = Button(self.parent,
                      text="Chose file",
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.on_autorun_folder_open)
        btn1.grid(row=1, column=2, pady=2, sticky='NSEW')

        btn2 = Button(self.parent,
                      text="Go to autorun folder",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.open_autorun_folder)
        btn2.grid(row=4, column=1, pady=2)

        btn3 = Button(self.parent,
                      text="Open Task Scheduler",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.open_task_sheduler)
        btn3.grid(row=5, column=1, pady=2)

        btn4 = Button(self.parent,
                      text="regedit HKEY_LOCAL_MACHINE\...\Run",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.registry_run_local_machine_open)
        btn4.grid(row=6, column=1, pady=2)

        btn5 = Button(self.parent,
                      text="regedit HKEY_LOCAL_MACHINE\...\RunOnce",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.registry_run_once_local_machine_open)
        btn5.grid(row=7, column=1, pady=2)

        btn6 = Button(self.parent,
                      text="regedit HKEY_CURRENT_USER\...\Run",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.registry_run_local_user_open)
        btn6.grid(row=8, column=1, pady=2)

        btn7 = Button(self.parent,
                      text="regedit HKEY_CURRENT_USER\...\RunOnce",
                      width=34,
                      foreground='#17212b',
                      background='#5288c1',
                      relief='flat',
                      command=self.registry_run_once_local_user_open)
        btn7.grid(row=9, column=1, pady=2)

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

    def on_click(self, event):
        self.entry.configure(state=NORMAL)
        self.entry.delete(0, END)
        self.entry.unbind('<Button-1>', self.clicked)

    def on_autorun_folder_open(self):
        ftypes = [('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        self.file_path = dlg.show()
        self.add_to_startup()

    def explore(self, path):
        # explorer would choke on forward slashes
        path = os.path.normpath(path)
        FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

        if os.path.isdir(path):
            subprocess.run([FILEBROWSER_PATH, path])
        elif os.path.isfile(path):
            subprocess.run([FILEBROWSER_PATH, '/select,', path])

    def open_autorun_folder(self):
        if platform.system() == "Windows":
            autorun_folder = rf'C:\Users\{self.USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\.'
            self.explore(autorun_folder)

    def add_to_startup(self):
        if platform.system() == "Windows":
            file_path = self.file_path
            bat_path = fr'C:\Users\{self.USER_NAME}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'
            autorun_filename = self.entry.get()
            if autorun_filename == 'Your custom file name':
                autorun_filename = file_path.split('/')[-1]

            new_file = bat_path + '\\' + f"{autorun_filename}_autorun.bat"
            user_answer = 'yes'
            if os.path.exists(new_file):
                user_answer = messagebox.askquestion(
                    'File exists!', 'Do you want to rewrite it?')
                print(user_answer)

            if user_answer == 'yes':
                with open(new_file, "w+") as bat_file:
                    if '.py' in self.file_path:
                        bat_file.write(fr'start pythonw "{file_path}"')
                    else:
                        bat_file.write(fr'start "{file_path}"')
                messagebox.showinfo('Done!', 'Done!')
            else:
                messagebox.showinfo('To Do:', 'Then enter custom name!')


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
