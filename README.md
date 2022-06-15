# Anyfile-Autorun-Creator
---

A utility for automatic creation Autorun files, as well as quickly moving to the main locations of Autorun entries for manual modification.


#### Features:
> ✅ So far only for Windows

> ✅ Adding startup entries to the autorun folder:  
<sub>Win + R</sub> <sub>shell:startup</sub>

> ✅ Adding startup entries to the registry

> ✅ Adding startup entries to the Task Scheduler

> ✅ Quick transition to all kinds of autorun places:
 - Task Scheduler
 - HKEY_LOCAL_MACHINE\\...\\Run
 - HKEY_CURRENT_USER\\...\\Run
 - HKEY_CURRENT_USER\\...\\RunOnce
 - HKEY_LOCAL_MACHINE\\...\\RunOnce

>Dark Mode:
<p align="center">
<img src="https://telegra.ph/file/c5743d5ec7a37a491aca7.png">
</p>

>Light Mode:
<p align="center">
<img src="https://telegra.ph/file/d46a01d590a2ea2d140f2.png">
</p>


#### Run/Build Instructions:
first install requirements.txt with pip  
```
> pip install -r requirements.txt
```


Next, you can run `autorun.py`

Or build the .exe file with the following command:
```
> pip install pyinstaller
> pyinstaller -F --onefile --noconsole ^
--clean --icon="customtkinter\assets\icon.ico" ^
--add-data=customtkinter;customtkinter "autorun.py"
```
The executable will be created in `dist` directory

Or you can just download already built executable [here](https://github.com/smyhlin/Anyfile-Autorun-Creator/releases)

#### ToDo:
> TODO


## Telegram Support:

[![ME](https://img.shields.io/badge/TG-ME-30302f?style=flat&logo=telegram)](https://t.me/s_myhlin)

#### LICENSE
- GPLv3