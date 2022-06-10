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

>Main window:
<p align="center">
<img src="https://telegra.ph/file/ed90783f64f34401f8382.png">
</p>

>Menubar:
<p align="center">
<img src="https://telegra.ph/file/b17634b9e311fe85bd35b.png">
</p>


#### Run/Build Instructions
first install requirements.txt with pip  
```
> pip install requirements.txt
```


Next, you can run `autorun.py`

Or build the .exe file with the following command:
```
> pip install pyinstaller 
> pyinstaller --onefile --noconsole autorun.py
```
The executable will be created in `dist` directory

Or you can just download already built executable [here](https://github.com/smyhlin/Anyfile-Autorun-Creator/releases)

#### ToDo:
> TODO


## Telegram Support:

[![ME](https://img.shields.io/badge/TG-ME-30302f?style=flat&logo=telegram)](https://t.me/s_myhlin)

#### LICENSE
- GPLv3