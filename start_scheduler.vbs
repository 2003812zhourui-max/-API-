' 开机静默启动领星调度平台（不弹黑窗）
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """C:\Users\Administrator\Documents\领星逆向\start_scheduler.bat""", 0, False
