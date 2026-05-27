$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Form Extractor.lnk")
$Shortcut.TargetPath = "pythonw.exe"
$Shortcut.Arguments = "app.py"
$Shortcut.WorkingDirectory = "C:\Users\CommShotMaster\Desktop\haitech\idea1"
$Shortcut.Description = "HaiTech Handwritten Form Extractor"
$Shortcut.IconLocation = "C:\Users\CommShotMaster\Desktop\haitech\idea1\icon.ico,0"
$Shortcut.Save()
Write-Host "Shortcut updated with HaiTech icon!"
