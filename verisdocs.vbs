Const NETHOOD = &H13&

Set objWSHShell = CreateObject(“Wscript.Shell”)
Set objShell = CreateObject(“Shell.Application”)
Set objFolder = objShell.Namespace(NETHOOD)
Set objFolderItem = objFolder.Self
strNetHood = objFolderItem.Path
strShortcutName = “Veris Docs”
strShortcutPath = “\\192.168.3.15\verisdocs”
Set objShortcut = objWSHShell.CreateShortcut _
    (strNetHood & “\” & strShortcutName & “.lnk”)
objShortcut.TargetPath = strShortcutPath
objShortcut.Save