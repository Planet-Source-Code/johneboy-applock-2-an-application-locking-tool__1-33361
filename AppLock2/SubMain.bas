Attribute VB_Name = "SubMain"
Option Explicit
Public Sub Main()
Dim EXEpath As String
Dim Path As String
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"
CreateKey "HKEY_CURRENT_USER\Software\AppLock2"
CreateKey "HKEY_CLASSES_ROOT\.lck"
SetStringValue "HKEY_CLASSES_ROOT\.lck", "", "Locked.App"
CreateKey "HKEY_CLASSES_ROOT\.lck\PersistentHandler"
SetStringValue "HKEY_CLASSES_ROOT\.lck\PersistentHandler", "", "{098f2470-bae0-11cd-b579-08002b30bfeb}"
CreateKey "HKEY_CLASSES_ROOT\Locked.App"
SetStringValue "HKEY_CLASSES_ROOT\Locked.App", "", "Locked Application"
SetStringValue "HKEY_CLASSES_ROOT\Locked.App", "NeverShowExt", ""
CreateKey "HKEY_CLASSES_ROOT\Locked.App\DefaultIcon"
SetStringValue "HKEY_CLASSES_ROOT\Locked.App\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe ,0"
CreateKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Open\Command"
SetStringValue "HKEY_CLASSES_ROOT\Locked.App\Shell\Open\Command", "", "" & EXEpath
CreateKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Unlock\Command"
SetStringValue "HKEY_CLASSES_ROOT\Locked.App\Shell\Unlock\Command", "", "" & EXEpath
SetStringValue "HKEY_CURRENT_USER\Software\AppLock2", "AppPath", "" & EXEpath
CreateKey "HKEY_CLASSES_ROOT\exefile\Shell\Lock\Command"
SetStringValue "HKEY_CLASSES_ROOT\exefile\Shell\Lock\Command", "", "" & EXEpath
If FirstRun = "No" Then

    Path = Command
        If Path = "" Then
        FrmHowTo.Show
        Else
        If Right(Path, 3) = "exe" Then
        LockFrm.Show
        Else
        If Right(Path, 3) = "bat" Or Right(Path, 3) = "reg" Then
        RunFileFrm.Show
        Else
        UnlockFrm.Show
        End If
        End If
         End If

Else
SetPassFrm.Show
End If
End Sub
