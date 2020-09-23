VERSION 5.00
Begin VB.Form AdvancedFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AppLock 2"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "AdvancedFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "File Types"
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "Lock"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Lock"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Done Yet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Lock access to run registry (*.reg) files."
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1140
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Lock access to run batch (*.bat) files."
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   420
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CMD.exe"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "Lock"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Main File (C:\WINNT\SYSTEM32\CMD.EXE)"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Backup File (C:\WINNT\SYSTEM32\DLLCACHE\CMD.EXE)"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock access to CMD.exe"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Regedt32.exe"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "Lock"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Main File (C:\WINNT\REGEDT32.EXE)"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Backup File (C:\WINNT\SYSTEM32\DLLCACHE\REGEDT32.EXE)"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock access to RegEdt32.exe"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Regedit.exe"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      Begin VB.CheckBox Check2 
         Caption         =   "Backup File (C:\WINNT\SYSTEM32\DLLCACHE\REGEDIT.EXE)"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Main File (C:\WINNT\REGEDIT.EXE)"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Lock"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock access to RegEdit.exe"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "These are some advanced options that will allow you to make AppLock more secure."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "AdvancedFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EXEpath As String


Private Sub Command1_Click()
On Error Resume Next
If Check1.Value = 0 And Check2.Value = 0 Then
MsgBox "You selected not to lock either file.", vbOKOnly + vbCritical, "Lock Failed."
Else
    If Check1.Value = 1 Then
    Name "c:\winnt\regedit.exe" As "c:\winnt\regedit.exe.lck"
    Command1.Enabled = False
    If Check2.Value = 1 Then
    Name "c:\winnt\system32\dllcache\regedit.exe" As "c:\winnt\system32\dllcache\regedit.exe.lck"
    
Command1.Enabled = False
End If
End If
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Check9.Value = 0 And Check8.Value = 0 Then
MsgBox "You selected not to lock either file.", vbOKOnly + vbCritical, "Lock Failed."
Else
If Check9.Value = 1 Then
Name "c:\winnt\system32\regedt32.exe" As "c:\winnt\system32\regedt32.exe.lck"
Command2.Enabled = False
If Check8.Value = 1 Then
Name "c:\winnt\system32\dllcache\regedt32.exe" As "c:\winnt\system32\dllcache\regedt32.exe.lck"
Command2.Enabled = False
End If
End If
End If

End Sub

Private Sub Command3_Click()
On Error Resume Next
If Check6.Value = 0 And Check5.Value = 0 Then
MsgBox "You selected not to lock either file.", vbOKOnly + vbCritical, "Lock Failed."
Else
If Check6.Value = 1 Then
Name "c:\winnt\system32\cmd.exe" As "c:\winnt\system32\cmd.exe.lck"
Command3.Enabled = False
If Check5.Value = 1 Then
Name "c:\winnt\system32\dllcache\cmd.exe" As "c:\winnt\system32\dllcache\cmd.exe.lck"
Command3.Enabled = False
End If
End If
End If
End Sub

Private Sub Command4_Click()
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"

SetStringValue "HKEY_CLASSES_ROOT\batfile\shell\open\command", "", "" & EXEpath
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"

SetStringValue "HKEY_CLASSES_ROOT\regfile\shell\open\command", "", "" & EXEpath
Command5.Enabled = False
End Sub

Private Sub Command6_Click()
FrmHowTo.Show
Unload Me
End Sub
