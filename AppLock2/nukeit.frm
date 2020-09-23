VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuke AppLock2"
   ClientHeight    =   750
   ClientLeft      =   6300
   ClientTop       =   4110
   ClientWidth     =   2865
   Icon            =   "nukeit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2865
   Begin VB.CommandButton Command1 
      Caption         =   "Nuke It"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DeleteKey "HKEY_CURRENT_USER\Software\AppLock2"
DeleteKey "HKEY_CLASSES_ROOT\.lck\PersistentHandler"
DeleteKey "HKEY_CLASSES_ROOT\.lck"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App\DefaultIcon"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Open\Command"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Open"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Unlock\Command"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App\Shell\Unlock"
DeleteKey "HKEY_CLASSES_ROOT\Locked.App"
DeleteKey "HKEY_CLASSES_ROOT\exefile\Shell\Lock\Command"
DeleteKey "HKEY_CLASSES_ROOT\exefile\Shell\Lock"
MsgBox "AppLock 2 has been nuked"
Unload Me
End Sub
