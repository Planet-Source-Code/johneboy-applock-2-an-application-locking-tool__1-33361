VERSION 5.00
Begin VB.Form RunFileFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AppLock"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "RunFileFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "Encrypted Pass From Reg"
      Top             =   7200
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Text            =   "Decrypted Pass From Reg"
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File to be run..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "RunFileFrm.frx":0CCA
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Full Path Of Application"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "RunFileFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function TDecrypt(iString)
Dim q As String
Dim zz As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim x As Variant
Dim f As Variant
Dim txt As String
Dim txt2 As String
    On Error GoTo uhohs
    q = ""
    zz = Left(iString, 3)
    a = Left(zz, 1)
    b = Mid(zz, 2, 1)
    c = Mid(zz, 3, 1)
    d = Right(iString, 1)
    a = Int(Asc(a)) 'key 1
    b = Int(Asc(b)) 'key 2
    c = Int(Asc(c)) 'key 3
    d = Int(Asc(d)) 'key 4
    txt = Left(iString, Len(iString) - 1)
    txt2 = Mid(txt, 4, Len(txt)) 'encrypted text
    e = 1
    For x = 1 To Len(txt2)
        f = Mid(txt2, x, 1)
        If e = 1 Then q = q & Chr(Asc(f) - a)
        If e = 2 Then q = q & Chr(Asc(f) - b)
        If e = 3 Then q = q & Chr(Asc(f) - c)
        If e = 4 Then q = q & Chr(Asc(f) - d)
        e = e + 1
        If e > 4 Then e = 1
    Next x
    TDecrypt = q
    Exit Function
uhohs:
    TDecrypt = "Error: Invalid text To Decrypt"
    Exit Function
End Function
Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1 = Text2 Then
Shell Label4.Caption
Else
MsgBox "Incorrect Password Entered. File will not be run.", vbOKOnly + vbCritical, "Password Error"
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim Path As String
Dim a, b, c, d
 
Path = Command
Label4.Caption = Path
Label2.Caption = Right$(Label4.Caption, (Len(Label4.Caption) - InStrRev(Label4.Caption, "\", -1, vbTextCompare)))

Text3.Text = GetPassword
Text2 = TDecrypt(Text3)
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Text1_Change()
Text1.Text = Replace(Text1.Text, " ", "")
If Text1 = "" Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call Command2_Click
End If
End Sub


