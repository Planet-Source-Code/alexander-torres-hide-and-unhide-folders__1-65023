VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Username and Password"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6915
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&REATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtCPass 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Confirm New Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2565
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdCreate_Click() 'creates new password
Dim nFile As Integer
Dim strPath As String

nFile = FreeFile
strPath = "c:\lstPathPass.dat"


txtPass = Trim(txtPass)
txtCPass = Trim(txtCPass)

If txtPass.Text = "" Or txtCPass.Text = "" Then

   MsgBox "Error creating new password. Verify that the information is correct.", vbCritical + vbOKOnly

   
   txtPass.Text = ""
   txtCPass.Text = ""
   txtPass.SetFocus

   Exit Sub

End If

If txtCPass = txtPass Then

Open strPath For Output As #nFile

Print #nFile, txtPass

Close #nFile

MsgBox "Password has been created.", vbInformation

Form1.Show
Unload Me
Unload Form2

Exit Sub

Else

MsgBox "Passwords don't match", vbCritical

Exit Sub

End If

End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 'centers form

End Sub

Private Sub Form_Unload(Cancel As Integer)

Form2.Show

End Sub

Private Sub txtCPass_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call cmdCreate_Click

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call cmdCreate_Click

End Sub

