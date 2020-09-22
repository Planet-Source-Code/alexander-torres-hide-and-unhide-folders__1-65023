VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log In"
   ClientHeight    =   1200
   ClientLeft      =   6300
   ClientTop       =   5085
   ClientWidth     =   4350
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4350
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&ENTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1215
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuChange 
      Caption         =   "&Change Password"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTemp As String 'This variable is use to temporarily hold the
'password in the file

Private Sub cmdCancel_Click() 'self explainatory

Unload Me

End

End Sub

Private Sub cmdEnter_Click()

If txtPass.Text = strTemp Then 'if the password the user types matches with the password
                               'in the variable strTemp then continue
   Unload Me
   
   Form1.Show
   
Else

   MsgBox "Password is incorrect. Try again.", vbCritical
   
   txtPass.Text = ""
   txtPass.SetFocus
   
End If

End Sub

Private Sub Form_Load()
Dim nFile As Integer
Dim strPath As String

On Error GoTo rutinerror

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 'center form

nFile = FreeFile
strPath = "c:\lstPathPass.dat" 'password file
                                          
Open strPath For Input As #nFile
                                                               
Do Until EOF(nFile)
                                         
Line Input #nFile, strTemp 'put password to strTemp
                                                                
strTemp = strTemp
                                                                
Loop
                                                               
Close #nFile

Exit Sub

'an error will occure when the program is run for the
'first time becuase the password file doesn't exists
rutinerror:

MsgBox "Password is (password)" + vbNewLine + "You can change your Password anytime.", vbInformation + vbOKOnly

AppUP 'calls sub which fixes the error

Call Form_Load 'reload the form again to retry loading the file

Exit Sub

End Sub

Private Sub mnuChange_Click()
Dim strPass As String

SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc 'sets "*" character for inputbox
 
'the inputbox is to confirm the old password
'if it is correct then form3 will show
'which is the form for changing the password
strPass = InputBox("To continue type in your old Password")

strPass = Trim(strPass)

If strPass = strTemp Then

   Form3.Show
   
   Me.Hide
   
Else

MsgBox "Password not found. Try again.", vbCritical

Exit Sub

End If

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then 'press enter

   Call cmdEnter_Click 'self explainatory
   
Else
   
   Exit Sub
   
End If

End Sub

Private Sub AppUP()                'This sub only is called the first time the app
Dim nFile As Integer               'is run. The reason is when the app first is run
Dim strPath As String              'on the form load event it tries to open the
                                   'password file and an error will
nFile = FreeFile                   'come because there is no file.
strPath = "c:\lstPathPass.dat"     'so AppUP will create that file an it will
                                   'create a temporary password "password"
Open strPath For Output As #nFile  'when this happens a msgbox will appear
                                   'telling the end user the temporary password
Print #nFile, "password"

Close #nFile

Exit Sub

End Sub
