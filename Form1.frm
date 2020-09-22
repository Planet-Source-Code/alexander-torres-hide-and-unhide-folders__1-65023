VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Folder Options                                                                                                       ver-2.1"
   ClientHeight    =   7590
   ClientLeft      =   4650
   ClientTop       =   2265
   ClientWidth     =   9705
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9705
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   3600
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Folder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Width           =   4455
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3150
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hidden Folders"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   2640
      TabIndex        =   4
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdUnhide 
      Caption         =   "&UNHIDE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   2
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&HIDE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HIDE AND UNHIDE FOLDERS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   9720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PATH:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'My name is Alex and I started programming a few months ago
'as a hobby so my knowlege in this is not good but I try
'This program what it does is hides and unhides folders
'I did it so I can hide personal folders that I don't want nobody seeing
'it is also password protected
'the codes for adding scrollbar to listbox and adding password char "*"
'are not mine.
'I know this is not professional or what not but the reason for
'posting it is to recive comments from different programmers
'so I can do better next time and to increase my knowlege in the programming world

'my e-mail: alexltkv@yahoo.com

Dim nFile As Integer  'open file declarations
Dim strPath As String

Dim ShowScroll As Boolean 'variable to stop timer

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub cmdHide_Click()

'makes sure there is a folder to hide
'to prevent errors
If txtPath.Text = "" Or txtPath.Text = "No folder selected to hide" Or txtPath.Text = "No folder selected to unhide" Then
   
   txtPath.Text = "No folder selected to hide"
   
   Exit Sub
   
End If

'trim spaces from path
List1.AddItem Trim(txtPath.Text)

fHide 'calls fHide sub

Save_List 'calls sub

Dir1.Refresh 'refreshes Directory
'this is optional

'if all goes well msgbox will popup
MsgBox "Folder" & " " & txtPath & " " & "is now hidden.", vbInformation + vbOKOnly

End Sub

Private Sub cmdUnhide_Click()

'makes sure there is a folder to hide
'to prevent errors
If txtPath.Text = "" Or txtPath.Text = "No folder selected to unhide" Or txtPath.Text = "No folder selected to hide" Then
   
   txtPath.Text = "No folder selected to unhide"
   
   Exit Sub
   
End If

fUnhide 'calls sub

'if error occures go to rutinerror
On Error GoTo rutinerror

Dim r As Integer 'declare variable for list index

r = List1.ListIndex

'removes path from listbox
'after it's been unhidden
If List1.ListCount = 1 Then txtPath.Text = ""

List1.RemoveItem (r)

Save_List 'calls sub

List1.Selected(0) = True 'selects first path in list

Dir1.Refresh 'optional

'if error then terminate sub
rutinerror: Exit Sub

End Sub

Private Sub Dir1_Change() 'self explainatory
File1 = Dir1
txtPath.Text = File1.Path
End Sub

Private Sub Drive1_Change() 'self explainatory
Dir1 = Drive1
End Sub

Private Sub File1_Click() 'puts the path from file to txtpath
txtPath.Text = File1.Path
End Sub

Private Sub Form_Load()
Dim i As String 'declared variable for
'input from file to text

ShowScroll = False 'this I use to stop the Timer

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 'center form

'from here to the next End Sub is pretty much self explainatory
'basicly what it does is load a file to a listbox
On Error GoTo rutinerror

nFile = FreeFile
strPath = "c:\lstPasth.ini"
i = List1.Text

Open strPath For Input As #nFile

Do Until EOF(nFile)

Line Input #nFile, i

List1.AddItem i

Loop

Close #nFile

Exit Sub

rutinerror: Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
nFile = FreeFile 'makes sure all files all not left open
'when program is closed. this is optional

Close #nFile

End Sub

Private Sub Form_Unload(Cancel As Integer) 'self explainatory

MsgBox "Program made by:  Alexander Torres" + vbNewLine + "Copyright (c) 2006", vbInformation + vbOKOnly

End

End Sub

Private Sub List1_Click()
txtPath = List1 'self explainatory
End Sub

Private Sub Timer1_Timer()

'I use a timer to display a Horizontal scrollbar on the listbox
'it checks every second for an item in the listbox
'if there is one then the scrollbar will show
If List1.ListCount > 0 Then

   Call AddHScroll(List1) 'calls sub from module
   
   ShowScroll = True 'use to stop the Timer
   'no need to have the timer checking evey second if the
   'listbox has a scrollbar already
   
Else

    Exit Sub
    
End If

If ShowScroll = True Then Timer1.Interval = 0

End Sub

Private Sub Save_List() 'codes for saving Path in listbox to file
Dim i As Integer

nFile = FreeFile
strPath = "c:\lstPasth.ini"

Open strPath For Output As #nFile

For i = 0 To List1.ListCount - 1

Print #nFile, List1.List(i)

Next i

Close #nFile

Exit Sub

End Sub

Private Sub fHide() 'hides the folder
Dim FileObject
Dim GFolder

    Set FileObject = CreateObject("Scripting.FileSystemObject")

      Set GFolder = FileObject.GetFolder(txtPath.Text)

GFolder.Attributes = -1

End Sub

Private Sub fUnhide() 'unhides folder
Dim FileObject
Dim GFolder

    Set FileObject = CreateObject("Scripting.FileSystemObject")

      Set GFolder = FileObject.GetFolder(txtPath.Text)

GFolder.Attributes = 0

End Sub
