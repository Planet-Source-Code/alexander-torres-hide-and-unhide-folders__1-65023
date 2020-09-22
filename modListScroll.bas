Attribute VB_Name = "Module1"
Option Explicit

'these are the codes for adding a scrollbar to the listbox

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LB_SETHORIZONTALEXTENT = &H194

Public Sub AddHScroll(List As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To List.ListCount - 1


        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(50))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0

End Sub
