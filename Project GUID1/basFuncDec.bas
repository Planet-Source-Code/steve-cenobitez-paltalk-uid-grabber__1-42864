Attribute VB_Name = "basFuncDec"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_SPACE = &H20
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Public Function GetText(Get_hWnd As Long) As String
Dim retVal As Long, lenTxt As Integer, retText As String
lenTxt = SendMessageLong(Get_hWnd, WM_GETTEXTLENGTH, 0&, 0&)
retText = String(lenTxt + 1, " ")
Call SendMessageByString(Get_hWnd, WM_GETTEXT, lenTxt + 1, retText)
GetText = Left(retText, lenTxt)
End Function
