VERSION 5.00
Begin VB.Form frmUID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hellbound UID: By Cenobitez"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   Icon            =   "frmUID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   900
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   -45
      Width           =   1590
      Begin VB.OptionButton optGetUser 
         Caption         =   "Get Username"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   1410
      End
      Begin VB.OptionButton optGetUID 
         Caption         =   "Get UID"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   900
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3465
      Top             =   495
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   1005
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   465
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Red or Dead UID Grabber                                                                        Coded by Cenobitez"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngHandle As Long
Dim strText As String
Dim lngHandleButton As Long
Dim lngHandles As Long

Private Sub cmdStart_Click()
If cmdStart.Caption = "Start" Then
cmdStart.Caption = "Stop"
Timer1.Enabled = True
Else
cmdStart.Caption = "Start"
Timer1.Enabled = False
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub Timer1_Timer()
Dim click As Integer, SplitSplit() As String, TheUsername As String
Dim splitArr() As String, TheText As String, intCount As Integer
    
    lngHandles = FindWindow("#32770", "Group Member Info")
        
        If lngHandles = 0 Then Exit Sub

    lngHandle = FindWindowEx(lngHandles, 0&, "Static", vbNullString)
    strText = GetText(lngHandle)
    lngHandleButton = FindWindowEx(lngHandles, 0&, "Button", vbNullString)
    
    splitArr() = Split(strText, Chr(9))
    intCount = UBound(splitArr())
    SplitSplit() = Split(splitArr(3), Chr(10))
    
    TheUsername = Trim(SplitSplit(0))
    TheText = splitArr(intCount)

    Clipboard.Clear

    If optGetUID.Value = True Then
        Clipboard.SetText TheText
    ElseIf optGetUser.Value = True Then
        Clipboard.SetText TheUsername
    Else
        MsgBox "Please Choose to get UID or Username!", vbCritical, "Oh Man You Messed Up Bad!!"
    End If

    click = SendMessage(lngHandleButton, WM_KEYDOWN, VK_SPACE, 0&) 'sends a left button down
    click = SendMessage(lngHandleButton, WM_KEYUP, VK_SPACE, 0&) 'sends a left button up
End Sub
