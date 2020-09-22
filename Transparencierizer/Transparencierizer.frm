VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Apeiron's Handy-Dandy Transparencierizer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   Icon            =   "Transparencierizer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTop 
      Caption         =   "Make Topmost"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pass mouse and keyboard events to whatever is underneath"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   4200
      Picture         =   "Transparencierizer.frx":08CA
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Text            =   "200"
      Top             =   840
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Opaque"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Transparent"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdTranparent 
      Caption         =   "Make Transparent"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      ToolTipText     =   "Click and drag over a window you want to make transparent"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Transparency Level (0-255)"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Window Caption"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Window Handle"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Const MF_CHECKED = &H8&
Const MF_APPEND = &H100&
Const TPM_LEFTALIGN = &H0&
Const MF_DISABLED = &H2&
Const MF_GRAYED = &H1&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const TPM_RETURNCMD = &H100&
Const TPM_RIGHTBUTTON = &H2&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const LWA_ALPHA = 2
Private Const ULW_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_OPAQUE = &H4
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean

Dim oldStyle As Long
Dim Dragging As Boolean
Private Sub ProcessMessages()
    Dim Message As Msg
    Do While Not bCancel
        WaitMessage
        
        If PeekMessage(Message, Me.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            ' This requires a little explanation
            ' The wParam is the number of the registered hotkey
            ' Look in the registerhotkey API call and the corresponding number is in hex (the 2nd argument)
            ' I just chose these numbers they could be different or you can add more hotkeys that way.
            
            If Message.wParam = 49151 Then ' Ctrl-O Opaque
                Option1(1).Value = True
                Transparent False
            ElseIf Message.wParam = 49150 Then ' Ctrl-O Opaque
                Option1(0).Value = True
                Transparent True
            ElseIf Message.wParam = 49149 Then ' ctrl-U to bring this program on top
                FormOnTop Me.hWnd, True
            ElseIf Message.wParam = 49148 Then ' ctrl-D to set this program normal again
                FormOnTop Me.hWnd, False
            End If
            
        End If
        
        DoEvents
    Loop
End Sub
Private Sub Form_Load()
    Dim ret As Long
    Picture1.Picture = Picture2.Picture
    bCancel = False
    ' Thanks to allapi for the hotkey stuff
    ret = RegisterHotKey(Me.hWnd, &HBFFF&, MOD_CONTROL, vbKeyO)
    ret = RegisterHotKey(Me.hWnd, &HBFFE&, MOD_CONTROL, vbKeyT)
    ret = RegisterHotKey(Me.hWnd, &HBFFD&, MOD_CONTROL, vbKeyU)
    ret = RegisterHotKey(Me.hWnd, &HBFFC&, MOD_CONTROL, vbKeyD)
    Me.Show
    ProcessMessages
End Sub
Private Sub Form_Unload(Cancel As Integer)
    bCancel = True
    Call UnregisterHotKey(Me.hWnd, &HBFFF&)
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And Not Dragging Then
        Dragging = True
        Me.MouseIcon = Picture2.Picture
        Me.MousePointer = 99 ' Set to custom.
        ' Erase picture from picCrossHair
        Picture1.Picture = Nothing
    End If
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And Dragging Then
        Dim tPA As POINTAPI
        Dim lhWnd As Long
        Dim sTitle As String * 255
        Dim sClass As String * 255
        Dim tRC As RECT
        Dim sParentTitle As String * 255
        Dim sParentClass As String * 255
        Dim lhWndParent As Long
        Dim sStyle As String
        Dim lRetVal As Long
                
        ' Get cursor position
        GetCursorPos tPA
        ' Get window handle from point
        lhWnd = WindowFromPoint(tPA.x, tPA.y)
        'Cruddy way but I'm in a hurry, Tunnel up to parent window
        lhWndParent = GetParent(lhWnd)
        Dim i As Integer
        For i = 0 To 10
          If lhWndParent = 0 Then
            Exit For
          Else
            lhWnd = lhWndParent
            lhWndParent = GetParent(lhWndParent)
          End If
        Next i
        Text1.Text = lhWnd
        ' Get window caption
        GetWindowText lhWnd, sTitle, 255
        Text2.Text = sTitle
    End If
End Sub

Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
  Dim wFlags As Long
  Dim placement As Long
  Const SWP_NOSIZE = &H1
  Const SWP_NOMOVE = &H2
  Const SWP_NOACTIVATE = &H10
  Const SWP_SHOWWINDOW = &H40
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
  Select Case bTopMost
  Case True
    placement = HWND_TOPMOST
  Case False
    placement = HWND_NOTOPMOST
  End Select
  SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags
End Sub

Public Sub Transparent(t As Boolean)
  
  If Text1.Text = "" Then Exit Sub
  
  If Option1(0).Value Then
      SetWindowLong Val(Text1.Text), GWL_EXSTYLE, oldStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
      FormOnTop Val(Text1.Text), CBool(chkTop.Value)
      SetLayeredWindowAttributes Val(Text1.Text), 0, Val(Text3.Text), LWA_ALPHA
      If Check1.Value = vbUnchecked Then
        SetWindowLong Val(Text1.Text), GWL_EXSTYLE, GetWindowLong(Val(Text1.Text), GWL_EXSTYLE) - WS_EX_TRANSPARENT
      End If
   Else
      FormOnTop Val(Text1.Text), False
      SetWindowLong Val(Text1.Text), GWL_EXSTYLE, GetWindowLong(Val(Text1.Text), GWL_EXSTYLE) And (Not (WS_EX_LAYERED Or WS_EX_TRANSPARENT))
      If Check1.Value = True Then
        SetWindowLong Val(Text1.Text), GWL_EXSTYLE, GetWindowLong(Val(Text1.Text), GWL_EXSTYLE) - WS_EX_TRANSPARENT
      End If
  End If
End Sub

Private Sub cmdTranparent_Click()

  Transparent True
  
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Picture1.MousePointer = vbDefault
  Picture1.Picture = Picture2.Picture
  Dragging = False
  Me.MousePointer = vbDefault
End Sub
