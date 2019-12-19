VERSION 5.00
Begin VB.Form 弹琴助手演奏窗口mini 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "宅爷自制弹琴助手-演奏窗口"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Shape 块 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   135
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape 低7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   2760
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 低6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   2280
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 低y 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   2400
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   1800
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 低3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   840
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 低2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   360
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 低w 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   480
      Top             =   0
      Width           =   480
   End
   Begin VB.Label 调整大小 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "="
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape 中q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   3360
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低u 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   2880
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低t 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   1920
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低r 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   1440
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低e 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   960
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 低q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "弹琴助手演奏窗口mini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''窗口置顶'''''''''''''''''''
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

''''''''''''''''''窗口可拖动''''''''''''''''''''''
Dim xa As Single, ya As Single
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub
'判断控件是否存在
Function fChkControls(frmObject As Form, strControlsName As String, Optional lngIndex As Long = -1) As Boolean
On Error GoTo Err
    Dim strContrName As String
    If lngIndex >= 0 Then
        strContrName = frmObject.Controls(strControlsName)(lngIndex).Name
    Else
        strContrName = frmObject.Controls(strControlsName).Name
    End If
    fChkControls = True
    Exit Function
Err:
    fChkControls = False
End Function
'鼠标经过窗口
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
调整大小.BackColor = &H80000002
End Sub
''''''''''''''''''''调整大小'''''''''''''''''''''''''
Private Sub 调整大小_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 设置窗口.开始演奏.Caption = "GO" Then
xa = X
ya = Y
'清空
Dim i As Long
i = 1
Do While (fChkControls(弹琴助手演奏窗口mini, "块", i) = True) '当块存在时
Unload 弹琴助手演奏窗口mini.块(i) '删除控件
i = i + 1
Loop
S = ""
End If
End Sub
'进行移动
Private Sub 调整大小_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 设置窗口.开始演奏.Caption = "GO" Then
If Button = 1 Then 调整大小.Move 调整大小.Left + X - xa, 调整大小.Top + Y - ya    '任意移动
'改变窗口大小
Me.Width = 调整大小.Left + 调整大小.Width
Me.Height = 调整大小.Top + 调整大小.Height
调整大小.BackColor = &H8080FF
'内部位置同步改变
'宽
Dim 宽 As Integer, 黑宽 As Integer
宽 = Me.Width / 8
黑宽 = 宽 / 2
低q.Width = 宽
低2.Width = 黑宽
低2.Left = 低q.Left + 低q.Width - 黑宽 / 2
低w.Width = 宽
低w.Left = 低q.Left + 低q.Width
低3.Width = 黑宽
低3.Left = 低w.Left + 低w.Width - 黑宽 / 2
低e.Width = 宽
低e.Left = 低w.Left + 低w.Width
低r.Width = 宽
低r.Left = 低e.Left + 低e.Width
低5.Width = 黑宽
低5.Left = 低r.Left + 低r.Width - 黑宽 / 2
低t.Width = 宽
低t.Left = 低r.Left + 低r.Width
低6.Width = 黑宽
低6.Left = 低t.Left + 低t.Width - 黑宽 / 2
低y.Width = 宽
低y.Left = 低t.Left + 低t.Width
低7.Width = 黑宽
低7.Left = 低y.Left + 低y.Width - 黑宽 / 2
低u.Width = 宽
低u.Left = 低y.Left + 低y.Width
中q.Width = 宽
中q.Left = 低u.Left + 低u.Width
'高
Dim 黑高 As Single
黑高 = 2005
低q.Height = Me.Height
低2.Top = Me.Height - 黑高
低w.Height = Me.Height
低3.Top = Me.Height - 黑高
低e.Height = Me.Height
低r.Height = Me.Height
低5.Top = Me.Height - 黑高
低t.Height = Me.Height
低6.Top = Me.Height - 黑高
低y.Height = Me.Height
低7.Top = Me.Height - 黑高
低u.Height = Me.Height
中q.Height = Me.Height

'初始方块位置
块(0).Top = 低2.Top
If 设置窗口.加载mini.Visible = True Then
Call 设置窗口.加载mini_Click
Else
Call 设置窗口.加载_Click
End If
End If
End Sub
'''''''''''''''''''''加载''''''''''''''''''''''
Private Sub Form_Load()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE     '设置窗口置顶
End Sub
