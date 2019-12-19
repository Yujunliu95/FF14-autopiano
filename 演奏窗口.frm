VERSION 5.00
Begin VB.Form 弹琴助手演奏窗口 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "宅爷自制弹琴助手-演奏窗口"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
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
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Shape 块 
      BackColor       =   &H00FFFFC0&
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
   Begin VB.Shape 高2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   7080
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 高3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   7560
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 高5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   8520
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 高6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   9000
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 高7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   9480
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 中2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   3720
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 中3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   4200
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 中5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   5160
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 中6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   5640
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape 中7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   6120
      Top             =   1040
      Width           =   255
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
      Left            =   10320
      TabIndex        =   0
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape 高u 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   9600
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高y 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   9120
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高t 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   8640
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高r 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   8160
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高e 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   7680
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高w 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   7200
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 高q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   6720
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中u 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   6240
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中y 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   5760
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中t 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   5280
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中r 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   4800
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中e 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   4320
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape 中w 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   3840
      Top             =   0
      Width           =   480
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
   Begin VB.Shape 高高q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   10080
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
Attribute VB_Name = "弹琴助手演奏窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''窗口置顶'''''''''''''''''''
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
Do While (fChkControls(弹琴助手演奏窗口, "块", i) = True) '当块存在时
Unload 弹琴助手演奏窗口.块(i) '删除控件
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
宽 = Me.Width / 22
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
中2.Width = 黑宽
中2.Left = 中q.Left + 中q.Width - 黑宽 / 2
中w.Width = 宽
中w.Left = 中q.Left + 中q.Width
中3.Width = 黑宽
中3.Left = 中w.Left + 中w.Width - 黑宽 / 2
中e.Width = 宽
中e.Left = 中w.Left + 中w.Width
中r.Width = 宽
中r.Left = 中e.Left + 中e.Width
中5.Width = 黑宽
中5.Left = 中r.Left + 中r.Width - 黑宽 / 2
中t.Width = 宽
中t.Left = 中r.Left + 中r.Width
中6.Width = 黑宽
中6.Left = 中t.Left + 中t.Width - 黑宽 / 2
中y.Width = 宽
中y.Left = 中t.Left + 中t.Width
中7.Width = 黑宽
中7.Left = 中y.Left + 中y.Width - 黑宽 / 2
中u.Width = 宽
中u.Left = 中y.Left + 中y.Width

高q.Width = 宽
高q.Left = 中u.Left + 中u.Width
高2.Width = 黑宽
高2.Left = 高q.Left + 高q.Width - 黑宽 / 2
高w.Width = 宽
高w.Left = 高q.Left + 高q.Width
高3.Width = 黑宽
高3.Left = 高w.Left + 高w.Width - 黑宽 / 2
高e.Width = 宽
高e.Left = 高w.Left + 高w.Width
高r.Width = 宽
高r.Left = 高e.Left + 高e.Width
高5.Width = 黑宽
高5.Left = 高r.Left + 高r.Width - 黑宽 / 2
高t.Width = 宽
高t.Left = 高r.Left + 高r.Width
高6.Width = 黑宽
高6.Left = 高t.Left + 高t.Width - 黑宽 / 2
高y.Width = 宽
高y.Left = 高t.Left + 高t.Width
高7.Width = 黑宽
高7.Left = 高y.Left + 高y.Width - 黑宽 / 2
高u.Width = 宽
高u.Left = 高y.Left + 高y.Width
高高q.Width = 宽
高高q.Left = 高u.Left + 高u.Width

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
中2.Top = Me.Height - 黑高
中w.Height = Me.Height
中3.Top = Me.Height - 黑高
中e.Height = Me.Height
中r.Height = Me.Height
中5.Top = Me.Height - 黑高
中t.Height = Me.Height
中6.Top = Me.Height - 黑高
中y.Height = Me.Height
中7.Top = Me.Height - 黑高
中u.Height = Me.Height

高q.Height = Me.Height
高2.Top = Me.Height - 黑高
高w.Height = Me.Height
高3.Top = Me.Height - 黑高
高e.Height = Me.Height
高r.Height = Me.Height
高5.Top = Me.Height - 黑高
高t.Height = Me.Height
高6.Top = Me.Height - 黑高
高y.Height = Me.Height
高7.Top = Me.Height - 黑高
高u.Height = Me.Height

高高q.Height = Me.Height

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
