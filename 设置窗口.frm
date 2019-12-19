VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form 设置窗口 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "宅爷自制弹琴助手"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "设置窗口.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   120
   End
   Begin VB.TextBox 乐谱路径 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox 软件图标 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "设置窗口.frx":1084A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label tips快捷键 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- Space -"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label 降落滑块 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label 速度滑块 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label 透明度滑块 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2580
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape 透明度底框 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label 加载mini 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape tips 
      BorderColor     =   &H80000002&
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label 按键模式 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "展开模式"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label 最顶部块 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label 已完成块数 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   5000
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.Label 降落 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label tips降落 
      BackStyle       =   0  'Transparent
      Caption         =   "降落速度"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label 速度 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label tips速度 
      BackStyle       =   0  'Transparent
      Caption         =   "乐谱速度"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label tips更新页 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "获取更多乐谱"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label 版本号 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "适用于最终幻想XIV  |  版权所有 宅爷  |  版本号 v1.4"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   240
      TabIndex        =   13
      Top             =   2830
      Width           =   4095
   End
   Begin VB.Label 加载 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.Label 开始演奏 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0FF&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label 打开乐谱 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "····"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3480
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label tips乐谱文件 
      BackStyle       =   0  'Transparent
      Caption         =   "乐谱文件"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape 鼠标穿透小圈 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   255
      Left            =   3160
      Shape           =   3  'Circle
      Top             =   605
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.Shape 鼠标穿透大圈 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   255
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   600
      Width           =   255
   End
   Begin VB.Label 鼠标穿透 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label tips鼠标穿透 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "鼠标穿透"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label 透明度 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label 窗口标题 
      BackStyle       =   0  'Transparent
      Caption         =   "宅爷自制弹琴助手"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label 关闭 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Label tips透明度 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   2800
      Width           =   4575
   End
   Begin VB.Shape 速度底框 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Shape 降落底框 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2460
      Width           =   1575
   End
End
Attribute VB_Name = "设置窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'鼠标穿透
Const WS_EX_TRANSPARENT As Long = &H20&
'窗口激活
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'获取窗口
Private Declare Function 取窗口句柄 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'窗口透明API
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'窗口透明常数
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2       '使用此参数，透明度有效，透明颜色无效
Const LWA_COLORKEY = &H1 '使用此参数，透明度无效，透明颜色有效
'读取乐谱变量
Dim S As String
Dim FreeNum As Integer
Dim 速度值 As Integer
Dim 块数 As Integer
'窗口可拖动
Dim xa As Single, ya As Single

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub
'鼠标经过主窗口
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
关闭.BackColor = &H80000002
按键模式.BackColor = &H80000002
打开乐谱.BackColor = &H80000002
加载.BackColor = &H80000002
加载mini.BackColor = &H80000002
开始演奏.BackColor = &HA0A0FF
透明度滑块.BackColor = &H80000002
速度滑块.BackColor = &H80000002
降落滑块.BackColor = &H80000002
鼠标穿透大圈.BorderColor = &H80000002
鼠标穿透小圈.BackColor = &H80000002
鼠标穿透小圈.BorderColor = &H80000002
tips鼠标穿透.ForeColor = &H80000012
tips更新页.ForeColor = &H80000002
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

''''标题可拖动
Private Sub 窗口标题_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub
Private Sub 窗口标题_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub
'模拟按钮的变色设置
Private Sub 关闭_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
关闭.BackColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub 按键模式_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
按键模式.BackColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub 打开乐谱_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
打开乐谱.BackColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub 加载_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
加载.BackColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub 加载mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
加载mini.BackColor = &H8080FF
End Sub

'模拟按钮的变色设置
Private Sub 开始演奏_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
开始演奏.BackColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub 鼠标穿透_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
鼠标穿透大圈.BorderColor = &H8080FF
鼠标穿透小圈.BackColor = &H8080FF
鼠标穿透小圈.BorderColor = &H8080FF
tips鼠标穿透.ForeColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub tips微博_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tips微博.ForeColor = &H8080FF
End Sub
'模拟按钮的变色设置
Private Sub tips更新页_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tips更新页.ForeColor = &H8080FF
End Sub
Private Sub Form_Initialize() '初始化
弹琴助手演奏窗口.Show
End Sub
Private Sub Form_Load() '加载
HooK ''热键
'''''''''''''''''''''窗体透明'''''''''''''''
Dim rtn As Long
弹琴助手演奏窗口.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
窗口句柄 = 取窗口句柄(vbNullString, "宅爷自制弹琴助手-演奏窗口")
rtn = GetWindowLong(窗口句柄, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong 窗口句柄, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes 窗口句柄, RGB(0, 0, 0), 128, LWA_ALPHA '整体窗口透明度
'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '控件不会透明
'RGB(0, 0, 0)参数就是要透明掉的颜色
End Sub

'''''''''''''''''''关闭事件''''''''''''''''''''
Private Sub 关闭_Click() '关闭按钮
     Unload Me   '这时就会调用UNLOAD事件
 End Sub
Public Sub Form_Unload(Cancel As Integer) '退出之前 关闭演奏窗口
Unload 弹琴助手演奏窗口
Unload 弹琴助手演奏窗口mini
UnHooK ''热键关
End Sub


''''''''''''''''''''''''''''''''''''自制滑块拖动'''''''''''''''''''''''''''''''''''''''
Private Sub 透明度滑块_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub 透明度滑块_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 透明度滑块.Move 透明度滑块.Left + X - xa ', 透明度滑块.Top + Y - ya   '左右移动
'判断位置 是否超出
If 透明度滑块.Left < 透明度底框.Left Then 透明度滑块.Left = 透明度底框.Left  '左超出
If 透明度滑块.Left + 透明度滑块.Width > 透明度底框.Left + 透明度底框.Width Then 透明度滑块.Left = 透明度底框.Left + 透明度底框.Width - 透明度滑块.Width
'判断值
Dim 透明度最小值 As Single, 透明度最大值 As Single, 透明度值 As Single
透明度最小值 = 0
透明度最大值 = 255
'滑块位置代表的最大值位置是 底框width-滑块width
透明度值 = (透明度滑块.Left - 透明度底框.Left) / (透明度底框.Width - 透明度滑块.Width) * (透明度最大值 - 透明度最小值)
'使用值进行窗口透明度调整
透明度.Caption = Int((透明度滑块.Left - 透明度底框.Left) / (透明度底框.Width - 透明度滑块.Width) * 100) & "%"
窗口句柄 = 取窗口句柄(vbNullString, "宅爷自制弹琴助手-演奏窗口")
rtn = GetWindowLong(窗口句柄, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong 窗口句柄, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes 窗口句柄, RGB(0, 0, 0), 透明度值, LWA_ALPHA
'按钮颜色选中时变化
透明度滑块.BackColor = &H8080FF
End Sub
'''''''''''''''''''''''''''''''''''自制滑块拖动'''''''''''''''''''''''''''''''''''''''
Private Sub 速度滑块_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub 速度滑块_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 速度滑块.Move 速度滑块.Left + X - xa ', 速度滑块.Top + Y - ya   '左右移动
'判断位置 是否超出
If 速度滑块.Left < 速度底框.Left Then 速度滑块.Left = 速度底框.Left  '左超出
If 速度滑块.Left + 速度滑块.Width > 速度底框.Left + 速度底框.Width Then 速度滑块.Left = 速度底框.Left + 速度底框.Width - 速度滑块.Width
'判断值
Dim 速度最小值 As Single, 速度最大值 As Single
速度最小值 = 1
速度最大值 = 10
'滑块位置代表的最大值位置是 底框width-滑块width
速度值 = Int((速度滑块.Left - 速度底框.Left) / (速度底框.Width - 速度滑块.Width) * (速度最大值 - 速度最小值)) + 1
速度.Caption = 速度值
'按钮颜色选中时变化
速度滑块.BackColor = &H8080FF
End Sub
'''''''''''''''''''''''''''''''''''自制滑块拖动'''''''''''''''''''''''''''''''''''''''
Private Sub 降落滑块_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub 降落滑块_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 降落滑块.Move 降落滑块.Left + X - xa ', 降落滑块.Top + Y - ya   '左右移动
'判断位置 是否超出
If 降落滑块.Left < 降落底框.Left Then 降落滑块.Left = 降落底框.Left  '左超出
If 降落滑块.Left + 降落滑块.Width > 降落底框.Left + 降落底框.Width Then 降落滑块.Left = 降落底框.Left + 降落底框.Width - 降落滑块.Width
'判断值
Dim 降落最小值 As Single, 降落最大值 As Single
降落最小值 = 1
降落最大值 = 10
'滑块位置代表的最大值位置是 底框width-滑块width
降落值 = Int((降落滑块.Left - 降落底框.Left) / (降落底框.Width - 降落滑块.Width) * (降落最大值 - 降落最小值)) + 1
降落.Caption = 降落值
'按钮颜色选中时变化
降落滑块.BackColor = &H8080FF
End Sub
'''''''''''''''''鼠标穿透开关'''''''''''''''''''
Private Sub 鼠标穿透_Click()
窗口句柄 = 取窗口句柄(vbNullString, "宅爷自制弹琴助手-演奏窗口")
If 鼠标穿透小圈.Visible = False Then '开启鼠标穿透
SetWindowLong 窗口句柄, GWL_EXSTYLE, GetWindowLong(窗口句柄, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
鼠标穿透小圈.Visible = True
Else '关闭鼠标穿透
SetWindowLong 窗口句柄, GWL_EXSTYLE, 0
鼠标穿透小圈.Visible = False
''''''未知原因透明度改变修复''''''
透明度最小值 = 0
透明度最大值 = 255
透明度值 = (透明度滑块.Left - 透明度底框.Left) / (透明度底框.Width - 透明度滑块.Width) * (透明度最大值 - 透明度最小值)
rtn = GetWindowLong(窗口句柄, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong 窗口句柄, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes 窗口句柄, RGB(0, 0, 0), 透明度值, LWA_ALPHA '窗口透明度
End If
End Sub
''''''''''''''''''打开乐谱''''''''''''''''''''''
Private Sub 打开乐谱_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Filter = "乐谱文件 (*.dat)|*.dat"
Me.CommonDialog1.ShowOpen
乐谱路径.Text = Me.CommonDialog1.FileName
If 设置窗口.加载mini.Visible = True Then
Call 设置窗口.加载mini_Click
Else
Call 设置窗口.加载_Click
End If
End Sub
Private Sub 乐谱路径_Change()

'List1.List(0) = 乐谱(1)
'总个数 UBound(乐谱) - LBound(乐谱) + 1
End Sub

''''''''''''''''''''''''''加载新乐谱'''''''''''''''''''''''''
Public Sub 加载_Click()
'判断之前是否有控件存在
Dim i As Long
i = 1
Do While (fChkControls(弹琴助手演奏窗口, "块", i) = True) '当块存在时
Unload 弹琴助手演奏窗口.块(i) '删除控件
i = i + 1
Loop
S = ""
''''''''''''开始新乐谱'''''''''''
'判断路径是否存在
If 乐谱路径.Text <> "" Then '存在
    Dim A As String
    '读取乐谱，S为总乐谱
    FreeNum = FreeFile
    Open 乐谱路径.Text For Input As #FreeNum
    Do While Not EOF(FreeNum) '读取一直到文件末尾
        Line Input #FreeNum, A
        S = S + "|" + A 'S用来保存整个文件
        If A满足某个条件 And Not EOF(FreeNum) Then
            Line Input #FreeNum, A '读取下一行的内容
            Exit Do '退出循环
        End If
    Loop
    Close #FreeNum
    '将S赋值给数组乐谱()
    Dim 乐谱() As String
    乐谱() = Split(S, "|")
    ''''''''''创建乐谱块''''''''
    Dim 宽 As Integer
    Dim B As Integer
    If (UBound(乐谱) - LBound(乐谱) + 1) Mod 2 = 1 Then '总数为奇数
    B = 0
    Else
    B = 1
    End If
    For i = 1 To (UBound(乐谱) - LBound(乐谱) + B) '循环
        If i Mod 2 = 1 Then '奇数
            Load 弹琴助手演奏窗口.块((i + 1) / 2) '创建一个新的块
            弹琴助手演奏窗口.块((i + 1) / 2).Visible = True  '显示
            弹琴助手演奏窗口.块((i + 1) / 2).ZOrder 0  '置顶
            'left
            Dim 左 As Integer
            Select Case 乐谱(i)
            Case "q"
                左 = 弹琴助手演奏窗口.中q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "2"
                左 = 弹琴助手演奏窗口.中2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "w"
                左 = 弹琴助手演奏窗口.中w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "3"
                左 = 弹琴助手演奏窗口.中3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "e"
                左 = 弹琴助手演奏窗口.中e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "r"
                左 = 弹琴助手演奏窗口.中r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "5"
                左 = 弹琴助手演奏窗口.中5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "t"
                左 = 弹琴助手演奏窗口.中t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "6"
                左 = 弹琴助手演奏窗口.中6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "y"
                左 = 弹琴助手演奏窗口.中y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "7"
                左 = 弹琴助手演奏窗口.中7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "u"
                左 = 弹琴助手演奏窗口.中u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "z"
                左 = 弹琴助手演奏窗口.低q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "x"
                左 = 弹琴助手演奏窗口.低2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "c"
                左 = 弹琴助手演奏窗口.低w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "v"
                左 = 弹琴助手演奏窗口.低3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "b"
                左 = 弹琴助手演奏窗口.低e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "n"
                左 = 弹琴助手演奏窗口.低r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "m"
                左 = 弹琴助手演奏窗口.低5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case ","
                左 = 弹琴助手演奏窗口.低t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "."
                左 = 弹琴助手演奏窗口.低6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "/"
                左 = 弹琴助手演奏窗口.低y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "["
                左 = 弹琴助手演奏窗口.低7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "]"
                左 = 弹琴助手演奏窗口.低u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "a"
                左 = 弹琴助手演奏窗口.高q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "s"
                左 = 弹琴助手演奏窗口.高2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "d"
                左 = 弹琴助手演奏窗口.高w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "f"
                左 = 弹琴助手演奏窗口.高3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "g"
                左 = 弹琴助手演奏窗口.高e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "h"
                左 = 弹琴助手演奏窗口.高r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "j"
                左 = 弹琴助手演奏窗口.高5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "k"
                左 = 弹琴助手演奏窗口.高t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "l"
                左 = 弹琴助手演奏窗口.高6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case ";"
                左 = 弹琴助手演奏窗口.高y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "'"
                左 = 弹琴助手演奏窗口.高7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "-"
                左 = 弹琴助手演奏窗口.高u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "="
                左 = 弹琴助手演奏窗口.高高q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case Else
                左 = 弹琴助手演奏窗口.低q.Left - 弹琴助手演奏窗口.低q.Width - 10  '向左超出
                宽 = 弹琴助手演奏窗口.低2.Width
            End Select
            弹琴助手演奏窗口.块((i + 1) / 2).Width = 宽 '宽度调整
            弹琴助手演奏窗口.块((i + 1) / 2).Left = 左 '左坐标调整
        Else '偶数 长度调整
            If IsNumeric(乐谱(i)) Then
                弹琴助手演奏窗口.块(i / 2).Height = 100 * 乐谱(i) * 速度.Caption * 降落.Caption
            Else
                弹琴助手演奏窗口.块(i / 2).Height = 1
            End If
            弹琴助手演奏窗口.块(i / 2).Top = 弹琴助手演奏窗口.块(i / 2 - 1).Top - 弹琴助手演奏窗口.块(i / 2).Height
            '顶部超出
            If 弹琴助手演奏窗口.块(i / 2).Top < 0 Then
                最顶部块.Caption = i
                Exit For
            End If
        End If
        
    Next
    开始演奏.Caption = "GO"
End If

End Sub
Private Sub tips更新页_Click()
Set ws = CreateObject("wscript.shell")
ws.run "explorer https://bbs.nga.cn/read.php?tid=17450001"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''开始演奏'''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub 开始演奏_Click()
If 按键模式.Caption = "展开模式" Then
If fChkControls(弹琴助手演奏窗口, "块", 1) = True Then
    If 开始演奏.Caption = "GO" Then
        Timer1.Enabled = True '计时器启用
        开始演奏.Caption = "暂停"
        已完成块数.Caption = 1
    Else
        Timer1.Enabled = False '计时器不启用
        开始演奏.Caption = "GO"
    End If
    Timer1.Interval = 50 '100毫秒启动一次
End If

Else '8键模式
If fChkControls(弹琴助手演奏窗口mini, "块", 1) = True Then
    If 开始演奏.Caption = "GO" Then
        Timer2.Enabled = True '计时器启用
        开始演奏.Caption = "暂停"
        已完成块数.Caption = 1
    Else
        Timer2.Enabled = False '计时器不启用
        开始演奏.Caption = "GO"
    End If
    Timer2.Interval = 50 '100毫秒启动一次
End If
End If

End Sub
Private Sub Timer1_Timer()
''''顶部溢出新建块
'计算已有的块总数
Dim 块总数 As Long
块总数 = 1
Do While (fChkControls(弹琴助手演奏窗口, "块", 块总数) = True) '当块存在时 得到块总数
块总数 = 块总数 + 1
Loop
块总数 = 块总数 - 1
If 弹琴助手演奏窗口.块(块总数).Top > 0 And 开始演奏.Caption <> "完毕" Then '''有新空余
    '将S赋值给数组乐谱()
    Dim 乐谱() As String
    乐谱() = Split(S, "|")
    ''''''''''创建乐谱块''''''''
    Dim 宽 As Integer
    Dim B As Integer
    If (UBound(乐谱) - LBound(乐谱) + 1) Mod 2 = 1 Then '总数为奇数
    B = 0
    Else
    B = 1
    End If

    Dim i As Integer
    i = 最顶部块.Caption + 1
    ''''读取结束

    'MsgBox 块总数
            Load 弹琴助手演奏窗口.块((i + 1) / 2) '创建一个新的块
            弹琴助手演奏窗口.块((i + 1) / 2).Visible = True  '显示
            弹琴助手演奏窗口.块((i + 1) / 2).ZOrder 0  '置顶
            'left
            Dim 左 As Integer
            Select Case 乐谱(i)
            Case "q"
                左 = 弹琴助手演奏窗口.中q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "2"
                左 = 弹琴助手演奏窗口.中2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "w"
                左 = 弹琴助手演奏窗口.中w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "3"
                左 = 弹琴助手演奏窗口.中3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "e"
                左 = 弹琴助手演奏窗口.中e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "r"
                左 = 弹琴助手演奏窗口.中r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "5"
                左 = 弹琴助手演奏窗口.中5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "t"
                左 = 弹琴助手演奏窗口.中t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "6"
                左 = 弹琴助手演奏窗口.中6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "y"
                左 = 弹琴助手演奏窗口.中y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "7"
                左 = 弹琴助手演奏窗口.中7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "u"
                左 = 弹琴助手演奏窗口.中u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "z"
                左 = 弹琴助手演奏窗口.低q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "x"
                左 = 弹琴助手演奏窗口.低2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "c"
                左 = 弹琴助手演奏窗口.低w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "v"
                左 = 弹琴助手演奏窗口.低3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "b"
                左 = 弹琴助手演奏窗口.低e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "n"
                左 = 弹琴助手演奏窗口.低r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "m"
                左 = 弹琴助手演奏窗口.低5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case ","
                左 = 弹琴助手演奏窗口.低t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "."
                左 = 弹琴助手演奏窗口.低6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "/"
                左 = 弹琴助手演奏窗口.低y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "["
                左 = 弹琴助手演奏窗口.低7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "]"
                左 = 弹琴助手演奏窗口.低u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "a"
                左 = 弹琴助手演奏窗口.高q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "s"
                左 = 弹琴助手演奏窗口.高2.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "d"
                左 = 弹琴助手演奏窗口.高w.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "f"
                左 = 弹琴助手演奏窗口.高3.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "g"
                左 = 弹琴助手演奏窗口.高e.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "h"
                左 = 弹琴助手演奏窗口.高r.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "j"
                左 = 弹琴助手演奏窗口.高5.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "k"
                左 = 弹琴助手演奏窗口.高t.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "l"
                左 = 弹琴助手演奏窗口.高6.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case ";"
                左 = 弹琴助手演奏窗口.高y.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "'"
                左 = 弹琴助手演奏窗口.高7.Left
                宽 = 弹琴助手演奏窗口.低2.Width
            Case "-"
                左 = 弹琴助手演奏窗口.高u.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case "="
                左 = 弹琴助手演奏窗口.高高q.Left
                宽 = 弹琴助手演奏窗口.低q.Width
            Case Else
                左 = 弹琴助手演奏窗口.低q.Left - 弹琴助手演奏窗口.低q.Width - 10  '向左超出
                宽 = 弹琴助手演奏窗口.低2.Width
            End Select
            弹琴助手演奏窗口.块((i + 1) / 2).Width = 宽 '宽度调整
            弹琴助手演奏窗口.块((i + 1) / 2).Left = 左 '左坐标调整
            i = i + 1
            If IsNumeric(乐谱(i)) Then
                弹琴助手演奏窗口.块(i / 2).Height = 100 * 乐谱(i) * 速度.Caption * 降落.Caption
            Else
                弹琴助手演奏窗口.块(i / 2).Height = 1
            End If
            弹琴助手演奏窗口.块(i / 2).Top = 弹琴助手演奏窗口.块(i / 2 - 1).Top - 弹琴助手演奏窗口.块(i / 2).Height
        'End If
        最顶部块.Caption = Int(最顶部块.Caption) + 2
        块总数 = 块总数 + 1
        If Int(最顶部块.Caption) + 1 > UBound(乐谱) - LBound(乐谱) + B Then 开始演奏.Caption = "完毕"
    'Next
End If
''''判断是否完成演奏
If 弹琴助手演奏窗口.块(块总数).Top > 弹琴助手演奏窗口.Height Then
Timer1.Enabled = False  '演奏完毕
开始演奏.Caption = "完毕"
End If
''''下降
For i = 已完成块数.Caption To 块总数
弹琴助手演奏窗口.块(i).Top = 弹琴助手演奏窗口.块(i).Top + 50 * 降落.Caption
If 弹琴助手演奏窗口.块(i).Top > 弹琴助手演奏窗口.Height Then 已完成块数.Caption = i
Next

End Sub

Private Sub 按键模式_Click()
If 开始演奏.Caption = "暂停" Then Call 开始演奏_Click
If 按键模式.Caption = "展开模式" Then
    Unload 弹琴助手演奏窗口
    弹琴助手演奏窗口mini.Show
    '''''''''''''''''''''窗体透明'''''''''''''''
    Dim rtn As Long
    弹琴助手演奏窗口mini.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
    窗口句柄 = 取窗口句柄(vbNullString, "宅爷自制弹琴助手-演奏窗口")
    rtn = GetWindowLong(窗口句柄, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong 窗口句柄, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes 窗口句柄, RGB(0, 0, 0), 128, LWA_ALPHA '整体窗口透明度
    'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '控件不会透明
    'RGB(0, 0, 0)参数就是要透明掉的颜色
    加载mini.Visible = True
    按键模式.Caption = "不展开"
    Call 设置窗口.加载mini_Click
Else
    Unload 弹琴助手演奏窗口mini
    弹琴助手演奏窗口.Show
    '''''''''''''''''''''窗体透明'''''''''''''''
    'Dim rtn As Long
    弹琴助手演奏窗口.BackColor = RGB(0, 0, 0) '设置一下窗口的颜色
    窗口句柄 = 取窗口句柄(vbNullString, "宅爷自制弹琴助手-演奏窗口")
    rtn = GetWindowLong(窗口句柄, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong 窗口句柄, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes 窗口句柄, RGB(0, 0, 0), 128, LWA_ALPHA '整体窗口透明度
    'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '控件不会透明
    'RGB(0, 0, 0)参数就是要透明掉的颜色
    加载mini.Visible = False
    按键模式.Caption = "展开模式"
    Call 设置窗口.加载_Click
End If
End Sub
''''''''''''''''''''''''''加载新乐谱'''''''''''''''''''''''''
Public Sub 加载mini_Click()
'判断之前是否有控件存在
Dim i As Long
i = 1
Do While (fChkControls(弹琴助手演奏窗口mini, "块", i) = True) '当块存在时
Unload 弹琴助手演奏窗口mini.块(i) '删除控件
i = i + 1
Loop
S = ""
''''''''''''开始新乐谱'''''''''''
'判断路径是否存在
If 乐谱路径.Text <> "" Then '存在
    Dim A As String
    '读取乐谱，S为总乐谱
    FreeNum = FreeFile
    Open 乐谱路径.Text For Input As #FreeNum
    Do While Not EOF(FreeNum) '读取一直到文件末尾
        Line Input #FreeNum, A
        S = S + "|" + A 'S用来保存整个文件
        If A满足某个条件 And Not EOF(FreeNum) Then
            Line Input #FreeNum, A '读取下一行的内容
            Exit Do '退出循环
        End If
    Loop
    Close #FreeNum
    '将S赋值给数组乐谱()
    Dim 乐谱() As String
    乐谱() = Split(S, "|")
    ''''''''''创建乐谱块''''''''
    Dim 宽 As Integer
    Dim B As Integer
    If (UBound(乐谱) - LBound(乐谱) + 1) Mod 2 = 1 Then '总数为奇数
    B = 0
    Else
    B = 1
    End If
    For i = 1 To (UBound(乐谱) - LBound(乐谱) + B) '循环
        If i Mod 2 = 1 Then '奇数
            Load 弹琴助手演奏窗口mini.块((i + 1) / 2) '创建一个新的块
            弹琴助手演奏窗口mini.块((i + 1) / 2).Visible = True  '显示
            弹琴助手演奏窗口mini.块((i + 1) / 2).ZOrder 0  '置顶
            'left
            Dim 左 As Integer
            Select Case 乐谱(i)
            Case "q"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "2"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "w"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "3"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "e"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "r"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "5"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "t"
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "6"
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "y"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "7"
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "u"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "z"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "x"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "c"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "v"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "b"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "n"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "m"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case ","
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "."
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "/"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "["
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "]"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "a"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "s"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "d"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "f"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "g"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "h"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "j"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "k"
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "l"
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case ";"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "'"
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "-"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "="
                左 = 弹琴助手演奏窗口mini.中q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case Else
                左 = 弹琴助手演奏窗口mini.低q.Left - 弹琴助手演奏窗口mini.低q.Width - 10  '向左超出
                宽 = 弹琴助手演奏窗口mini.低2.Width
            End Select
            弹琴助手演奏窗口mini.块((i + 1) / 2).Width = 宽 '宽度调整
            弹琴助手演奏窗口mini.块((i + 1) / 2).Left = 左 '左坐标调整
        Else '偶数 长度调整
            If IsNumeric(乐谱(i)) Then
                弹琴助手演奏窗口mini.块(i / 2).Height = 100 * 乐谱(i) * 速度.Caption * 降落.Caption
            Else
                弹琴助手演奏窗口mini.块(i / 2).Height = 1
            End If
            弹琴助手演奏窗口mini.块(i / 2).Top = 弹琴助手演奏窗口mini.块(i / 2 - 1).Top - 弹琴助手演奏窗口mini.块(i / 2).Height
            '顶部超出
            If 弹琴助手演奏窗口mini.块(i / 2).Top < 0 Then
                最顶部块.Caption = i
                Exit For
            End If
        End If
        
    Next
    开始演奏.Caption = "GO"
End If
End Sub

Private Sub Timer2_Timer()
''''顶部溢出新建块
'计算已有的块总数
Dim 块总数 As Long
块总数 = 1
Do While (fChkControls(弹琴助手演奏窗口mini, "块", 块总数) = True) '当块存在时 得到块总数
块总数 = 块总数 + 1
Loop
块总数 = 块总数 - 1
If 弹琴助手演奏窗口mini.块(块总数).Top > 0 And 开始演奏.Caption <> "完毕" Then '''有新空余
    '将S赋值给数组乐谱()
    Dim 乐谱() As String
    乐谱() = Split(S, "|")
    ''''''''''创建乐谱块''''''''
    Dim 宽 As Integer
    Dim B As Integer
    If (UBound(乐谱) - LBound(乐谱) + 1) Mod 2 = 1 Then '总数为奇数
    B = 0
    Else
    B = 1
    End If

    Dim i As Integer
    i = 最顶部块.Caption + 1
    ''''读取结束

    'MsgBox 块总数
            Load 弹琴助手演奏窗口mini.块((i + 1) / 2) '创建一个新的块
            弹琴助手演奏窗口mini.块((i + 1) / 2).Visible = True  '显示
            弹琴助手演奏窗口mini.块((i + 1) / 2).ZOrder 0  '置顶
            'left
            Dim 左 As Integer
            Select Case 乐谱(i)
            Case "q"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "2"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "w"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "3"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "e"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "r"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "5"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "t"
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "6"
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "y"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "7"
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
            Case "u"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
            Case "z"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "x"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "c"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "v"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "b"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "n"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "m"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case ","
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "."
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "/"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "["
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "]"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HFFFFC0
            Case "a"
                左 = 弹琴助手演奏窗口mini.低q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "s"
                左 = 弹琴助手演奏窗口mini.低2.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "d"
                左 = 弹琴助手演奏窗口mini.低w.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "f"
                左 = 弹琴助手演奏窗口mini.低3.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "g"
                左 = 弹琴助手演奏窗口mini.低e.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "h"
                左 = 弹琴助手演奏窗口mini.低r.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "j"
                左 = 弹琴助手演奏窗口mini.低5.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "k"
                左 = 弹琴助手演奏窗口mini.低t.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "l"
                左 = 弹琴助手演奏窗口mini.低6.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case ";"
                左 = 弹琴助手演奏窗口mini.低y.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "'"
                左 = 弹琴助手演奏窗口mini.低7.Left
                宽 = 弹琴助手演奏窗口mini.低2.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "-"
                左 = 弹琴助手演奏窗口mini.低u.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case "="
                左 = 弹琴助手演奏窗口mini.中q.Left
                宽 = 弹琴助手演奏窗口mini.低q.Width
                弹琴助手演奏窗口mini.块((i + 1) / 2).BackColor = &HC0C0FF
            Case Else
                左 = 弹琴助手演奏窗口mini.低q.Left - 弹琴助手演奏窗口mini.低q.Width - 10  '向左超出
                宽 = 弹琴助手演奏窗口mini.低2.Width
            End Select
            弹琴助手演奏窗口mini.块((i + 1) / 2).Width = 宽 '宽度调整
            弹琴助手演奏窗口mini.块((i + 1) / 2).Left = 左 '左坐标调整
            i = i + 1
            If IsNumeric(乐谱(i)) Then
                弹琴助手演奏窗口mini.块(i / 2).Height = 100 * 乐谱(i) * 速度.Caption * 降落.Caption
            Else
                弹琴助手演奏窗口mini.块(i / 2).Height = 1
            End If
            弹琴助手演奏窗口mini.块(i / 2).Top = 弹琴助手演奏窗口mini.块(i / 2 - 1).Top - 弹琴助手演奏窗口mini.块(i / 2).Height
        'End If
        最顶部块.Caption = Int(最顶部块.Caption) + 2
        块总数 = 块总数 + 1
        If Int(最顶部块.Caption) + 1 > UBound(乐谱) - LBound(乐谱) + B Then 开始演奏.Caption = "完毕"
    'Next
End If
''''判断是否完成演奏
If 弹琴助手演奏窗口mini.块(块总数).Top > 弹琴助手演奏窗口mini.Height Then
Timer1.Enabled = False  '演奏完毕
开始演奏.Caption = "完毕"
End If
''''下降
For i = 已完成块数.Caption To 块总数
弹琴助手演奏窗口mini.块(i).Top = 弹琴助手演奏窗口mini.块(i).Top + 50 * 降落.Caption
If 弹琴助手演奏窗口mini.块(i).Top > 弹琴助手演奏窗口mini.Height Then 已完成块数.Caption = i
Next

End Sub
