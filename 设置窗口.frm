VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ���ô��� 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "լү���Ƶ�������"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "���ô���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   StartUpPosition =   2  '��Ļ����
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
   Begin VB.TextBox ����·�� 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.PictureBox ���ͼ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "���ô���.frx":1084A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label tips��ݼ� 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- Space -"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ���们�� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label �ٶȻ��� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ͸���Ȼ��� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "||||"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Shape ͸���ȵ׿� 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label ����mini 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ����ģʽ 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "չ��ģʽ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label ����� 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label ����ɿ��� 
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
   Begin VB.Label ���� 
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
   Begin VB.Label tips���� 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ٶ�"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label �ٶ� 
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
   Begin VB.Label tips�ٶ� 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ٶ�"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label tips����ҳ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ȡ��������"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label �汾�� 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������ջ���XIV  |  ��Ȩ���� լү  |  �汾�� v1.4"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ���� 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ��ʼ���� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0FF&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ������ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label tips�����ļ� 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ļ�"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape ��괩͸СȦ 
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
   Begin VB.Shape ��괩͸��Ȧ 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   255
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   600
      Width           =   255
   End
   Begin VB.Label ��괩͸ 
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
   Begin VB.Label tips��괩͸ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��괩͸"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label ͸���� 
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
   Begin VB.Label ���ڱ��� 
      BackStyle       =   0  'Transparent
      Caption         =   "լү���Ƶ�������"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label �ر� 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label tips͸���� 
      BackStyle       =   0  'Transparent
      Caption         =   "͸����"
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
   Begin VB.Shape �ٶȵ׿� 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Shape ����׿� 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   135
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2460
      Width           =   1575
   End
End
Attribute VB_Name = "���ô���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��괩͸
Const WS_EX_TRANSPARENT As Long = &H20&
'���ڼ���
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'��ȡ����
Private Declare Function ȡ���ھ�� Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'����͸��API
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'����͸������
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2       'ʹ�ô˲�����͸������Ч��͸����ɫ��Ч
Const LWA_COLORKEY = &H1 'ʹ�ô˲�����͸������Ч��͸����ɫ��Ч
'��ȡ���ױ���
Dim S As String
Dim FreeNum As Integer
Dim �ٶ�ֵ As Integer
Dim ���� As Integer
'���ڿ��϶�
Dim xa As Single, ya As Single

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub
'��꾭��������
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
�ر�.BackColor = &H80000002
����ģʽ.BackColor = &H80000002
������.BackColor = &H80000002
����.BackColor = &H80000002
����mini.BackColor = &H80000002
��ʼ����.BackColor = &HA0A0FF
͸���Ȼ���.BackColor = &H80000002
�ٶȻ���.BackColor = &H80000002
���们��.BackColor = &H80000002
��괩͸��Ȧ.BorderColor = &H80000002
��괩͸СȦ.BackColor = &H80000002
��괩͸СȦ.BorderColor = &H80000002
tips��괩͸.ForeColor = &H80000012
tips����ҳ.ForeColor = &H80000002
End Sub
'�жϿؼ��Ƿ����
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

''''������϶�
Private Sub ���ڱ���_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub
Private Sub ���ڱ���_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub �ر�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
�ر�.BackColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub ����ģʽ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����ģʽ.BackColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub ������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
������.BackColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub ����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����.BackColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub ����mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����mini.BackColor = &H8080FF
End Sub

'ģ�ⰴť�ı�ɫ����
Private Sub ��ʼ����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
��ʼ����.BackColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub ��괩͸_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
��괩͸��Ȧ.BorderColor = &H8080FF
��괩͸СȦ.BackColor = &H8080FF
��괩͸СȦ.BorderColor = &H8080FF
tips��괩͸.ForeColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub tips΢��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tips΢��.ForeColor = &H8080FF
End Sub
'ģ�ⰴť�ı�ɫ����
Private Sub tips����ҳ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tips����ҳ.ForeColor = &H8080FF
End Sub
Private Sub Form_Initialize() '��ʼ��
�����������ര��.Show
End Sub
Private Sub Form_Load() '����
HooK ''�ȼ�
'''''''''''''''''''''����͸��'''''''''''''''
Dim rtn As Long
�����������ര��.BackColor = RGB(0, 0, 0) '����һ�´��ڵ���ɫ
���ھ�� = ȡ���ھ��(vbNullString, "լү���Ƶ�������-���ര��")
rtn = GetWindowLong(���ھ��, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong ���ھ��, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes ���ھ��, RGB(0, 0, 0), 128, LWA_ALPHA '���崰��͸����
'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '�ؼ�����͸��
'RGB(0, 0, 0)��������Ҫ͸��������ɫ
End Sub

'''''''''''''''''''�ر��¼�''''''''''''''''''''
Private Sub �ر�_Click() '�رհ�ť
     Unload Me   '��ʱ�ͻ����UNLOAD�¼�
 End Sub
Public Sub Form_Unload(Cancel As Integer) '�˳�֮ǰ �ر����ര��
Unload �����������ര��
Unload �����������ര��mini
UnHooK ''�ȼ���
End Sub


''''''''''''''''''''''''''''''''''''���ƻ����϶�'''''''''''''''''''''''''''''''''''''''
Private Sub ͸���Ȼ���_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub ͸���Ȼ���_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ͸���Ȼ���.Move ͸���Ȼ���.Left + X - xa ', ͸���Ȼ���.Top + Y - ya   '�����ƶ�
'�ж�λ�� �Ƿ񳬳�
If ͸���Ȼ���.Left < ͸���ȵ׿�.Left Then ͸���Ȼ���.Left = ͸���ȵ׿�.Left  '�󳬳�
If ͸���Ȼ���.Left + ͸���Ȼ���.Width > ͸���ȵ׿�.Left + ͸���ȵ׿�.Width Then ͸���Ȼ���.Left = ͸���ȵ׿�.Left + ͸���ȵ׿�.Width - ͸���Ȼ���.Width
'�ж�ֵ
Dim ͸������Сֵ As Single, ͸�������ֵ As Single, ͸����ֵ As Single
͸������Сֵ = 0
͸�������ֵ = 255
'����λ�ô�������ֵλ���� �׿�width-����width
͸����ֵ = (͸���Ȼ���.Left - ͸���ȵ׿�.Left) / (͸���ȵ׿�.Width - ͸���Ȼ���.Width) * (͸�������ֵ - ͸������Сֵ)
'ʹ��ֵ���д���͸���ȵ���
͸����.Caption = Int((͸���Ȼ���.Left - ͸���ȵ׿�.Left) / (͸���ȵ׿�.Width - ͸���Ȼ���.Width) * 100) & "%"
���ھ�� = ȡ���ھ��(vbNullString, "լү���Ƶ�������-���ര��")
rtn = GetWindowLong(���ھ��, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong ���ھ��, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes ���ھ��, RGB(0, 0, 0), ͸����ֵ, LWA_ALPHA
'��ť��ɫѡ��ʱ�仯
͸���Ȼ���.BackColor = &H8080FF
End Sub
'''''''''''''''''''''''''''''''''''���ƻ����϶�'''''''''''''''''''''''''''''''''''''''
Private Sub �ٶȻ���_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub �ٶȻ���_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then �ٶȻ���.Move �ٶȻ���.Left + X - xa ', �ٶȻ���.Top + Y - ya   '�����ƶ�
'�ж�λ�� �Ƿ񳬳�
If �ٶȻ���.Left < �ٶȵ׿�.Left Then �ٶȻ���.Left = �ٶȵ׿�.Left  '�󳬳�
If �ٶȻ���.Left + �ٶȻ���.Width > �ٶȵ׿�.Left + �ٶȵ׿�.Width Then �ٶȻ���.Left = �ٶȵ׿�.Left + �ٶȵ׿�.Width - �ٶȻ���.Width
'�ж�ֵ
Dim �ٶ���Сֵ As Single, �ٶ����ֵ As Single
�ٶ���Сֵ = 1
�ٶ����ֵ = 10
'����λ�ô�������ֵλ���� �׿�width-����width
�ٶ�ֵ = Int((�ٶȻ���.Left - �ٶȵ׿�.Left) / (�ٶȵ׿�.Width - �ٶȻ���.Width) * (�ٶ����ֵ - �ٶ���Сֵ)) + 1
�ٶ�.Caption = �ٶ�ֵ
'��ť��ɫѡ��ʱ�仯
�ٶȻ���.BackColor = &H8080FF
End Sub
'''''''''''''''''''''''''''''''''''���ƻ����϶�'''''''''''''''''''''''''''''''''''''''
Private Sub ���们��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
'ya = Y
End Sub
Private Sub ���们��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ���们��.Move ���们��.Left + X - xa ', ���们��.Top + Y - ya   '�����ƶ�
'�ж�λ�� �Ƿ񳬳�
If ���们��.Left < ����׿�.Left Then ���们��.Left = ����׿�.Left  '�󳬳�
If ���们��.Left + ���们��.Width > ����׿�.Left + ����׿�.Width Then ���们��.Left = ����׿�.Left + ����׿�.Width - ���们��.Width
'�ж�ֵ
Dim ������Сֵ As Single, �������ֵ As Single
������Сֵ = 1
�������ֵ = 10
'����λ�ô�������ֵλ���� �׿�width-����width
����ֵ = Int((���们��.Left - ����׿�.Left) / (����׿�.Width - ���们��.Width) * (�������ֵ - ������Сֵ)) + 1
����.Caption = ����ֵ
'��ť��ɫѡ��ʱ�仯
���们��.BackColor = &H8080FF
End Sub
'''''''''''''''''��괩͸����'''''''''''''''''''
Private Sub ��괩͸_Click()
���ھ�� = ȡ���ھ��(vbNullString, "լү���Ƶ�������-���ര��")
If ��괩͸СȦ.Visible = False Then '������괩͸
SetWindowLong ���ھ��, GWL_EXSTYLE, GetWindowLong(���ھ��, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
��괩͸СȦ.Visible = True
Else '�ر���괩͸
SetWindowLong ���ھ��, GWL_EXSTYLE, 0
��괩͸СȦ.Visible = False
''''''δ֪ԭ��͸���ȸı��޸�''''''
͸������Сֵ = 0
͸�������ֵ = 255
͸����ֵ = (͸���Ȼ���.Left - ͸���ȵ׿�.Left) / (͸���ȵ׿�.Width - ͸���Ȼ���.Width) * (͸�������ֵ - ͸������Сֵ)
rtn = GetWindowLong(���ھ��, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong ���ھ��, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes ���ھ��, RGB(0, 0, 0), ͸����ֵ, LWA_ALPHA '����͸����
End If
End Sub
''''''''''''''''''������''''''''''''''''''''''
Private Sub ������_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Filter = "�����ļ� (*.dat)|*.dat"
Me.CommonDialog1.ShowOpen
����·��.Text = Me.CommonDialog1.FileName
If ���ô���.����mini.Visible = True Then
Call ���ô���.����mini_Click
Else
Call ���ô���.����_Click
End If
End Sub
Private Sub ����·��_Change()

'List1.List(0) = ����(1)
'�ܸ��� UBound(����) - LBound(����) + 1
End Sub

''''''''''''''''''''''''''����������'''''''''''''''''''''''''
Public Sub ����_Click()
'�ж�֮ǰ�Ƿ��пؼ�����
Dim i As Long
i = 1
Do While (fChkControls(�����������ര��, "��", i) = True) '�������ʱ
Unload �����������ര��.��(i) 'ɾ���ؼ�
i = i + 1
Loop
S = ""
''''''''''''��ʼ������'''''''''''
'�ж�·���Ƿ����
If ����·��.Text <> "" Then '����
    Dim A As String
    '��ȡ���ף�SΪ������
    FreeNum = FreeFile
    Open ����·��.Text For Input As #FreeNum
    Do While Not EOF(FreeNum) '��ȡһֱ���ļ�ĩβ
        Line Input #FreeNum, A
        S = S + "|" + A 'S�������������ļ�
        If A����ĳ������ And Not EOF(FreeNum) Then
            Line Input #FreeNum, A '��ȡ��һ�е�����
            Exit Do '�˳�ѭ��
        End If
    Loop
    Close #FreeNum
    '��S��ֵ����������()
    Dim ����() As String
    ����() = Split(S, "|")
    ''''''''''�������׿�''''''''
    Dim �� As Integer
    Dim B As Integer
    If (UBound(����) - LBound(����) + 1) Mod 2 = 1 Then '����Ϊ����
    B = 0
    Else
    B = 1
    End If
    For i = 1 To (UBound(����) - LBound(����) + B) 'ѭ��
        If i Mod 2 = 1 Then '����
            Load �����������ര��.��((i + 1) / 2) '����һ���µĿ�
            �����������ര��.��((i + 1) / 2).Visible = True  '��ʾ
            �����������ര��.��((i + 1) / 2).ZOrder 0  '�ö�
            'left
            Dim �� As Integer
            Select Case ����(i)
            Case "q"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "2"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "w"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "3"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "e"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "r"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "5"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case "t"
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "6"
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case "y"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "7"
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "u"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "z"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "x"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "c"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "v"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "b"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "n"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "m"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case ","
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "."
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case "/"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "["
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "]"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "a"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "s"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "d"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "f"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "g"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "h"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "j"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case "k"
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "l"
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case ";"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "'"
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "-"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "="
                �� = �����������ര��.�߸�q.Left
                �� = �����������ര��.��q.Width
            Case Else
                �� = �����������ര��.��q.Left - �����������ര��.��q.Width - 10  '���󳬳�
                �� = �����������ര��.��2.Width
            End Select
            �����������ര��.��((i + 1) / 2).Width = �� '��ȵ���
            �����������ര��.��((i + 1) / 2).Left = �� '���������
        Else 'ż�� ���ȵ���
            If IsNumeric(����(i)) Then
                �����������ര��.��(i / 2).Height = 100 * ����(i) * �ٶ�.Caption * ����.Caption
            Else
                �����������ര��.��(i / 2).Height = 1
            End If
            �����������ര��.��(i / 2).Top = �����������ര��.��(i / 2 - 1).Top - �����������ര��.��(i / 2).Height
            '��������
            If �����������ര��.��(i / 2).Top < 0 Then
                �����.Caption = i
                Exit For
            End If
        End If
        
    Next
    ��ʼ����.Caption = "GO"
End If

End Sub
Private Sub tips����ҳ_Click()
Set ws = CreateObject("wscript.shell")
ws.run "explorer https://bbs.nga.cn/read.php?tid=17450001"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''��ʼ����'''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ��ʼ����_Click()
If ����ģʽ.Caption = "չ��ģʽ" Then
If fChkControls(�����������ര��, "��", 1) = True Then
    If ��ʼ����.Caption = "GO" Then
        Timer1.Enabled = True '��ʱ������
        ��ʼ����.Caption = "��ͣ"
        ����ɿ���.Caption = 1
    Else
        Timer1.Enabled = False '��ʱ��������
        ��ʼ����.Caption = "GO"
    End If
    Timer1.Interval = 50 '100��������һ��
End If

Else '8��ģʽ
If fChkControls(�����������ര��mini, "��", 1) = True Then
    If ��ʼ����.Caption = "GO" Then
        Timer2.Enabled = True '��ʱ������
        ��ʼ����.Caption = "��ͣ"
        ����ɿ���.Caption = 1
    Else
        Timer2.Enabled = False '��ʱ��������
        ��ʼ����.Caption = "GO"
    End If
    Timer2.Interval = 50 '100��������һ��
End If
End If

End Sub
Private Sub Timer1_Timer()
''''��������½���
'�������еĿ�����
Dim ������ As Long
������ = 1
Do While (fChkControls(�����������ര��, "��", ������) = True) '�������ʱ �õ�������
������ = ������ + 1
Loop
������ = ������ - 1
If �����������ര��.��(������).Top > 0 And ��ʼ����.Caption <> "���" Then '''���¿���
    '��S��ֵ����������()
    Dim ����() As String
    ����() = Split(S, "|")
    ''''''''''�������׿�''''''''
    Dim �� As Integer
    Dim B As Integer
    If (UBound(����) - LBound(����) + 1) Mod 2 = 1 Then '����Ϊ����
    B = 0
    Else
    B = 1
    End If

    Dim i As Integer
    i = �����.Caption + 1
    ''''��ȡ����

    'MsgBox ������
            Load �����������ര��.��((i + 1) / 2) '����һ���µĿ�
            �����������ര��.��((i + 1) / 2).Visible = True  '��ʾ
            �����������ര��.��((i + 1) / 2).ZOrder 0  '�ö�
            'left
            Dim �� As Integer
            Select Case ����(i)
            Case "q"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "2"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "w"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "3"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "e"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "r"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "5"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case "t"
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "6"
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case "y"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "7"
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "u"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "z"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "x"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "c"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "v"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "b"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "n"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "m"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case ","
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "."
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case "/"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "["
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "]"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "a"
                �� = �����������ര��.��q.Left
                �� = �����������ര��.��q.Width
            Case "s"
                �� = �����������ര��.��2.Left
                �� = �����������ര��.��2.Width
            Case "d"
                �� = �����������ര��.��w.Left
                �� = �����������ര��.��q.Width
            Case "f"
                �� = �����������ര��.��3.Left
                �� = �����������ര��.��2.Width
            Case "g"
                �� = �����������ര��.��e.Left
                �� = �����������ര��.��q.Width
            Case "h"
                �� = �����������ര��.��r.Left
                �� = �����������ര��.��q.Width
            Case "j"
                �� = �����������ര��.��5.Left
                �� = �����������ര��.��2.Width
            Case "k"
                �� = �����������ര��.��t.Left
                �� = �����������ര��.��q.Width
            Case "l"
                �� = �����������ര��.��6.Left
                �� = �����������ര��.��2.Width
            Case ";"
                �� = �����������ര��.��y.Left
                �� = �����������ര��.��q.Width
            Case "'"
                �� = �����������ര��.��7.Left
                �� = �����������ര��.��2.Width
            Case "-"
                �� = �����������ര��.��u.Left
                �� = �����������ര��.��q.Width
            Case "="
                �� = �����������ര��.�߸�q.Left
                �� = �����������ര��.��q.Width
            Case Else
                �� = �����������ര��.��q.Left - �����������ര��.��q.Width - 10  '���󳬳�
                �� = �����������ര��.��2.Width
            End Select
            �����������ര��.��((i + 1) / 2).Width = �� '��ȵ���
            �����������ര��.��((i + 1) / 2).Left = �� '���������
            i = i + 1
            If IsNumeric(����(i)) Then
                �����������ര��.��(i / 2).Height = 100 * ����(i) * �ٶ�.Caption * ����.Caption
            Else
                �����������ര��.��(i / 2).Height = 1
            End If
            �����������ര��.��(i / 2).Top = �����������ര��.��(i / 2 - 1).Top - �����������ര��.��(i / 2).Height
        'End If
        �����.Caption = Int(�����.Caption) + 2
        ������ = ������ + 1
        If Int(�����.Caption) + 1 > UBound(����) - LBound(����) + B Then ��ʼ����.Caption = "���"
    'Next
End If
''''�ж��Ƿ��������
If �����������ര��.��(������).Top > �����������ര��.Height Then
Timer1.Enabled = False  '�������
��ʼ����.Caption = "���"
End If
''''�½�
For i = ����ɿ���.Caption To ������
�����������ര��.��(i).Top = �����������ര��.��(i).Top + 50 * ����.Caption
If �����������ര��.��(i).Top > �����������ര��.Height Then ����ɿ���.Caption = i
Next

End Sub

Private Sub ����ģʽ_Click()
If ��ʼ����.Caption = "��ͣ" Then Call ��ʼ����_Click
If ����ģʽ.Caption = "չ��ģʽ" Then
    Unload �����������ര��
    �����������ര��mini.Show
    '''''''''''''''''''''����͸��'''''''''''''''
    Dim rtn As Long
    �����������ര��mini.BackColor = RGB(0, 0, 0) '����һ�´��ڵ���ɫ
    ���ھ�� = ȡ���ھ��(vbNullString, "լү���Ƶ�������-���ര��")
    rtn = GetWindowLong(���ھ��, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong ���ھ��, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes ���ھ��, RGB(0, 0, 0), 128, LWA_ALPHA '���崰��͸����
    'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '�ؼ�����͸��
    'RGB(0, 0, 0)��������Ҫ͸��������ɫ
    ����mini.Visible = True
    ����ģʽ.Caption = "��չ��"
    Call ���ô���.����mini_Click
Else
    Unload �����������ര��mini
    �����������ര��.Show
    '''''''''''''''''''''����͸��'''''''''''''''
    'Dim rtn As Long
    �����������ര��.BackColor = RGB(0, 0, 0) '����һ�´��ڵ���ɫ
    ���ھ�� = ȡ���ھ��(vbNullString, "լү���Ƶ�������-���ര��")
    rtn = GetWindowLong(���ھ��, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong ���ھ��, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes ���ھ��, RGB(0, 0, 0), 128, LWA_ALPHA '���崰��͸����
    'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_COLORKEY '�ؼ�����͸��
    'RGB(0, 0, 0)��������Ҫ͸��������ɫ
    ����mini.Visible = False
    ����ģʽ.Caption = "չ��ģʽ"
    Call ���ô���.����_Click
End If
End Sub
''''''''''''''''''''''''''����������'''''''''''''''''''''''''
Public Sub ����mini_Click()
'�ж�֮ǰ�Ƿ��пؼ�����
Dim i As Long
i = 1
Do While (fChkControls(�����������ര��mini, "��", i) = True) '�������ʱ
Unload �����������ര��mini.��(i) 'ɾ���ؼ�
i = i + 1
Loop
S = ""
''''''''''''��ʼ������'''''''''''
'�ж�·���Ƿ����
If ����·��.Text <> "" Then '����
    Dim A As String
    '��ȡ���ף�SΪ������
    FreeNum = FreeFile
    Open ����·��.Text For Input As #FreeNum
    Do While Not EOF(FreeNum) '��ȡһֱ���ļ�ĩβ
        Line Input #FreeNum, A
        S = S + "|" + A 'S�������������ļ�
        If A����ĳ������ And Not EOF(FreeNum) Then
            Line Input #FreeNum, A '��ȡ��һ�е�����
            Exit Do '�˳�ѭ��
        End If
    Loop
    Close #FreeNum
    '��S��ֵ����������()
    Dim ����() As String
    ����() = Split(S, "|")
    ''''''''''�������׿�''''''''
    Dim �� As Integer
    Dim B As Integer
    If (UBound(����) - LBound(����) + 1) Mod 2 = 1 Then '����Ϊ����
    B = 0
    Else
    B = 1
    End If
    For i = 1 To (UBound(����) - LBound(����) + B) 'ѭ��
        If i Mod 2 = 1 Then '����
            Load �����������ര��mini.��((i + 1) / 2) '����һ���µĿ�
            �����������ര��mini.��((i + 1) / 2).Visible = True  '��ʾ
            �����������ര��mini.��((i + 1) / 2).ZOrder 0  '�ö�
            'left
            Dim �� As Integer
            Select Case ����(i)
            Case "q"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
            Case "2"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
            Case "w"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
            Case "3"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
            Case "e"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
            Case "r"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
            Case "5"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
            Case "t"
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
            Case "6"
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
            Case "y"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
            Case "7"
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
            Case "u"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
            Case "z"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "x"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "c"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "v"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "b"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "n"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "m"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case ","
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "."
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "/"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "["
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "]"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "a"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "s"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "d"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "f"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "g"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "h"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "j"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "k"
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "l"
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case ";"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "'"
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "-"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "="
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case Else
                �� = �����������ര��mini.��q.Left - �����������ര��mini.��q.Width - 10  '���󳬳�
                �� = �����������ര��mini.��2.Width
            End Select
            �����������ര��mini.��((i + 1) / 2).Width = �� '��ȵ���
            �����������ര��mini.��((i + 1) / 2).Left = �� '���������
        Else 'ż�� ���ȵ���
            If IsNumeric(����(i)) Then
                �����������ര��mini.��(i / 2).Height = 100 * ����(i) * �ٶ�.Caption * ����.Caption
            Else
                �����������ര��mini.��(i / 2).Height = 1
            End If
            �����������ര��mini.��(i / 2).Top = �����������ര��mini.��(i / 2 - 1).Top - �����������ര��mini.��(i / 2).Height
            '��������
            If �����������ര��mini.��(i / 2).Top < 0 Then
                �����.Caption = i
                Exit For
            End If
        End If
        
    Next
    ��ʼ����.Caption = "GO"
End If
End Sub

Private Sub Timer2_Timer()
''''��������½���
'�������еĿ�����
Dim ������ As Long
������ = 1
Do While (fChkControls(�����������ര��mini, "��", ������) = True) '�������ʱ �õ�������
������ = ������ + 1
Loop
������ = ������ - 1
If �����������ര��mini.��(������).Top > 0 And ��ʼ����.Caption <> "���" Then '''���¿���
    '��S��ֵ����������()
    Dim ����() As String
    ����() = Split(S, "|")
    ''''''''''�������׿�''''''''
    Dim �� As Integer
    Dim B As Integer
    If (UBound(����) - LBound(����) + 1) Mod 2 = 1 Then '����Ϊ����
    B = 0
    Else
    B = 1
    End If

    Dim i As Integer
    i = �����.Caption + 1
    ''''��ȡ����

    'MsgBox ������
            Load �����������ര��mini.��((i + 1) / 2) '����һ���µĿ�
            �����������ര��mini.��((i + 1) / 2).Visible = True  '��ʾ
            �����������ര��mini.��((i + 1) / 2).ZOrder 0  '�ö�
            'left
            Dim �� As Integer
            Select Case ����(i)
            Case "q"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
            Case "2"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
            Case "w"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
            Case "3"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
            Case "e"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
            Case "r"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
            Case "5"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
            Case "t"
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
            Case "6"
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
            Case "y"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
            Case "7"
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
            Case "u"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
            Case "z"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "x"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "c"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "v"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "b"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "n"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "m"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case ","
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "."
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "/"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "["
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "]"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HFFFFC0
            Case "a"
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "s"
                �� = �����������ര��mini.��2.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "d"
                �� = �����������ര��mini.��w.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "f"
                �� = �����������ര��mini.��3.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "g"
                �� = �����������ര��mini.��e.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "h"
                �� = �����������ര��mini.��r.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "j"
                �� = �����������ര��mini.��5.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "k"
                �� = �����������ര��mini.��t.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "l"
                �� = �����������ര��mini.��6.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case ";"
                �� = �����������ര��mini.��y.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "'"
                �� = �����������ര��mini.��7.Left
                �� = �����������ര��mini.��2.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "-"
                �� = �����������ര��mini.��u.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case "="
                �� = �����������ര��mini.��q.Left
                �� = �����������ര��mini.��q.Width
                �����������ര��mini.��((i + 1) / 2).BackColor = &HC0C0FF
            Case Else
                �� = �����������ര��mini.��q.Left - �����������ര��mini.��q.Width - 10  '���󳬳�
                �� = �����������ര��mini.��2.Width
            End Select
            �����������ര��mini.��((i + 1) / 2).Width = �� '��ȵ���
            �����������ര��mini.��((i + 1) / 2).Left = �� '���������
            i = i + 1
            If IsNumeric(����(i)) Then
                �����������ര��mini.��(i / 2).Height = 100 * ����(i) * �ٶ�.Caption * ����.Caption
            Else
                �����������ര��mini.��(i / 2).Height = 1
            End If
            �����������ര��mini.��(i / 2).Top = �����������ര��mini.��(i / 2 - 1).Top - �����������ര��mini.��(i / 2).Height
        'End If
        �����.Caption = Int(�����.Caption) + 2
        ������ = ������ + 1
        If Int(�����.Caption) + 1 > UBound(����) - LBound(����) + B Then ��ʼ����.Caption = "���"
    'Next
End If
''''�ж��Ƿ��������
If �����������ര��mini.��(������).Top > �����������ര��mini.Height Then
Timer1.Enabled = False  '�������
��ʼ����.Caption = "���"
End If
''''�½�
For i = ����ɿ���.Caption To ������
�����������ര��mini.��(i).Top = �����������ര��mini.��(i).Top + 50 * ����.Caption
If �����������ര��mini.��(i).Top > �����������ര��mini.Height Then ����ɿ���.Caption = i
Next

End Sub
