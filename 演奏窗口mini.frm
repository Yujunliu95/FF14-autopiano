VERSION 5.00
Begin VB.Form �����������ര��mini 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "լү���Ƶ�������-���ര��"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Shape �� 
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
   Begin VB.Shape ��7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   2760
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape ��6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   2280
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape ��y 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   2400
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   1800
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape ��3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   840
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape ��2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   360
      Top             =   1040
      Width           =   255
   End
   Begin VB.Shape ��w 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   480
      Top             =   0
      Width           =   480
   End
   Begin VB.Label ������С 
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
   Begin VB.Shape ��q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   3360
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��u 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   2880
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��t 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   1920
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��r 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   1440
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��e 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   960
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape ��q 
      BackColor       =   &H00303030&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "�����������ര��mini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''�����ö�'''''''''''''''''''
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

''''''''''''''''''���ڿ��϶�''''''''''''''''''''''
Dim xa As Single, ya As Single
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
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
'��꾭������
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
������С.BackColor = &H80000002
End Sub
''''''''''''''''''''������С'''''''''''''''''''''''''
Private Sub ������С_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ���ô���.��ʼ����.Caption = "GO" Then
xa = X
ya = Y
'���
Dim i As Long
i = 1
Do While (fChkControls(�����������ര��mini, "��", i) = True) '�������ʱ
Unload �����������ര��mini.��(i) 'ɾ���ؼ�
i = i + 1
Loop
S = ""
End If
End Sub
'�����ƶ�
Private Sub ������С_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ���ô���.��ʼ����.Caption = "GO" Then
If Button = 1 Then ������С.Move ������С.Left + X - xa, ������С.Top + Y - ya    '�����ƶ�
'�ı䴰�ڴ�С
Me.Width = ������С.Left + ������С.Width
Me.Height = ������С.Top + ������С.Height
������С.BackColor = &H8080FF
'�ڲ�λ��ͬ���ı�
'��
Dim �� As Integer, �ڿ� As Integer
�� = Me.Width / 8
�ڿ� = �� / 2
��q.Width = ��
��2.Width = �ڿ�
��2.Left = ��q.Left + ��q.Width - �ڿ� / 2
��w.Width = ��
��w.Left = ��q.Left + ��q.Width
��3.Width = �ڿ�
��3.Left = ��w.Left + ��w.Width - �ڿ� / 2
��e.Width = ��
��e.Left = ��w.Left + ��w.Width
��r.Width = ��
��r.Left = ��e.Left + ��e.Width
��5.Width = �ڿ�
��5.Left = ��r.Left + ��r.Width - �ڿ� / 2
��t.Width = ��
��t.Left = ��r.Left + ��r.Width
��6.Width = �ڿ�
��6.Left = ��t.Left + ��t.Width - �ڿ� / 2
��y.Width = ��
��y.Left = ��t.Left + ��t.Width
��7.Width = �ڿ�
��7.Left = ��y.Left + ��y.Width - �ڿ� / 2
��u.Width = ��
��u.Left = ��y.Left + ��y.Width
��q.Width = ��
��q.Left = ��u.Left + ��u.Width
'��
Dim �ڸ� As Single
�ڸ� = 2005
��q.Height = Me.Height
��2.Top = Me.Height - �ڸ�
��w.Height = Me.Height
��3.Top = Me.Height - �ڸ�
��e.Height = Me.Height
��r.Height = Me.Height
��5.Top = Me.Height - �ڸ�
��t.Height = Me.Height
��6.Top = Me.Height - �ڸ�
��y.Height = Me.Height
��7.Top = Me.Height - �ڸ�
��u.Height = Me.Height
��q.Height = Me.Height

'��ʼ����λ��
��(0).Top = ��2.Top
If ���ô���.����mini.Visible = True Then
Call ���ô���.����mini_Click
Else
Call ���ô���.����_Click
End If
End If
End Sub
'''''''''''''''''''''����''''''''''''''''''''''
Private Sub Form_Load()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE     '���ô����ö�
End Sub
