Attribute VB_Name = "modTray"
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 0 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 0 7
' Description :
'     History :
'*****************************************************************
Option Explicit

Private Declare Function Shell_NotifyIcon _
                Lib "shell32.dll" _
                Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                           lpData As NOTIFYICONDATA) As Long

'---------- dwMessage����������NIM_ADD��NIM_DELETE��NIM_MODIFY ��ʶ��֮һ----------
Private Const NIM_ADD = &H0    '��������������һ��ͼ��

Private Const NIM_DELETE = &H2    'ɾ���������е�һ��ͼ��

Private Const NIM_MODIFY = &H1    '�޸��������и�ͼ����Ϣ

Private Const NIM_SETFOCUS = &H3

Private Const NIM_SETVERSION = &H4

Private Const NIF_MESSAGE = &H1    'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ

Private Const NIF_ICON = &H2

Private Const NIF_TIP = &H4

Private Const NIF_STATE = &H8

Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1

Private Const NIS_SHAREDICON = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const WM_LBUTTONDOWN = &H201

Private Const WM_LBUTTONUP = &H202

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_RBUTTONDOWN = &H204

Private Const WM_RBUTTONUP = &H205

Private Const WM_RBUTTONDBLCLK = &H206

Private Const WM_MBUTTONDOWN = &H207

Private Const WM_MBUTTONUP = &H208

Private Const WM_MBUTTONDBLCLK = &H209

Private Const NOTIFYICON_VERSION = 3    '���

Private Const NOTIFYICON_OLDVERSION = 0    'Win95 ��������ʽ

'ϵͳ��������
Private Type NOTIFYICONDATA

    cbSize As Long    '�����ݽṹ�Ĵ�С
    hwnd As Long    '������������ͼ��Ĵ��ھ��
    uID As Long    '�������������ͼ��ı�ʶ
    uFlags As Long    '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
    '   NIF_MESSAGE ��ʾ���Ϳ�����Ϣ��
    '   NIF_ICON��ʾ��ʾ�������е�ͼ�ꣻ
    '   NIF_TIP��ʾ�������е�ͼ���ж�̬��ʾ��
    uCallbackMessage As Long    '������ͼ��ͨ�������û����򽻻���Ϣ���������Ϣ�Ĵ�����hWnd����
    hIcon As Long    '�������е�ͼ��Ŀ��ƾ��
    szTip As String * 128    'ͼ�����ʾ��Ϣ����Ҫ����������ʾ��Ϣ����һ��Ҫ128���ԣ�Ϊ64���޷��������ݣ��������ܶ�������ԭ����
    '������ʾ��Ϣ����
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256    '������ʾ����
    uTimeout As Long    '������ʾ��ʾʱ��
    szInfoTitle As String * 64    '������ʾ����
    dwInfoFlags As Long    '������ʾ���ͣ��� NIIF_*** ����

End Type

Public Enum ico '������ʾ����

    NIIF_NONE = &H0     'w��ͼ�� = &H0      '
    NIIF_INFO = &H1     'x��Ϣͼ�� = &H1    '
    NIIF_WARNING = &H2  'j����ͼ�� = &H2    '
    NIIF_ERROR = &H3    'z����ͼ�� = &H3    '
    NIIF_GUID = &H4     't����ͼ�� = &H4    '

End Enum

Private IconData As NOTIFYICONDATA

Public Function AddToTray(ByVal frm As Form, _
                          ByVal Tip As String, _
                          Optional ByVal TrayIco As ico = 0) As Long

    '����ϵͳ����ͼ��
    With IconData
        .cbSize = Len(IconData)
        .hwnd = frm.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE    '��Ӧ����¼� 'WM_LBUTTONDOWN

        If TrayIco = 0 Then
            .hIcon = frm.Icon    'Ĭ��Ϊ���ڵ�ͼ��
        Else
            .hIcon = TrayIco

        End If

        .szTip = Tip & vbNullChar

    End With

    AddToTray = Shell_NotifyIcon(NIM_ADD, IconData)   '��������ͼ��

End Function

Public Sub MouseOnTray(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lMsg As Single

    lMsg = X / Screen.TwipsPerPixelX

    Select Case lMsg

        Case WM_LBUTTONUP                   '���̧��

        Case WM_LBUTTONDBLCLK               '���˫��
            FrmMain.MidScreenShow

        Case WM_RBUTTONUP                   '�Ҽ�̧��
            FrmMain.PopupMenu FrmSetting.munFlowerPassword

        Case WM_RBUTTONDBLCLK               '�Ҽ�˫��

        Case WM_MBUTTONUP                   '�м�̧��

        Case WM_MBUTTONDBLCLK               '�м�˫��

        Case 1028                           '����������ݹرհ�ť

        Case 1029                           '����������ݴ��屾��

    End Select

End Sub

Public Function RemoveFromTray() As Long
    RemoveFromTray = Shell_NotifyIcon(NIM_DELETE, IconData)   'ж������ͼ��

End Function

Public Function SetTrayMsgbox(ByVal MsgInfo As String, _
                              ByVal MsgFlags As Integer, _
                              ByVal MsgTitle As String, _
                              Optional ByVal TrayIco As ico = 0) As Long

    '    "ϵͳ����������ʾ���ֲ��ó���128���ַ���"
    With IconData
        .szInfoTitle = MsgTitle & Chr$(0)
        .szInfo = MsgInfo & Chr$(0)
        .dwInfoFlags = MsgFlags

        If TrayIco <> 0 Then
            .hIcon = TrayIco    '��������ͼ��

        End If

    End With

    SetTrayMsgbox = Shell_NotifyIcon(NIM_MODIFY, IconData)   '�޸�����ͼ�꼰�����Ϣ

End Function
