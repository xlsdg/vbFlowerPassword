Attribute VB_Name = "modOnTop"
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

'���������Ϊ����ָ��һ����λ�ú�״̬����Ҳ�ɸı䴰�����ڲ������б��е�λ�á��ú�����DeferWindowPos�������ƣ�ֻ�������������������ֳ����ģ���vb��ʹ�ã����vb���壬��������win32�����λ���С���������������״̬�����б�Ҫ������һ�����ദ��ģ�����������״̬
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const SWP_NOACTIVATE = &H10

Private Const SWP_SHOWWINDOW = &H40

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Function SetWinByPoint(ByVal WinHwnd As Long, _
                              ByVal point_x As Long, _
                              ByVal point_y As Long) As Long
    SetWinByPoint = SetWindowPos(WinHwnd, HWND_TOPMOST, point_x, point_y, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW)

End Function

Public Function SetWinOnTop(ByVal WinHwnd As Long) As Long
    SetWinOnTop = SetWindowPos(WinHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Function

Public Function UnSetWinOnTop(ByVal WinHwnd As Long) As Long
    UnSetWinOnTop = SetWindowPos(WinHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

End Function
