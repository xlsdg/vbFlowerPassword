Attribute VB_Name = "modTransparent"
Option Explicit

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Const GWL_EXSTYLE = (-20)

Private Const LWA_COLORKEY  As Long = &H1

Private Const LWA_ALPHA     As Long = &H2

Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByVal crKey As Long, _
                                  ByVal bAlpha As Long, _
                                  ByVal dwFlags As Long) As Long

Public Function SetFrmTransparent(ByVal frmHwnd As Long, _
                                  Optional ByVal intPercent As Integer = 192) As Long

    Dim lngStyle As Long

    lngStyle = GetWindowLong(frmHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetWindowLong frmHwnd, GWL_EXSTYLE, lngStyle
    '�����е�͸��ɫ
    'SetFrmTransparent = SetLayeredWindowAttributes(frmHwnd, RGB(255, 255, 255), 0, LWA_COLORKEY)
    '��������ʾ�����е�͸��ɫ,���ڶ���������ʾ͸��ɫ,������RGB������ָ����ɫֵ
    '����͸����
    SetFrmTransparent = SetLayeredWindowAttributes(frmHwnd, 0, intPercent, LWA_ALPHA)

    '�Ѵ������óɰ�͸����ʽ,�ڶ���������ʾ͸���̶�,ȡֵ��Χ 0 - 255.Ϊ0ʱ����һ��ȫ͸���Ĵ�����
End Function

Public Function UnSetFrmTransparent(ByVal frmHwnd As Long) As Long

    Dim lngStyle As Long

    lngStyle = GetWindowLong(frmHwnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED)
    UnSetFrmTransparent = SetWindowLong(frmHwnd, GWL_EXSTYLE, lngStyle)

End Function
