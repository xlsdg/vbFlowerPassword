Attribute VB_Name = "modDPI"
'96 DPI �� TwipsPerPixelX TwipsPerPixelY Ϊ 15 --- ��DPIΪ96ʱ,15羵���1����
'120 DPI �� TwipsPerPixelX TwipsPerPixelY Ϊ 12 --- ��DPIΪ120ʱ,12羵���1����
'��ô���� ÿ�� 1 DPI ��+8
'------------
'�������߶���[��96DPI�²��]:2145�[143����,Y]  �����:8715�[581����,X]
'�����ṩһ����ʽ:1 ���� = 1440 TPI / 96 DPI = 15 �
'����X����=1440/DPIֵ=Y�;
'####################################################################################################################################
Option Explicit

Private Declare Function GetDeviceCaps _
                Lib "gdi32.dll" (ByVal hdc As Long, _
                                 ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X

Private FormOldWidth  As Long

'���洰���ԭʼ���
Private FormOldHeight As Long
'���洰���ԭʼ�߶�
    
'�ڵ���ResizeFormǰ�ȵ��ñ�����
Public Sub ResizeInit(FormName As Form)

    Dim Obj As Control

    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight

    On Error Resume Next

    For Each Obj In FormName

        Obj.Tag = Obj.Left & "   " & Obj.Top & "   " & Obj.Width & "   " & Obj.Height & "   "
    Next Obj

    On Error GoTo 0

End Sub
    
'�������ı���ڸ�Ԫ���Ĵ�С��
'�ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)

    Dim Pos(4) As Double

    Dim i      As Long, TempPos       As Long, StartPos       As Long

    Dim Obj    As Control

    Dim ScaleX As Double, ScaleY       As Double

    ScaleX = FormName.ScaleWidth / FormOldWidth
    '���洰�������ű���
    ScaleY = FormName.ScaleHeight / FormOldHeight

    '���洰��߶����ű���
    On Error Resume Next

    For Each Obj In FormName

        StartPos = 1

        For i = 0 To 4
            '��ȡ�ؼ���ԭʼλ�����С
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)

            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0

            End If

            '���ݿؼ���ԭʼλ�ü�����ı��С
            '�ı����Կؼ����¶�λ��ı��С
            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i
    Next Obj

    On Error GoTo 0

End Sub
    
Private Sub Form_Activate()

    Dim aa   As Long

    Dim hdc0 As Long

    hdc0 = GetDC(0)
    aa = GetDeviceCaps(hdc0, LOGPIXELSX) '���DPIֵ

    Dim x As Integer

    x = 1440 / aa 'X�=1����
    Me.Height = 143 * x
    Me.Width = 581 * x
    Image1.Height = 114 * x
    Image1.Width = 581 * x

End Sub
    
