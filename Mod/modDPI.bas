Attribute VB_Name = "modDPI"
'96 DPI 下 TwipsPerPixelX TwipsPerPixelY 为 15 --- 即DPI为96时,15缇等于1像素
'120 DPI 下 TwipsPerPixelX TwipsPerPixelY 为 12 --- 即DPI为120时,12缇等于1像素
'这么看来 每高 1 DPI 就+8
'------------
'这个窗体高度是[在96DPI下测得]:2145缇[143像素,Y]  宽度是:8715缇[581像素,X]
'在这提供一个公式:1 像素 = 1440 TPI / 96 DPI = 15 缇
'所以X像素=1440/DPI值=Y缇;
'####################################################################################################################################
Option Explicit

Private Declare Function GetDeviceCaps _
                Lib "gdi32.dll" (ByVal hdc As Long, _
                                 ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X

Private FormOldWidth  As Long

'保存窗体的原始宽度
Private FormOldHeight As Long
'保存窗体的原始高度
    
'在调用ResizeForm前先调用本函数
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
    
'按比例改变表单内各元件的大小，
'在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)

    Dim Pos(4) As Double

    Dim i      As Long, TempPos       As Long, StartPos       As Long

    Dim Obj    As Control

    Dim ScaleX As Double, ScaleY       As Double

    ScaleX = FormName.ScaleWidth / FormOldWidth
    '保存窗体宽度缩放比例
    ScaleY = FormName.ScaleHeight / FormOldHeight

    '保存窗体高度缩放比例
    On Error Resume Next

    For Each Obj In FormName

        StartPos = 1

        For i = 0 To 4
            '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)

            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0

            End If

            '根据控件的原始位置及窗体改变大小
            '的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i
    Next Obj

    On Error GoTo 0

End Sub
    
Private Sub Form_Activate()

    Dim aa   As Long

    Dim hdc0 As Long

    hdc0 = GetDC(0)
    aa = GetDeviceCaps(hdc0, LOGPIXELSX) '获得DPI值

    Dim x As Integer

    x = 1440 / aa 'X缇=1像素
    Me.Height = 143 * x
    Me.Width = 581 * x
    Image1.Height = 114 * x
    Image1.Width = 581 * x

End Sub
    
