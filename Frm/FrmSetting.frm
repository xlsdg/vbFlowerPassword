VERSION 5.00
Begin VB.Form FrmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ����"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   16335
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   6
      Left            =   11520
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4575
      Begin VB.CheckBox chkPasswordLength 
         Caption         =   "�Զ�����������λ�� (Ĭ��Ϊ16λ):"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   3135
      End
      Begin VB.ComboBox comDiyLength 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox comPwdLength 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   100
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�"
         Height          =   360
         Left            =   1440
         TabIndex        =   38
         Top             =   2640
         Width           =   1110
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "���"
         Height          =   360
         Left            =   1440
         TabIndex        =   40
         Top             =   3120
         Width           =   1110
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��"
         Height          =   360
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���"
         Height          =   360
         Left            =   120
         TabIndex        =   37
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtDiyKey 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ListBox lstDiyKey 
         Height          =   2205
         Left            =   2760
         TabIndex        =   41
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Line line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   4440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblDiyLength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���볤�� :"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblDiyKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ִ��� :"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������볤������ :"
         Height          =   195
         Left            =   2760
         TabIndex        =   49
         Top             =   840
         Width           =   1545
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   4
      Left            =   6720
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4575
      Begin VB.CheckBox chkProtection 
         Caption         =   "��������������������ȡ�������� (�Ƽ�)"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkAutoUseDomain 
         Caption         =   "�Զ����������ַ����վ�����������ִ���"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "���������������������ڲ�ʱ, �Զ���ȡ��ַ����վ������Ϊ���ִ���"
         Top             =   600
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkAutoCopy 
         Caption         =   "�Զ������ɵ����븴�Ƶ�������"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox chkDomainSuffix 
         Caption         =   "������վ������׺ (��: .com , .net , .org , ...)"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   1560
         Width           =   3915
      End
      Begin VB.CheckBox chkAutoCheckClipboard 
         Caption         =   "�Զ�ʶ������帴�Ƶ���ַ"
         Height          =   255
         Left            =   480
         TabIndex        =   30
         ToolTipText     =   "���ڲ����Զ���ȡ��ַ����վ�����������, ���������ֶ�����URL���������ɻ����Զ�ʶ��������Ϊ���ִ���"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkUseMouseHook 
         Caption         =   "ʹ��ȫ����갴���ҹ���ʽ������λ���������"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "���ʹ��ģ�ⰴ�����뷽ʽʱ, ��ǰ����״̬�����������������"
         Top             =   2520
         Width           =   4215
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   1
      Left            =   6720
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Width           =   4575
      Begin VB.HScrollBar HScrollTransparent 
         Enabled         =   0   'False
         Height          =   240
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   16
         Top             =   120
         Value           =   192
         Width           =   1455
      End
      Begin VB.CheckBox chkShowPassword 
         Caption         =   "�����������������������ʾ (����)"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CheckBox chkShowSwitch 
         Caption         =   "�ڻ����������ʾ "" �������� "" ��ݿ���"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   3735
      End
      Begin VB.OptionButton optClodeToHide 
         Caption         =   "���ص�������֪ͨ���� , ���˳�����"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   2640
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optCloseToExit 
         Caption         =   "�˳�����"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "��������������͸����Ч��"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   2895
      End
      Begin VB.CheckBox chkShowCode 
         Caption         =   "�ڻ����������ʾ���ɵ�����"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "192"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3600
         TabIndex        =   52
         Top             =   360
         Width           =   225
      End
      Begin VB.Line line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   4440
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblAskClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�رջ��������ʱ :"
         Height          =   195
         Left            =   480
         TabIndex        =   47
         Top             =   2250
         Width           =   1545
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   1800
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   240
      Width           =   4575
      Begin VB.CheckBox chkShowHelp 
         Caption         =   "��������ʱ��������������ʾ��ȫ���ȼ�����"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.ComboBox comHKey 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1875
         Width           =   855
      End
      Begin VB.CheckBox chkAlt 
         Caption         =   "Alt"
         Height          =   255
         Left            =   2790
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox chkWin 
         Caption         =   "Win"
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   1920
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkCtrl 
         Caption         =   "Ctrl"
         Height          =   255
         Left            =   1170
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox chkShift 
         Caption         =   "Shift"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkAutoMini 
         Caption         =   "��������ʱΪ���Զ���С����������֪ͨ����"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   4215
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "���������ʼ�ձ�������������ǰ��"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.Label lblHotKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����������ʾȫ���ȼ� :"
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   1560
         Width           =   2085
      End
      Begin VB.Line line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   4440
         X2              =   120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblDonation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1560
         MouseIcon       =   "FrmSetting.frx":43B2
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lblUpdate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   480
         MouseIcon       =   "FrmSetting.frx":4508
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2880
         Width           =   720
      End
      Begin VB.Image imgxlsdg 
         Height          =   300
         Left            =   4200
         MouseIcon       =   "FrmSetting.frx":465E
         MousePointer    =   99  'Custom
         Picture         =   "FrmSetting.frx":47B4
         ToolTipText     =   "xLsDg"
         Top             =   2760
         Width           =   300
      End
      Begin VB.Image imgJohnnyJian 
         Height          =   300
         Left            =   2760
         MouseIcon       =   "FrmSetting.frx":4C39
         MousePointer    =   99  'Custom
         Picture         =   "FrmSetting.frx":4D8F
         ToolTipText     =   "JohnnyJian"
         Top             =   2760
         Width           =   300
      End
      Begin VB.Image imgKenshin 
         Height          =   300
         Left            =   3240
         MouseIcon       =   "FrmSetting.frx":4F6A
         MousePointer    =   99  'Custom
         Picture         =   "FrmSetting.frx":50C0
         ToolTipText     =   "Kenshin"
         Top             =   2760
         Width           =   300
      End
      Begin VB.Image imgkisexu 
         Height          =   300
         Left            =   3720
         MouseIcon       =   "FrmSetting.frx":5401
         MousePointer    =   99  'Custom
         Picture         =   "FrmSetting.frx":5557
         ToolTipText     =   "KiseXu"
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2011-2012 FlowerPassword.com All rights reserved."
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   50
         MouseIcon       =   "FrmSetting.frx":5855
         TabIndex        =   48
         Top             =   3240
         Width           =   4500
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   3
      Left            =   11520
      ScaleHeight     =   3615
      ScaleWidth      =   4575
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox comPrefixCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   600
         TabIndex        =   24
         Top             =   1080
         Width           =   3720
      End
      Begin VB.CheckBox chkPrefix 
         Caption         =   "ǰ׺ ( �������� + ���ִ��� )"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkSuffix 
         Caption         =   "��׺ ( ���ִ��� + �������� )"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox comSuffixCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   600
         TabIndex        =   26
         Top             =   2160
         Width           =   3720
      End
      Begin VB.CheckBox chkAutoAddUserCode 
         Caption         =   "�Զ����沢Ĭ�������ϴ�ʹ�õĸ������� :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "�Զ�������������ӵ����ִ�����Ӧλ��"
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label lblExample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� : "
         Enabled         =   0   'False
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   3000
         Width           =   510
      End
   End
   Begin VB.ListBox lstSet 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   360
      TabIndex        =   3
      Top             =   380
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "�ָ�����Ĭ������"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   360
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   1110
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ȷ��"
      Height          =   360
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Line line1 
      BorderColor     =   &H80000010&
      X1              =   1680
      X2              =   1680
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Shape saeSetting 
      BorderColor     =   &H8000000C&
      Height          =   3855
      Left            =   120
      Top             =   120
      Width           =   6375
   End
   Begin VB.Menu munFlowerPassword 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu munShow 
         Caption         =   "��ʾ"
      End
      Begin VB.Menu munLine 
         Caption         =   "-"
      End
      Begin VB.Menu munSetting 
         Caption         =   "����"
      End
      Begin VB.Menu munExit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 0 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 1 2
' Description :
'     History :
'*****************************************************************
Option Explicit

Private tmpSel As Long, tmpHotKeyValue As Long, tmpKeyValue As Integer

Private Sub ShowExample()

    Dim temp As String

    temp = "google"

    If chkPrefix.value = Checked Then
        temp = comPrefixCode.Text + temp

    End If

    If chkSuffix.value = Checked Then
        temp = temp + comSuffixCode.Text

    End If

    lblExample.Caption = "���� : " + temp

End Sub

Private Sub chkAlt_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkAlwaysOnTop_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkAutoAddUserCode_Click()
    cmdApply.Enabled = True

    If chkAutoAddUserCode.value = Checked Then
        lblExample.Enabled = True
        chkPrefix.Enabled = True
        chkSuffix.Enabled = True
        'Call chkPrefix_Click
        'Call chkSuffix_Click
        FrmMain.chkAddUserCode.value = Checked
    Else
        lblExample.Enabled = False
        chkPrefix.Enabled = False
        chkSuffix.Enabled = False
        'comPrefixCode.Enabled = False
        'comSuffixCode.Enabled = False
        FrmMain.chkAddUserCode.value = Unchecked

    End If
    
    Call chkPrefix_Click
    Call chkSuffix_Click
    
End Sub

Private Sub chkAutoCheckClipboard_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkAutoCopy_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkAutoMini_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkAutoUseDomain_Click()
    cmdApply.Enabled = True

    If chkAutoUseDomain.value = Checked Then
        chkDomainSuffix.Enabled = True
        chkAutoCheckClipboard.Enabled = True
    Else
        chkDomainSuffix.Enabled = False
        chkAutoCheckClipboard.Enabled = False

    End If

End Sub

Private Sub chkCtrl_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkDomainSuffix_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkPasswordLength_Click()

    If chkPasswordLength.value = Checked Then
        If Not isDiyWordLength Then
            cmdApply.Enabled = True

            If MsgBox("�˲�����������ɵ�����λ���뻨��ͬ������Ӧ�ò�ͬ, ȷ��Ҫʹ���Զ�����������λ����?", vbYesNo + vbDefaultButton2 + vbSystemModal + vbExclamation) = vbYes Then
                isDiyWordLength = True
                comPwdLength.Enabled = True
            Else
                chkPasswordLength.value = Unchecked

            End If

        End If

    Else
        isDiyWordLength = False
        comPwdLength.Enabled = False

    End If

End Sub

Private Sub chkPrefix_Click()
    cmdApply.Enabled = True

    If chkPrefix.value = Checked Then
        If chkAutoAddUserCode.value = Checked Then
            comPrefixCode.Enabled = True
        Else
            comPrefixCode.Enabled = False
    
        End If
    Else
        comPrefixCode.Enabled = False

    End If

    Call ShowExample

End Sub

Private Sub chkProtection_Click()
    cmdApply.Enabled = True

    If chkProtection.value = Unchecked Then
        If isProtect Then
            If MsgBox("�˲������ܻ���ɼ������뱻��ȡ, ȷ��Ҫȡ�����������������ǿ�ͱ���������?", vbYesNo + vbDefaultButton2 + vbSystemModal + vbExclamation) = vbYes Then
                isProtect = False
                UnProtectTextBox FrmMain.txtPassword.hwnd
            Else
                chkProtection.value = Checked

            End If

        End If

    Else
        isProtect = True

        If ProtectTextBox(FrmMain.txtPassword.hwnd) = 0 Then
            MsgBox "�������ü�����������򱣻���ʩʧ��!", vbCritical + vbOKOnly + vbSystemModal

        End If

    End If

End Sub

Private Sub chkShift_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkShowCode_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkShowHelp_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkShowPassword_Click()
    cmdApply.Enabled = True

    If chkShowPassword.value = Checked Then
        If Not isShowPassword Then
            If MsgBox("�˲������ܻ���ɼ������뱻��ȡ, ȷ��Ҫ�����ĵķ�ʽ��ʾ����������?", vbYesNo + vbDefaultButton2 + vbSystemModal + vbExclamation) = vbYes Then
                isShowPassword = True
            Else
                chkShowPassword.value = Unchecked

            End If

        End If

    Else
        isShowPassword = False

    End If

End Sub

Private Sub chkShowSwitch_Click()
    cmdApply.Enabled = True

End Sub

Private Sub chkSuffix_Click()
    cmdApply.Enabled = True

    If chkSuffix.value = Checked Then
        If chkAutoAddUserCode.value = Checked Then
            comSuffixCode.Enabled = True
        Else
            comSuffixCode.Enabled = False
        End If
    Else
        comSuffixCode.Enabled = False
    
    End If

    Call ShowExample

End Sub

Private Sub chkTransparent_Click()
    cmdApply.Enabled = True

    If chkTransparent.value = Checked Then
        HScrollTransparent.Enabled = True
        lblValue.Enabled = True
    Else
        HScrollTransparent.Enabled = False
        lblValue.Enabled = False

    End If

End Sub

Private Sub chkUseMouseHook_Click()
    cmdApply.Enabled = True

    If chkUseMouseHook.value = Checked Then
        If Not isUseMouseHook Then
            If MsgBox("�˲�����ϵͳ�������, ���ܻ�����ȫ�������. ȷ��Ҫʹ�øù�����?" & vbCrLf & vbCrLf & "ע: �����ù��ܺ�, Ϊ����׼ȷģ�ⰴ������, ������������������������Զ�����.", vbQuestion + vbYesNo + vbDefaultButton2 + vbSystemModal) = vbYes Then
                isUseMouseHook = True
            Else
                chkUseMouseHook.value = Unchecked

            End If

        End If

    Else
        isUseMouseHook = False

    End If

End Sub

Private Sub chkWin_Click()
    cmdApply.Enabled = True

End Sub

Private Sub cmdAdd_Click()
    cmdApply.Enabled = True

    If Len(Trim(txtDiyKey.Text)) > 0 Then
        If CheckListbox(lstDiyKey, txtDiyKey.Text) Then
            If MsgBox("�����ִ����Ѿ������б���, �Ƿ������޸������볤�� ?", vbQuestion + vbYesNo + vbDefaultButton1 + vbSystemModal) = vbYes Then

                Dim Index As Long

                For Index = 0 To lstDiyKey.ListCount - 1

                    Dim strCode() As String

                    strCode = Split(lstDiyKey.List(Index), vbTab, -1, vbBinaryCompare)

                    If strCode(0) = txtDiyKey.Text Then
                        lstDiyKey.RemoveItem Index
                        lstDiyKey.AddItem txtDiyKey.Text + vbTab + comDiyLength.Text, Index
                        lstDiyKey.ListIndex = Index
                        Exit For

                    End If

                Next

            End If

        Else
            lstDiyKey.AddItem txtDiyKey.Text + vbTab + comDiyLength.Text
            lstDiyKey.ListIndex = lstDiyKey.ListCount - 1
            txtDiyKey.Text = ""
            txtDiyKey.SetFocus

        End If

    End If

End Sub

Private Sub cmdApply_Click()
    cmdApply.Enabled = False
    Call SaveSetting
    Call FrmMain.ShowResult

End Sub

Private Sub cmdCancel_Click()
    Call LoadSetting
    Call FrmMain.ShowResult
    Unload Me

End Sub

Private Sub cmdClear_Click()
    cmdApply.Enabled = True

    If MsgBox("ȷ��Ҫ��� <�������볤������> �б��ȫ��������?", vbQuestion + vbYesNo + vbDefaultButton2 + vbSystemModal) = vbYes Then
        lstDiyKey.Clear

    End If

End Sub

Private Sub cmdConfirm_Click()
    Call cmdApply_Click

    If HotKeyValue <> tmpHotKeyValue Or KeyValue <> tmpKeyValue Then
        If HotKeyValue = -1 And tmpHotKeyValue = &H8 And tmpKeyValue = vbKeyS Then
        ElseIf tmpHotKeyValue = -1 And HotKeyValue = &H8 And KeyValue = vbKeyS Then
        Else
            UnSetHotKey FrmMain.hwnd

            If SetHotKey(FrmMain.hwnd) = 0 Then
                MsgBox "����ע��ϵͳȫ���ȼ� (" + getHotKeyText + ") ʧ��, �����������ȼ�!", vbCritical + vbOKOnly + vbSystemModal

            End If

        End If

    End If

    Unload Me

End Sub

Private Sub cmdDefault_Click()

    Dim strSettingPath As String, isExist As Boolean

    strSettingPath = App.Path + "\Config.ini"
    isExist = False

    If Dir(strSettingPath, vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
        isExist = True
        Name strSettingPath As strSettingPath + ".bak"

    End If

    Call LoadSetting
    Call FrmMain.ShowResult

    If isExist Then
        Name strSettingPath + ".bak" As strSettingPath

    End If

End Sub

Private Sub cmdDelete_Click()
    cmdApply.Enabled = True

    If lstDiyKey.ListIndex >= 0 Then
        lstDiyKey.RemoveItem lstDiyKey.ListIndex

    End If

End Sub

Private Sub cmdModify_Click()
    cmdApply.Enabled = True

    If lstDiyKey.ListIndex >= 0 Then
        If Len(Trim(txtDiyKey.Text)) > 0 Then

            Dim temp As Long

            temp = lstDiyKey.ListIndex
            lstDiyKey.RemoveItem temp
            lstDiyKey.AddItem txtDiyKey.Text + vbTab + comDiyLength.Text, temp
            lstDiyKey.ListIndex = temp

        End If

    End If

End Sub

Private Sub comDiyLength_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Len(Trim(txtDiyKey.Text)) > 0 Then
            Call cmdAdd_Click

        End If

    End If

End Sub

Private Sub comHKey_Click()
    cmdApply.Enabled = True

End Sub

Private Sub comPrefixCode_Change()

    If Not isUserCodeLoading Then cbBox_Change comPrefixCode

End Sub

Private Sub comPrefixCode_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
    cbBox_KeyDown comPrefixCode, KeyCode, Shift

End Sub

Private Sub comPrefixCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ShowExample

End Sub

Private Sub comPrefixCode_LostFocus()
    cbBox_LostFocus comPrefixCode

End Sub

Private Sub comPwdLength_Click()
    cmdApply.Enabled = True

End Sub

Private Sub comSuffixCode_Change()

    If Not isUserCodeLoading Then cbBox_Change comSuffixCode

End Sub

Private Sub comSuffixCode_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
    cbBox_KeyDown comSuffixCode, KeyCode, Shift

End Sub

Private Sub comSuffixCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ShowExample

End Sub

Private Sub comSuffixCode_LostFocus()
    cbBox_LostFocus comSuffixCode

End Sub

Private Sub Form_Load()
    Me.Height = 5055: Me.Width = 6705
    Call Init_UI
    Call LoadSetting
    cmdApply.Enabled = False
    tmpHotKeyValue = HotKeyValue
    tmpKeyValue = KeyValue

End Sub

Private Sub HScrollTransparent_Change()
    cmdApply.Enabled = True
    lblValue.Caption = HScrollTransparent.value

End Sub

Private Sub HScrollTransparent_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
    lblValue.Caption = HScrollTransparent.value

End Sub

Private Sub HScrollTransparent_Scroll()
    cmdApply.Enabled = True
    lblValue.Caption = HScrollTransparent.value

End Sub

Private Sub ImgJohnnyJian_Click()
    OpenWebsite "http://johnnyjian.iteye.com/"

End Sub

Private Sub ImgKenshin_Click()
    OpenWebsite "http://www.k-zone.cn/zblog/"

End Sub

Private Sub Imgkisexu_Click()
    OpenWebsite "http://kisexu.com/"

End Sub

Private Sub Imgxlsdg_Click()
    OpenWebsite "http://blog.csdn.net/xlsdg"

End Sub

Private Sub Init_UI()
    picBack(0).Visible = True
    picBack(0).Left = 1800: picBack(0).Top = 240
    picBack(1).Visible = False
    picBack(1).Left = 1800: picBack(1).Top = 240
    picBack(3).Visible = False
    picBack(3).Left = 1800: picBack(3).Top = 240
    picBack(4).Visible = False
    picBack(4).Left = 1800: picBack(4).Top = 240
    picBack(6).Visible = False
    picBack(6).Left = 1800: picBack(6).Top = 240
    lstSet.AddItem "��������", 0
    lstSet.AddItem "�������", 1
    lstSet.AddItem "------------", 2
    lstSet.AddItem "��������", 3
    lstSet.AddItem "����ѡ��", 4
    lstSet.AddItem "------------", 5
    lstSet.AddItem "���볤��", 6
    lstSet.ListIndex = 0: tmpSel = 0

    Dim i As Long

    For i = 0 To 9
        comHKey.AddItem i
    Next

    For i = Asc("A") To Asc("Z")
        comHKey.AddItem Chr(i)
    Next

    For i = 6 To 32
        comPwdLength.AddItem i
        comDiyLength.AddItem i
    Next
    comDiyLength.ListIndex = 16 - 6

End Sub

Private Sub lblDonation_Click()
    OpenWebsite "http://kisexu.com/go/huamidonation"

End Sub

Private Sub lblUpdate_Click()
    OpenWebsite "http://code.google.com/p/flower-password-windows/wiki/ChangeLog"

End Sub

Private Sub lstDiyKey_Click()

    Dim i As Long

    For i = 0 To lstDiyKey.ListCount - 1

        If lstDiyKey.Selected(i) Then

            Dim strCode() As String

            strCode = Split(lstDiyKey.List(i), vbTab, -1, vbBinaryCompare)
            txtDiyKey.Text = strCode(0)
            comDiyLength.ListIndex = strCode(1) - 6

        End If

    Next

End Sub

Private Sub LstSet_Click()

    Dim i As Long

    For i = 0 To lstSet.ListCount - 1

        If lstSet.Selected(i) Then
            If InStr(lstSet.Text, "-") > 0 Then
                lstSet.ListIndex = tmpSel
            Else
                picBack(tmpSel).Visible = False
                picBack(i).Visible = True
                tmpSel = i

            End If

        End If

    Next

End Sub

Private Sub munExit_Click()
    Unload Me
    Unload FrmMain

End Sub

Private Sub munSetting_Click()
    Me.Show

End Sub

Private Sub munShow_Click()

    If Me.Visible Then
        MsgBox "���ȹر�ϵͳ���ô���, �ٳ�����ʾ���������!", vbCritical + vbOKOnly + vbSystemModal
    Else
        Call FrmMain.MidScreenShow

    End If

End Sub

Private Sub optClodeToHide_Click()
    cmdApply.Enabled = True

End Sub

Private Sub optCloseToExit_Click()
    cmdApply.Enabled = True

End Sub

Private Sub picUserCode_Click()

    If comSuffixCode.Enabled Then comSuffixCode.SetFocus

End Sub

Private Sub txtDiyKey_GotFocus()
    txtDiyKey.SelStart = 0
    txtDiyKey.SelLength = Len(txtDiyKey.Text)

End Sub

Private Sub txtDiyKey_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Len(Trim(txtDiyKey.Text)) > 0 Then
            Call cmdAdd_Click

        End If

    End If

End Sub
