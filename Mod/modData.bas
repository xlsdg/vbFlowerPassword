Attribute VB_Name = "modData"
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

Public Function LoadDomains() As String

    Dim bytData() As Byte

    bytData = LoadResData("DOMAINS", "DATA")
    LoadDomains = StrConv(bytData, vbUnicode)

End Function

Public Function LoadPasswords() As String

    Dim bytData() As Byte

    bytData = LoadResData("PASSWORDS", "DATA")
    LoadPasswords = StrConv(bytData, vbUnicode)

End Function
