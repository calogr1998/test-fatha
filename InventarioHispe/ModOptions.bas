Attribute VB_Name = "modOptions"
' Catalyst Internet Mail Control 4.5
' Copyright 2002-2006, Catalyst Development Corporation
' All rights reserved
'
' This product is licensed to you pursuant to the terms of the
' Catalyst license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'
Option Explicit

Private Const KEY_PRODUCT = "SocketTools Examples"

Global g_strSenderName As String
Global g_strSenderAddress As String
Global g_strOrganization As String
Global g_bRelayMessages As Boolean
Global g_strRelayServer As String
Global g_nRelayPort As Long

Sub LoadOptions()
    On Error Resume Next
    g_strSenderName = GetSetting(KEY_PRODUCT, App.Title, "SenderName")
    g_strSenderAddress = GetSetting(KEY_PRODUCT, App.Title, "SenderAddress")
    g_strOrganization = GetSetting(KEY_PRODUCT, App.Title, "Organization")
    g_bRelayMessages = CBool(GetSetting(KEY_PRODUCT, App.Title, "RelayMessages", "False"))
    g_strRelayServer = GetSetting(KEY_PRODUCT, App.Title, "RelayServer")
    g_nRelayPort = CLng(GetSetting(KEY_PRODUCT, App.Title, "RelayPort", "25"))
End Sub

Sub SaveOptions()
    SaveSetting KEY_PRODUCT, App.Title, "SenderName", g_strSenderName
    SaveSetting KEY_PRODUCT, App.Title, "SenderAddress", g_strSenderAddress
    SaveSetting KEY_PRODUCT, App.Title, "Organization", g_strOrganization
    SaveSetting KEY_PRODUCT, App.Title, "RelayMessages", CStr(g_bRelayMessages)
    SaveSetting KEY_PRODUCT, App.Title, "RelayServer", g_strRelayServer
    SaveSetting KEY_PRODUCT, App.Title, "RelayPort", CStr(g_nRelayPort)
End Sub
