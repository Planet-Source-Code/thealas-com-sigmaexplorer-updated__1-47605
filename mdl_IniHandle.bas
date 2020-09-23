Attribute VB_Name = "mdl_IniHandle"
'
' Module for advanced Get/Write PrivateProfileString, writing and
' reading INI files, including error handle.
' Written by Sala Bojan
' Copyright(C) Hallsoft 2003, All rights reserved
'

Option Explicit
Const WindowName As String = "mdl_inihandle"


' This is all you need for r/w, just have a look at LoadInfo on how to use getppstring !
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub SaveInformation(Filename As String, Section As String, KeyName As String, Value As String)
'
' Put data to an INI file
'

On Error GoTo e
    WritePrivateProfileString Section, KeyName, Value, Filename
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "ini_handle", "saveinformation": Resume Next
End Sub
Public Function LoadInformation(Filename As String, Section As String, KeyName As String) As String
'
' Get data form an INI file, "" for error
'

On Error GoTo e
    Dim strResult As String * 150, G&
    G = GetPrivateProfileString(Section, KeyName, Filename, strResult, Len(strResult), Filename)
    LoadInformation = Trim(strResult)
    If G = 93 Then
        LoadInformation = ""
    End If
Exit Function
e:
    LoadInformation = ""
    Err_Raise Err.Number, Err.Description, "ini_handle", "loadinformation": Resume Next
End Function
Public Function OI(KeyValue As String, Optional Default As String) As String
'
' Opens Setting from SETTINGS.INI in app dir
'
    OI = LoadInformation(App.Path & "\SETTINGS.INI", "Settings", KeyValue)
    If OI = "" Then OI = Default
End Function

Public Sub SI(KeyValue As String, Value As String)
'
' Saves Setting
'
    SaveInformation App.Path & "\SETTINGS.INI", "Settings", KeyValue, Value
End Sub


