Attribute VB_Name = "mdl_ErrHandle"
'
' Module for writing error log file
' Written by Sala Bojan
' Copyright(C) Hallsoft 2003, All rights reserved
'

Option Explicit
Const WindowName As String = "mdl_errhandle"


Public Sub Log_Append(Text As String)
'
' Appends a text at "PROGRAM.LOG" file in appdir
' Use err_raise for errors
'
    
    ' Use this stuff for putting text in error log
    Open App.Path & "\PROGRAM.LOG" For Append As #2
        Print #2, "DATE: " & Date & ",TIME: " & Time & " - " & Text
    Close #2
e:
    Exit Sub
End Sub

Public Sub Log_Clear()
'
' Clears a log file
'

    ' This clears the file
    Open App.Path & "\PROGRAM.LOG" For Output As #3
        Print #3, ""
    Close #3
End Sub

Public Sub Err_Raise(Err_Num As String, Err_Description As String, Err_Module As String, Err_Function As String)
'
' Writes a detailed error information to a log file
' Use err object for getting info
'
    Log_Append "ERROR " & Err_Num & " - " & Err_Module & "\" & Err_Function & " ::: " & Err_Description
End Sub

