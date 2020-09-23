Attribute VB_Name = "mdl_TextHandle"
Option Explicit
Const WindowName As String = "mdl_texthandle"
'
' Module for handling richtext control.
' Written by Sala Bojan
' Copyright(C) Hallsoft 2003, All rights reserved
'


Public Function Get_TotalLines(RichTextBox As RichTextBox) As Long
'
' Gets total lines from richtextbox
' use it in selchange event
'
    
    Dim TotalLines&
    TotalLines = SendMessage(RichTextBox.hWnd, EM_GETLINECOUNT, 0, 0&)
    Get_TotalLines = Format(TotalLines, "###,###,###,###")
End Function

Public Function Get_CurrentLine(RichTextBox As RichTextBox) As Long
'
' Same, it gets current line
'
    Dim CurrentLine&
    CurrentLine = SendMessage(RichTextBox.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    Get_CurrentLine = Format(CurrentLine, "###,###,###,###")
End Function

