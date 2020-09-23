Attribute VB_Name = "mdl_ZipHandle"
Option Explicit
Const WindowName As String = "mdl_ziphandle"
'
' IMPORTANT:
' This code is just using zipit.dll, the newest version of Info-Zip Library
' so the REAL code is made in C, but some Dana Seaman has made UNZIP code in
' VB by using unace.dll and unrar.zip, I have that code, but this one is much
' easier to understand, and I've included all the needed dlls in the zip file.
' There is also licence.txt for conditions, if you plan to distribute the dlls.
' You only need this module for working with it.
' Written by Sala Bojan
' Copyright(C) Hallsoft 2003, All rights reserved
' Special thanks to Richsoft Computing
'

Public Declare Function AddFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Public Declare Function ExtractFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal ExtrDir As String, ByVal UseDirInfo As Boolean, ByVal Overwrite As Boolean, ByVal Action As Integer) As Boolean
Public Declare Function DeleteFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String) As Boolean

Public Type ZIPPROPS
    Version As Integer
    Flag As Integer
    CompressionMethod As Integer
    Time As Integer
    Date As Integer
    CRC32 As Long
    CompressedSize As Long
    UncompressedSize As Long
    FileNameLength As Integer
    ExtraFieldLength As Integer
    Filename As String
End Type
Public Enum ZIPLEVEL
    zipStore = 0
    zipLevel1 = 1
    zipSuperFast = 2
    zipFast = 3
    zipLevel4 = 4
    zipNormal = 5
    zipLevel6 = 6
    zipLevel7 = 7
    zipLevel8 = 8
    zipMax = 9
End Enum
Public Enum ZIPACTION
    zipDefault = 1
    zipFreshen = 2
    zipUpdate = 3
End Enum

Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50

Public Archive As New Collection
Public ZipFile As ZIPPROPS
Public CompLevel As ZIPLEVEL
Public ArchiveFilename As String
Public Sub Zip_ReadArchive(ZipFileName As String)
'
' Reads an archive adding ints files to Archive collection
'

    Dim Sig&, ZipStream&, Res&, zFile As ZIPPROPS, Name$, I&
    
    If Not win_Function_Exist("zipit.dll", "AddFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "ExtractFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "DeleteFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub

    For I = Archive.Count To 1 Step -1
        Archive.Remove I
    Next I
    ZipStream = FreeFile
    Open ZipFileName For Binary As ZipStream
        Do While True
            Get ZipStream, , Sig
            'See if the file header has been found
            If Sig = LocalFileHeaderSig Then
                'Read each part of the file header
                Get ZipStream, , ZipFile.Version
                Get ZipStream, , ZipFile.Flag
                Get ZipStream, , ZipFile.CompressionMethod
                Get ZipStream, , ZipFile.Time
                Get ZipStream, , ZipFile.Date
                Get ZipStream, , ZipFile.CRC32
                Get ZipStream, , ZipFile.CompressedSize
                Get ZipStream, , ZipFile.UncompressedSize
                Get ZipStream, , ZipFile.FileNameLength
                Get ZipStream, , ZipFile.ExtraFieldLength
                'Get the filename
                'Set up a empty string so the right number of
                'bytes is read
                Name = String$(ZipFile.FileNameLength, " ")
                Get ZipStream, , Name
                ZipFile.Filename = Mid$(Name, 1, ZipFile.FileNameLength)
                'Move on through the archive
                'Skipping extra space, and compressed data
                Seek ZipStream, (Seek(ZipStream) + ZipFile.ExtraFieldLength)
                Seek ZipStream, (Seek(ZipStream) + ZipFile.CompressedSize)
                'Add the fileinfo to the collection
                AddEntry ZipFile
            Else
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    'All the filenames have been found so
                    'exit the loop
                    Exit Do
                Else
                    If Sig = EndCentralDirSig Then
                        Exit Do
                    End If
                End If
            End If
        Loop
    Close ZipStream
    ArchiveFilename = ZipFileName
End Sub
Public Function Zip_GetEntry(ByVal Index As Long) As mcls_ZipEntry
'
' Gets properties for a spec. filename (archive index)
'
    Set Zip_GetEntry = Archive(Index)
End Function
Private Sub AddEntry(zFile As ZIPPROPS)
    Dim xFile As New mcls_ZipEntry
'
' Adds a file from the archive into the collection
' It does not add entry that are just folders
'
    If File_ParseName(zFile.Filename) <> "" Then
        xFile.Version = zFile.Version
        xFile.Flag = zFile.Flag
        xFile.CompressionMethod = zFile.CompressionMethod
        xFile.CRC32 = zFile.CRC32
        xFile.FileDateTime = GetDateTime(zFile.Date, zFile.Time)
        xFile.CompressedSize = zFile.CompressedSize
        xFile.UncompressedSize = zFile.UncompressedSize
        xFile.FileNameLength = zFile.FileNameLength
        xFile.Filename = zFile.Filename
        xFile.ExtraFieldLength = zFile.ExtraFieldLength
    End If
    Archive.Add xFile
End Sub
Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
'
' Converts the file date/time dos stamp from the archive
' in to a normal date/time string
'
    Dim r As Long
    Dim FTime As FILETIME
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String
    ' Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    ' Convert the file time into a standard time
    r = FileTimeToSystemTime(FTime, Sys)
    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond
    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function

Public Sub Zip_Add2Archive(ZipFileName As String, Files As Collection, Action As ZIPACTION, StorePathInfo As Boolean, RecurseSubFolders As Boolean, UseDOS83 As Boolean, CompressionLevel As ZIPLEVEL)
'
' Archive creation/modification sub
'
    Dim I&, Result&, FilesToAdd As Collection
    
    If Not win_Function_Exist("zipit.dll", "AddFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "ExtractFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "DeleteFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    
    'Check to see if there are any files in the archive
    'if not delete the file so there are not error messages
    If Archive.Count = 0 Then
        If Dir$(ZipFileName, vbHidden Or vbSystem Or vbReadOnly) <> "" Then
            Kill ZipFileName
        End If
    End If
    Set FilesToAdd = FindFiles(Files, RecurseSubFolders)
    frm_Main.pb_ProgressBar.Max = FilesToAdd.Count
    For I = 1 To FilesToAdd.Count
        DoEvents
        frm_Main.pb_ProgressBar.Value = I
        If AddFile(ZipFileName, FilesToAdd(I), StorePathInfo, UseDOS83, Action, CompressionLevel) Then
            Result = Result + 1
        Else
            MsgBox "Failed to add " & FilesToAdd(I) & " !", vbCritical, "Error"
        End If
    Next I
    frm_Main.pb_ProgressBar.Value = 0
End Sub
Private Function FindFiles(Files As Collection, Recurse As Boolean)
'
' Finds all the files matching the specification
' RECURSIVE FOLDER SEARCH NOT YET IMPLEMENTED
'
    Dim Result As New Collection, Path$, r$
    Dim I As Long
    For I = 1 To Files.Count
        Path = File_ParsePath(Files(I))
        r = Dir$(Files(I), vbHidden Or vbSystem Or vbReadOnly)
        Do Until r = ""
            Result.Add Path & r
            r = Dir$()
        Loop
    Next I
    Set FindFiles = Result
End Function

Public Sub Zip_DeleteFiles(Files As Collection)
'
' Deletes files from an open archive
'
    Dim FilesToDelete As Collection
    Dim I As Long
    Set FilesToDelete = SelectFiles(Files)
    'Extract each file in turn
    For I = 1 To FilesToDelete.Count
        DoEvents
        If Not DeleteFile(ArchiveFilename, FilesToDelete(I)) Then
            MsgBox "Failed to delete file: " & FilesToDelete(I), vbExclamation, "Damn..."
        End If
    Next I
End Sub

Private Function SelectFiles(Files As Collection) As Collection
'
' Selects files from a wildcard specification
' Wildcards only corrispond to the filename and not the path
'
    Dim I As Long
    Dim j As Long
    Dim Result As New Collection
    For I = 1 To Files.Count
        For j = 1 To Archive.Count
            'Check the pattern, ignoring case
            If LCase$(File_ParseName(Zip_GetEntry(j).Filename)) Like LCase$(Files(I)) Then
                Result.Add Zip_GetEntry(j).Filename
            End If
        Next j
    Next I
    Set SelectFiles = Result
End Function
Public Sub Zip_Extract(Files As Collection, ByVal Action As ZIPACTION, ByVal UsePathInfo As Boolean, ByVal Overwrite As Boolean, ByVal Path As String)
'
' Extracts open archive to spec. dir
'
    Dim FilesToExtract As Collection
    Dim I As Long
    Dim Result As Long
    
    If Not win_Function_Exist("zipit.dll", "AddFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "ExtractFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    If Not win_Function_Exist("zipit.dll", "DeleteFile") Then MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    
    'First find the files which match the patterns
    'specified in the collection
    Set FilesToExtract = SelectFiles(Files)
    frm_Main.pb_ProgressBar.Max = FilesToExtract.Count + 1
    For I = 1 To FilesToExtract.Count
        frm_Main.pb_ProgressBar.Value = I
        DoEvents
        If Not ExtractFile(ArchiveFilename, CStr(FilesToExtract(I)), Path, UsePathInfo, Overwrite, Action) Then
            MsgBox "Failed to extract file: " & FilesToExtract(I) & " !", vbExclamation, "Damn..."
        End If
    Next I
    frm_Main.pb_ProgressBar.Value = 0
End Sub


