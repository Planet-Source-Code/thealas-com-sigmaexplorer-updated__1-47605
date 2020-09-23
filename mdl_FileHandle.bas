Attribute VB_Name = "mdl_FileHandle"
'
' Module for handling files/folders, including few windows informations
' Replacement for Microsoft(R) ScriptingRuntime library
' Written by Sala Bojan
' Copyright(C) Hallsoft 2003, All rights reserved
'

Option Explicit
Const ModuleName As String = "mdl_filehandle"

Private Const gstrQUOTE$ = """" ' Funny thing you know...
Public FindFileCol As New Collection


Public Function Icon_AddToImageList(FileName As String, FType As String, IML As ImageList) As Long
'
' Adds icon to imagelist from any filename or directory,
' and stores its extension (not exe or ico format).
' Use it to specify an icon for listview when adding
' files to it.
'

On Error GoTo e
    Dim I&
    If IsNumeric(FType) Then FType = "XXX"  'It wont take numeric value as extension
    If FType = "EXE" Or FType = "ICO" Then 'If it is EXE or ICO, then make duplicates
        Call Icon_Extract(FileName, frm_Main.pic_Icon)
        Icon_AddToImageList = IML.ListImages.Add(, , frm_Main.pic_Icon.Image).Index
    Else 'If not, or folder, then get the icon
        For I = 1 To IML.ListImages.Count
            If IML.ListImages(I).Key = FType Then ' We have found the dup, so exit
                Icon_AddToImageList = I
                Exit Function
            End If
        Next I
        Call Icon_Extract(FileName, frm_Main.pic_Icon)
        Icon_AddToImageList = IML.ListImages.Add(, FType, frm_Main.pic_Icon.Image).Index
    End If
Exit Function
e:
    If Err = 7 Then MsgBox "No more memory for this folder. You must lower the icon cashe memory slider.", vbCritical, "Memory error": Unload frm_Main
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "Icon_AddToImageList": Resume Next
End Function
Public Sub Icon_Extract(FileName As String, PictureBox As PictureBox)
'
' Draws small icon to a picturebox from specified file
'
On Error GoTo e
    'It will just get an icon and draw it to the picturebox, pretty simple:
    Dim Icon As Long
    'It can be large, but we dont need some nice icons, just files
    Icon = SHGetFileInfo(FileName, 0&, IFileInfo, Len(IFileInfo), IFlags Or SHGFI_SMALLICON)
    If Icon <> 0 Then
      With PictureBox
        .Picture = LoadPicture("")
        Icon = ImageList_Draw(Icon, IFileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
      End With
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "mdl_FileHandle", "Icon_Extract": Resume Next
End Sub
Public Sub Icon_ExtractLarge(FileName As String, PictureBox As PictureBox)
'
' Same, but for large icon.
'

On Error GoTo e
    Dim Icon As Long
    Icon = SHGetFileInfo(FileName, 0&, IFileInfo, Len(IFileInfo), IFlags Or SHGFI_LARGEICON)
    If Icon <> 0 Then
      With PictureBox
        .Picture = LoadPicture("")
        Icon = ImageList_Draw(Icon, IFileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, "mdl_FileHandle", "Icon_ExtractLarge": Resume Next
End Sub

Public Function file_StripPath(t As String) As String
'
' Gets the last directory from path (t)
'

    Dim X%, ct%, Y$
    file_StripPath = t
    X = InStr(t, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, t, "\")
    Loop
    If ct > 0 Then file_StripPath = Mid$(t, ct + 1)
End Function

Public Function f_GetPath(Optional Side As Integer) As String
    If Side = 0 Then Side = frm_Main.ifSelected
    Select Case Side
    Case 1
        If Right(frm_Main.dir_Left.Path, 1) = "\" Then
            f_GetPath = Left(frm_Main.dir_Left.Path, Len(frm_Main.dir_Left.Path) - 1)
        Else
            f_GetPath = frm_Main.dir_Left.Path
        End If
    Case 2
        If Right(frm_Main.dir_Right.Path, 1) = "\" Then
            f_GetPath = Left(frm_Main.dir_Right.Path, Len(frm_Main.dir_Right.Path) - 1)
        Else
            f_GetPath = frm_Main.dir_Right.Path
        End If
    End Select
End Function
Public Function File_Open(FileName As String, Action As String) As Long
'
' Opens any file in its associated program
'

On Error GoTo e
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    File_Open = ShellExecute(Scr_hDC, Action, FileName, "", Left(FileName, 3), 1)
    If File_Open = 31 Then MsgBox "Failed to open this file, the associated program might not exist !", vbCritical, "Crap..."
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "File_Open": Resume Next
End Function

Public Function File_GetLWSize(ListView As ListView, sCol As Long, Optional FCount As Long) As Single
'
' Calculates a size of all selected files in LW
' sCol is column that holds size information
'
    Dim I&, fSize!, D$, C&
    With ListView
        While I < .ListItems.Count
            I = I + 1
            If .ListItems(I).Selected Then
                D = .ListItems(I).SubItems(2)
                If Not D = "" Then fSize = fSize + CSng(D): C = C + 1
            End If
        Wend
    End With
    FCount = C
    File_GetLWSize = fSize
End Function
Public Function File_GetLWSizeTotal(ListView As ListView, sCol As Long) As Long
'
' Gets the size of all files in ListView
'
    Dim I&, fSize&, D$
    With ListView
        While I < .ListItems.Count
            I = I + 1
            D = .ListItems(I).SubItems(2)
            If Not D = "" Then fSize = fSize + D
        Wend
    End With
    File_GetLWSizeTotal = fSize
End Function

Public Function File_Move(srcPath As String, DstPath As String) As Long
'
' Moves the file or directory to spec. path
'

On Error GoTo e
    Dim FileOperation As SHFILEOPSTRUCT    'Wanted operation
    
    srcPath = srcPath & Chr$(0) & Chr$(0)
    With FileOperation
       .wFunc = 1 'Move
       .pFrom = srcPath
       .pTo = DstPath
       .fFlags = FOF_SILENT
       If frm_Dirs.ch_NoConfirmation.Value Then .fFlags = .fFlags Or FOF_NOCONFIRMATION
       If frm_Dirs.ch_Rename.Value Then .fFlags = .fFlags Or FOF_RENAMEONCOLLISION
    End With
    File_Move = SHFileOperation(FileOperation)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_move": Resume Next
End Function
Public Function File_Copy(srcPath As String, DstPath As String) As Long
'
' Copy the file or dir to spec. path
'

On Error GoTo e
    Dim FileOperation As SHFILEOPSTRUCT    'Wanted operation
    
    srcPath = srcPath & Chr$(0) & Chr$(0)
    With FileOperation
       .wFunc = FO_COPY 'Copy
       .pFrom = srcPath
       .pTo = DstPath
       .fFlags = FOF_SILENT
       If frm_Dirs.ch_NoConfirmation.Value Then .fFlags = .fFlags Or FOF_NOCONFIRMATION
       If frm_Dirs.ch_Rename.Value Then .fFlags = .fFlags Or FOF_RENAMEONCOLLISION
    End With
    File_Copy = SHFileOperation(FileOperation)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_copy": Resume Next
End Function

Public Function File_Ex(FileName As String) As Boolean
'
' Checks if spec. file exists
'

On Error GoTo e
    Dim RetCode As Integer
    Dim OpenFileStructure As OFSTRUCT

    Const OF_EXIST = &H4000
    Const FILE_NOT_FOUND = 2

    RetCode = OpenFile(FileName$, OpenFileStructure, OF_EXIST)
    If OpenFileStructure.nErrCode = FILE_NOT_FOUND Then
        File_Ex = False
    Else
        If Not OpenFileStructure.nErrCode = 5 Then
            If Not OpenFileStructure.nErrCode = 3 Then
                File_Ex = True
            End If
        End If
    End If
    
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_ex": Resume Next
End Function

Public Function File_ParseName(Path As String) As String
'
' Gets a filename from path
'

On Error GoTo e
    Dim A
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            File_ParseName = Mid$(Path, A + 1)
            Exit Function
        End If
    Next A
    File_ParseName = Path
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_parsename": Resume Next
End Function


Public Function File_Delete(srcPath As String) As Long
'
' Deletes a file or dir
'

On Error GoTo e
    Dim FileOperation As SHFILEOPSTRUCT    'Wanted operation
    
    srcPath = srcPath & Chr$(0) & Chr$(0)
    With FileOperation
       .wFunc = FO_DELETE 'Del
       .pFrom = srcPath
       .fFlags = FOF_SILENT Or FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
    End With
    File_Delete = SHFileOperation(FileOperation)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_delete": Resume Next
End Function
Public Sub File_CreateLink(ByVal strLinkPath As String, ByVal strGroupName As String, ByVal strLinkArguments As String, ByVal strLinkName As String, ByVal fPrivate As Boolean, sParent As String, Optional ByVal fLog As Boolean = True)
'
' Creates a shortcut
'
    Dim lREt       As Boolean   ' Return
    strLinkName = strUnQuoteString(strLinkName)
    strLinkPath = strUnQuoteString(strLinkPath)
    If StrPtr(strLinkArguments) = 0 Then strLinkArguments = ""
    lREt = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments, fPrivate, sParent)    ' the path should never be enclosed in double quotes
End Sub


Private Function strUnQuoteString(ByVal strQuotedString As String)
    ' For use in filesearch
    ' It removes '?"'...
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
            ' It's quoted.  Get rid of the quotes.
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function



Public Function File_ParsePath(Path As String) As String
'
' Gets path from file
'

On Error GoTo e
    Dim A&
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            If Mid$(Path, A, 1) = "\" Then
                File_ParsePath = LCase$(Left$(Path, A - 1) & "\")
            Else
                File_ParsePath = LCase$(Left$(Path, A - 1) & "/")
            End If
            Exit Function
        End If
    Next A
Exit Function
e:
    Err_Raise Err.Number, Err.Description, "mdl_filehandle", "file_parsepath": Resume Next
End Function

Public Function File_ShowProps(FileName As String, hWnd As Long) As Long
'
' Shows properties dialog
'
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = hWnd
        .lpVerb = "properties" ' I only know this one... if u know more tell me (file related) !
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
    File_ShowProps = SEI.hInstApp
End Function

Public Function File_FindFiles(FileSpec As String, Optional Recursive As Boolean = True)
'
' Finds the files and puts them to collection: FindFileCol
' FileSpec must be path and search string: "C:\windows\*.exe"
' Special thanks to Intech Solutions !
'

    Static lLevel As Long
  
    lLevel = lLevel + 1
    If lLevel = 1 Then
        Set FindFileCol = Nothing
        Set FindFileCol = New Collection
    End If
    On Error GoTo 0
    
    Dim lFind As Long, lMatch As Long
    Dim tInfo As WIN32_FIND_DATA
    ' Scan Subdirs First
    If Recursive Then
        Dim sDirSpec As String
        Dim sSpec As String
        
        sSpec = File_ParseName(FileSpec)
        sDirSpec = File_ParsePath(FileSpec)
        lFind = FindFirstFile(sDirSpec & "*.*", tInfo)
        lMatch = 99
        Do While lFind > 0 And lMatch > 0
            If (tInfo.dwFileAttributes And vbDirectory) Then  '** > 0
                Dim sDirName As String
                sDirName = sNT(tInfo.cFileName)
                If sDirName <> "." And sDirName <> ".." Then
                    File_FindFiles sDirSpec & sDirName & "\" & sSpec, Recursive
                End If
            End If
            lMatch = FindNextFile(lFind, tInfo)
        Loop
        FindClose lFind
    End If
    
    lFind = FindFirstFile(FileSpec, tInfo)
    lMatch = 99
    Do While lFind > 0 And lMatch > 0
        If Not (tInfo.dwFileAttributes And vbDirectory) > 0 Then
            FindFileCol.Add File_ParsePath(FileSpec) & sNT(tInfo.cFileName)
        End If
        lMatch = FindNextFile(lFind, tInfo)
        On Error Resume Next
        On Error GoTo 0
    Loop
    FindClose lFind
    lLevel = lLevel - 1
End Function
Private Function sNT(sString As String) As String
    ' For use in filesearch
    Dim iNullLoc As Integer
    iNullLoc = InStr(sString, Chr(0))
    If iNullLoc > 0 Then
        sNT = Left(sString, iNullLoc - 1)
    Else
        sNT = sString
    End If
End Function


Public Function File_FixPath(Path As String, Optional AddSlash As Boolean = False) As String
'
' Adds "\" or removes it to filename
'

    If Right(Path, 1) = "\" Then
        If AddSlash Then
            File_FixPath = Path
        Else
            File_FixPath = Left(Path, Len(Path) - 1)
        End If
    Else
        If AddSlash Then
            File_FixPath = File_FixPath & "\"
        Else
            File_FixPath = Path
        End If
    End If
End Function

Public Sub image_GetInfo(ByVal FileName As String, Optional X As Long, Optional Y As Long, Optional Depth As Long)
'
' Gets image size, and color depth, a big part is not mine !
'

    Dim bBuf(65535) As Byte
    Dim iFN%, iType&
    
    iFN = FreeFile
    Open FileName For Binary As iFN
        Get #iFN, 1, bBuf()
    Close iFN
    
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        ' GIF
        iType = 1
        X = Mult(bBuf(6), bBuf(7))
        Y = Mult(bBuf(8), bBuf(9))
        ' get bit depth
        Depth = (bBuf(10) And 7) + 1
    End If
    
    If bBuf(0) = 66 And bBuf(1) = 77 Then
        ' BMP
        iType = 2
        X = Mult(bBuf(18), bBuf(19))
        Y = Mult(bBuf(22), bBuf(23))
        ' get bit depth
        Depth = bBuf(28)
    End If

    If iType = 0 Then
    ' If the file is not one of the above type then
    ' check to see if it is a JPEG file
        Dim lPos&
        Do
            ' loop through looking for the byte sequence FF,D8,FF
            ' which marks the begining of a JPEG file
            ' lPos will be left at the postion of the start
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) Or (lPos >= 65525) Then Exit Do
            ' move our pointer up
            lPos = lPos + 1
            ' and continue
        Loop
        lPos = lPos + 2
        If lPos >= 65525 Then Exit Sub
        Do
        ' Loop through the markers until we find the one
        ' starting with FF,C0 which is the block containing the
        ' image information
            Do
                ' Loop until we find the beginning of the next marker
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= 65525 Then Exit Sub
            Loop
            ' Move pointer up
            lPos = lPos + 1
            Select Case bBuf(lPos)
                Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, &HCD To &HCF
                ' we found the right block
                Exit Do
            End Select
            ' otherwise keep looking
            lPos = lPos + Mult(bBuf(lPos + 2), bBuf(lPos + 1))
            ' check for end of buffer
            If lPos >= 65525 Then Exit Sub
        Loop
        
        ' If we've gotten this far it is a JPEG and we are ready
        ' to grab the information.
        iType = 3
        X = Mult(bBuf(lPos + 5), bBuf(lPos + 4))
        Y = Mult(bBuf(lPos + 7), bBuf(lPos + 6))
        ' get the color depth
        Depth = bBuf(lPos + 8) * 8
    End If
End Sub
Private Function Mult(lsb As Byte, msb As Byte) As Long
    Mult = lsb + (msb * CLng(256))
End Function

Public Function win_Is95() As Boolean
'
' Checks if OS is Windows95
'

    Dim s As OSVERSIONINFOEX
    s.dwOSVersionInfoSize = Len(s)
    GetVersionEx s
    If s.dwMajorVersion = 4 And s.dwMinorVersion = 0 And s.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS Then
        win_Is95 = True
    Else
        win_Is95 = False
    End If
End Function

Public Function win_Function_Exist(sModule As String, sFunction As String) As Boolean
'
' Checks if spec. function exists.
' be sure to add .dll at the end :)
'

    Dim hHandle As Long
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        If GetProcAddress(hHandle, sFunction) = 0 Then
            win_Function_Exist = False
        Else
            win_Function_Exist = True
        End If
        FreeLibrary hHandle
    Else
        If GetProcAddress(hHandle, sFunction) <> 0 Then
            win_Function_Exist = True
        End If
    End If
End Function
Public Function win_C32to64(ByVal lLo As Long, ByVal lHi As Long) As Double
'
' Gets a Double from LargeInt
'
    
    Dim dLo As Double
    Dim dHi As Double
    
    If lLo < 0 Then
        dLo = (2 ^ 32) + lLo
    Else
        dLo = lLo
    End If
    If lHi < 0 Then
        dHi = (2 ^ 32) + lHi
    Else
        dHi = lHi
    End If
    
    win_C32to64 = (dLo + (dHi * (2 ^ 32)))
End Function

Public Function win_GetDirectory() As String
'
' Gets a windows directory
'
    Dim tWin$, L&
    tWin = String$(256, 0)
    L = GetWindowsDirectory(tWin, Len(tWin))
    win_GetDirectory = Left(tWin, L)
End Function


Public Function win_SystemInfo(sInfo As SYSTEMINFO, sDrive As String)
'
' Simple system info
'
On Error GoTo e
    ' This will get all the info, just select the drive
    Dim dType&, SectorsPerCluster&, BytesPerSector&, TotalClusters&, FreeClusters&
    Dim FreeSpace#, TotalSpace#, TotalSectors&, FreeSectors&
    Dim FreeSpaceEx#, liFS As LARGE_INTEGER
    Dim TotalSpaceEx#, liTS As LARGE_INTEGER
    Dim TotalFreeSpaceEx#, liTFS As LARGE_INTEGER
    Dim vName$, vID&, vFile$, vFileFlags&, mcl&
    Dim cName$, mStatus As MEMORYSTATUS, mPageFile!, cPos!
    
    cPos = 1
    dType = GetDriveType(sDrive)
    Select Case dType
        Case DRIVE_UNKNOWN: sInfo.sysDriveType = "Unknown"
        Case DRIVE_NO_ROOT_DIR: sInfo.sysDriveType = "No Root Dir"
        Case DRIVE_REMOVABLE: sInfo.sysDriveType = "Removable"
        Case DRIVE_FIXED: sInfo.sysDriveType = "Fixed"
        Case DRIVE_REMOTE: sInfo.sysDriveType = "Remote"
        Case DRIVE_RAMDISK: sInfo.sysDriveType = "Ram Disk"
        Case DRIVE_CDROM: sInfo.sysDriveType = "CD Rom"
    End Select
    
    cPos = 2
    GetDiskFreeSpace sDrive, SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters
    ' Calculate the drive space
    TotalSectors = SectorsPerCluster * TotalClusters
    FreeSectors = SectorsPerCluster * FreeClusters
    FreeSpace = (FreeSectors * BytesPerSector) / 1048576
    TotalSpace = (TotalSectors * BytesPerSector) / 1048576
    ' if old win, then use the normal api
    If Not win_Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") Then
        cPos = 2.1
        sInfo.sysFreeSpace = FreeSpace
        sInfo.sysTotalClusters = TotalSpace
    ' Else, use the advanced stuff
    Else
        cPos = 2.2
        GetDiskFreeSpaceEx sDrive, liFS, liTS, liTFS
        ' You just need to convert the high and low values
        FreeSpaceEx = win_C32to64(liFS.LowPart, liFS.HighPart)
        TotalSpaceEx = win_C32to64(liTS.LowPart, liTS.HighPart)
        sInfo.sysFreeSpace = FreeSpaceEx
        sInfo.sysTotalSpace = TotalSpaceEx
    End If
    sInfo.sysTotalClusters = TotalClusters
    sInfo.sysTotalSectors = TotalSectors
    sInfo.sysBytesPerSector = BytesPerSector
    sInfo.sysSectorsPerCluster = SectorsPerCluster
    
    cPos = 3
    vName = String(256, 0) ' Fill it, if you will use Len
    vFile = String(256, 0)
    GetVolumeInformation sDrive, vName, Len(vName), vID, mcl, vFileFlags, vFile, Len(vFile)
    sInfo.sysDriveName = vName
    sInfo.sysFileSystem = vFile
    sInfo.sysSerialID = vID
    
    cPos = 4
    cName = Space(32)
    GetComputerName cName, 32
    sInfo.sysComputerName = cName
    
    ' The memo status may be false on some computers
    cPos = 5
    GlobalMemoryStatus mStatus
    sInfo.sysPsysicalMemory = (CDbl(mStatus.dwAvailPhys) * 100) / mStatus.dwTotalPhys
    sInfo.sysVirtualMemory = (CDbl(mStatus.dwAvailVirtual) * 100) / mStatus.dwTotalVirtual
    sInfo.sysPageFile = (CDbl(mStatus.dwAvailPageFile) * 100) / mStatus.dwTotalPageFile
    sInfo.sysMemoryLoad = mStatus.dwMemoryLoad
Exit Function
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "win_systeminfo: " & cPos
    Resume Next
End Function
Public Function GetFindFile(FileName) As WIN32_FIND_DATA
'
' Easier way to find fileinfo
'

    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    plngFirstFileHwnd = FindFirstFile(FileName, Win32Data)
    ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
        GetFindFile.cFileName = "Error"   ' If file was not found
    Else
        GetFindFile = Win32Data
    End If
    plngRtn = FindClose(plngFirstFileHwnd) '
End Function

Public Function Zeros(Num As Long) As String
'
' Creates 00021 from 21
'
    Dim I&, N$
    For I = 1 To 6 - Len(CStr(Num))
        N = N & "0"
    Next I
    Zeros = N & Num
End Function
