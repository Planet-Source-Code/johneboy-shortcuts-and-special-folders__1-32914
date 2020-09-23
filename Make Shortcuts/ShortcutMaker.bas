Attribute VB_Name = "ShortcutMaker"
Const CSIDL_DESKTOP = &H0
Const CSIDL_RECENT = &H8
Const CSIDL_STARTMENU = &HB

Public Type HoopDee
   cb As Long
   abID As Byte
End Type

Public Type ITEMIDLIST
   mkid As HoopDee
End Type

Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As Any) As Long

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        Path$ = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
Public Function strFile(FilePath As String, AppName As String) As String
Dim intcount As Integer
Dim intcount2 As Integer

For intcount = Len(FilePath) To 1 Step -1
If Mid(FilePath, intcount, 1) = "\" Then
intcount = 1
Else
intcount2 = intcount2 + 1
End If
Next intcount
AppName = Right(FilePath, intcount2)
End Function


Public Function MakeDesktopShortcut(OrigFile As String)
Dim AppName As String
Call SHAddToRecentDocs(2, OrigFile)
   SleepEx 200, False
strFile OrigFile, AppName
FileCopy "" + GetSpecialfolder(CSIDL_RECENT) + "\" + AppName + ".lnk", "" + GetSpecialfolder(CSIDL_DESKTOP) + "\" + AppName + ".lnk"
End Function

Public Function MakeStartMenuShortcut(OrigFile As String)
Dim AppName As String
Call SHAddToRecentDocs(2, OrigFile)
   SleepEx 200, False
strFile OrigFile, AppName
FileCopy "" + GetSpecialfolder(CSIDL_RECENT) + "\" + AppName + ".lnk", "" + GetSpecialfolder(CSIDL_STARTMENU) + "\" + AppName + ".lnk"
End Function

Public Function MakeStartMenuFolderShortcut(OrigFile As String, FolderName As String)
Dim AppName As String
Call SHAddToRecentDocs(2, OrigFile)
CreateDirectoryEx "" + GetSpecialfolder(CSIDL_STARTMENU) + "\Programs", "" + GetSpecialfolder(CSIDL_STARTMENU) + "\programs\" + FolderName + "", ByVal 0&
   SleepEx 200, False
strFile OrigFile, AppName
FileCopy "" + GetSpecialfolder(CSIDL_RECENT) + "\" + AppName + ".lnk", "" + GetSpecialfolder(CSIDL_STARTMENU) + "\Programs\" + FolderName + "\" + AppName + ".lnk"
End Function
