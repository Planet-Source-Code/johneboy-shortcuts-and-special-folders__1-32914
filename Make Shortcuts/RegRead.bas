Attribute VB_Name = "RegRead"


Private Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long
     
     Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long


Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long


Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hkey As Long) As Long


Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long


Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hkey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long


Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hkey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long


Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
    Private Const REG_NONE = 0
    Private Const REG_SZ = 1
    Private Const REG_EXPAND_SZ = 2
    Private Const REG_BINARY = 3
    Private Const REG_DWORD = 4
    Private Const REG_DWORD_LITTLE_ENDIAN = 4
    Private Const REG_DWORD_BIG_ENDIAN = 5
    Private Const REG_LINK = 6
    Private Const REG_MULTI_SZ = 7
    Private Const REG_RESOURCE_LIST = 8
    Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
    Private Const REG_CREATED_NEW_KEY = &H1
    Private Const REG_OPENED_EXISTING_KEY = &H2
    Private Const REG_WHOLE_HIVE_VOLATILE = &H1
    Private Const REG_REFRESH_HIVE = &H2
    Private Const REG_NOTIFY_CHANGE_NAME = &H1
    Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4
    Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
    Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
    Private Const REG_OPTION_RESERVED = 0
    Private Const REG_OPTION_NON_VOLATILE = 0
    Private Const REG_OPTION_VOLATILE = 1
    Private Const REG_OPTION_CREATE_LINK = 2
    Private Const REG_OPTION_BACKUP_RESTORE = 4
    Private Const STANDARD_RIGHTS_READ = &H20000
    Private Const STANDARD_RIGHTS_WRITE = &H20000
    Private Const STANDARD_RIGHTS_EXECUTE = &H20000
    Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Private Const STANDARD_RIGHTS_ALL = &H1F0000
    Private Const delete = &H10000
    Private Const READ_CONTROL = &H20000
    Private Const WRITE_DAC = &H40000
    Private Const WRITE_OWNER = &H80000
    Private Const SYNCHRONIZE = &H100000
    Private Const KEY_QUERY_VALUE = &H1
    Private Const KEY_SET_VALUE = &H2
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_NOTIFY = &H10
    Private Const KEY_CREATE_LINK = &H20
    Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))


Private Function GetString(hkey As Long, strpath As String, strValue As String)
    Dim keyhand&
    Dim DataType&
    r = RegOpenKey(hkey, strpath, keyhand&)
    GetString = RegQueryStringValue(keyhand&, strValue)
    r = RegCloseKey(keyhand&)
End Function


Function RegQueryStringValue(ByVal hkey As Long, ByVal strValueName As String)
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    On Error GoTo 0
    lResult = RegQueryValueEx(hkey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hkey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)


            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function


Private Sub SaveKey(hkey As Long, strpath As String)
    Dim keyhand&
    r = RegCreateKey(hkey, strpath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Private Function SaveString(hkey As Long, strpath As String, strValue As String, strdata As String)
    Dim keyhand&
    r = RegCreateKey(hkey, strpath, keyhand&)
    r = RegSetValueEx(keyhand&, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand&)
End Function


Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))


    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function



Function CheckSpecialFolderCount() As String
    CheckSpecialFolderCount = GetString(HKEY_CURRENT_USER, "Software\SpecialFolders", "AlertForUpdate")
End Function
