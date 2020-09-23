VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shortcut without VB6STKIT plus Special Folders!"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3600
      TabIndex        =   33
      Top             =   7200
      Width           =   1455
   End
   Begin VB.ListBox GUIDLst 
      Height          =   1620
      Left            =   840
      TabIndex        =   32
      Top             =   5520
      Width           =   7095
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select if a folder"
      Height          =   195
      Left            =   6840
      TabIndex        =   31
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton GoSpecial 
      Caption         =   "Go"
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   25
      Text            =   "Open's Registry Editor"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desktop"
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   3720
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "My Computer"
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Control Panel"
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   19
      Text            =   "C:\WINDOWS\REGEDIT.EXE,0"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Text            =   "Special File Example"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Text            =   "C:\WINDOWS\REGEDIT.EXE"
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox FolderNameTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "My App"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "StartMenu/Programs"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "StartMenu"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desktop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox UninLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "C:\Program Files\MyApp\uninstall.exe"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox HelpLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "C:\Program Files\MyApp\MyApp.hlp"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox AppLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "C:\Program Files\MyApp\MyApp.exe"
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton GoShortCut 
      Caption         =   "Go"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label GUIDLbl 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6600
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "The reason I have this here is to make it a little easier to get rid of the test icons you make."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   5160
      Width           =   6735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created Special Folders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   8880
      Y1              =   4670
      Y2              =   4670
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Info Tip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      ToolTipText     =   "This is a sample Info Tip"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Special Folder in.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   24
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "File/Folder Icon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "File/Folder Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "File or Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Special File/Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Shortcuts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4335
      X2              =   4335
      Y1              =   0
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   4680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "StartMenu Folder Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Shortcuts on.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Help File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long

Public Function CreateGUID() As String
    Dim id(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    If CoCreateGuid(id(0)) = 0 Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    Else
        MsgBox "Error while creating GUID!"
    End If
End Function
Function xListRemoveSelected(listbox As listbox)
Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Function

Private Sub Form_Load()
    Dim Path As String, strSave As String
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\REGEDIT.EXE"
    Text1 = Path
    Text3 = Path & ",0"
    GUIDLbl = CreateGUID
    
    Dim i As Integer
 Set Reg = New RegistryRoutines
    
    Reg.hkey = HKEY_CURRENT_MACHINE
    Reg.KeyRoot = MainKeyRoot
    Reg.Subkey = MainSubKey
    If Not Reg.KeyExists Then Reg.CreateKey



    Dim KeyCollection As Collection
    Dim Object As Variant
        Set KeyCollection = Reg.EnumRegistryKeys(HKEY_CURRENT_USER, "Software\SpecialFolders")
      
        For Each Object In KeyCollection
            GUIDLst.AddItem Object
        Next
        Set KeyCollection = Nothing

CreateKey "HKEY_CURRENT_USER\Software\SpecialFolders"

End Sub


Private Sub GoShortCut_Click()
If Check1.Value = 1 Then
MakeDesktopShortcut AppLocTxt
MakeDesktopShortcut UninLocTxt
MakeDesktopShortcut HelpLocTxt
End If

If Check2.Value = 1 Then
MakeStartMenuShortcut AppLocTxt
MakeStartMenuShortcut UninLocTxt
MakeStartMenuShortcut HelpLocTxt
End If

If Check3.Value = 1 Then
MakeStartMenuFolderShortcut AppLocTxt, FolderNameTxt
MakeStartMenuFolderShortcut UninLocTxt, FolderNameTxt
MakeStartMenuFolderShortcut HelpLocTxt, FolderNameTxt
End If

End Sub

Private Sub GoSpecial_Click()
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\DefaultIcon"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\InProcServer32"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\Shell\Open\Command"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\ShellEx\PropertySheetHandlers\" & GUIDLbl
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\ShellFolder"
 CreateKey "HKEY_CURRENT_USER\Software\SpecialFolders\{" & GUIDLbl & "}"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}", "", "" + Text2.Text + ""
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}", "InfoTip", "" + Text4.Text + ""
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\DefaultIcon", "", "" + Text3.Text + ""
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\InProcServer32", "", "Shell32.dll"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\InProcServer32", "ThreadingModel", "Apartment"
 
 If Check7.Value = 0 Then
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\Shell\Open\Command", "", "" + Text1.Text + ""
 Else
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\Shell\Open\Command", "", "explorer , " + Text1.Text + ""
 End If
 
 SetStringValue "HKEY_CURRENT_USER\Software\SpecialFolders\{" & GUIDLbl & "}", "" + Text2.Text + "", ""
 SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{" & GUIDLbl & "}\ShellFolder", "Attributes", Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)

If Check6.Value = 1 Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{" & GUIDLbl & "}"
End If

If Check5.Value = 1 Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{" & GUIDLbl & "}"
End If

If Check4.Value = 1 Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{" & GUIDLbl & "}"
End If

GUIDLbl = CreateGUID
Call Form_Load
End Sub

Private Sub Remove_Click()
Dim PosDel As Variant
PosDel = MsgBox("Are you sure you want to delete the Special Folder?", vbYesNo + vbCritical, "Confirmation")
If PosDel = vbYes Then

DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" & GUIDLst.Text
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\" & GUIDLst.Text
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\" & GUIDLst.Text
DeleteKey "HKEY_CURRENT_USER\Software\SpecialFolders\" & GUIDLst.Text
Call xListRemoveSelected(GUIDLst)
End If
End Sub
