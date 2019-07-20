VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IdeaMIS"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8625
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtDetails 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmMain.frx":16692
      Top             =   2400
      Width           =   7455
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "配置路径"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FF0000&
      Picture         =   "frmMain.frx":16696
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   5880
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其它"
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "品味"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   4770
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "关闭(&x)"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "保存"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Tag             =   " "
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "映射"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   4140
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "奇物类"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3510
      Width           =   850
   End
   Begin VB.OptionButton Option1 
      Caption         =   "观察"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "建模"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2250
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "词句"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "编程"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   990
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "商机"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtContent 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":26EE0
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "补充理念文本信息:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lable_time 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3840
      TabIndex        =   12
      Top             =   80
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "当前时间"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim choose_caption(0 To 8) As String
 Dim s As String
 Dim M_TOSAVEPARENT_DIR As String
 'In general section
 
Private Declare Function CreateDirectory Lib "Kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
    Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Const CREATE_NEW = 1
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "Kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long


Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
   ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const BIF_NEWDIALOGSTYLE = &H40
Const BIF_EDITBOX = &H10
Const BIF_USENEWUI = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String,lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As _
'String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'  Function SaveINI Lib "Kernel32" Alias _
'"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
'lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
'Function GetINI(AppName As String, KeyName As String, filename As String) As String
'Dim RetStr As String
'RetStr = String(10000, Chr(0))
'GetINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), filename))
'End Function

Public Function BrowseForFolder(Optional sTitle As String = "请选择文件夹") As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = 0 ' Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
       sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
       iNull = InStr(sPath, vbNullChar)
        If iNull Then
          sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseForFolder = sPath
End Function
 

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
    s = s + Format(Now, "YYYYmmddHHMM")  ' YYYYmmddHHMMSS
    Dim Security As SECURITY_ATTRIBUTES
    'Create a directory
    Dim ideaDir As String
    ideaDir = IIf(Right(M_TOSAVEPARENT_DIR, 1) = "\", M_TOSAVEPARENT_DIR, M_TOSAVEPARENT_DIR + "\") + s
    ret& = CreateDirectory(ideaDir, Security)
    'If CreateDirectory returns 0, the function has failed
    If ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
    ret& = 0
    ' save txtdetails text
    Dim str As String, detail_name As String
    str = Trim(txtDetails.Text)
    If str = "" Then Exit Sub
    detail_name = ideaDir + "\detail.txt"
    
 ' Method One
    Dim nHandle As Integer, fName As String
    fName = detail_name
    nHandle = FreeFile
    Open fName For Output As #nHandle
    Print #nHandle, Trim(str)
    Close nHandle
 MsgBox "保存成功", vbInformation + vbOKOnly
 
    
' Method Two
'    Dim lngHandle As Long
'    lngHandle = CreateFile(detail_name, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_NEW, 0, 0)
'    WriteFile hNewFile, str, Len(str), Ret, ByVal 0&
'    CloseHandle lngHandle
    
' Method Three
  '  Shell "cmd ", vbNormalNoFocus
  '  SendKeys " echo " + s + " > " + detail_name
 '   SendKeys vbEnter
 '   MsgBox str
End Sub

Private Sub cmdSet_Click()
  Dim sPath_ini As String
  sPath_ini = BrowseForFolder()
  SaveSetting App.EXEName, "Idea_Save_RootDir", sPath_ini, ""
  
End Sub
 
Private Sub Form_Load()
 Dim i As Integer
 i = 0
 choose_caption(0) = "商机"
  choose_caption(1) = "编程"
 choose_caption(2) = "词句"
 choose_caption(3) = "建模"
 choose_caption(4) = "观察"
 choose_caption(5) = "奇物类"
 choose_caption(6) = "映射"
 choose_caption(7) = "品味"
 choose_caption(8) = "其它"
 
 M_TOSAVEPARENT_DIR = GetSetting(App.EXEName, "Idea_Save_RootDir", "default_dir", CurDir)
 s = ""
  For i = LBound(choose_caption) To UBound(choose_caption)
     s = s + " " + choose_caption(i)
  Next
  Call Timer1_Timer
  
  Option1.Item(0).Value = True
  txtContent.MaxLength = 100
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim cap As String
 cap = "Idea灵感" + choose_caption(Index)
  Me.txtContent = cap
  Me.txtContent.SelStart = Len(cap)
  Me.txtContent.SelText = "" 'Right(cap, 1)

End Sub

Private Sub Timer1_Timer()
  lable_time.Caption = Format(Now, "YYYYmmdd HH:MM:SS")
End Sub

Private Sub txtContent_LostFocus()
   s = Replace(Trim(txtContent.Text), vbCrLf, "")
End Sub
