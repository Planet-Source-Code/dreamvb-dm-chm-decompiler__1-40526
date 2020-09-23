VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Install DM CHM Decompiler V1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcan 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   5
      Top             =   1605
      Width           =   1215
   End
   Begin VB.CommandButton cmdext 
      Caption         =   "&Install"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1605
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4545
      TabIndex        =   3
      Top             =   630
      Width           =   585
   End
   Begin VB.TextBox txtextpath 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   645
      Width           =   3210
   End
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   255
      TabIndex        =   6
      Top             =   1185
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Extract To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   705
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose were you like to install the program."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const m_Filename = "CHMcomp.exe" ' the name of the file we want to install


Function FixPath(lzPath As String) As String
    ' Adds a blackslash to a path if needed
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function GetWinDir() As String
Dim Iret As Long, StrBuff As String
    StrBuff = Space$(512) ' Create a buffer for our string
    Iret = GetWindowsDirectory(StrBuff, 512) ' Get the windows default path
    GetWinDir = Left(StrBuff, Iret) ' Extract the part that we need
End Function

Function GetFolder(ByVal hWndOwner As Long, ByVal TitleText As String) As String
Dim BrowseF As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim lzPath As String
    Const MAX_PATH = 260
    BrowseF.hOwner = hWndOwner
    BrowseF.lpszTitle = TitleText
    BrowseF.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(BrowseF)
    lzPath = String(MAX_PATH, vbNullChar)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal lzPath)
    If RetVal Then
        GetFolder = Left(lzPath, InStr(lzPath, vbNullChar) - 1)
    End If
End Function

Private Sub cmdcan_Click()
    Unload frmmain  ' Unload the program
End Sub

Private Sub cmdext_Click()
Dim sData As String, OutFile As String, tFile As Long
   ' First we will install the main program file
    OutFile = txtextpath.Text & m_Filename
    sData = StrConv(LoadResData(101, "CUSTOM"), vbUnicode) ' Get the file form the resource file
    tFile = FreeFile  ' Pointer to free file
    Open OutFile For Binary As #tFile ' Open the new file in binary mode
        Put #tFile, , sData ' Write the file data to the new file
    Close #tFile    ' Close the file
    sData = ""  ' Clean up
    
    ' Now we need to add our registery keys
    Sleep 100 * 3
    lblstat.Caption = "Installing " & m_Filename & " Please wait...." ' update the install status caption
    ' Next we will add the needed registery keys
    Sleep 100 * 3 ' This is just a little delay I added in let the user know the program is doing something
    lblstat.Caption = "Adding registery keys Please wait...." ' update the install status caption
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, "chm.file\shell\Decompile\command"
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, "chm.file\shell\Decompile\command", "", Chr$(34) & OutFile & Chr$(34) & " %1 -d"
    MsgBox "Successful completed the install operation.", vbInformation
    Unload frmmain
End Sub

Private Sub Command1_Click()
Dim Folname As String
    Folname = Trim$(GetFolder(frmmain.hWnd, "Please choose a folder to install to:"))
    If Len(Folname) = 0 Then
        ' The code above shows the browse for folder dialog and
        ' check that the user has selected a folder
        txtextpath.Text = FixPath(GetWinDir)
    Else
        txtextpath.Text = FixPath(Folname) ' Update the text box with new folder path
    End If
    
End Sub

Private Sub Form_Load()
    frmmain.Icon = Nothing ' Remove the forms icon
    txtextpath.Text = FixPath(GetWinDir) ' Update the textbox the default windows location
    
End Sub

