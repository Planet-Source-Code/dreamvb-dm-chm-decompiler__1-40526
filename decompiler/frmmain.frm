VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose a Folder"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdabout 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3315
      TabIndex        =   3
      Top             =   3015
      Width           =   1155
   End
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
      Height          =   405
      Left            =   3315
      TabIndex        =   4
      Top             =   3585
      Width           =   1155
   End
   Begin VB.CommandButton cmdnewdir 
      Caption         =   "&New Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3315
      TabIndex        =   2
      Top             =   2445
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Decompile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3285
      TabIndex        =   1
      Top             =   1890
      Width           =   1155
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   165
      TabIndex        =   5
      Top             =   975
      Width           =   2790
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   510
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DM CHM Decompiler V1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   3300
      TabIndex        =   7
      Top             =   855
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3585
      Picture         =   "frmmain.frx":08CA
      Top             =   285
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose a location:"
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
      Left            =   165
      TabIndex        =   6
      Top             =   165
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   3135
      X2              =   3135
      Y1              =   570
      Y2              =   3990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   3120
      X2              =   3120
      Y1              =   570
      Y2              =   3990
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const STATUS_PENDING = &H103
Private Const STILL_ACTIVE = STATUS_PENDING

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Dim DirPath As String, sCommand As String, m_filename As String

Public Function SHWait(ByVal ProgID As Long) As Boolean
Dim mExitID As Long, hdlProg As Long
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgID)
    GetExitCodeProcess hdlProg, mExitID
    Do While mExitID = STILL_ACTIVE
        DoEvents
        GetExitCodeProcess hdlProg, mExitID
    Loop
    CloseHandle hdlProg
    SHWait = mExitID
End Function

Function FixPath(lzPath As String) As String
    ' Adds a blackslash to a path if needed
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Sub cmdabout_Click()
Dim msg As String
' this just displays an about with info about the program.
    msg = msg & "DM CHM Unpacker V1" & vbNewLine _
    & "A Freeware CHM Help Decompiler for Win98,Win2000 and WinXP" & vbNewLine & vbNewLine _
    & "Made and designed by Dreamvb"
    MsgBox msg, vbInformation, "About....."
    msg = ""
    
End Sub

Private Sub cmdcan_Click()
    DirName = ""    ' Clean out the dir path
    Unload frmmain  ' unload the program
End Sub

Private Sub cmdnewdir_Click()
Dim DirName As String
On Error GoTo MkDirErr

    DirName = Trim$(InputBox("Please enter a new folder name", "Create new folder"))
    ' the above line displays a inputbox asking the user for a new folder name
    If Len(DirName) <= 0 Then
        ' This check to see if they have entered in a folder name or not
        ' if not we show then a error message telling them they need to include a folder name
        MsgBox "You must enter a new folder name", vbInformation, "Create new folder"
        Exit Sub
    Else
        MkDir DirPath & DirName ' Create the new folder
        MsgBox "The new folder has now been created", vbInformation, "Create new folder"
        Dir1.Refresh ' Update the Dir box
        Exit Sub
MkDirErr:
        If Err Then
            DirName = "" ' Clean up
            MsgBox Err.Description, vbExclamation, "Error_" & Err.Number
            ' The above line means there was an error createing the new folder
        End If
    End If
End Sub

Private Sub cmdok_Click()
Dim RetVal As Long, ExecuteCmd As String, ans, mFile As String
On Error Resume Next

    ExecuteCmd = "hh.exe -decompile " & DirPath & " " & m_filename
    RetVal = Shell("command.com /c" & ExecuteCmd, vbHide)
    If Not SHWait(RetVal) Then
        ans = MsgBox("The CHM file you selected has now been unpacked" _
        & vbNewLine & "whould you like to view the files now", vbYesNo Or vbQuestion)
        If ans = vbNo Then
            Exit Sub
        Else
            Shell "explorer.exe " & DirPath, vbNormalFocus
        End If
    End If

    
End Sub

Private Sub Dir1_Change()
    DirPath = FixPath(Dir1.Path) ' Assign DirPath with the dir's path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveError
    Dir1.Path = Drive1.Drive ' Update the dir path
    Exit Sub
DriveError:
    If Err Then
        ' Display error message to use
        MsgBox Err.Description & " " & UCase$(Drive1.Drive) & vbNewLine & vbNewLine _
        & "Please try agian.", vbExclamation, "Error_" & Err.Number
    End If
End Sub

Private Sub Form_Load()
    sCommand = Trim$(Command$) ' This gets the file name to decompile
    frmmain.Icon = Nothing ' remove the forms icon
    DirPath = FixPath(Dir1.Path) ' update the Dirpath Varible
    If Not Right$(UCase$(sCommand), 2) = "-D" Then MsgBox "An invalid parameter was passed program will now end", vbCritical, "Error_48": End
    m_filename = Trim$(Left(sCommand, Len(sCommand) - 2)) ' EXtract the filename to decompile
    sCommand = "" ' clean up
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing ' free up memory and exit program
End Sub
