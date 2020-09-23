VERSION 5.00
Begin VB.Form frmMount 
   Caption         =   "Drive Mounter Professional"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Select Action"
      Height          =   665
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5175
      Begin VB.CommandButton Command5 
         Caption         =   "Quit"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Unmount"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mount"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Physical Path"
      Height          =   665
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   275
         Left            =   4560
         TabIndex        =   4
         Top             =   250
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   250
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Virtual Drive Letter"
      Height          =   665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuProgHelp 
         Caption         =   "&Program Help"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Sub MountVirtualDrive(strVirtualDrive, strPhysicPath)
    Shell "subst.exe " & strVirtualDrive & Chr(32) & strPhysicPath, vbHide
End Sub

Sub UnMountVirtualDrive(strVirtualDrive)
    Shell "subst.exe " & strVirtualDrive & " /d", vbHide
End Sub

Private Sub Command2_Click()
    Dim tBuf As String * 100
    'The subst.exe program only supports short dos-pathnames
    a = GetShortPathName(BrowseForFolder(Me.hWnd, "Select Folder"), tBuf, 100)
    Text1.Text = Left(tBuf, a)
End Sub

Private Sub Command3_Click()
    If Text1.Text = "" Then
        MsgBox "First Select Physical path", vbCritical
        Exit Sub
    End If
    
    MountVirtualDrive Combo1.Text, Text1.Text
    MsgBox "Virtual drive " & Combo1.Text & " mounted", vbInformation
End Sub

Private Sub Command4_Click()
    UnMountVirtualDrive Combo2.Text
    MsgBox "Virtual drive " & Combo2.Text & " unmounted", vbInformation
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetDrives
    FillList
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Sub FillList()
    For i% = 1 To 26
        If GetDriveType(Drives(i%)) = 1 Then
            Combo1.AddItem UCase(Drives(i%))
        End If
        Combo2.AddItem UCase(Drives(i%))
    Next i%
End Sub

Private Sub MnuAbout_Click()
    MsgBox "Drive Mounter Professional" & vbCrLf & "Created by SonyMS" & vbCrLf & "Mail to: sonysmsone@softhome.net", vbInformation
End Sub

Private Sub MnuProgHelp_Click()
    frmHelp.Show
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub
