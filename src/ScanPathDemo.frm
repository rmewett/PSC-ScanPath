VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScanPath Demo"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMask 
      Height          =   315
      Left            =   540
      TabIndex        =   5
      Top             =   1140
      Width           =   5505
   End
   Begin VB.CheckBox chkSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Return Files in Sorted Sequence"
      Height          =   195
      Left            =   2100
      TabIndex        =   13
      Top             =   1980
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkSubFolders 
      Alignment       =   1  'Right Justify
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   2100
      TabIndex        =   12
      Top             =   1770
      Width           =   3015
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   1305
      Left            =   5250
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1500
      Width           =   1395
   End
   Begin VB.ListBox lstFiles 
      Height          =   3570
      Left            =   60
      TabIndex        =   15
      Top             =   2910
      Width           =   6555
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6060
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Width           =   555
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   2190
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1770
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   2610
      Width           =   1365
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Text            =   "c:\"
      Top             =   750
      Width           =   5475
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ScanPathDemo.frx":1194
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6585
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1170
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##############################################################################################
'Purpose:   This project is a Demo for my cScanpath Class

'           This Class scans a specified path and returns the files it finds.
'           It has fairly comprehensive Filters. You can select files by:
'           Attributes (Normal, Hidden, Read Only, System etc)
'           File Size (>, <, Range)
'           File Date (From, To, Range)
'           File Extensions (multiple allowed i.e. *.txt;*.dat;*.tmp)

'           You can optionally scan sub-folders

'           To keep the demo simple I have only used attributes & Extensions
'           for Filter. For full example of this Class see WipeIt3 submission
       
'Author:    Richard Mewett ©2005
'##############################################################################################

Private Const MAX_LISTBOXITEMS = 32767

'We must declare with WithEvents to process the files returned
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1

Private Sub cmdBrowse_Click()
    txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)
End Sub

Private Sub cmdScan_Click()
    If cmdScan.Caption = "Scan" Then
        'Create our Scan Object
        Set SP = New cScanPath
        
        cmdScan.Caption = "Stop"
        Screen.MousePointer = vbHourglass
        lstFiles.Clear
        
        'Set the Scan properties
        With SP
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScan txtPath, chkSubFolders.Value, chkSort.Value
        End With
        
        With lstFiles
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
        End With
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If
End Sub


Private Sub Form_Load()
    With cboMask
        .AddItem "*.*"
        .AddItem "*.dll;*.exe;*.ocx"
        .AddItem "*.doc;*.mdb;*.xls"
        .AddItem "*.bmp;*.gif;*.jpg;*.tif"
        .AddItem "*.bas;*.cls;*.ctl;*.frm;*.vbp"
        .ListIndex = 0
    End With
End Sub


Private Sub SP_DirMatch(Directory As String, Path As String)
    'This Event fires for each Folder found
End Sub

Private Sub SP_FileMatch(Filename As String, Path As String)
    'This Event fires for each File found
    
    With lstFiles
        'Make sure we do not exced the item limit of the ListBox control!
        If .ListCount < MAX_LISTBOXITEMS Then
            .AddItem Path & Filename
            
            If (.ListCount Mod 10) = 0 Then
                'Scroll the list every 10th file
                .ListIndex = .NewIndex
            End If
        Else
            'Stop the scan object
            SP.StopScan
            
            MsgBox "The scan had been stopped (the ListBox Control is full!)", vbInformation
        End If
    End With
End Sub


