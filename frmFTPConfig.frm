VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFTPConfig 
   Caption         =   "FTP Transfer Tool"
   ClientHeight    =   7755
   ClientLeft      =   3600
   ClientTop       =   1605
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   8475
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   7500
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAddSchedule 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      Picture         =   "frmFTPConfig.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Add to Schedule"
      Top             =   4680
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDestOption 
      Caption         =   "Destination"
      Height          =   975
      Left            =   6120
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
      Begin VB.OptionButton optFile 
         Caption         =   "File"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDatabase 
         Caption         =   "Database"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   ">>>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      ToolTipText     =   "Transfer"
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame fraRaw 
      Caption         =   "Local File Destination"
      Height          =   2895
      Left            =   6120
      TabIndex        =   13
      Top             =   2280
      Width           =   2175
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Remote File Location"
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4815
      Begin VB.ListBox lstItems 
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdUp 
         Height          =   375
         Left            =   4200
         Picture         =   "frmFTPConfig.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Up One Level"
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraConnect 
      Caption         =   "Connection Settings"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblUser 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblAddress 
         Caption         =   "FTP://"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraDB 
      Caption         =   "Database Destination"
      Height          =   2895
      Left            =   6120
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSett 
      Caption         =   "Settings"
      Begin VB.Menu mnuSettSched 
         Caption         =   "Schedule"
      End
   End
End
Attribute VB_Name = "frmFTPConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Module Variables
Private mFTP As FTP_JPC.CFTP_JPC    'DLL created to perform FTP.
Private mrecSchedule As ADODB.Recordset
Private mstrFile As String

Private Sub cmdAddSchedule_Click()
    'Add the data entry to the FTP List.
    AddSchedule
End Sub

Private Sub cmdConnect_Click()
'Local Variables
Dim bRet As Boolean

    MousePointer = vbHourglass
    
    'Connect to the specified server.
    bRet = mFTP.FTPConnect(CStr(txtAddress.Text), _
            CStr(txtUser.Text), CStr(txtPassword.Text))
    
    'If the cnnection is successful, navigate and acquire the
    'remote directory list.
    If bRet = True Then
        StatusChange ("Connection was successful.")
        NavigateRemote
    Else
        StatusChange ("Connection was not successful.")
    End If
    
    MousePointer = vbDefault
    EnableUI (bRet)
    
End Sub

Private Sub cmdDisconnect_Click()
    mFTP.FTPDisconnect
    ClearItems
    EnableUI (False)
End Sub



Private Sub cmdTransfer_Click()
'Local Variables.
Dim strRemoteFile As String
Dim strLocalFile As String

    'Uses the relative address that has already
    'Been established via the NavigateRemote procedure
    'to transfer the file
    strRemoteFile = CStr(lstItems.Text)
    'The local file is a contatenation of the Dir1.Path
    'and the slected item.
    strLocalFile = CStr(Dir1.Path & "\" & lstItems.Text)
    
    'Attempt to transfer the file and update the Status
    'bar accordingly.
    If mFTP.FTPTransfer(strRemoteFile, strLocalFile) Then
        StatusChange ("Transfer was successful.")
    Else
        StatusChange ("Transfer was not successful.")
    End If

End Sub

Private Sub cmdUp_Click()
    'Set the text of txtDirectory to the parent directory and
    'display the contents.
    txtDirectory.Text = Left(txtDirectory.Text, _
            InStrRev(txtDirectory.Text, "/", -1, vbBinaryCompare) - 1)
    NavigateRemote
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
    'Load the selcted record into ther Data Entry section.
    EstablishBindings
End Sub

Private Sub Drive1_Change()
    'If Drive1 changes reassociate Dir1 to the new drive.
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Set mFTP = New FTP_JPC.CFTP_JPC
    staStatus.Panels(1).Width = 3000
    mstrFile = App.Path & "\" & "Schedule.xml"
    LoadSchedule
End Sub

Sub EnableUI(bEnabled As Boolean)
    cmdConnect.Enabled = Not bEnabled
    cmdDisconnect.Enabled = bEnabled
    cmdAddSchedule.Enabled = bEnabled
    cmdTransfer.Enabled = bEnabled
    DataGrid1.Enabled = Not bEnabled
End Sub

Private Sub lstItems_DblClick()
    'This fails to work correctly if a file is selected.
    'The file name is appended to the txtDirectory, when this should
    'not be allowed to occur.
    txtDirectory.Text = txtDirectory.Text & "/" & lstItems.Text
    NavigateRemote
End Sub

Sub ListItems()
'Local Variables.
Dim dirlist As New Collection
Dim i As Integer

    'Acquire the current directory list.
    Set dirlist = mFTP.DirectoryList

    'Clear existing item list and repopulate with new
    'list informaiton.
    lstItems.Clear
    
    For i = 1 To dirlist.Count
        lstItems.AddItem dirlist.Item(i)
    Next
    
End Sub

Sub ClearItems()
'Local Variables.
Dim ctl As Control

    'Clear the text of qualified controls.
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then ctl.Text = ""
        If TypeOf ctl Is ListBox Then ctl.Clear
        If TypeOf ctl Is StatusBar Then ctl.Panels(1).Text = ""
    Next

End Sub
Sub SelectDestination()
    'Determine to show File or Database transfer.
    fraRaw.Visible = optFile.Value
    fraDB.Visible = optDatabase.Value
    '*** Note DB transfer is still in progress at this time.  I am
    'looking for a way to use APIs or a Control to leverage
    'something similar to the VB Data Link Libraries Dialog to capture
    'the desired destination.  Stay tuned.
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    'Save the File Transfer List to the .xml file.
    mrecSchedule.Save mstrFile, adPersistXML
End Sub

Private Sub mnuSettSched_Click()
    'Open the Scheduler.
    With frmScheduler
        .Caption = mnuSettSched.Caption
        .Height = 3075
        .Width = 3630
        .Show vbModal
    End With
End Sub

Private Sub optDatabase_Click()
    SelectDestination
End Sub

Private Sub optFile_Click()
    SelectDestination
End Sub

'*********************************************************************
'Purpose:   Load the file from the Saved .xml file or create a new
'           one.
'Scope:     Private
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Sub LoadSchedule()
'Local Variables.
Dim fso As Object

    'Instantiate Objects.
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mrecSchedule = New ADODB.Recordset
    
    'If the dat file already exists, use it.  If not create a new
    'one with the elements required to support the program.
    If fso.FileExists(mstrFile) Then
        mrecSchedule.Open mstrFile
    Else
        mrecSchedule.Fields.Append "Remote Server", adVarChar, 128
        mrecSchedule.Fields.Append "Remote UID", adVarChar, 128
        mrecSchedule.Fields.Append "Remote PWD", adVarChar, 128
        mrecSchedule.Fields.Append "Remote Path", adVarChar, 128
        mrecSchedule.Fields.Append "Remote File", adVarChar, 128
        mrecSchedule.Fields.Append "Local Path", adVarChar, 128
        mrecSchedule.Fields.Append "Local File", adVarChar, 128
        mrecSchedule.Open
    End If
    
    'Establish databinding between the recordset and the Data Grid.
    'This enables changes to the Data Grid to cascade directly to
    'the recordset with little effort.
    Set DataGrid1.DataSource = mrecSchedule

End Sub

'*********************************************************************
'Purpose:   Add the data entry to the File Transfer List.
'Scope:     Private
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Private Sub AddSchedule()

    With mrecSchedule
        .AddNew
        .Fields("Remote Server") = CStr(txtAddress.Text)
        .Fields("Remote UID") = CStr(txtUser.Text)
        .Fields("Remote PWD") = CStr(txtPassword.Text)
        .Fields("Remote Path") = CStr(txtDirectory.Text)
        .Fields("Remote File") = CStr(lstItems.Text)
        .Fields("Local Path") = CStr(Dir1.Path)
        .Fields("Local File") = CStr(lstItems.Text)
        .Update
    End With
    
End Sub

'*********************************************************************
'Purpose:   Places the information selcted from the File Transfer List
'           on the Data entry portion of the form.
'Scope:     Private
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Private Sub EstablishBindings()

    txtAddress.Text = mrecSchedule.Fields("Remote Server")
    txtUser.Text = mrecSchedule.Fields("Remote UID")
    txtPassword.Text = mrecSchedule.Fields("Remote PWD")
    txtDirectory.Text = mrecSchedule.Fields("Remote Path")
    'This portion will be required after connectioon is established.
    'lstItems.Text = Test
    Dir1.Path = mrecSchedule.Fields("Local Path")
    
End Sub

'*********************************************************************
'Purpose:   Navigate the remote file structure.
'Scope:     Private
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Private Sub NavigateRemote()

'Local Variables.
Dim strDir As String

    strDir = txtDirectory.Text
    
    'If strDirectory is empty, we should be looking at the root.
    If strDir = "" Then strDir = "/"

    If mFTP.ChangeDirectory(strDir) = True Then
        ListItems
    End If

End Sub

'*********************************************************************
'Purpose:   Handle status changes.
'Scope:     Private
'Inputs:    strMessage(Required)
'Returns:   Nothing
'*********************************************************************
Private Sub StatusChange(strMessage As String)
    
    With staStatus.Panels(1)
        .Text = strMessage
    End With
    
End Sub
