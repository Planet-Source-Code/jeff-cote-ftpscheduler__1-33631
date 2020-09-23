VERSION 5.00
Begin VB.Form frmScheduler 
   Caption         =   "Scheduler"
   ClientHeight    =   2670
   ClientLeft      =   6360
   ClientTop       =   5025
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3510
   Begin VB.OptionButton optDOW 
      Caption         =   "Sunday"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Saturday"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Friday"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Thursday"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Wednesday"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Tuesday"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optDOW 
      Caption         =   "Monday"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Remove an Existing Job."
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Add New"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Add a New Job."
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Caption         =   "Time:"
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRemove_Click()
'Local Variables
Dim strResponse As String

'Disallow non-numeric entries.  An emty string signifies
'cancel.
Do
    strResponse = InputBox("Enter the Job ID that you would like to delete.", _
                    "Delete Scheduled Job", "1")
Loop While Not IsNumeric(strResponse) And strResponse <> ""

'Attempt to delete the entry.
If strResponse <> "" Then
    DelSchedApp (CLng(strResponse))
End If

End Sub

Private Sub cmdUpdate_Click()

    'Add the user entry to the scheduler, after determining which
    'day it should be executed on.
    SchedApp CStr(App.Path & "\FTPAutoExec.exe"), DOWVal, CStr(txtTime.Text)

End Sub

'*********************************************************************
'Purpose:   Acquire the day of the week that has been selected.
'Scope:     Private
'Inputs:    Nothing
'Returns:   DOWVal as Byte
'*********************************************************************
Private Function DOWVal() As Byte

'Local Variables.
Dim ctl As Control

'Iterate through the forms controls to determine which
'OptionButton has been selected.
    For Each ctl In Me.Controls
        If TypeOf ctl Is OptionButton Then
            If ctl.Value = True Then
                DOWVal = CByte(ctl.Index)
                Exit Function
            End If
        End If
    Next

End Function
