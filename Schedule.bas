Attribute VB_Name = "Schedule"
Option Explicit

'Add a new Job.
Declare Function NetScheduleJobAdd Lib "netapi32.dll" _
(ByVal Servername As String, Buffer As Any, Jobid As Long) As Long
'Delete a Job.
Declare Function NetScheduleJobDel Lib "netapi32.dll" _
(ByVal Servername As String, ByVal MinJobid As Long, ByVal MaxJobid As Long) As Long

'Types.
'AT_INFO is used to set the particulars about the job.
Type AT_INFO
    JobTime     As Long
    DaysOfMonth As Long
    DaysOfWeek  As Byte
    Flags       As Byte
    Command     As String
End Type

'AT_ENUM is used to gather information about jobs that
'already exist in the scheduler.
Type AT_ENUM
    Jobid       As Long
    JobTime     As Long
    DaysOfMonth As Long
    DaysOfWeek  As Byte
    Flags       As Byte
    Command     As String
End Type

'Constants
Public Const JOB_RUN_PERIODICALLY = &H1
Public Const JOB_NONINTERACTIVE = &H10
Public Const NERR_Success = 0

'*********************************************************************
'Purpose:   Adds the requested job to the scheduler.
'Scope:     Public
'Inputs:    strApplication(Required): App to execute.
'           bytDay(Optional): Day to execute.
'           strTime(Optional): Time to execute.
'Returns:   Nothing
'Notes:     In the Win2K environment you are required to
'           submit the user name and password to operate the scheduled
'           task as.  I have yet to resolve this issue.
'*********************************************************************
Public Sub SchedApp(ByVal strApplication As String, Optional ByVal bytDay As Byte, _
                Optional ByVal strTime As String)
                
'Local Variables.
Dim lngJobID As Long
Dim udtAtInfo As AT_INFO
    
    'Establish AT_INFO settings that will be passed to the API.
    With udtAtInfo
        .JobTime = JT(strTime)
        .DaysOfWeek = DOW(bytDay)
        .Flags = JOB_NONINTERACTIVE + JOB_RUN_PERIODICALLY
        .Command = StrConv(strApplication, vbUnicode)
    End With
    
    'Determine if NetScheduleJobAdd was successful, MsgBox accordingly.
    If NetScheduleJobAdd(vbNullString, udtAtInfo, lngJobID) = NERR_Success Then
        MsgBox "Job ID: " & lngJobID & " was created successfully.", vbOKOnly
    Else
        MsgBox "The job was not successfully created.", vbCritical
    End If

End Sub

'*********************************************************************
'Purpose:   Removes the requested job from the scheduler.
'Scope:     Public
'Inputs:    lngJob(Required): Job ID to delete.
'Returns:   Nothing
'*********************************************************************
Sub DelSchedApp(ByVal lngJob As Long)
    
    'Determine if NetScheduleJobDel was successful, MsgBox accordingly.
    If NetScheduleJobDel(vbNullString, lngJob, lngJob) = NERR_Success Then
        MsgBox "Job ID: " & lngJob & " was deleted.", vbOKOnly
    Else
        MsgBox "The job was not deleted.", vbCritical
    End If
    
End Sub

'*********************************************************************
'Purpose:   Convert Day of Week to API recognized format.
'Scope:     Private
'Inputs:    bytDay(Optional): Day of Week (0-6).
'Returns:   DOW as Byte
'*********************************************************************
Private Function DOW(Optional bytDay As Byte = 0) As Byte

    'Convert the day of the week to the format recognized by the API.
    DOW = 2 ^ bytDay
    
End Function

'*********************************************************************
'Purpose:   Convert Time to API recognized format.
'Scope:     Private
'Inputs:    strTime(Optional): Time on 24H.
'Returns:   JT as Long
'*********************************************************************
Private Function JT(Optional strTime As String = "00:00") As Long

    'Need to handle Empty strings passed off by textboxes.
    If strTime = "" Then strTime = "00:00"
    'Reformat the Time, just in case...
    strTime = Format(strTime, "hh:mm")
    'Convert the time to the format recognized by the API.
    JT = (Hour(strTime) * 3600 + Minute(strTime) * 60) * 1000
    
End Function



