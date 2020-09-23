Attribute VB_Name = "ModAutoExec"
'*********************************************************************
'Purpose:   This exe is used to execute the automated transfer of files.
'Scope:     Self Contained exe.
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Sub Main()
'Local Variables
Dim ftp As FTP_JPC.CFTP_JPC 'DLL created to perform FTP.
Dim rec As ADODB.Recordset
Dim fso As Object
Dim strFile As String
Dim boolRet As Boolean      'Future use in logging.

'Instantiate variables.
Set ftp = New FTP_JPC.CFTP_JPC
Set rec = New ADODB.Recordset
Set fso = CreateObject("Scripting.FileSystemObject")
strFile = App.Path & "\Schedule.xml"

'If the file actually existis perform the execution.
If fso.FileExists(strFile) Then
    'Open the .xml File.
    rec.Open strFile
    
    'Step through each item in the list and transfer the requested files.
    Do While Not rec.EOF
        'Open the connection to the server.
        If ftp.FTPConnect(rec("Remote Server"), rec("Remote UID"), rec("Remote PWD")) = True Then
            'Change the directory.
            If ftp.ChangeDirectory(rec("Remote Path")) = True Then
                'Transfer the file.
                If ftp.FTPTransfer(rec("Remote File"), rec("Local Path") & "\" & rec("Local File")) Then
                    ftp.FTPDisconnect
                    boolRet = True
                Else
                    boolRet = False
                End If
            Else
                boolRet = False
            End If
        Else
            boolRet = False
        End If
        
    rec.MoveNext
    Loop
    
    rec.Close
    
End If

End Sub
