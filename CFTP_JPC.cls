VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFTP_JPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declared Constants
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_SYNCH = 0
Private Const INTERNET_FLAG_ASYNCH = 1
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const FTP_FLAG_PASSIVE = &H8000000
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3
Private Const strUserAgent As String = "CFTPHandler vb wininet"
Private Const ERROR_NO_MORE_FILES = 18
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003

'Types
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

'Class variables.
Private hOpen As Long
Private hConnection As Long

'Establishes an open internet connection.
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

'Establishes a connection to a specific address.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

'Closes Connections.
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hIniet As Long) As Integer

'Navigation Functions
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, _
ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
(ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Boolean

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
(ByVal hFtpSession As Long, ByVal lpszDirectory As String, lpdwCurrentDirectory As String)

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
(ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
(lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, _
ByVal fFailExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
ByVal dwContext As Long) As Boolean

'*********************************************************************
'Purpose:   Open an Internet Connection based on client application
'           request.
'Scope:     Public Method
'Inputs:    strServer(Required): FTP Server to connect to.
'           strUser(Required): User account.
'           strPassword(Required): User password.
'Returns:   FTPCOnnect as Boolean
'*********************************************************************
Public Function FTPConnect(ByVal strServer As String, ByVal strUser As String, ByVal strPassword As String) As Boolean
    'Open an internet session.
    hOpen = InternetOpen(strUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, INTERNET_FLAG_SYNCH)
    If hOpen = 0 Then
        FTPConnect = False
        Exit Function
    Else
        FTPConnect = True
    End If
    'If the Internet is accessed, open an FTP session.
    hConnection = InternetConnect(hOpen, strServer, 0, strUser, strPassword, INTERNET_SERVICE_FTP, FTP_FLAG_PASSIVE, 0)
    If hConnection = 0 Then
        FTPConnect = False
        Exit Function
    Else
        FTPConnect = True
    End If
    'Return true if totally successful, false if not.
End Function

'*********************************************************************
'Purpose:   Disconnect active sessions.
'Scope:     Public Method
'Inputs:    Nothing
'Returns:   Nothing
'*********************************************************************
Public Function FTPDisconnect()

    'Close handles if they are open.
    If hConnection <> 0 Then InternetCloseHandle (hConnection)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    
    'Set values = 0.
    hConnection = 0
    hOpen = 0
    
End Function

'*********************************************************************
'Purpose:   Acquire the directory List of the active directory.
'Scope:     Public Property Get
'Inputs:    Nothing
'Returns:   DirectoryList as Collection
'*********************************************************************
Public Property Get DirectoryList() As Collection

'Local Variables
Dim colItems As New Collection
Dim hFind As Long
Dim pdata As WIN32_FIND_DATA
Dim strItemData As String
Dim bRet As Boolean
    
    pdata.cFileName = String(260, 0)
    hFind = FtpFindFirstFile(hConnection, "*.*", pdata, 0, 0)
    
    'Add items to the collection until, the API returns the
    'ERROR_NO_MORE_FILES error.
    If hFind <> False Then
        Do While Not Err.LastDllError = ERROR_NO_MORE_FILES
            InternetFindNextFile hFind, pdata
            If Err.LastDllError = ERROR_NO_MORE_FILES Then Exit Do
            strItemData = Left(pdata.cFileName, InStr(1, pdata.cFileName, String(1, 0), vbBinaryCompare) - 1)
            colItems.Add strItemData
        Loop
    End If
    
    'Return the directory list.
    Set DirectoryList = colItems
    'Close the open handle.
    InternetCloseHandle (hFind)
    
End Property

'*********************************************************************
'Purpose:   Change the directory on the Remote Server.
'Scope:     Public Method
'Inputs:    strDirectory(Required): Directory to navigate to.
'Returns:   ChangeDirectory as Boolean
'*********************************************************************
Public Function ChangeDirectory(ByVal strDirectory As String) As Boolean
    'Empty string means that the directory
    'does not need to be changed.
    If strDirectory = "" Then
        ChangeDirectory = False
        Exit Function
    Else
        'Attempt to change the directory on the Remote Server.
        If FtpSetCurrentDirectory(hConnection, strDirectory) = False Then
            ChangeDirectory = False
            Exit Function
        End If
    End If
    
    ChangeDirectory = True
    'Return true if totally successful, false if not.
End Function

'*********************************************************************
'Purpose:   Transfer File to the desired destination.
'Scope:     Public Method
'Inputs:    strRemoteFile(Required): Name of the File in the active
'                                       directory.
'           strLocalFile(Required): Fully qualified path and filename
'                                   of the desintation of the transfer.
'Returns:   FTPTransfer as Boolean
'*********************************************************************
Public Function FTPTransfer(strRemoteFile As String, _
                            strLocalFile As String) As Boolean
                            
    'Attempt to transfer the file.
    FTPTransfer = FtpGetFile(hConnection, strRemoteFile, _
                strLocalFile, False, INTERNET_FLAG_RELOAD, _
                FTP_TRANSFER_TYPE_ASCII, 0)

End Function

'***********************************************************
'*** Removed due to Err return isssues through functions.***
'***********************************************************
'Private Function ErrorOut(Optional lngError As Long, _
'                Optional lngDLLError As Long, _
'                Optional strCallFunc As String)
    'Variables
'    Dim lngInternetError As Long
'    Dim lngLength As Long
'    Dim strBuffer As String
'    Dim strErr As String
'    Dim lngErr As Long
    
    'If the error is an internet error collect the internet
    'error, else collect the err information.
'    If lngDLLError = ERROR_INTERNET_EXTENDED_ERROR Then
        'Determine the length of the response.
'        InternetGetLastResponseInfo lngInternetError, _
                                    vbNullString, _
                                    lngLength
'        strBuffer = String(lngLength + 1, 0)
        'Acquire the response.
'        InternetGetLastResponseInfo lngInternetError, _
                                    strBuffer, _
                                    lngLength
        'Establish return message and number.
'        strErr = "FTP Extended Error.  Details: " & lngInternetError & " " & strBuffer
'        lngErr = vbObjectError + lngInternetError
'    Else
        'Establish return message and number.
'        strErr = "Unspecified Error.  Details:" & Err.Number & " " & Err.Description
'        lngErr = vbObjectError + Err.Number
'    End If
    
    'Raise error.
'    Err.Raise lngErr, strCallFunc, strErr
'End Function
