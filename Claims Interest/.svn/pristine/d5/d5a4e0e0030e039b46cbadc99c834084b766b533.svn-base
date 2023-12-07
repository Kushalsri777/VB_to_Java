Attribute VB_Name = "modAppLog"
' Module     : modAppLog
' Description:
' Procedures : fnLogClose()
'              fnLogOpen()
'              fnLogPrune()
'              fnLogWrite(ByVal pstrLogEntry As String, ByVal pstrProcNm As String)
'
' Called by  : MDIForm_Unload() in frmMDIMain
'
' Modified   :
'   01/2002  BAW Copied from SPUDS/SCUDS. This was edited slightly to
'                save the logfile in the CSIDL_PERSONAL folder rather than
'                the CSIDL_LOCAL_APPDATA folder.
' --------------------------------------------------
Option Explicit
Option Compare Binary

Private Const mcstrName             As String = "modAppLog."


' The following determines how wide each entry in the log file should be
Private mlngLogMaxLineSize As Long

Private Const mcstrLogFileName As String = "ClaimsLog.Log"

Private mTs As TextStream


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Function fnGetAppLogFileFQ() As String
    ' Comments  : Returns the fully qualified filename of the application log file
    ' Parameters: None
    '
    ' Called By : mnuHelpViewApplicationLogFile_Click() in frmMDIMain
    '
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnGetAppLogFileFQ"
    Dim strPath                 As String

    ' Get the path to where Per User non-roaming data is stored. This path
    ' will be created if it doesn't already exist.
    strPath = fnGetSpecialFolder(0, CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
    fnGetAppLogFileFQ = fnBuildQualifiedFileName(strPath, mcstrLogFileName)
    
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Function
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnLogClose()
    ' Comments  : Closes the application Log File
    ' Parameters: None
    ' Called By : MDIForm_Unload() in frmMDIMain
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLogClose"

    fnLogWrite "***End***", cstrCurrentProc

    If Not (mTs Is Nothing) Then
        mTs.Close
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject mTs
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnLogOpen()
    ' Comments  : Creates or opens a text file called ClaimsLog.Log to keep track
    '             processing throughout each session. This log
    '             file will be truncated when it exceeds a certain size, to ensure
    '             it never consumes too much space on the user's hard drive.
    ' Parameters: None
    '
    ' Called By : MDIForm_Load() in frmMDIMain
    '
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc       As String = "fnLogOpen"
    Const cstrEqualSign         As String = "="
    Const cintForAppending      As Integer = 8
    Const cintTristateFalse     As Integer = 0
    Dim strLogFile              As String
    Dim fso                     As Scripting.FileSystemObject
    Dim strPath                 As String

    ' Prune the log file so it doesn't consume the user's hard drive
    fnLogPrune

    Set fso = New Scripting.FileSystemObject
    
    ' Get the path to where Per User non-roaming data is stored. This path
    ' will be created if it doesn't already exist.
    strPath = fnGetSpecialFolder(0, CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
    strLogFile = fnBuildQualifiedFileName(strPath, mcstrLogFileName)
    If (fso.FileExists(strLogFile)) Then
        ' Open the existing file
        Set mTs = fso.OpenTextFile(strLogFile, cintForAppending, True)
    Else
        ' Create the file
        Set mTs = fso.CreateTextFile(strLogFile, cintTristateFalse)
    End If
    
    ' Set how wide each log entry should be, based on whether a verbose log
    ' was requested on the command line
    ' Note: This assumes that the gbLogVerbose boolean was set prior to
    '       calling *this* function (i.e. in Sub Main)
    If Not gbLogVerbose Then
        mlngLogMaxLineSize = 85
    Else
        mlngLogMaxLineSize = 200
    End If


    fnLogWrite gcstrBlankEntry, cstrCurrentProc
    fnLogWrite gcstrBlankEntry, cstrCurrentProc
    fnLogWrite String$(30, cstrEqualSign) & " NEW SESSION " & String$(30, cstrEqualSign), cstrCurrentProc
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject fso
    
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub



'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnLogPrune()
    ' Comments  : This procedure is called each time the application starts. If it
    '             detects that the log file has exceeded a certain size, it
    '             prunes it to a specified smaller size. This ensures the log file
    '             will never consume too much space on the user's hard drive.
    ' Parameters: None
    '
    ' Called By : Form_Load() in frmMain
    '
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc   As String = "fnLogPrune"
    Const clngMaxFileLength As Long = 128000
    Const cintLinesToKeep   As Integer = 300
    Const cintForReading    As Integer = 1
    
    Dim astrLines()         As String
    Dim fso                 As Scripting.FileSystemObject
    Dim lngIndex            As Long
    Dim strlines            As String
    Dim strLogFile          As String
    Dim strPath             As String
    Dim ts                  As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    
    strPath = fnGetSpecialFolder(0, CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
    strLogFile = fnBuildQualifiedFileName(strPath, mcstrLogFileName)
    
    ' Open the log file if it exists; otherwise create one.
    If (fso.FileExists(strLogFile)) Then
        Set ts = fso.OpenTextFile(strLogFile, cintForReading, True)
    Else
        Set ts = fso.CreateTextFile(strLogFile, True)
    End If
    
    ' --------------------------------------------------------------
    ' If the log file has exceeded 128k (clngMaxFileLength) in size,
    ' prune it to 300 (cintLinesToKeep) lines
    ' --------------------------------------------------------------
    If FileLen(strLogFile) > clngMaxFileLength Then
        ' Read the entire file, then split it into an array of lines
        strlines = ts.ReadAll
        astrLines = Split(strlines, vbCrLf)
        ts.Close

        ' With the file in memory, delete and recreate the file again, so
        ' its new contents will reflect only the post-pruning contents
        fso.DeleteFile strLogFile
        Set ts = fso.CreateTextFile(strLogFile, True)

        ' Write the last 300 (cintLinesToKeep) lines to the new log file
        For lngIndex = UBound(astrLines) - cintLinesToKeep To UBound(astrLines)
            ts.WriteLine astrLines(lngIndex)
        Next lngIndex
        
        ts.WriteLine ("On " & CStr(Date) & " at " & CStr(Time) & " the log file was pruned.")
        ts.Close
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler

    ' Clean-up statements go here
    fnFreeObject ts
    fnFreeObject fso

    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
    


'////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub fnLogWrite(ByVal strLogEntry As String, ByVal strProcName As String)
    ' Comments  : Write a Line to the application log file to show a running
    '             tally of application timing/processing
    ' Parameters: strLogEntry = What to show in the log file
    '             strProcName = The name of the procedure to show in log entry
    '
    ' Called By : Every procedure that needs to log something
    '
    ' Modified  :
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    Const cstrCurrentProc As String = "fnLogWrite"
    
    If Not (mTs Is Nothing) Then
        ' Add "..." if the string is long enough to get truncated in a moment
        ' "3" is the length of the truncation marker ("...")
        If Len(strLogEntry) > mlngLogMaxLineSize Then
            strLogEntry = Left$(strLogEntry, mlngLogMaxLineSize - 3) & "..."
        End If

        ' Pad the log entry string with trailing spaces, to make it easier to read
        strLogEntry = fnPadRightString(strLogEntry, mlngLogMaxLineSize, " ")

        mTs.WriteLine CStr(Date) + " " + CStr(Time) + " " + strLogEntry + _
                     "      [Proc=" + strProcName + "]"
    End If
PROC_EXIT:
    On Error GoTo 0     ' disable error handler
    ' Clean-up statements go here
    If gerhApp.ErrNum <> gclngERR_NUM_DEFAULT Then
        gerhApp.PropagateError mcstrName & cstrCurrentProc
    End If
    Exit Sub
PROC_ERR:
    Select Case Err.Number
        'Case statements for expected errors go here
        Case Else
            gerhApp.SaveErrObjectData mcstrName & cstrCurrentProc
    End Select
    Resume PROC_EXIT
End Sub
