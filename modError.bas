Attribute VB_Name = "modError"
'===============================================================================
' modError - Central error handling support module
' Provides centralized error handling and support for logging errors to the
' event log.
'
' Version   Date        User            Notes
'   1.0     11/16/00    Mel Grubb II    Initial version
'   1.1     11/29/00    Mel Grubb II    Added error handlers
'                                       Applied new coding standards
'===============================================================================
Option Explicit

'===============================================================================
' Constants
'===============================================================================
Private Const mc_strModuleID As String = "modError."    ' Used to identify the location of errors

'===============================================================================
' Enumerations
'===============================================================================
Public Enum ErrorBase                       ' Assigns blocks of error numbers to different classes
    Common = vbObjectError + 512            ' Errors not specific to one class
End Enum

Public Enum eCommonErrors
    [_Base] = ErrorBase.Common
End Enum


'===============================================================================
' AppVersion - Standardize the formatting of the application version number
'
' Arguments: None
'
' Notes:
'===============================================================================
Public Function AppVersion() As String
    On Error GoTo ErrorHandler
    AppVersion = App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
    Exit Function

ErrorHandler:
    AppVersion = "<Error>"

End Function


'==============================================================================
' ProcessError - Logs the specified error to the NT error log.
'
' Parameters:
'   objErr (IN) - the error to be logged
'   strModuleID (IN) - the module or method name where the error occurred.
'   bReraiseError (IN) - True if the error should be reraised; False otherwise.
'
' Notes:
'==============================================================================
Public Sub ProcessError(ByRef objErr As ErrObject, Optional ByVal ModuleID As String, Optional ByVal blnReraiseError As Boolean = False)
    Dim strMessage As String
    Dim strTitle As String
    
    ' Build the simple error string for the dialog
    strMessage = "Error Number = " & Err.Number & " (0x" & Hex$(Err.Number) & ")" & vbCrLf _
        & "Description = " & Err.Description & vbCrLf _
        & "Source = " & objErr.Source
    If (Len(ModuleID) > 0) Then strMessage = strMessage & vbCrLf & "Module = " & ModuleID

    ' Show the error dialog
    strTitle = App.Title & " [" & AppVersion() & "]"
    MsgBox strMessage, vbOKOnly, strTitle
    
    ' Expand the error before logging
    strMessage = strTitle & vbCrLf & strMessage
    
    ' Log the error to the event log or log file, and the debug window
    App.LogEvent strMessage, vbLogEventTypeError
    Debug.Print vbCrLf & strMessage

    ' Reraise the error if necessary
    If (blnReraiseError) Then
        ReraiseError objErr, ModuleID
    End If

    ' The next line will only be executed in Debug mode while in the IDE.
    ' It causes the application to stop so that the programmer can debug.
    Debug.Assert StopInIDE() = True

ExitHandler:
    ' Release any screen locks
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub


'===========================================================================
' StopInIDE - Causes a stop, but only in development mode
'
' Arguments: None
'
' Notes:
'===========================================================================
Private Function StopInIDE() As Boolean
    On Error GoTo ExitHandler
    
    Stop
    StopInIDE = True
    
ExitHandler:
    Exit Function

End Function


'==============================================================================
' ReraiseError - reraises the specified error.
'
' Parameters:
'   objErr (IN) - the error to be reraised
'   strModuleID (IN) - the module or method name where the error occurred.
'
' Notes:
'==============================================================================
Private Sub ReraiseError(objErr As ErrObject, Optional ByVal strModuleID As String = "")
    On Error Resume Next
    If (Len(strModuleID) > 0) Then
        Err.Raise objErr.Number, strModuleID & vbCrLf & objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext
    Else
        Err.Raise objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext
    End If
End Sub
