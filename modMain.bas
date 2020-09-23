Attribute VB_Name = "modMain"
'==============================================================================
' modMain - Child Class Builder Utility Main Module
'
' Version   Date        User            Notes
'   1.0     04/20/01    Mel Grubb II    Initial version
'
' Notes: Some of these routines are gross hacks intended to make up for some of
' shortcomings of the Extensibility model, such as the inability to easily
' determine the Data Type of a Member, or the lack of an easily accessable
' Parameters collection.  This is by no means the most elegant code ever
' written, but sometimes brute force is simply the only way.
'==============================================================================
Option Explicit


'==============================================================================
' Global Constants
'==============================================================================
Public Const g_strAppTitle = "Child Class Builder Utility"


'==============================================================================
' Global Types
'==============================================================================
Public Type gudt_Member
    Name As String
    Type As ge_MemberTypes
    Declaration As String
    DataType As String
End Type


'==============================================================================
' Global Enumerations
'==============================================================================
Public Enum ge_MemberTypes
    emtFunction
    emtPropertyGet
    emtPropertyLet
    emtPropertySet
    emtSub
End Enum


'==============================================================================
' Global Variables
'==============================================================================
Public g_objVBInstance As VBIDE.VBE


'==============================================================================
' Private Member Constants
'==============================================================================
Private Const mc_strModuleID = "modMain."


'==============================================================================
' AppendLines - Append a string onto the end of a code module with support for
' C-style formatting command expansion.
'
' Arguments:
'   CodeModule - The VB Code module to append the line to
'   Line - The text to append
'
' Notes: The C-style formatting commands \n and \t are supported to make for
' somewhat smaller code in the calling functions.  This by no means makes the
' routine more efficient, but it does make the code easier to read and
' understand.
'==============================================================================
Public Sub AppendLines(Module As VBIDE.CodeModule, Line As String)
    On Error GoTo ErrorHandler
    
    With Module
        .InsertLines .CountOfLines + 1, Replace(Replace(Line, "\n", vbCrLf), "\t", vbTab)
    End With
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "AppendLines"
    
End Sub


'==============================================================================
' DataType - Returns the apparant data type of a Function or Property.
'
' Arguments:
'   Declaration - The declaration line of the Function or Property
'
' Notes:
'==============================================================================
Public Function DataType(ByRef Member As gudt_Member) As String
    On Error GoTo ErrorHandler
    Dim strDeclaration As String
    
    With Member
        Select Case .Type
            Case emtFunction, emtPropertyGet
                ' First, chop off everything before the final Parenthesis
                strDeclaration = Mid$(.Declaration, InStrRev(.Declaration, ")") + 1)
                
                ' Now grab the last word on the line
                DataType = Mid$(strDeclaration, InStrRev(strDeclaration, " ") + 1)
                
            Case emtPropertyLet, emtPropertySet
                DataType = Mid$(.Declaration, InStrRev(.Declaration, " ") + 1, Len(.Declaration) - InStrRev(.Declaration, " ") - 1)
        End Select
    End With
    Exit Function

ErrorHandler:
    ProcessError Err, mc_strModuleID & "DataType"
    
End Function


'==============================================================================
' Declaration - Returns the declaration of a Function, Sub, or Property.
'
' Arguments:
'   Module - A CodeModule object containing the Function, Sub, or Property.
'   MemberName - The name of the member to retrieve.
'   MemberType - The type of member to retrieve.
'
' Notes: If the declaration is split onto multiple lines, then a mutliple-line
' result will be returned.  If the member does not exist, an empty string will
' be returned.
'==============================================================================
Public Function Declaration(Module As VBIDE.CodeModule, MemberName As String, MemberType As vbext_ProcKind) As String
    On Error GoTo ErrorHandler
    Dim strDeclaration As String
    Dim intLine As Integer

    strDeclaration = Trim$(Module.Lines(Module.ProcBodyLine(MemberName, MemberType), 1))
    Do While Right$(strDeclaration, 1) = "_"
        intLine = intLine + 1
        strDeclaration = strDeclaration & vbCrLf & vbTab & Trim$(Module.Lines(Module.ProcBodyLine(MemberName, MemberType) + intLine, 1))
    Loop

    ' Return the declaration
    Declaration = strDeclaration
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case 35 ' Member not found
            Declaration = ""
            Err.Clear
        
        Case Else
            ProcessError Err, mc_strModuleID & "Declaration"
    End Select
End Function


'==============================================================================
' Parameters - Returns the parameter section of a declaration, with or without
' declaration keywords.
'
' Arguments:
'   Declaration - The declaration line for the Function, Sub, or Property
'
' Notes:
'==============================================================================
Public Function Parameters(ByVal Declaration As String, Optional StripKeywords As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim astrParameters() As String
    Dim intIndex As Integer
    Dim intPos As Integer
    Dim strTemp As String
    Dim strParameters As String

    ' First, get just the parentheses
    Declaration = Mid$(Declaration, InStr(Declaration, "(") + 1)
    Declaration = Left$(Declaration, InStrRev(Declaration, ")") - 1)
    
    If StripKeywords Then
        Debug.Print Declaration
        If Declaration = "" Then
            ' There was nothing between the parentheses
            Parameters = Declaration
        Else
            ' Kill off any extra keywords that don't belong in a call
            Declaration = Replace$(Declaration, "ByRef ", "")
            Declaration = Replace$(Declaration, "ByVal ", "")
            Declaration = Replace$(Declaration, "Optional ", "")
            Declaration = Replace$(Declaration, "ParamArray ", "")
    
            ' Split up the parameters into an array
            astrParameters = Split(Declaration, ", ")
                    
            For intIndex = 0 To UBound(astrParameters)
                ' Kill off any line continuations and/or tabs and then grab the first word
                strTemp = Trim$(Replace$(Replace$(astrParameters(intIndex), "_" & vbCrLf, ""), vbTab, ""))
                intPos = InStr(strTemp, " ")
                strParameters = strParameters & IIf(intPos > 0, Left$(strTemp, intPos - 1), strTemp) & ", "
            Next intIndex

            ' Strip off that last ", " and return the result
            Parameters = Left$(strParameters, Len(strParameters) - 2)
        End If
    Else
        Parameters = Declaration
    End If
    Exit Function

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Parameters"
    
End Function
