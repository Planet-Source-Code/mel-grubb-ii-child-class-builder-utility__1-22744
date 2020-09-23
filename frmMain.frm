VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Child Class Builder"
   ClientHeight    =   4800
   ClientLeft      =   2190
   ClientTop       =   2235
   ClientWidth     =   8610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8100
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   60
      ScaleHeight     =   4695
      ScaleWidth      =   2055
      TabIndex        =   3
      Top             =   60
      Width           =   2055
      Begin VB.TextBox txtChildClass 
         Height          =   285
         Left            =   855
         TabIndex        =   5
         Top             =   0
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstBaseClasses 
         Height          =   4335
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgMembers"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Base class"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLabel 
         Caption         =   "&New class:"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.PictureBox picSplitter 
      Height          =   4695
      Left            =   2160
      ScaleHeight     =   4635
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   60
      Width           =   75
   End
   Begin MSComctlLib.ImageList imgMembers 
      Left            =   8040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B24
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10BE
            Key             =   "Method"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMembers 
      Height          =   4275
      Left            =   2340
      TabIndex        =   1
      Top             =   420
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgMembers"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsMembers 
      Height          =   4695
      Left            =   2280
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Methods"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&All"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileUpdate 
         Caption         =   "&Update Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTab 
         Caption         =   "&Properties"
         Index           =   1
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "&Methods"
         Index           =   2
      End
      Begin VB.Menu mnuViewTab 
         Caption         =   "&All"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmMain - Child Class Builder Utility add-in form
'
' Version   Date        User            Notes
'   1.0     04/20/01    Mel Grubb II    Initial version
'==============================================================================
Option Explicit

Private Const mc_strModuleID = "frmMain."

'==============================================================================
' Private Member Enumerations
'==============================================================================
Private Enum me_Tabs
    etProperties = 1                    ' Set to 1 to match tab indexes
    etMethods
    etAll
End Enum


'==============================================================================
' Private Member Variables
'==============================================================================
Private ma_udtMembers() As gudt_Member
Private m_objSplitter As CSplitDDC
Private m_blnDirty As Boolean


'==============================================================================
' RefreshMemberArray - Loads the member array with the properties and methods
' of the currently selected base class
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub RefreshMemberArray()
    On Error GoTo ErrorHandler
    Dim objModule As VBIDE.CodeModule
    Dim objMember As VBIDE.Member
    Dim udtMember As gudt_Member
    Dim intUbound As Integer
    Dim strTemp As String
    
    ' Clear out any previous array contents
    Erase ma_udtMembers

    Set objModule = g_objVBInstance.ActiveVBProject.VBComponents(lstBaseClasses.SelectedItem.Text).CodeModule
    For Each objMember In objModule.Members
        If objMember.Scope = vbext_Public Then
            ' Public members of the base class are part of the interface
            If objMember.Type = vbext_mt_Method Then
                ' Add the Method to the array
                ReDim Preserve ma_udtMembers(intUbound)
                With ma_udtMembers(intUbound)
                    .Name = objMember.Name
                    .Declaration = Declaration(objModule, objMember.Name, vbext_pk_Proc)
                    If InStr(.Declaration, "Function ") Then
                        .Type = emtFunction
                        .DataType = DataType(ma_udtMembers(intUbound))
                    Else
                        .Type = emtSub
                    End If
                End With
                intUbound = intUbound + 1

            ElseIf objMember.Type = vbext_mt_Property Then
                ' Check for Property Get
                strTemp = Declaration(objModule, objMember.Name, vbext_pk_Get)
                If Not (strTemp = "") Then
                    ' Add the Property Get to the array
                    ReDim Preserve ma_udtMembers(intUbound)
                    With ma_udtMembers(intUbound)
                        .Name = objMember.Name
                        .Type = emtPropertyGet
                        .Declaration = strTemp
                        .DataType = DataType(ma_udtMembers(intUbound))
                    End With
                    intUbound = intUbound + 1
                End If

                ' Check for Property Let
                strTemp = Declaration(objModule, objMember.Name, vbext_pk_Let)
                If Not (strTemp = "") Then
                    ' Add the Property Let to the array
                    ReDim Preserve ma_udtMembers(intUbound)
                    With ma_udtMembers(intUbound)
                        .Name = objMember.Name
                        .Type = emtPropertyLet
                        .Declaration = strTemp
                        .DataType = DataType(ma_udtMembers(intUbound))
                    End With
                    intUbound = intUbound + 1
                End If

                ' Check for Property Set
                strTemp = Declaration(objModule, objMember.Name, vbext_pk_Set)
                If Not (strTemp = "") Then
                    ' Add the Property Set to the array
                    ReDim Preserve ma_udtMembers(intUbound)
                    With ma_udtMembers(intUbound)
                        .Name = objMember.Name
                        .Type = emtPropertySet
                        .Declaration = strTemp
                        .DataType = DataType(ma_udtMembers(intUbound))
                    End With
                    intUbound = intUbound + 1
                End If
            End If
        End If
    Next objMember
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "RefreshMemberArray"

End Sub


'==============================================================================
' RefreshMemberList - Refreshes the contents of the member list based on the
' current contents of the member array
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub RefreshMemberList()
    On Error GoTo ErrorHandler
    Dim objMember As gudt_Member
    Dim intIndex As Integer
    Dim intTabIndex As Integer
    
    ' Clear out the list before filling it
    lstMembers.ListItems.Clear
    
    ' Get which tab is selected
    intTabIndex = tsMembers.SelectedItem.Index
        
    ' Fill in the methods and properties
    For intIndex = 0 To UBound(ma_udtMembers)
        objMember = ma_udtMembers(intIndex)
        Select Case objMember.Type
            Case emtFunction
                If intTabIndex = etAll Or intTabIndex = etMethods Then
                    ' Add the method
                    With lstMembers.ListItems.Add
                        .Text = objMember.Name
                        .SmallIcon = "Method"
                        .SubItems(1) = ma_udtMembers(intIndex).DataType
                        .SubItems(2) = "Function"
                    End With
                End If
            
            Case emtSub
                If intTabIndex = etAll Or intTabIndex = etMethods Then
                    ' Add the method
                    With lstMembers.ListItems.Add
                        .Text = objMember.Name
                        .SmallIcon = "Method"
                        .SubItems(1) = ma_udtMembers(intIndex).DataType
                        .SubItems(2) = "Sub"
                    End With
                End If

            Case emtPropertyGet
                If intTabIndex = etAll Or intTabIndex = etProperties Then
                    ' Add the property
                    With lstMembers.ListItems.Add
                        .Text = objMember.Name & "_Get"
                        .SmallIcon = "Property"
                        .SubItems(1) = ma_udtMembers(intIndex).DataType
                        .SubItems(2) = "Property Get"
                    End With
                End If

            Case emtPropertyLet
                If intTabIndex = etAll Or intTabIndex = etProperties Then
                    ' Add the property
                    With lstMembers.ListItems.Add
                        .Text = objMember.Name & "_Let"
                        .SmallIcon = "Property"
                        .SubItems(1) = ma_udtMembers(intIndex).DataType
                        .SubItems(2) = "Property Let"
                    End With
                End If

            Case emtPropertySet
                If intTabIndex = etAll Or intTabIndex = etProperties Then
                    ' Add the property
                    With lstMembers.ListItems.Add
                        .Text = objMember.Name & "_Set"
                        .SmallIcon = "Property"
                        .SubItems(1) = ma_udtMembers(intIndex).DataType
                        .SubItems(2) = "Property Set"
                    End With
                End If
        End Select
    Next intIndex
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 9 ' Subscript out of range, user selected a tab before selecting a base class
            Exit Sub
        
        Case Else
            ProcessError Err, mc_strModuleID & "RefreshMemberList"
    End Select
End Sub


'==============================================================================
' UpdateProject - Create the new class and add the properties and methods
'
' Arguments: None
'
' Returns: True if creation was successful, False otherwise
'==============================================================================
Private Function UpdateProject() As Boolean
    On Error GoTo ErrorHandler
    Dim objComponent As VBComponent
    Dim objModule As VBIDE.CodeModule
    Dim strBaseClass As String
    Dim udtMember As gudt_Member
    Dim intIndex As Integer
    Dim strMember As String                 ' We will build the current member in this string
    Dim strTemp As String
    
    ' Get the name of the base class we are inheriting
    strBaseClass = lstBaseClasses.SelectedItem.Text

    ' Create a new class if possible
    Set objComponent = g_objVBInstance.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    On Error Resume Next
    objComponent.Name = txtChildClass.Text
    If Not (Err.Number = 0) Then
        ' There is already a class with that name
        MsgBox "'" & txtChildClass.Text & "' already exists!", vbOKOnly + vbExclamation, "Child Class Builder"
        g_objVBInstance.ActiveVBProject.VBComponents.Remove objComponent
        UpdateProject = False
        Exit Function
    End If
    On Error GoTo 0

    ' Class created, now add the contents
    Set objModule = objComponent.CodeModule
        
    ' Implement the base class, and create its member variable
    AppendLines objModule, "Implements " & strBaseClass & "\n\nPrivate m_objSuper As " & strBaseClass

    ' Insert Overloaded members
    For intIndex = UBound(ma_udtMembers) To 0 Step -1 ' To preserve original function order
        udtMember = ma_udtMembers(intIndex)
        
        With udtMember
            Select Case .Type
                Case emtFunction
                    AppendLines objModule, "\n\n" & .Declaration _
                        & "\n\t" & .Name & " = m_objSuper." & .Name & "(" & Parameters(.Declaration, True) & ")\n" _
                        & "End Function"
                    
                Case emtSub
                    AppendLines objModule, "\n\n" & .Declaration _
                        & "\n\tm_objSuper." & .Name & " " & Parameters(.Declaration, True) _
                        & "\nEnd Sub"
                    
                Case emtPropertyGet
                    AppendLines objModule, "\n\n" & .Declaration _
                        & "\n\t" & .Name & " = m_objSuper." & .Name & "(" & Parameters(.Declaration, True) & ")\n" _
                        & "End Property"
                    
                Case emtPropertyLet
                    AppendLines objModule, "\n\n" & .Declaration _
                        & "\tm_objSuper." & .Name & " = " & Parameters(.Declaration, True) _
                        & "\nEnd Property"
    
                Case emtPropertySet
                    AppendLines objModule, "\n\n" & .Declaration _
                        & "\tSet m_objSuper." & .Name & " = " & Parameters(.Declaration, True) _
                        & "\nEnd Property"
            End Select
        End With
    Next intIndex

    ' Add Initialize and Terminate events
    AppendLines objModule, "\n\nPrivate Sub Class_Initialize()\n\tSet m_objSuper = New " & strBaseClass & "\nEnd Sub"
    AppendLines objModule, "\n\nPrivate Sub Class_Terminate()\n\tSet m_objSuper = Nothing\nEnd Sub"

    ' Insert interface members
    For intIndex = UBound(ma_udtMembers) To 0 Step -1 ' To preserve original function order
        udtMember = ma_udtMembers(intIndex)
        
        With udtMember
            ' Build function name
            strTemp = strBaseClass & "_" & .Name
            
            Select Case .Type
                Case emtFunction
                    AppendLines objModule, "\n\nPrivate Function " & strTemp & "(" & Parameters(.Declaration, False) & ") As " & .DataType _
                        & "\n\t" & strTemp & " = " & .Name & "(" & Parameters(.Declaration, True) & ")\n" _
                        & "End Function"
                    
                Case emtSub
                    AppendLines objModule, "\n\nPrivate Sub " & strTemp & "(" & Parameters(udtMember.Declaration, False) & ")\n" _
                        & "\t" & udtMember.Name & " " & Parameters(udtMember.Declaration, True) & vbCrLf _
                        & "End Sub"
                    
                Case emtPropertyGet
                    AppendLines objModule, "\n\nPrivate Property Get " & strTemp & "(" & Parameters(.Declaration, False) & ") As " & .DataType _
                        & "\n\t" & strTemp & " = " & .Name & "(" & Parameters(.Declaration, True) & ")\n" _
                        & "End Property"
                        
                Case emtPropertyLet
                    AppendLines objModule, "\n\nPrivate Property Let " & strTemp & "(" & Parameters(.Declaration, False) & ")\n" _
                        & "\t" & .Name & " = " & Parameters(.Declaration, True) & vbCrLf _
                        & "End Property"
                
                Case emtPropertySet
                    AppendLines objModule, "\n\nPrivate Property Set " & strTemp & "(" & Parameters(.Declaration, False) & ")\n" _
                        & "\tSet " & .Name & " = " & Parameters(.Declaration, True) & vbCrLf _
                        & "End Property"
            End Select
        End With
    Next intIndex
    m_blnDirty = False
    Exit Function
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "UpdateProject"

End Function


'==============================================================================
' Form_Load - The form is about to be shown, or someone is accessing a control
' property.
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim objComponent As VBComponent

    ' Initialize picIcon (Used by the connect class to fill the menu item icon)
    picIcon.Picture = Me.Icon
    
    ' Show the All tab
    Set tsMembers.SelectedItem = tsMembers.Tabs(etAll)

    ' Add all classes to the base-class list
    If Not g_objVBInstance.ActiveVBProject Is Nothing Then
        lstBaseClasses.ListItems.Clear
        For Each objComponent In g_objVBInstance.ActiveVBProject.VBComponents
            If objComponent.Type = vbext_ct_ClassModule Then
                With lstBaseClasses.ListItems.Add
                    .Text = objComponent.Name
                    .SmallIcon = "Class"
                End With
            End If
        Next objComponent
    End If
    
    ' Initialize the splitter bar
    Set m_objSplitter = New CSplitDDC
    With m_objSplitter
        .SplitObject = picSplitter
        .Orientation = espVertical
        .Border(espbLeft) = 32
        .Border(espbRight) = 32
    End With
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load"
    
End Sub


'==============================================================================
' Form_QueryUnload - The user is trying to close the form
'==============================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandler
    
    If m_blnDirty And Not (txtChildClass.Text = "") Then
        Select Case MsgBox("Update Project with changes?", vbYesNoCancel + vbQuestion, "Child Class Builder")
            Case vbYes
                UpdateProject
            
            Case vbCancel
                Cancel = True
            
        End Select
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_QueryUnload"
    
End Sub


'==============================================================================
' Form_Resize - Move/Resize form controls
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    Dim lngRight As Long
    Dim intBorder As Integer
    
    intBorder = 60
    With picSplitter
        ' Move/Resize the splitter
        lngRight = .Left + .Width
        .Height = Me.ScaleHeight
        
        ' Move/Resize the left frame
        picLeft.Move 0, intBorder, .Left, (.Height - intBorder)
        txtChildClass.Width = picLeft.ScaleWidth - txtChildClass.Left - 10
        With lstBaseClasses
            .Move 0, .Top, picLeft.ScaleWidth, picLeft.ScaleHeight - .Top
            .ColumnHeaders(1).Width = .Width - 90
        End With

        ' Move/Resize the right frame
        tsMembers.Move lngRight, intBorder, (Me.ScaleWidth - lngRight), (.Height - tsMembers.Top)
        lstMembers.Move (tsMembers.Left + intBorder), lstMembers.Top, (tsMembers.Width - (2 * intBorder)), (tsMembers.Height - lstMembers.Top)
    End With
End Sub


'==============================================================================
' lstBaseClasses_Click - The user changed the base-class selection
'==============================================================================
Private Sub lstBaseClasses_Click()
    On Error GoTo ErrorHandler
    
    ' Reload the member array with the members of the new class
    RefreshMemberArray
    
    ' Reload the member list
    RefreshMemberList
    
    ' Reset the dirty flag
    m_blnDirty = True
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "lstBaseClasses_Click"

End Sub


'==============================================================================
' mnuFileReplace_Click - Create the new child class
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub mnuFileReplace_Click()
    On Error GoTo ErrorHandler
    
    UpdateProject
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileReplace_Click"
    
End Sub


'==============================================================================
' mnuFileExit_Click - The user wants to leave
'==============================================================================
Private Sub mnuFileExit_Click()
    On Error Resume Next
    
    Unload Me
End Sub


'==============================================================================
' mnuHelpAbout_Click - Show the About form
'
' Arguments: None
'
' Notes:
'==============================================================================
Private Sub mnuHelpAbout_Click()
    On Error GoTo ErrorHandler
    
    frmAbout.Show vbModal, Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuHelpAbout_Click"

End Sub


'===============================================================================
' mnuViewTab_Click - The user has selected a different view from the menu
'
' Arguments:
'   Index - Which view should we set
'
' Notes: The indexes on the menu items must be kept in sync with the tab numbers
' for this to work correctly.
'===============================================================================
Private Sub mnuViewTab_Click(Index As Integer)
    On Error GoTo ErrorHandler
    
    Set tsMembers.SelectedItem = tsMembers.Tabs(Index)
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuViewTab_Click"
    
End Sub


'===============================================================================
' picSplitter_MouseDown - Redirect splitter MouseDowns to the Splitter object
'
' Arguments:
'   Button (IN) - Indicates which button was pressed
'   Shift (IN) - Indicated the state of the Shift, Ctrl, and Alt keys
'   X,Y (IN) - Indicate the position of the mouse when the even occurred
'
' Notes:
'===============================================================================
Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    
    m_objSplitter.SplitterMouseDown Me.hwnd, X, Y
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "picSplitter_MouseDown"
        
End Sub


'===============================================================================
' picSplitter_MouseMove - Redirect splitter MouseMoves to the Splitter object
'
' Arguments:
'   Button (IN) - Indicates the state of the mouse buttons
'   Shift (IN) - Indicated the state of the Shift, Ctrl, and Alt keys
'   X,Y (IN) - Indicate the position of the mouse when the even occurred
'
' Notes:
'===============================================================================
Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler

    m_objSplitter.SplitterMouseMove X, Y
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "picSplitter_MouseMove"

End Sub


'===============================================================================
' picSplitter_MouseUp - Redirect splitter mouse up events to the Splitter object
'
' Arguments:
'   Button (IN) - Indicates which button was released
'   Shift (IN) - Indicated the state of the Shift, Ctrl, and Alt keys
'   X,Y (IN) - Indicate the position of the mouse when the even occurred
'
' Notes:
'===============================================================================
Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler

    If (m_objSplitter.SplitterMouseUp(X, Y)) Then Form_Resize
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "picSplitter_MouseUp"

End Sub


'===============================================================================
' tsMembers_Click - The tabstrip selection may have changed
'===============================================================================
Private Sub tsMembers_Click()
    On Error GoTo ErrorHandler
    
    RefreshMemberList
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "tsMembers_Click"

End Sub

