VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10050
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11745
   _ExtentX        =   20717
   _ExtentY        =   17727
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Child Class Builder Utility"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================================
' Connect - Add-in connect class designer
'
' Version   Date        User            Notes
'   1.0     04/20/01    Mel Grubb II    Initial version
'==============================================================================
Option Explicit


'==============================================================================
' Private Member Constants
'==============================================================================
Private Const mc_strModuleID = "Connect."

'==============================================================================
' Private Member Variables
'==============================================================================
Private m_objMenuItem As Office.CommandBarButton
Private WithEvents m_objMenuEvents As CommandBarEvents
Attribute m_objMenuEvents.VB_VarHelpID = -1


'==============================================================================
' AddinInstance_OnConnection - Occurs when an add-in is connected to the
' Visual Basic IDE, either through the Add-In Manager dialog box or another
' add-in
'
' Arguments:
'   Application - An object representing the instance of the current Visual
'       Basic session.
'   ConnectMode - Indicates how the addin was started
'   AddInInst - An AddIn object representing the instance of the add-in.
'   Custom()  An array of variant expressions to hold user-defined data.
'
' Notes:
'==============================================================================
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo ErrorHandler
    Dim objAddinMenu As Office.CommandBar
    
    ' Set the global application reference
    Set g_objVBInstance = Application
    
    ' Get a handle to the Add-in menu
    Set objAddinMenu = g_objVBInstance.CommandBars("Add-Ins")
    If Not (objAddinMenu Is Nothing) Then
        ' Add new menu item to add-ins menu
        Set m_objMenuItem = objAddinMenu.Controls.Add(MsoControlType.msoControlButton)
        With m_objMenuItem
            .Caption = g_strAppTitle
            Clipboard.SetData frmMain.picIcon.Image
            .PasteFace
            Clipboard.Clear
        End With
        
        ' Set the reference to the event handler
        Set m_objMenuEvents = g_objVBInstance.Events.CommandBarEvents(m_objMenuItem)
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "AddinInstance_OnConnection"

End Sub


'==============================================================================
' AddinInstance_OnDisconnection
'==============================================================================
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next

    ' Remove the add-in menu item
    m_objMenuItem.Delete

    ' Unload the interface
    Unload frmMain
End Sub


'==============================================================================
' m_objMenuEvents_Click
'==============================================================================
Private Sub m_objMenuEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo ErrorHandler
    
    frmMain.Show
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "m_objMenuEvents_Click"
    
End Sub
