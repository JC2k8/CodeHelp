VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Plugin const
Private Const CH_LONGNAME As String = "Complexity Metrics for VB"
Private Const CH_DESCRIPTION As String = "Calculates various code complexity metrics."
Private Const CH_COPYRIGHT As String = "clayreimann@gmail.com"

Private Const TOOL_GUID As String = "{7140818C-0E51-4DE8-A2DA-873F6E102D5A}"

Private m_CHCorePtr As Long
Private m_Enabled As Boolean

Private m_ToolWindow    As Window
Private m_MetricsPanel  As MetricsPanel
Private m_LastComponent As VBComponent
Private WithEvents m_ComponentEvents As VBComponentsEvents
Attribute m_ComponentEvents.VB_VarHelpID = -1

Implements ICHPlugin

Private Property Let ICHPlugin_CHCore(ByVal RHS As Long)
    'Save the Pointer for later use
    m_CHCorePtr = RHS
End Property

Private Property Get ICHPlugin_CopyRight() As String
    ICHPlugin_CopyRight = CH_COPYRIGHT
End Property

Private Property Get ICHPlugin_Description() As String
    ICHPlugin_Description = CH_DESCRIPTION
End Property

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   m_Enabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   ICHPlugin_Enabled = m_Enabled
End Property

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean
'Enable/disable help button in plugin manager
End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean
'Enable/disable properties button in plugin manager
End Property

Private Property Get ICHPlugin_LongName() As String
    ICHPlugin_LongName = CH_LONGNAME
End Property

Private Property Get ICHPlugin_Name() As String
    ICHPlugin_Name = "CHCodeMetrics"
End Property

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, custom() As Variant)
    Dim oActiveWin As Window
    Dim oCHCore As ICHCore
    
    Set oCHCore = GetCHCore(m_CHCorePtr)
    
    Set m_ComponentEvents = oCHCore.VBE.Events.VBComponentsEvents(Nothing)
    Set m_ToolWindow = oCHCore.VBE.Windows.CreateToolWindow(oCHCore.AddInInst, "CHCodeComplexity.MetricsPanel", _
                                                            "Metrics", TOOL_GUID, m_MetricsPanel)
    
    Call m_MetricsPanel.Initalize(oCHCore.VBE)
    m_ToolWindow.Visible = True
    
    Set oCHCore = Nothing
    Set oActiveWin = Nothing
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, custom() As Variant)
    'Do your clean up here'
    Set m_LastComponent = Nothing
    Set m_MetricsPanel = Nothing
    Set m_ToolWindow = Nothing
    Set m_ComponentEvents = Nothing
End Sub

' for a discussion of wParam and lParam see https://msdn.microsoft.com/en-us/library/windows/desktop/ms644984(v=vs.85).aspx
Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, lReturn As Long, wParam As Long, lParam As Long)
    Select Case wParam
        Case vbKeyM
            If (lParam And &HC0000000) = 0 Then ' check for key down messages only
                If GetAsyncKeyState(vbKeyControl) Then
                    lReturn = 1
                    bHandled = True
                    If Not m_LastComponent Is Nothing Then Call m_MetricsPanel.UpdateCodeMetrics(m_LastComponent.CodeModule)
                End If
            End If
            
    End Select
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, _
                                bHandled As Boolean, lReturn As Long)
    'subclassed message goes here
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                                    bHandled As Boolean, placeSubclass As Boolean, before As Boolean)
    'Hook msg goes here
End Sub

Private Sub ICHPlugin_ShowHelp()
'Show help instruction when user click on help button in plugin manager
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
'show property dialog when user click on properties button in plugin manager
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub m_ComponentEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    Set m_LastComponent = VBComponent
    Call CalculateMetrics(VBComponent)
End Sub

Private Sub m_ComponentEvents_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    Set m_LastComponent = VBComponent
    Call CalculateMetrics(VBComponent)
End Sub

Private Sub CalculateMetrics(comp As VBComponent)
    Call m_MetricsPanel.UpdateCodeMetrics(comp.CodeModule)
End Sub
