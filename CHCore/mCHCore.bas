Attribute VB_Name = "mCHCore"
Option Explicit

Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'Public vars
Public HookMon As HookMonitor

'for passing to manual oPlugin connect/disconnect, not used but need to be pass along
Public customVar() As Variant

'this is a pointer to Connect object, we'll need it when we re-enable a oPlugin at runtime
Public gPtr As Long

Private m_Plugins               As Plugins

Public Property Get Plugins() As Plugins
    Set Plugins = m_Plugins
End Property

Public Property Set Plugins(value As Plugins)
    Set m_Plugins = value
End Property



