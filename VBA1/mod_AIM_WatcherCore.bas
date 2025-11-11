Attribute VB_Name = "mod_AIM_WatcherCore"
'=== Standard Module: mod_AIM_WatcherCore ===
Option Explicit
Public AIM_AppEvents As clsAIM_AppEvents

'---------------------------------------------------------------
' Automatically start watcher when Excel loads PERSONAL.XLSB
'---------------------------------------------------------------
Public Sub Auto_Open()
    Set AIM_AppEvents = New clsAIM_AppEvents
    Set AIM_AppEvents.App = Application
    Application.StatusBar = "? AIM Two-Way Sync active"
End Sub

'---------------------------------------------------------------
' Manual reload (use if sync ever stops)
'---------------------------------------------------------------
Public Sub Reset_AIM_Watcher()
    Set AIM_AppEvents = New clsAIM_AppEvents
    Set AIM_AppEvents.App = Application
    Application.StatusBar = "?? AIM Two-Way Sync reloaded"
End Sub

'---------------------------------------------------------------
' Manual stop (disables watcher)
'---------------------------------------------------------------
Public Sub Stop_AIM_Watcher()
    Set AIM_AppEvents = Nothing
    Application.StatusBar = "?? AIM Two-Way Sync disabled"
End Sub


