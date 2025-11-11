'=== Standard Module: mod_AIM_WatcherCore ===
Option Explicit
Public AIM_AppEvents As clsAIM_AppEvents

' Initialize watcher on Excel startup
Public Sub Auto_Open()
    Set AIM_AppEvents = New clsAIM_AppEvents
    Set AIM_AppEvents.App = Application
    Application.StatusBar = "✅ AIM Two-Way Sync active"
End Sub

' Manual reload command
Public Sub Reset_AIM_Watcher()
    Set AIM_AppEvents = New clsAIM_AppEvents
    Set AIM_AppEvents.App = Application
    Application.StatusBar = "🔄 AIM Two-Way Sync reloaded"
End Sub

' Manual stop command
Public Sub Stop_AIM_Watcher()
    Set AIM_AppEvents = Nothing
    Application.StatusBar = "🛑 AIM Two-Way Sync disabled"
End Sub
