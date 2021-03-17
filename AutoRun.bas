Attribute VB_Name = "AutoRun"
Option Explicit

'============================================================
'                   AUTO RUN ON STARTUP
'============================================================

'These procedures are trigger automatically

Dim ToolMon As Monitor_ToolId

'Fires when a new document is opened
Sub AutoOpen()
    'Initialize the Monitor_ToolId class
    Set ToolMon = New Monitor_ToolId
End Sub
