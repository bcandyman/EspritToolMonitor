VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Monitor_ToolId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================
'         CLASS TO MONITOR CHANGES IN TOOL TECHNOLOGY
'============================================================

'This class will monitor the Esprit document for changes within a tool technology.
'Any change that is made to a tool will trigger the toolId to be recaculated.
'This class will currently only monitor End Mill type tools. All other tool type ToolIds will NOT be modified

Dim WithEvents Tools  As Tools
Attribute Tools.VB_VarHelpID = -1

'Initialize the class
Private Sub Class_Initialize()
    Set Tools = DocTools
End Sub

'Runs when a tool has been modified
Private Sub Tools_OnModify(ByVal OldName As String, ByVal CLCode As Long, ByVal Tool As Object)
    Dim T   As EspritTechnology.Tool
    Set T = Tool
        
    Select Case T.ToolStyle
        Case espMillToolBallMill
            T.ToolID = GetMillBallMillToolId(T)
        Case espMillToolBullNose
            T.ToolID = GetMillBullMillToolId(T)
        Case espMillToolEndMill
            T.ToolID = GetMillEndMillToolId(T)
    End Select
End Sub

'Returns the Esprit Document
Private Function Doc() As Esprit.Document
    Set Doc = App.Document
End Function

'Returns the Esprit Application
Private Function App() As Esprit.Application
    Set App = Esprit.Application
End Function

'Returns the Esprit Tools collection
Private Function DocTools() As EspritTools.Tools
    Set DocTools = Doc.Tools
End Function

'Returns the EspritTechnology.TechnologyUtility Object
Private Function TechUtil() As EspritTechnology.TechnologyUtility
    Set TechUtil = Doc.TechnologyUtility
End Function

'Returns the ToolId for a Ball Mill
Private Function GetMillBallMillToolId(ToolTech As EspritTechnology.ToolMillBallMill) As String
    Dim D   As String
    Dim F   As String
    Dim P   As String
    With ToolTech
        D = CustomFormat(.ToolDiameter)
        P = CustomFormat(.ToolLength)
        F = CustomFormat(.CuttingLength)
    End With
    GetMillBallMillToolId = "BM_" & D & "DIA_" & F & "FL_" & P & "HO"
End Function

'Returns the ToolId for a Bull Nosed Mill
Private Function GetMillBullMillToolId(ToolTech As EspritTechnology.ToolMillBullNose) As String
    Dim C   As String
    Dim D   As String
    Dim F   As String
    Dim P   As String
    With ToolTech
        C = CustomFormat(.ToolRadius)
        D = CustomFormat(.ToolDiameter)
        F = CustomFormat(.CuttingLength)
        P = CustomFormat(.ToolLength)
    End With
    GetMillBullMillToolId = "BN_" & D & "DIA_" & C & "CR_" & F & "FL_" & P & "HO"
End Function

'Returns the ToolId for a Square End Mill
Private Function GetMillEndMillToolId(ToolTech As EspritTechnology.ToolMillEndMill) As String
    Dim D   As String
    Dim F   As String
    Dim P   As String
    With ToolTech
        D = CustomFormat(.ToolDiameter)
        P = CustomFormat(.ToolLength)
        F = CustomFormat(.CuttingLength)
    End With
    GetMillEndMillToolId = "EM_" & D & "DIA_" & F & "FL_" & P & "HO"
End Function

'Applies a standard format to a value
Private Function CustomFormat(Value As Variant) As String
    CustomFormat = Format(Value, "###.00")
End Function
