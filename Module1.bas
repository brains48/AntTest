Attribute VB_Name = "Module1"
Option Explicit

Sub DisplayName()

Dim intCount As Integer

For intCount = 1 To 3
MsgBox Sheets("Ant").Range("A1")
Next intCount

End Sub
