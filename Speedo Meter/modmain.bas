Attribute VB_Name = "modMain"
Option Explicit

Public v(0 To 2) As clsVUmeter

Public Sub Initialize()
    Set v(0) = New clsVUmeter
    Set v(1) = New clsVUmeter
    Set v(2) = New clsVUmeter
End Sub

Public Sub Shutdown()
    Set v(0) = Nothing
    Set v(1) = Nothing
    Set v(2) = Nothing
End Sub

Sub Main()
    Initialize
    
    Form1.Show
End Sub
