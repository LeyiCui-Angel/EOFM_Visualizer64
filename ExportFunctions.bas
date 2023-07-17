Attribute VB_Name = "ExportFunctions"
Option Explicit

Private Function BrowseFolder(Optional Caption As String, Optional InitialFolder As String) As String
    Dim SH As Object
    Dim F As Object
    
    Set SH = CreateObject("Shell.Application")
    Set F = SH.BrowseForFolder(0&, Caption, &H1, InitialFolder)
    
    If Not F Is Nothing Then
        BrowseFolder = F.Self.path
    End If
End Function

Public Sub ExportAllToPNG()
    Dim MyFolder As String
    Dim i As Integer
    
    MyFolder = BrowseFolder("Please select a folder.")
    
    For i = 1 To ActiveDocument.Pages.Count
        With ActiveDocument.Pages(i)
            .Export (MyFolder & "\" & Replace(.Name, ":", "-") & ".png")
        End With
    Next i
End Sub
