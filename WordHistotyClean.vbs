If MsgBox ("Clear the file open history in Word. Are you sure?", 289, "Confirmation") = 1 Then
    With WScript.CreateObject("Word.Application")
        While (.RecentFiles.Count > 0)
            .RecentFiles(1).Delete
        Wend
		.Quit()
    End With
    MsgBox "Finished.", 64, "Information"
End If
