If MsgBox ("Clear the file open history in Excel. Are you sure?", 289, "Confirmation") = 1 Then
    With WScript.CreateObject("Excel.Application")
        While (.RecentFiles.Count > 0)
            .RecentFiles(1).Delete
        Wend
		.Quit()
    End With
    MsgBox "Finished.", 64, "Information"
End If
