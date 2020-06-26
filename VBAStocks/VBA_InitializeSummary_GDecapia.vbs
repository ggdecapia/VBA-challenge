Attribute VB_Name = "Module2"
'This subroutine clears the summary columns positions I to Q
Sub initialize()
    
    For Each ws In Worksheets
        ws.Range("I1:Q1048576").Value = ""
        ws.Range("I1:Q1048576").Interior.ColorIndex = 0
    Next ws
    
End Sub


