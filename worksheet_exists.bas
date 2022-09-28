Attribute VB_Name = "worksheet_exists"
Option Explicit

Public Function funcWorksheetExists(ByVal strSheetName As String) As Boolean
    funcWorksheetExists = False
    If strSheetName <> "" Then
        Dim i As Integer
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = strSheetName Then
                funcWorksheetExists = True
            End If
        Next i
    End If
End Function

Private Sub TestWorksheetExists()
    'Place your cursor in this procedure and click the play button.
    'Make sure you have the Immediate Window showing (Ctrl + G)
    Debug.Print "Worksheet named ""Sheet1"": " & funcWorksheetExists("Sheet1")
    
    Dim strSheetName As String
    strSheetName = "Sheet2"
    Debug.Print "Worksheet named """ & strSheetName & """: " & funcWorksheetExists(strSheetName)
    
    Debug.Print "Worksheet named ""SalesFigures"": " & funcWorksheetExists("SalesFigures")
End Sub
