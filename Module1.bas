Attribute VB_Name = "Module1"


Function Populate_to_Excel(sForm As Form)
Dim XLApp As Excel.Application
Dim XLWorkBook As Excel.Workbook
Dim Rows, Cols As Integer
Dim i, h, g As Integer
Dim New_Col As Boolean

    If sForm.fg.Rows <= 1 Then
        MsgBox "No Data to extract", vbInformation, App.Title
        Exit Function
    End If
    
    Set XLApp = CreateObject("Excel.application")
    Set XLWorkBook = XLApp.Workbooks.Add
    
    Dim New_Column As Boolean
    
    With sForm.fg
        Rows = .Rows
        Cols = .Cols
        i = 0
        g = 0
        For h = 0 To Cols - 1
            For i = 0 To Rows - 1
                If .ColIsVisible(h) = False Then
                    XLApp.Cells(i + 1, g + 1).Value = .TextMatrix(i, h)
                    New_Column = True
                Else
                    New_Column = False
                End If
                
            Next i
            If New_Column = True Then g = g + 1
        Next h
    End With
    
    XLApp.Rows(1).Font.Bold = True
    XLApp.Cells.Select
    XLApp.Columns.AutoFit
    XLApp.Cells(1, 1).Select
    XLApp.Application.Visible = True
    
    Set XLWorkBook = Nothing
    Set XLApp = Nothing
End Function
