Imports Microsoft.Office.Interop
'
'See https://support.microsoft.com/en-us/help/302094/how-to-automate-excel-from-visual-basic-net-to-fill-or-to-obtain-data
'
Public Class Form1
    'Keep the application object and the workbook object global, so you can  
    'retrieve the data in Button2_Click that was set in Button1_Click.
    Dim objApp As Excel.Application
    Dim objBook As Excel._Workbook

    Private Sub Button1_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button1.Click
        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range

        ' Create a new instance of Excel and start a new workbook.
        objApp = New Excel.Application()
        objBooks = objApp.Workbooks
        objBook = objBooks.Add
        objSheets = objBook.Worksheets
        objSheet = objSheets(1)

        'Get the range where the starting cell has the address
        'm_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
        range = objSheet.Range("A1", Reflection.Missing.Value)
        range = range.Resize(5, 5)

        If (Me.FillWithStrings.Checked = False) Then
            'Create an array.
            Dim saRet(5, 5) As Double

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5

                    'Put a counter in the cell.
                    saRet(iRow, iCol) = iRow * iCol
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet

        Else
            'Create an array.
            Dim saRet(5, 5) As String

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5
                    'Put the row and column address in the cell.
                    saRet(iRow, iCol) = iRow.ToString() + "|" + iCol.ToString()
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet
        End If

        'Return control of Excel to the user.
        objApp.Visible = True
        objApp.UserControl = True

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
        objBooks = Nothing
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button2.Click
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range

        'Get a reference to the first sheet of the workbook.
        On Error GoTo ExcelNotRunning
        objSheets = objBook.Worksheets
        objSheet = objSheets(1)
ExcelNotRunning:
        If (Not (Err.Number = 0)) Then
            MessageBox.Show("Cannot find the Excel workbook.  Try clicking Button1 to " +
            "create an Excel workbook with data before running Button2.",
            "Missing Workbook?")

            'We cannot automate Excel if we cannot find the data we created, 
            'so leave the subroutine.
            Exit Sub
        End If

        'Get a range of data.
        range = objSheet.Range("A1", "E5")

        'Retrieve the data from the range.
        Dim saRet(,) As Object
        saRet = range.Value

        'Determine the dimensions of the array.
        Dim iRows As Long
        Dim iCols As Long
        iRows = saRet.GetUpperBound(0)
        iCols = saRet.GetUpperBound(1)

        'Build a string that contains the data of the array.
        Dim valueString As String
        valueString = "Array Data" + vbCrLf

        Dim rowCounter As Long
        Dim colCounter As Long
        For rowCounter = 1 To iRows
            For colCounter = 1 To iCols

                'Write the next value into the string.
                valueString = String.Concat(valueString,
                    saRet(rowCounter, colCounter).ToString() + ", ")

            Next colCounter

            'Write in a new line.
            valueString = String.Concat(valueString, vbCrLf)
        Next rowCounter

        'Report the value of the array.
        MessageBox.Show(valueString, "Array Values")

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
    End Sub
End Class
