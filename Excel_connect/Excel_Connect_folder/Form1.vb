Imports Microsoft.Office.Interop
'
'See https://support.microsoft.com/en-us/help/302094/how-to-automate-excel-from-visual-basic-net-to-fill-or-to-obtain-data
'


Public Class Form1
    'https://social.msdn.microsoft.com/Forums/vstudio/en-US/03d76f3d-b91a-4707-89ce-ab752c6823e2/excel-workbook-problem-in-vbnet?forum=exceldev
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With DataGridView1
            .ColumnCount = 3
            .Rows.Clear()
            .Rows.Add(4)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            .Columns(0).HeaderText = "H1"
            .Columns(1).HeaderText = "H2"
            .Columns(2).HeaderText = "H3"

            For i = 0 To .RowCount - 1
                .Rows(i).Cells(0).Value = "00"
                .Rows(i).Cells(1).Value = "01"
                .Rows(i).Cells(2).Value = "02"
            Next
        End With

        With DataGridView2
            .ColumnCount = 3
            .Rows.Clear()
            .Rows.Add(4)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            .Columns(0).HeaderText = "H1"
            .Columns(1).HeaderText = "H2"
            .Columns(2).HeaderText = "H3"
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "Please select a PSD file"
        OpenFileDialog1.InitialDirectory = "C:\Repos\MS-excel-connection\Excel_connect\"
        OpenFileDialog1.Filter = "PSD Files|*.xlsx"
        OpenFileDialog1.FileName = "PSD_*.xlsx"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Retrieve_xls_file(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button1.Click
        Save_excel_file()
    End Sub

    Private Sub Save_excel_file()
        Dim xlApp As Excel.Application
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim range As Excel.Range = Nothing

        ' Create a new instance of Excel and start a new workbook.
        xlApp = New Excel.Application()
        xlBooks = xlApp.Workbooks
        xlBook = xlBooks.Add()
        xlSheet = CType(xlBook.ActiveSheet, Excel.Worksheet)
        xlSheet.Name = "Save_excel_file"

        'Get the range where the starting cell has the address
        'm_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
        range = xlSheet.Range("A1", Reflection.Missing.Value)
        range = range.Resize(15, 2)

        'Create an array.
        Dim saRet(15, 2) As String

        'Fill the array.
        Dim random As New Random()
        For iRow = 0 To 15
            saRet(iRow, 0) = random.Next(0, 50).ToString()
            saRet(iRow, 1) = random.Next(10, 150).ToString()
        Next iRow

        'Set the range value to the array.
        range.Value = saRet

        'Return control of Excel to the user.
        xlApp.Visible = True
        xlApp.UserControl = True
        xlBook.Close(SaveChanges:=True)
        xlApp.Quit()
        xlApp = Nothing
    End Sub

    Private Sub Retrieve_xls_file(xl_filename As String)
        Dim xlApp As New Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim saRet As Object(,)

        If My.Computer.FileSystem.FileExists(xl_filename) Then
            xlApp = New Excel.Application
            xlBook = xlApp.Workbooks.Open(xl_filename, IgnoreReadOnlyRecommended:=True, ReadOnly:=False, Editable:=True)

            xlSheet = CType(xlBook.ActiveSheet, Excel.Worksheet)
            xlSheet.Name = "Retrieve_xls_file"

            range = xlSheet.Range("A2", "B300")         'Get a range of data.
            saRet = CType(range.Value, Object(,))       'Retrieve the data from the range.

            '---- Lose the empty cells --
            Dim colcnt As Integer
            For rowC = 1 To saRet.GetUpperBound(0)
                If Not IsNothing(saRet(rowC, 1)) Then
                    colcnt += 1
                End If
            Next

            '---- Write the retrieved data to the DGV -----
            With DataGridView2
                .Rows.Clear()
                .Rows.Add(colcnt)       'Resize the dgv
                For rowC = 1 To colcnt
                    For colC = 1 To saRet.GetUpperBound(1)
                        .Rows(rowC - 1).Cells(colC - 1).Value = saRet(rowC, colC)
                    Next
                Next
            End With

            xlBook.Close(SaveChanges:=False)
            xlApp.Quit()
            xlApp = Nothing
        Else
            MsgBox("File does not exist")
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DGV_to_file()
    End Sub
    Private Sub DGV_to_file()
        Dim xl_filename As String = "C:\Temp\PSD_Typical_tst.xlsx"
        Dim xlApp As Excel.Application
        Dim xlBooks As Excel.Workbooks = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        'Dim range As Excel.Range = Nothing

        ' Create a new instance of Excel and start a new workbook.
        xlApp = New Excel.Application()
        xlBooks = xlApp.Workbooks
        xlBook = xlBooks.Add()
        xlSheet = CType(xlBook.ActiveSheet, Excel.Worksheet)
        xlSheet.Name = "DGV_to_file"

        '  If DataGridView1.DataSource IsNot Nothing Then
        Dim i, j As Integer
        For i = 1 To DataGridView1.RowCount - 1
            For j = 1 To DataGridView1.ColumnCount
                xlSheet.Cells(i + 1, j) = "33"  ' DataGridView1.Rows(i - 1).Cells(j - 1).Value
            Next
        Next


        xlBook.SaveAs(xl_filename, Excel.XlFileFormat.xlOpenXMLWorkbook)
        xlBook.Close(SaveChanges:=True, xl_filename)

        xlApp.Visible = True
        xlApp.UserControl = True
        xlApp.Quit()
        xlApp = Nothing
    End Sub

End Class
