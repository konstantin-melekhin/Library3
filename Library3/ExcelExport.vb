Imports System.Data.SqlClient

Public Module ExcelExport
    'экспорт в excel
    Public Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" (
        ByVal hWnd As Integer, ByVal lpOperation As String,
        ByVal lpFile As String, ByVal lpParameters As String,
        ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Public myFile As String
    Public Sub exportExcel(ByVal grdView As DataGridView, ByVal fileName As String,
      ByVal fileExtension As String, ByVal filePath As String)

        ' Choose the path, name, and extension for the Excel file
        myFile = filePath & "\" & fileName & fileExtension
        Try
            ' Open the file and write the headers
            Dim fs As New IO.StreamWriter(myFile, False)
            fs.WriteLine("<?xml version=""1.0""?>")
            fs.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
            fs.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
            ' Create the styles for the worksheet
            fs.WriteLine("  <ss:Styles>")
            ' Style for the column headers
            fs.WriteLine("    <ss:Style ss:ID=""1"">")
            fs.WriteLine("      <ss:Font ss:Bold=""1""/>")
            fs.WriteLine("      <ss:Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" " &
                "ss:WrapText=""1""/>")
            fs.WriteLine("      <ss:Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
            fs.WriteLine("    </ss:Style>")
            ' Style for the column information
            fs.WriteLine("    <ss:Style ss:ID=""2"">")
            fs.WriteLine("      <ss:Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>")
            fs.WriteLine("    </ss:Style>")
            fs.WriteLine("  </ss:Styles>")

            ' Write the worksheet contents
            fs.WriteLine("<ss:Worksheet ss:Name=""Report"">")
            fs.WriteLine("  <ss:Table>")
            For i As Integer = 0 To grdView.Columns.Count - 1
                fs.WriteLine(String.Format("    <ss:Column ss:Width=""{0}""/>",
                grdView.Columns.Item(i).Width))
            Next
            fs.WriteLine("    <ss:Row>")
            For i As Integer = 0 To grdView.Columns.Count - 1
                fs.WriteLine(String.Format("      <ss:Cell ss:StyleID=""1"">" &
                    "<ss:Data ss:Type=""String"">{0}</ss:Data></ss:Cell>",
                    grdView.Columns.Item(i).HeaderText))
            Next
            fs.WriteLine("    </ss:Row>")

            ' Check for an empty row at the end due to Adding allowed on the DataGridView
            Dim subtractBy As Integer, cellText As String
            If grdView.AllowUserToAddRows = True Then subtractBy = 2 Else subtractBy = 1
            ' Write contents for each cell
            For i As Integer = 0 To grdView.RowCount - subtractBy
                fs.WriteLine(String.Format("    <ss:Row ss:Height=""{0}"">",
                    grdView.Rows(i).Height))
                For intCol As Integer = 0 To grdView.Columns.Count - 1
                    cellText = grdView.Item(intCol, i).Value
                    ' Check for null cell and change it to empty to avoid error
                    If cellText = vbNullString Then cellText = ""
                    fs.WriteLine(String.Format("      <ss:Cell ss:StyleID=""2"">" &
                        "<ss:Data ss:Type=""String"">{0}</ss:Data></ss:Cell>",
                        cellText.ToString))
                Next
                fs.WriteLine("    </ss:Row>")
            Next

            ' Close up the document
            fs.WriteLine("  </ss:Table>")
            fs.WriteLine("</ss:Worksheet>")
            fs.WriteLine("</ss:Workbook>")
            fs.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'Open the file in Microsoft Excel
        '10 = SW_SHOWDEFAULT
        'ShellEx(Me.Handle, "Open", myFile, "", "", 10)
    End Sub

End Module
