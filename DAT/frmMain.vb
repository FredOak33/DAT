Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions
Imports System.Drawing.Printing


Public Class frmMain
    WithEvents bsUser As New BindingSource
    WithEvents bsEquip As New BindingSource
    WithEvents bsDevice As New BindingSource
    WithEvents bsVirt As New BindingSource
    Dim WithEvents mPrintDocument As New PrintDocument
    Dim mPrintBitMap As Bitmap

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Added new lines to test the date of the VZW_HN
        'If it hasn't been created today create a new one
        'If it has already been created today use the current one

        Dim strLastModified As String
        'Date Last Created
        strLastModified = System.IO.File.GetLastWriteTime("S:\Corp\DAT\VZW_HN.csv").ToShortDateString()

        If strLastModified <> Today Then
            'If VZW_HN is an older file create a new one from VZW_HN_Report.csv
            ConvertCSVToDataSet("S:\Corp\DAT\VZW_HN_Report.csv")
        Else
            'If the file is from today, use VZW_HN.csv
            csvToDatatable_2("S:\Corp\DAT\VZW_HN.csv", ",")
        End If

        btnDiscover.Text = "DISCOVER"

    End Sub
    Public Function csvToDatatable_2(ByVal filename As String, ByVal separator As String)
        Dim dt As New System.Data.DataTable
        Dim firstLine As Boolean = True
        If IO.File.Exists(filename) Then
            Using sr As New StreamReader(filename)
                While Not sr.EndOfStream
                    If firstLine Then
                        firstLine = False
                        Dim cols = sr.ReadLine.Split(separator)
                        For Each col In cols
                            dt.Columns.Add(New DataColumn(col, GetType(String)))
                        Next
                    Else
                        Dim data() As String = sr.ReadLine.Split(separator)
                        For i = 0 To 27
                            data(i) = data(i).Replace("""", "")
                        Next
                        dt.Rows.Add(data.ToArray)
                    End If
                End While
            End Using
        End If
        bsUser.DataSource = dt
        Return dt
    End Function
    Private Sub BtnDiscover_Click(sender As Object, e As EventArgs) Handles btnDiscover.Click

        Dim SSql As String = ""
        Dim asset As New DataSet

        Dim dt As New DataTable("Assets")

        btnDiscover.Text = "PROCESSING"
        btnDiscover.BackColor = Color.Blue
        btnDiscover.Enabled = False
        Dim i As Integer = 1
        Do Until i = 5000
            i += 1
        Loop
        btnDiscover.Enabled = True


        KaceReadOne("S:\Corp\DAT\DAT Device Summary Report.csv")
        KaceReadTwo("S:\Corp\DAT\DAT Device Summary Report.csv")

        Try

            dgAM.DataSource = bsDevice
            bsDevice.Filter = String.Format("[Assigned User ID] Like '%{0}%'", txtID.Text)
            dgAM.DataSource = bsDevice

            dgUser.DataSource = bsUser
            bsUser.Filter = String.Format("[User ID] Like '%{0}%'", txtID.Text)
            dgUser.DataSource = bsUser

            dgVirt.DataSource = bsVirt
            bsVirt.Filter = String.Format("([User Name] Like '%{0}%' OR [Comments] Like '%{0}%') AND [System Manufacturer] = 'VMware, Inc.'", txtID.Text)
            dgVirt.DataSource = bsVirt

            '====================================================================================

            dgUser.Columns("Contact Number").Visible = False
            dgUser.Columns("Contact Name").Visible = False
            dgUser.Columns("Contract activation Date").Visible = False
            dgUser.Columns("Contract length").Visible = False
            dgUser.Columns("Total current charges").Visible = False
            dgUser.Columns("Original early termination fee").Visible = False
            dgUser.Columns("Current device ID - 4G only").Visible = False
            dgUser.Columns("Device ID DEC").Visible = False
            dgUser.Columns("Device ID HEX").Visible = False
            dgUser.Columns("Upgrade eligibility date").Visible = False
            dgUser.Columns("NE2 date").Visible = False
            dgUser.Columns("Early upgrade indicator").Visible = False
            dgUser.Columns("Device manufacturer").Visible = False
            dgUser.Columns("Early upgrade indicator").Visible = False
            dgUser.Columns("Price plan ID").Visible = False
            dgUser.Columns("Voice Enabled Device").Visible = False
            dgUser.Columns("Early upgrade indicator").Visible = False
            dgUser.Columns("Device SIM 4G (Y/N)").Visible = False
            dgUser.Columns("Shipped device ID").Visible = False
            dgUser.Columns("Email address").Visible = False
            dgUser.Columns("User ID").Width = 50
            dgUser.Columns("Cost Center").Width = 50
            dgUser.Columns("Billing cycle date").Width = 75
            dgUser.Columns("Contract End Date").Width = 75
            dgUser.Columns("SIM").Width = 150
            dgUser.Columns("User Name").Width = 150
            dgUser.Columns("Device model").Width = 175

            btnDiscover.Text = "DISCOVER"
            btnDiscover.BackColor = Color.ForestGreen
            btnDiscover.Enabled = False
            i = 1
            Do Until i = 5000
                i += 1
            Loop
            btnDiscover.Enabled = True


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub
    Public Function ConvertCSVToDataSet(ByVal filetable As String) As DataSet

        Dim mydataSet As New DataSet
        Dim LNO As Integer = 0

        'Initialize a dataset
        mydataSet.Tables.Add("Verizon")

        'Pre-define the field headers And set their order
        mydataSet.Tables("Verizon").Columns.Add("Wireless Number").SetOrdinal(0)
        mydataSet.Tables("Verizon").Columns.Add("Account Number").SetOrdinal(1)
        mydataSet.Tables("Verizon").Columns.Add("Billing cycle date").SetOrdinal(2)
        mydataSet.Tables("Verizon").Columns.Add("Contact Name").SetOrdinal(3)
        mydataSet.Tables("Verizon").Columns.Add("Contact Number").SetOrdinal(4)
        mydataSet.Tables("Verizon").Columns.Add("User ID").SetOrdinal(5)
        mydataSet.Tables("Verizon").Columns.Add("User name").SetOrdinal(6)
        mydataSet.Tables("Verizon").Columns.Add("Contract activation Date").SetOrdinal(7)
        mydataSet.Tables("Verizon").Columns.Add("Contract End Date").SetOrdinal(8)
        mydataSet.Tables("Verizon").Columns.Add("Contract length").SetOrdinal(9)
        mydataSet.Tables("Verizon").Columns.Add("Total current charges").SetOrdinal(10)
        mydataSet.Tables("Verizon").Columns.Add("Original early termination fee").SetOrdinal(11)
        mydataSet.Tables("Verizon").Columns.Add("Cost center").SetOrdinal(12)
        mydataSet.Tables("Verizon").Columns.Add("Current device ID - 4G only").SetOrdinal(13)
        mydataSet.Tables("Verizon").Columns.Add("Device SIM 4G (Y/N)").SetOrdinal(14)
        mydataSet.Tables("Verizon").Columns.Add("Device ID DEC").SetOrdinal(15)
        mydataSet.Tables("Verizon").Columns.Add("Device ID HEX").SetOrdinal(16)
        mydataSet.Tables("Verizon").Columns.Add("Upgrade eligibility date").SetOrdinal(17)
        mydataSet.Tables("Verizon").Columns.Add("NE2 date").SetOrdinal(18)
        mydataSet.Tables("Verizon").Columns.Add("Early upgrade indicator").SetOrdinal(19)
        mydataSet.Tables("Verizon").Columns.Add("Device manufacturer").SetOrdinal(20)
        mydataSet.Tables("Verizon").Columns.Add("Device model").SetOrdinal(21)
        mydataSet.Tables("Verizon").Columns.Add("Email address").SetOrdinal(22)
        mydataSet.Tables("Verizon").Columns.Add("Price plan ID").SetOrdinal(23)
        mydataSet.Tables("Verizon").Columns.Add("Price plan description").SetOrdinal(24)
        mydataSet.Tables("Verizon").Columns.Add("SIM").SetOrdinal(25)
        mydataSet.Tables("Verizon").Columns.Add("Shipped device ID").SetOrdinal(26)


        'Parse each row indocument using a comma deliminter
        Try
            Using Parser As New TextFieldParser(filetable)
                Parser.TextFieldType = FileIO.FieldType.Delimited
                Parser.SetDelimiters(",")
                While Not Parser.EndOfData
                    Dim New_row As DataRow = mydataSet.Tables("Verizon").Rows.Add()
                    Dim Ritems() As Object = Parser.ReadFields
                    New_row.ItemArray = Ritems
                End While
            End Using
        Catch exception As Exception
            MessageBox.Show(exception.Message)
        End Try

        'Delete the first 13 rows as they do Not contain any record data
        For i As Int32 = 1 To 14
            mydataSet.Tables(0).Rows.RemoveAt(0)
        Next

        'Insert New field And shift the position of the fields after the insertion
        mydataSet.Tables(0).Columns("Price plan description").SetOrdinal(25)
        mydataSet.Tables("Verizon").Columns.Add("Voice Enabled Device").SetOrdinal(24)
        mydataSet.Tables(0).Columns("SIM").SetOrdinal(26)
        mydataSet.Tables(0).Columns("Shipped device ID").SetOrdinal(27)

        'Switch From a dataset to a dataview And sort by the one field
        Dim objDv As New DataView(mydataSet.Tables(0))
        objDv.Sort = "Price plan description"

        'Remove the Total row that comes to the top after the sort
        Dim Row As DataRowView
        For Each Row In objDv
            If Row("Wireless Number") = "Total" Then
                Row.Delete()
            End If
        Next
        'Accept the above changes to the datview
        objDv.Table.AcceptChanges()

        '=============================================================

        Dim drv As DataRowView
        'Populate the newly added field based on the value in plan description
        For Each drv In objDv
            drv.BeginEdit()
            If Microsoft.VisualBasic.Left(drv.Item("Price plan description"), 3) = "MOB" Then
                drv.Item("Voice Enabled Device") = "N"
            Else
                drv.Item("Voice Enabled Device") = "Y"
            End If
            drv.EndEdit()
        Next
        'Accept the above changes to the datview
        objDv.Table.AcceptChanges()

        'Accept the above changes to the datview
        objDv.Table.AcceptChanges()

        objDv.Sort = "User ID"
        objDv.Table.AcceptChanges()
        '=================================================================================================================


        objDv.Sort = "Wireless Number"
        objDv.Table.AcceptChanges()

        'Eliminate the xA0 encoding in the following 3 fields
        For Each drv In objDv
            drv.BeginEdit()
            drv.Item("User ID") = Regex.Replace(drv.Item("User ID"), "[^A-Za-z0-9\-/]", "")
            drv.EndEdit()
        Next
        For Each drv In objDv
            drv.BeginEdit()
            drv.Item("Cost Center") = Regex.Replace(drv.Item("Cost Center"), "[^A-Za-z0-9\-/]", "")
            drv.EndEdit()
        Next
        For Each drv In objDv
            drv.BeginEdit()
            drv.Item("Current device ID - 4G only") = Regex.Replace(drv.Item("Current device ID - 4G only"), "[^A-Za-z0-9\-/]", "")
            drv.EndEdit()
        Next
        '======================================================================================================
        'Export("S:\Corp\DAT\VZW_HN_Report.csv", mydataSet.Tables("Verizon"))
        Export("S:\Corp\DAT\VZW_HN.csv", mydataSet.Tables("Verizon"))
        '======================================================================================================

        bsUser.DataSource = objDv
        'dgUser.DataSource = bsUser

        Return mydataSet

    End Function

    Public Sub Export(ByVal path As String, ByVal table As DataTable)
        Dim output As New StreamWriter(path, False, UnicodeEncoding.ASCII) '.Default)
        Dim delim As String

        ' Write out the header row
        delim = ""
        For Each col As DataColumn In table.Columns
            output.Write(delim)
            output.Write(col.ColumnName)
            delim = ","
        Next
        output.WriteLine()

        ' write out each data row
        For Each row As DataRow In table.Rows
            delim = ""
            For Each value As Object In row.ItemArray
                output.Write(delim)
                If TypeOf value Is String Then
                    output.Write(""""c) ' thats four double quotes and a c
                    output.Write(value)

                    output.Write(""""c) ' thats four double quotes and a c
                Else
                    output.Write(value)
                End If
                delim = ","
            Next
            output.WriteLine()
        Next

        output.Close()

    End Sub

    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        'ExportTxt()
        ' Copy the form's image into a bitmap.
        mPrintBitMap = New Bitmap(Me.Width, Me.Width)
        Dim lRect As System.Drawing.Rectangle
        lRect.Width = Me.Width
        lRect.Height = Me.Width
        Me.DrawToBitmap(mPrintBitMap, lRect)


        ' Make a PrintDocument and print.
        mPrintDocument = New PrintDocument
        mPrintDocument.DefaultPageSettings.Landscape = True
        mPrintDocument.Print()
        'PrintDocument1.DefaultPageSettings.Landscape = True
        'PrintDocument1.Print()
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim w As Integer = System.Math.Max(dgAM.Width, dgUser.Width)
        Dim h As Integer = dgAM.Height + dgUser.Height
        Dim bmp As Bitmap = New Bitmap(w, h)
        Dim r As Rectangle = New Rectangle(0, 0, dgAM.Width, dgAM.Height)
        dgAM.DrawToBitmap(bmp, r)
        r.Y = dgAM.Height

        r.Width = dgUser.Width
        r.Height = dgUser.Height
        dgUser.DrawToBitmap(bmp, r)
        e.Graphics.DrawImage(bmp, e.MarginBounds)

    End Sub

    Private Sub m_PrintDocument_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles mPrintDocument.PrintPage
        ' Draw the image centered.
        Dim lWidth As Integer = e.MarginBounds.X + (e.MarginBounds.Width - mPrintBitMap.Width) \ 2
        Dim lHeight As Integer = e.MarginBounds.Y + (e.MarginBounds.Height - mPrintBitMap.Height) \ 2
        e.Graphics.DrawImage(mPrintBitMap, lWidth, lHeight)

        ' There's only one page.
        e.HasMorePages = False
    End Sub

    'Private Sub ExportTxt()
    '    Using writer As New System.IO.StreamWriter("C:\Users\y15645\audit.txt")
    '        For row As Integer = 0 To dgAM.RowCount - 2
    '            For col As Integer = 0 To dgAM.ColumnCount - 1
    '                writer.WriteLine(dgAM.Rows(row).Cells(col).Value.ToString)
    '            Next
    '        Next
    '        For row As Integer = 0 To dgUser.RowCount - 2
    '            For col As Integer = 0 To dgUser.ColumnCount - 1
    '                writer.WriteLine(dgUser.Rows(row).Cells(col).Value.ToString)
    '            Next
    '        Next
    '    End Using
    'End Sub

    Private Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        '====================================================
        'EXPORT REMEDY DATAGRID TO EXCEL USER H: DRIVE
        '====================================================
        Dim value As Cursor
        Dim datSQL As New DataSet
        'Dim CLFR As ClassTag = New ClassTag
        '===========================================================================================================
        'verifying the data grid view having data or not
        If ((dgAM.Columns.Count = 0) Or (dgAM.Rows.Count = 0)) Then
            Exit Sub
        End If
        'Creating dataset to export
        Dim dset As New DataSet
        'add table to dataset
        dset.Tables.Add()
        'add column to that table
        For i As Integer = 0 To dgAM.ColumnCount - 1
            dset.Tables(0).Columns.Add(dgAM.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr1 As DataRow
        For i As Integer = 0 To dgAM.RowCount - 1
            dr1 = dset.Tables(0).NewRow
            For j As Integer = 0 To dgAM.Columns.Count - 1
                dr1(j) = dgAM.Rows(i).Cells(j).Value
            Next
            dset.Tables(0).Rows.Add(dr1)
        Next
        '=======================================================================================================================================
        If ((dgUser.Columns.Count = 0) Or (dgUser.Rows.Count = 0)) Then
            Exit Sub
        End If
        'add table to dataset
        dset.Tables.Add("User")
        'add column to that table
        For i As Integer = 0 To dgUser.ColumnCount - 1
            dset.Tables("User").Columns.Add(dgUser.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr2 As DataRow
        For i As Integer = 0 To dgUser.RowCount - 1
            dr2 = dset.Tables("User").NewRow
            For j As Integer = 0 To dgUser.Columns.Count - 1
                dr2(j) = dgUser.Rows(i).Cells(j).Value
            Next
            dset.Tables("User").Rows.Add(dr2)
        Next
        '=======================================================================================================================================
        'Virtual
        '=======================================================================================================================================
        If ((dgVirt.Columns.Count = 0) Or (dgVirt.Rows.Count = 0)) Then
            Exit Sub
        End If
        'add table to dataset
        dset.Tables.Add("Virtual")
        'add column to that table
        For i As Integer = 0 To dgVirt.ColumnCount - 1
            dset.Tables("Virtual").Columns.Add(dgVirt.Columns(i).HeaderText)
        Next
        'add rows to the table
        Dim dr3 As DataRow
        For i As Integer = 0 To dgVirt.RowCount - 1
            dr3 = dset.Tables("Virtual").NewRow
            For j As Integer = 0 To dgVirt.Columns.Count - 1
                dr3(j) = dgVirt.Rows(i).Cells(j).Value
            Next
            dset.Tables("Virtual").Rows.Add(dr3)
        Next
        '=======================================================================================================================================

        Dim excel = CreateObject("Excel.Application")
        Dim wBook = excel.Workbooks.Add
        Dim wSheet = wBook.Worksheets(1)
        Dim dt As System.Data.DataTable = dset.Tables(0)
        Dim dt2 As System.Data.DataTable = dset.Tables("User")
        Dim dt3 As System.Data.DataTable = dset.Tables("AM")
        Dim dt4 As System.Data.DataTable = dset.Tables("Virtual")
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim R As Integer = 2
        Dim C As Integer = 0
        'Populate Worksheet with AM Data
        For Each dc In dt.Columns
            colIndex = colIndex + 1
            excel.Cells(1, colIndex) = dc.ColumnName
        Next
        C = colIndex

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            R = R + 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next
        'Populate worksheet with User Data 
        R = R + 1
        colIndex = 0
        For Each dc In dt2.Columns
            colIndex = colIndex + 1
            excel.Cells(R, colIndex) = dc.ColumnName
        Next
        rowIndex = R - 1
        For Each dr In dt2.Rows
            rowIndex = rowIndex + 1
            R = R + 1
            colIndex = 0
            For Each dc In dt2.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next

        'Populate worksheet with Virtual Data 
        R = R + 1
        colIndex = 0
        For Each dc In dt4.Columns
            colIndex = colIndex + 1
            excel.Cells(R, colIndex) = dc.ColumnName
        Next
        rowIndex = R - 1
        For Each dr In dt4.Rows
            rowIndex = rowIndex + 1
            R = R + 1
            colIndex = 0
            For Each dc In dt4.Columns
                colIndex = colIndex + 1
                excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next

        wSheet.Columns.AutoFit()

        'Save sheet rather then display
        Dim strFileName As String = "H:\Equipment_Phone_" & txtID.Text & ".xls"
        Dim blnFileOpen As Boolean = False
        Try
            Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
            fileTemp.Close()
        Catch ex As Exception
            blnFileOpen = False
        End Try
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If
        wBook.SaveAs(strFileName)
        '=============================================================================================================
        value = Cursors.Arrow
        wBook.Close()

        releaseObject(excel)
        releaseObject(wBook)
        releaseObject(wSheet)
        'excel.Quit()

        MsgBox("Excel file exported to H:\Equipment_Phone_" & txtID.Text & ".xls", MsgBoxStyle.Information, "EXCEL CREATED")

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function KaceReadOne(ByVal filetable As String) As DataSet
        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("S:\Corp\DAT\DAT Device Summary Report.csv")
        Dim sr As IO.StreamReader = New IO.StreamReader("S:\Corp\DAT\DAT Device Summary Report.csv", System.Text.Encoding.Unicode)
        Dim mydataSet As New DataSet

        'Initialize a dataset
        mydataSet.Tables.Add("Virt")
        'End If
        mydataSet.Tables("Virt").Columns.Add("No").SetOrdinal(0)
        mydataSet.Tables("Virt").Columns.Add("SystemName").SetOrdinal(1)
        mydataSet.Tables("Virt").Columns.Add("User Full Name").SetOrdinal(2)
        mydataSet.Tables("Virt").Columns.Add("User Name").SetOrdinal(3)
        mydataSet.Tables("Virt").Columns.Add("Service Tag").SetOrdinal(4)
        mydataSet.Tables("Virt").Columns.Add("System Manufacturer").SetOrdinal(5)
        mydataSet.Tables("Virt").Columns.Add("System Model").SetOrdinal(6)
        mydataSet.Tables("Virt").Columns.Add("Assigned User").SetOrdinal(7)
        mydataSet.Tables("Virt").Columns.Add("Comments").SetOrdinal(8)
        mydataSet.Tables("Virt").Columns.Add("Assigned User ID").SetOrdinal(9)
        mydataSet.Tables("Virt").Columns.Add("USER_NAME").SetOrdinal(10)
        mydataSet.Tables("Virt").Columns.Add("USER_LOGGED").SetOrdinal(11)
        mydataSet.Tables("Virt").Columns.Add("MAC").SetOrdinal(12)
        mydataSet.Tables("Virt").Columns.Add("Assigned User LanID").SetOrdinal(13)
        mydataSet.Tables("Virt").Columns.Add("USER_NAME2").SetOrdinal(14)
        mydataSet.Tables("Virt").Columns.Add("USER_LOGGED2").SetOrdinal(15)
        mydataSet.Tables("Virt").Columns.Add("MAC2").SetOrdinal(16)
        mydataSet.Tables("Virt").Columns.Add("USER_LOGGED3").SetOrdinal(17)
        Try
            Using Parser As New TextFieldParser(filetable)
                Parser.TextFieldType = FileIO.FieldType.Delimited
                Parser.SetDelimiters(",")
                While Not Parser.EndOfData
                    Dim New_row As DataRow = mydataSet.Tables("Virt").Rows.Add()
                    Dim Ritems() As Object = Parser.ReadFields
                    New_row.ItemArray = Ritems
                End While
            End Using
        Catch exception As Exception
            MessageBox.Show(exception.Message)
        End Try
        'Switch From a dataset to a dataview And sort by the one field
        Dim objDv2 As New DataView(mydataSet.Tables(0))
        'Accept the above changes to the datview
        objDv2.Table.AcceptChanges()
        bsVirt.DataSource = objDv2
        Return mydataSet
    End Function

    Public Function KaceReadTwo(ByVal filetable As String) As DataSet
        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("S:\Corp\DAT\DAT Device Summary Report.csv")
        Dim sr As IO.StreamReader = New IO.StreamReader("S:\Corp\DAT\DAT Device Summary Report.csv", System.Text.Encoding.Unicode)
        Dim mydataSet As New DataSet

        'Initialize a dataset
        mydataSet.Tables.Add("Device")
        'End If
        mydataSet.Tables("Device").Columns.Add("No").SetOrdinal(0)
        mydataSet.Tables("Device").Columns.Add("SystemName").SetOrdinal(1)
        mydataSet.Tables("Device").Columns.Add("User Full Name").SetOrdinal(2)
        mydataSet.Tables("Device").Columns.Add("User Name").SetOrdinal(3)
        mydataSet.Tables("Device").Columns.Add("Service Tag").SetOrdinal(4)
        mydataSet.Tables("Device").Columns.Add("System Manufacturer").SetOrdinal(5)
        mydataSet.Tables("Device").Columns.Add("System Model").SetOrdinal(6)
        mydataSet.Tables("Device").Columns.Add("Assigned User").SetOrdinal(7)
        mydataSet.Tables("Device").Columns.Add("Comments").SetOrdinal(8)
        mydataSet.Tables("Device").Columns.Add("Assigned User ID").SetOrdinal(9)
        mydataSet.Tables("Device").Columns.Add("USER_NAME").SetOrdinal(10)
        mydataSet.Tables("Device").Columns.Add("USER_LOGGED").SetOrdinal(11)
        mydataSet.Tables("Device").Columns.Add("MAC").SetOrdinal(12)
        mydataSet.Tables("Device").Columns.Add("Assigned User LanID").SetOrdinal(13)
        mydataSet.Tables("Device").Columns.Add("USER_NAME2").SetOrdinal(14)
        mydataSet.Tables("Device").Columns.Add("USER_LOGGED2").SetOrdinal(15)
        mydataSet.Tables("Device").Columns.Add("MAC2").SetOrdinal(16)
        mydataSet.Tables("Device").Columns.Add("USER_LOGGED3").SetOrdinal(17)
        Try
            Using Parser As New TextFieldParser(filetable)
                Parser.TextFieldType = FileIO.FieldType.Delimited
                Parser.SetDelimiters(",")
                While Not Parser.EndOfData
                    Dim New_row As DataRow = mydataSet.Tables("Device").Rows.Add()
                    Dim Ritems() As Object = Parser.ReadFields
                    New_row.ItemArray = Ritems
                End While
            End Using
        Catch exception As Exception
            MessageBox.Show(exception.Message)
        End Try
        'Switch From a dataset to a dataview And sort by the one field
        Dim objDv As New DataView(mydataSet.Tables(0))
        'Accept the above changes to the datview
        objDv.Table.AcceptChanges()

        bsDevice.DataSource = objDv

        Return mydataSet
    End Function

End Class