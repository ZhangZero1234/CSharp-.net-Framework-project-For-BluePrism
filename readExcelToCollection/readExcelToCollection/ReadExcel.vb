Public Class ReadExcel
    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim raXL As Microsoft.Office.Interop.Excel.Range

    Function OpenFile(ByVal filePath As String) As Integer
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlBook = xlApp.Workbooks.Open(filePath)
        OpenFile = 0
    End Function
    'xlApp = New Microsoft.Office.Interop.Excel.Application
    'xlBook = xlApp.Workbooks.Open(filePath)

    Function getContract(ByVal filePath As String) As Object
        'Dim xlApp As Microsoft.Office.Interop.Excel.Application
        'Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim raXL As Microsoft.Office.Interop.Excel.Range

        Dim dt As New System.Data.DataTable()
        Dim dr As DataRow
        Dim dc As DataColumn
        dc = New DataColumn("Area", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Service Center", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Resp svc ofc", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Division", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contract status", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Customer No", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Customer sign date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Acceptance date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contract start date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contract end date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Competency", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Channel ind", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Legal Contract NO", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Opportunity ID", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contract Amount", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Currency id", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Title", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Description", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contact name", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contact Phone No", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Contact User ID", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Print invoice summary only", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("checkBox", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        'DataR = DataCol.NewRow()
        'DataR(0) = noDocument.GetItemValue("TContractNumber")(0)

        'xlApp = New Microsoft.Office.Interop.Excel.Application
        'xlBook = xlApp.Workbooks.Open(filePath)
        xlSheet = xlBook.Worksheets(1)
        Dim ranges = xlSheet.Range("A1", "Q61")
        'If xlSheet.Range("B29").Text="詳細内容" Then
        'test = xlSheet.Range("F44").Value
        'xlSheet.Range("C44").Offset(0,3).Value = "hahahaha"
        'test = xlSheet.CheckBoxes("Check Box 15").Value
        'End If
        dr = dt.NewRow()
        For Each raXL In ranges
            Select Case raXL.Text
                Case "Area"
                    'a = raXL.Offset(0,3).Value
                    dr(0) = raXL.Offset(0, 3).Value
                Case "Service Center"
                    'b = raXL.Offset(0,2).Value
                    dr(1) = raXL.Offset(0, 2).Value
                Case "Resp svc ofc"
                    'c = raXL.Offset(0,3).Value
                    dr(2) = raXL.Offset(0, 3).Value
                Case "Division"
                    'd = raXL.Offset(0,2).Value
                    dr(3) = raXL.Offset(0, 2).Value
                Case "Contract status"
                    'e = raXL.Offset(0,3).Value
                    dr(4) = raXL.Offset(0, 3).Value
                Case "Customer No."
                    'f = raXL.Offset(0,2).Value
                    dr(5) = raXL.Offset(0, 2).Value
                Case "Customer sign date"
                    'g = raXL.Offset(0,3).Value
                    dr(6) = raXL.Offset(0, 3).Text
                Case "Acceptance date"
                    'h = raXL.Offset(0,2).Value
                    dr(7) = raXL.Offset(0, 2).Text
                Case "Contract start date"
                    'i = raXL.Offset(0,3).Value
                    dr(8) = raXL.Offset(0, 3).Text
                Case "Contract end date"
                    'j = raXL.Offset(0,2).Value
                    dr(9) = raXL.Offset(0, 2).Text
                Case "Competency"
                    'k = raXL.Offset(0,3).Value
                    dr(10) = raXL.Offset(0, 3).Value
                Case "Channel ind"
                    'l = raXL.Offset(0,2).Value
                    dr(11) = raXL.Offset(0, 2).Value
                Case "Legal Contract NO."
                    'm = raXL.Offset(0,2).Value
                    dr(12) = raXL.Offset(0, 2).Value
                Case "Opportunity ID"
                    'n = raXL.Offset(0,3).Value
                    dr(13) = raXL.Offset(0, 3).Value
                Case "Contract Amount"
                    'o = raXL.Offset(0,2).Value
                    dr(14) = raXL.Offset(0, 2).Value
                Case "Currency id"
                    'p = raXL.Offset(0,3).Value
                    dr(15) = raXL.Offset(0, 3).Value
                Case "Title　"
                    'q = raXL.Offset(0,3).Value
                    dr(16) = raXL.Offset(0, 3).Value
                Case "Description"
                    'r = raXL.Offset(0,3).Value
                    dr(17) = raXL.Offset(0, 3).Value
                Case "Contact name"
                    's = raXL.Offset(0,3).Value
                    dr(18) = raXL.Offset(0, 3).Value
                Case "Contact Phone No"
                    't = raXL.Offset(0,2).Value
                    dr(19) = raXL.Offset(0, 2).Value
                Case "Contact User ID"
                    'u = raXL.Offset(0,3).Value
                    dr(20) = raXL.Offset(0, 3).Value
                Case "Print invoice summary only"
                    'v = raXL.Offset(1,7).Value
                    dr(21) = raXL.Offset(1, 7).Value
                    If xlSheet.CheckBoxes("Check Box 15").Value > 0 Then
                        dr(22) = "Yes"
                    Else
                        dr(22) = "No"
                    End If
                    'dr(22) = xlSheet.CheckBoxes("Check Box 15").Value
            End Select
        Next

        dt.Rows.Add(dr)

        'xlBook = Nothing
        xlSheet = Nothing
        raXL = Nothing
        'xlApp.Quit()
        'xlApp = Nothing
        getContract = dt
    End Function

    Function getSoftLayer(ByVal filePath As String) As Object
        'Dim xlApp As Microsoft.Office.Interop.Excel.Application
        'Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim raXL As Microsoft.Office.Interop.Excel.Range

        Dim dt As New System.Data.DataTable()
        Dim dr As DataRow
        Dim dc As DataColumn

        dc = New DataColumn("Work Start date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Work End date", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Service Type", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Labor Amount", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Customer number", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Status", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Business type", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Owning RSO", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Owning Div", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Opportunity ID", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Offering ID", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Billing Frequency", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Invoice to customer", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("ADU", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("BGC", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Title", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Description", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Customer project", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Finance Indicator", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Channel ind", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Print invoice summary only", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("checkBox", Type.GetType("System.String"))
        dt.Columns.Add(dc)
        'DataR = DataCol.NewRow()
        'DataR(0) = noDocument.GetItemValue("TContractNumber")(0)

        'xlApp = New Microsoft.Office.Interop.Excel.Application
        'xlBook = xlApp.Workbooks.Open(filePath)
        xlSheet = xlBook.Worksheets(2)
        Dim ranges = xlSheet.Range("A1", "M76")
        dr = dt.NewRow()
        For Each raXL In ranges
            Select Case raXL.Text
                Case "Work Start date"
                    dr(0) = raXL.Offset(0, 3).Text
                Case "Work End date"
                    dr(1) = raXL.Offset(0, 2).Text
                Case "Service Type"
                    dr(2) = raXL.Offset(0, 3).Value
                Case "Labor Amount"
                    dr(3) = raXL.Offset(0, 2).Value
                Case "Customer number"
                    dr(4) = raXL.Offset(0, 3).Value
                Case "Status"
                    dr(5) = raXL.Offset(0, 3).Value
                Case "Business type"
                    dr(6) = raXL.Offset(0, 2).Value
                Case "Owning RSO"
                    dr(7) = raXL.Offset(0, 3).Value
                Case "Owning Div"
                    dr(8) = raXL.Offset(0, 2).Value
                Case "Opportunity ID"
                    dr(9) = raXL.Offset(0, 3).Value
                Case "Offering ID"
                    dr(10) = raXL.Offset(0, 3).Value
                Case "Billing Frequency"
                    dr(11) = raXL.Offset(0, 2).Value
                Case "Invoice to customer"
                    dr(12) = raXL.Offset(0, 3).Value
                Case "ADU"
                    dr(13) = raXL.Offset(0, 3).Value
                Case "BGC"
                    dr(14) = raXL.Offset(0, 2).Value
                Case "Title"
                    dr(15) = raXL.Offset(0, 3).Value
                Case "Description"
                    dr(16) = raXL.Offset(0, 3).Value
                Case "Customer project"
                    dr(17) = raXL.Offset(0, 3).Value
                Case "Finance Indicator"
                    dr(18) = raXL.Offset(0, 3).Value
                Case "Channel ind"
                    dr(19) = raXL.Offset(0, 3).Value
                Case "Print invoice summary only"
                    dr(20) = raXL.Offset(1, 7).Value
                    If xlSheet.CheckBoxes("Check Box 1").Value > 0 Then
                        dr(21) = "Yes"
                    Else
                        dr(21) = "No"
                    End If
            End Select
        Next

        dt.Rows.Add(dr)
        xlBook = Nothing
        xlSheet = Nothing
        raXL = Nothing
        'xlApp.Quit()
        xlApp = Nothing
        getSoftLayer = dt
    End Function
End Class




