Public Class Notes
    Dim DataCol As New DataTable()
    Dim DataR As DataRow
    Dim DataC As DataColumn
    Dim repeatYearUrl(255) As String
    Dim repeatMonthUrl(255) As String
    Dim rtitem As Object
    Dim noSession As Object
    Dim noDatabase As Object
    Dim noview As Object
    Dim noDocument As Object
    Function Settings(ByVal notes As System.String， ByVal serverName As System.String, ByVal dbName As System.String, ByVal viewName As System.String) As Integer
        'noSession = CreateObject("notes.NotesSession")
        'noDatabase = noSession.GETDATABASE("D19DBR15/19/A/IBM", "m_dir\maxdevz\ymaxBn.nsf")
        'noview = noDatabase.getView("SlinkII/Dalian-BP")
        'noDocument = noview.GetLastDocument
        noSession = CreateObject(notes)
        noDatabase = noSession.GETDATABASE(serverName, dbName)
        noview = noDatabase.getView(viewName)
        noDocument = noview.GetLastDocument
        Settings = 0
    End Function


    Dim DataYear As New DataTable()
    Dim DataYearR As DataRow
    Dim DataYearC As DataColumn
    Dim arr(2) As Object

    Public Function SetDataTable() As Integer
        DataC = New DataColumn("contracN", Type.GetType("System.String"))
        DataCol.Columns.Add(DataC)
        DataC = Nothing

        DataYearC = New DataColumn("contracN", Type.GetType("System.String"))
        DataYear.Columns.Add(DataYearC)
        DataYearC = Nothing

        DataC = New DataColumn("urls", Type.GetType("System.String"))
        DataCol.Columns.Add(DataC)
        DataC = Nothing

        DataYearC = New DataColumn("urls", Type.GetType("System.String"))
        DataYear.Columns.Add(DataYearC)
        DataYearC = Nothing

        DataC = New DataColumn("status", Type.GetType("System.String"))
        DataCol.Columns.Add(DataC)
        DataC = Nothing

        DataYearC = New DataColumn("status", Type.GetType("System.String"))
        DataYear.Columns.Add(DataYearC)
        DataYearC = Nothing

        SetDataTable = 0
    End Function

    Public Function GetData() As Object
        'Dim e As Integer
        Dim Count As Integer = 0
        Dim flagYear As Boolean
        Dim i As Integer
        Dim flagMonth As Boolean

        While noDocument IsNot Nothing
            If noDocument.GetItemValue("Status")(0) = "確認待ち" Then
                If noDocument.GetItemValue("TSLINKStatus")(0) = "契約締結済み" Then
                    If InStr(noDocument.GetItemValue("Cnt_cn")(0), "Tool") > 0 Then
                        If InStr(noDocument.GetItemValue("CntrA")(0), "年額") > 0 Then
                            flagYear = True
                            For i = 0 To UBound(repeatYearUrl) - 1
                                If repeatYearUrl(i) = "" Then
                                    Exit For
                                End If
                                If repeatYearUrl(i) = noDocument.GetItemValue("TSlinkNo")(0) Then
                                    flagYear = False
                                End If
                            Next
                            If flagYear Then

                                For i = 0 To UBound(repeatYearUrl) - 1
                                    If repeatYearUrl(i) = "" Then
                                        repeatYearUrl(i) = noDocument.GetItemValue("TSlinkNo")(0)
                                        Exit For
                                    End If
                                Next
                                DataR = DataCol.NewRow()
                                DataR(0) = noDocument.GetItemValue("TContractNumber")(0)
                                DataR(1) = noDocument.GetItemValue("TContractURL")(0)
                                DataR(2) = False
                                DataCol.Rows.Add(DataR)
                                DataR = Nothing
                                'Count = Count+1
                                'If Count>0
                                '	b = (repeatYearUrl(0)=noDocument.GetItemValue("TSlinkNo")(0))
                                '	Exit While
                                'End If
                            End If
                        End If

                        If InStr(noDocument.GetItemValue("CntrA")(0), "月額") > 0 Then
                            flagMonth = True
                            For i = 0 To UBound(repeatMonthUrl) - 1
                                If repeatMonthUrl(i) = "" Then
                                    Exit For
                                End If
                                If repeatMonthUrl(i) = noDocument.GetItemValue("TSlinkNo")(0) Then
                                    flagMonth = False
                                End If
                            Next
                            If flagMonth Then

                                For i = 0 To UBound(repeatMonthUrl) - 1
                                    If repeatMonthUrl(i) = "" Then
                                        repeatMonthUrl(i) = noDocument.GetItemValue("TSlinkNo")(0)
                                        Exit For
                                    End If
                                Next
                                DataYearR = DataYear.NewRow()
                                DataYearR(0) = noDocument.GetItemValue("TContractNumber")(0)
                                DataYearR(1) = noDocument.GetItemValue("TContractURL")(0)
                                DataYearR(2) = False
                                DataYear.Rows.Add(DataYearR)
                                DataYearR = Nothing
                            End If
                        End If
                    End If
                End If
            End If
            noDocument = noview.GetPrevDocument(noDocument)
        End While
        arr(0) = DataCol
        arr(1) = DataYear
        GetData = arr
    End Function
    Public Function MarkStatus(ByVal collection As Object, ByVal viewName As System.String, ByVal PDFPath As String) As Integer
        Dim ColRowsNum As Integer
        Dim NumDoc As Integer
        'Dim bufferCol As Object
        Dim buffer As Object
        ColRowsNum = collection.Rows.Count

        For i = 0 To ColRowsNum - 1
            noview = noDatabase.getView(viewName)
            NumDoc = noview.FTSearch(collection.Rows(i)("contracN").ToString(), 0)
            If NumDoc > 0 Then
                buffer = noview.GetFirstDocument()
                While buffer IsNot Nothing
                    If collection.Rows(i)("status").ToString() = "True" Then
                        'buffer.ReplaceItemValue("Cnt_cn", buffer.GetItemValue("Cnt_cn")(0) & "完成")
                        buffer.ReplaceItemValue("Cnt_cn", Replace(buffer.GetItemValue("Cnt_cn")(0), "Tool", "Tool☆"))
                        rtitem = buffer.GetFirstItem("Cnt_mt")
                        buffer.ReplaceItemValue("Status", "確認中")
                        buffer.ReplaceItemValue("StatusCode", "G")
                        buffer.ReplaceItemValue("StatusN", "4")
                        Call rtitem.EmbedObject(1454, "", PDFPath & "\" & collection.Rows(i)("contracN").ToString() & ".pdf")
                    Else
                        'buffer.ReplaceItemValue("Cnt_cn", buffer.GetItemValue("Cnt_cn")(0) & "未完成")
                        buffer.ReplaceItemValue("Cnt_cn", Replace(buffer.GetItemValue("Cnt_cn")(0), "Tool", "Tool【未処理】"))
                    End If
                    Call buffer.Save(True, True)
                    buffer = noview.GetNextDocument(buffer)
                End While
            End If
            Call noview.Clear
        Next
        MarkStatus = noview.FTSearch(collection.Rows(0)("contracN").ToString(), 0)
    End Function

    Public Function Clear() As Integer
        repeatYearUrl = Nothing
        noSession = Nothing
        noDatabase = Nothing
        noDocument = Nothing
        Clear = 0
    End Function
End Class

