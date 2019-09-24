Imports SergeyDll
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.Utils
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid
Imports System.Web.Script.Serialization
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports DevExpress.XtraPrinting


Public Class ReadyComplectation
    Public Structure TPodrInfo
        Public namePodr As String
        Public listGroups As String
        Public listIds As String
        Public Sub New(ByVal _namePodr As String, ByVal _listGroups As String, ByVal _listIds As String)
            namePodr = _namePodr
            listGroups = _listGroups
            listIds = _listIds
        End Sub
    End Structure

    Private query As String
    Private dt As DataTable
    Private arrPodr As Dictionary(Of Integer, TPodrInfo)

    Private Sub ReadyComplectation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gxOrdersView1.Tag = 1
        CType(Me.ParentForm, MainForm).protect.Protect(Me)
        refreshGrid()
        CType(Me.ParentForm, MainForm).mainProgressPanel.Visible = False
        gxOrdersView1.Tag = 0

        Timer.Start()
    End Sub

    Private Sub refreshGrid()
        gxOrdersView1.Tag = 1
        Dim arrOrdersAll As New Dictionary(Of Integer, DataTable)
        Dim arrOrdersReady As New Dictionary(Of Integer, DataTable)
        Dim idsKatMcInPodr As String
        Dim listPord = "-1", listOrders As String = "-1"
        Dim dtStatuses, dtAllDetails, dtZakazPord, dtIsCompl, dtBuf As DataTable
        Dim idPodr, idGroupMC, idZakaz As Integer
        Dim bufInfo As TPodrInfo

        arrPodr = New Dictionary(Of Integer, TPodrInfo)

        readyComplProgressPanel.Visible = True
        Me.Refresh()
        '=========== Получаем участки пользователя ===========================================
        query = "SELECT UsersParam.idValue, KatPodr.NamePodr, KatPodr.listOfKatMcId, BondKatPodrAndGroupMC.idGroupMC " + _
                "FROM UsersParam INNER JOIN " + _
                    "KatPodr ON UsersParam.idValue = KatPodr.id LEFT JOIN " + _
                    "BondKatPodrAndGroupMC ON KatPodr.id = BondKatPodrAndGroupMC.idKatPodr " + _
                "WHERE idUsers = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND isPodr = 1"
        dt = ClassDbWorkBase.FillDataTable(query)
        For i As Integer = 0 To dt.Rows.Count - 1
            idPodr = dt.Rows(i)("idValue")
            idGroupMC = IIf(IsDBNull(dt.Rows(i)("idGroupMC")), -1, dt.Rows(i)("idGroupMC"))
            idsKatMcInPodr = IIf(IsDBNull(dt.Rows(i)("listOfKatMcId")) OrElse String.IsNullOrEmpty(dt.Rows(i)("listOfKatMcId")), "-1", dt.Rows(i)("listOfKatMcId"))

            If (Not arrPodr.ContainsKey(idPodr)) Then
                listPord += IIf(String.IsNullOrEmpty(listPord), "", ",") + idPodr.ToString()
                arrPodr.Add(idPodr, New TPodrInfo(dt.Rows(i)("NamePodr").ToString(), idGroupMC, idsKatMcInPodr))
            Else
                bufInfo = arrPodr(idPodr)
                bufInfo.listGroups += IIf(String.IsNullOrEmpty(bufInfo.listGroups), idGroupMC.ToString(), "," + idGroupMC.ToString())
                arrPodr(idPodr) = bufInfo
            End If
        Next

        '=========== Основная таблица ========================================================
        query = "SELECT idZakaz, CONVERT(DATETIME, ShipmentDate, 104) AS 'Дата отгрузки', idCargo AS 'Погрузочный', NameOrg AS 'Дилер', CASE WHEN Nomer > '800' THEN ('Пробный ' + Nomer) ELSE Nomer END as 'Номер заказа' , " + _
                 "StatusName AS 'Статус заказа', DataZak AS 'Дата заказа', OrderPrim AS 'Примечание', Binary " + _
                "FROM ( " + _
                    "SELECT Zakaz.id AS idZakaz, CONVERT(VARCHAR(10), DataOut, 104) AS ShipmentDate, " + _
                        "Zakaz.idCargo, KatOrg.NameOrg + ' (' + KatOrg.Remark + ')' AS NameOrg, Zakaz.Nomer, Status.StatusName, Zakaz.DataZak, " + _
                        "Zakaz.PrimPar + ' ' + Zakaz.Prim AS OrderPrim, Zakaz.Binary " + _
                    "FROM KatOrg INNER JOIN  Zakaz ON KatOrg.id = Zakaz.idOrg INNER JOIN Status ON Zakaz.Status = Status.id LEFT OUTER JOIN Brak ON Zakaz.id = Brak.idZakaz " + _
                    "WHERE(Zakaz.Status < 90 And Zakaz.Status >= 3) AND (Zakaz.Status < 6)" + _
                 ") as z " + _
                "GROUP BY idZakaz, ShipmentDate, idCargo, NameOrg, Nomer, StatusName, DataZak, OrderPrim, Binary " + _
                "ORDER BY 'Дата отгрузки' DESC, idCargo, NameOrg, Binary"
        dt = ClassDbWorkBase.FillDataTable(query)

        For i As Integer = 0 To dt.Rows.Count - 1
            listOrders += IIf(String.IsNullOrEmpty(listOrders), "", ",") + dt.Rows(i)("idZakaz").ToString()
        Next

        Dim col As DataColumn = dt.Columns.Add("Дата комплектации", Type.GetType("System.String"))
        col.SetOrdinal(1)
        For i As Integer = 0 To dt.Rows.Count - 1
            dt.Rows(i)("Дата комплектации") = Convert.ToDateTime(dt.Rows(i)("Дата отгрузки")).AddDays(-12).ToString().Split(" ")(0)
        Next

        dtZakazPord = New DataTable

        '=========== добавление столбцов с участками =========================================
        For Each podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr
            dt.Columns.Add(podr.Value.namePodr, Type.GetType("System.String"))
            dtZakazPord.Columns.Add(podr.Value.namePodr, Type.GetType("System.String"))
            dtZakazPord.Columns.Add(podr.Value.namePodr + "dateUpdate", Type.GetType("System.String"))
            dtZakazPord.Columns.Add(podr.Value.namePodr + "countCompl", Type.GetType("System.String"))
            dtZakazPord.Columns.Add(podr.Value.namePodr + "isCompl", Type.GetType("System.String"))
        Next

        dtZakazPord.Rows.Add({})
        For i As Integer = 0 To dtZakazPord.Columns.Count - 1
            dtZakazPord.Rows(0)(i) = 0
        Next

        '=========== Получаем количество отданных в работу деталей============================
        query = "SELECT OrderMark.idZakaz, KatPodr.NamePodr, MAx(OrderMark.dateUpdate) AS dateUpdate, COUNT(OrderMark.idPodr) AS countReady, " + _
                        "OrderMark.isCompl, count(OrderMark.isCompl) AS countIsCompl " + _
                "FROM OrderMark INNER JOIN KatPodr ON OrderMark.idPodr = KatPodr.id " + _
                "WHERE OrderMark.idZakaz IN (" + listOrders + ") AND idPodr IN (" + listPord + ") AND " + _
                    "(OrderMark.idStatus = 304 Or (OrderMark.idStatus = 305 And OrderMark.isCompl = 1) or (OrderMark.idStatus = 302 And OrderMark.isCompl = 1)) " + _
                "GROUP BY OrderMark.idZakaz, KatPodr.NamePodr, OrderMark.isCompl"
        dtStatuses = ClassDbWorkBase.FillDataTable(query)

        For i As Integer = 0 To dtStatuses.Rows.Count - 1
            If (Not arrOrdersReady.ContainsKey(dtStatuses.Rows(i)("idZakaz"))) Then
                arrOrdersReady.Add(dtStatuses.Rows(i)("idZakaz"), dtZakazPord.Copy())
            End If

            dtBuf = arrOrdersReady(dtStatuses.Rows(i)("idZakaz"))
            dtBuf.Rows(0)(dtStatuses.Rows(i)("NamePodr").ToString()) += dtStatuses.Rows(i)("countReady")
            dtBuf.Rows(0)(dtStatuses.Rows(i)("NamePodr").ToString() + "dateUpdate") = dtStatuses.Rows(i)("dateUpdate")
            'If (dtStatuses.Rows(i)("isCompl")) Then
            'dtBuf.Rows(0)(dtStatuses.Rows(i)("NamePodr").ToString() + "isCompl") = dtStatuses.Rows(i)("countIsCompl")
            'End If
            arrOrdersReady(dtStatuses.Rows(i)("idZakaz")) = dtBuf.Copy()
        Next


        '=========== Получаем общее количество деталей========================================
        Dim sumCaseCategory = String.Empty, caseCategory As String = String.Empty
        Dim iNumber = 1, bufnUmber As Integer = 1

        For Each podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr
            sumCaseCategory += ", SUM(Category" + iNumber.ToString() + ") as '" + podr.Value.namePodr + "' "
            caseCategory += ", CASE WHEN (KatMC.idGrMC IN ( " + podr.Value.listGroups + " ) OR KatMC.id IN ( " + podr.Value.listIds + " )) THEN COUNT(SpZak.id) ELSE 0 END AS 'Category" + iNumber.ToString() + " ' "
            iNumber += 1
        Next
        query = "SELECT idZakaz " + sumCaseCategory + " " + _
                "FROM (SELECT dbo.SpZak.idZakaz " + caseCategory + _
                    "FROM dbo.KatMC INNER JOIN " + _
                "dbo.SpZak ON dbo.SpZak.idMC = dbo.KatMC.id INNER JOIN " + _
                "Zakaz ON Zakaz.id = SpZak.idZakaz " + _
                    "WHERE  (dbo.SpZak.idZakaz IN (" + listOrders + ")) " + _
                    "GROUP BY SpZak.idZakaz, KatMC.idGrMC, KatMC.id) AS z " + _
                "GROUP BY idZakaz "
        dtAllDetails = ClassDbWorkBase.FillDataTable(query)

        For Each podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr
            For j As Integer = 0 To dtAllDetails.Rows.Count - 1
                If (Not arrOrdersAll.ContainsKey(dtAllDetails.Rows(j)("idZakaz"))) Then
                    arrOrdersAll.Add(dtAllDetails.Rows(j)("idZakaz"), dtZakazPord.Copy())
                End If

                dtBuf = arrOrdersAll(dtAllDetails.Rows(j)("idZakaz"))
                dtBuf.Rows(0)(podr.Value.namePodr) = dtAllDetails.Rows(j)(podr.Value.namePodr)
                arrOrdersAll(dtAllDetails.Rows(j)("idZakaz")) = dtBuf.Copy()
            Next
        Next

        '========== Получаем детали, которые скомплектованы, но на них нет отметки передан/не пердан в работу====================
        query = "SELECT idZakaz " + sumCaseCategory + " " + _
                "FROM (SELECT dbo.SpZak.idZakaz " + caseCategory + _
                    "FROM KatMC INNER JOIN " + _
                        "SpZak ON SpZak.idMC = KatMC.id INNER JOIN " + _
                        "Zakaz ON Zakaz.id = SpZak.idZakaz INNER JOIN " + _
                        "OrderMark ON OrderMark.idSpZak = SpZak.id " + _
                    "WHERE  (SpZak.idZakaz IN (" + listOrders + ") AND OrderMark.idStatus = " + ClassCommon.STATUS_COMPL.ToString() + ") " + _
                    "GROUP BY SpZak.idZakaz, KatMC.idGrMC, KatMC.id) AS z " + _
                "GROUP BY idZakaz"
        dtIsCompl = ClassDbWorkBase.FillDataTable(query)

        For Each podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr
            For j As Integer = 0 To dtIsCompl.Rows.Count - 1
                If (Not arrOrdersReady.ContainsKey(dtIsCompl.Rows(j)("idZakaz"))) Then
                    arrOrdersReady.Add(dtIsCompl.Rows(j)("idZakaz"), dtZakazPord.Copy())
                End If
                dtBuf = arrOrdersReady(dtIsCompl.Rows(j)("idZakaz"))
                dtBuf.Rows(0)(podr.Value.namePodr + "isCompl") = dtIsCompl.Rows(j)(podr.Value.namePodr)
                arrOrdersReady(dtIsCompl.Rows(j)("idZakaz")) = dtBuf.Copy()
            Next
        Next

        '=========== Получаем заказы со статусом ПРОБЛЕМА=====================================
        query = "SELECT OrderMark.idZakaz, KatPodr.NamePodr, COUNT(OrderMark.idPodr) AS cnt " + _
                "FROM OrderMark INNER JOIN KatPodr ON OrderMark.idPodr = KatPodr.id " + _
                "WHERE OrderMark.idZakaz IN (" + listOrders + ") AND idPodr IN (" + listPord + ") AND (OrderMark.idStatus = 307) " + _
                "GROUP BY OrderMark.idZakaz, KatPodr.NamePodr"
        dtStatuses = ClassDbWorkBase.FillDataTable(query)
        For i As Integer = 0 To dtStatuses.Rows.Count - 1
            If (Not arrOrdersReady.ContainsKey(dtStatuses.Rows(i)("idZakaz"))) Then
                arrOrdersReady.Add(dtStatuses.Rows(i)("idZakaz"), dtZakazPord.Copy())
            End If

            dtBuf = arrOrdersReady(dtStatuses.Rows(i)("idZakaz"))
            dtBuf.Rows(0)(dtStatuses.Rows(i)("NamePodr").ToString()) = -1
            arrOrdersReady(dtStatuses.Rows(i)("idZakaz")) = dtBuf.Copy()
        Next

        '=========== записываем данные в участки
        For i As Integer = 0 To dt.Rows.Count - 1
            idZakaz = dt.Rows(i)("idZakaz")

            If (idZakaz = 2143739972) Then
                idZakaz = 2143739972
            End If

            For Each podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr
                If (arrOrdersAll.ContainsKey(idZakaz)) Then
                    If (arrOrdersAll(idZakaz).Rows(0)(podr.Value.namePodr) <> 0) Then

                        If (arrOrdersReady.ContainsKey(idZakaz) AndAlso arrOrdersAll(idZakaz).Rows(0)(podr.Value.namePodr) = arrOrdersReady(idZakaz).Rows(0)(podr.Value.namePodr + "isCompl")) Then
                            dt.Rows(i)(podr.Value.namePodr) = "КОМПЛ"
                        ElseIf (Not arrOrdersReady.ContainsKey(idZakaz) OrElse arrOrdersReady(idZakaz).Rows(0)(podr.Value.namePodr) = 0) Then
                            dt.Rows(i)(podr.Value.namePodr) = "НЗ"
                        ElseIf (arrOrdersReady(idZakaz).Rows(0)(podr.Value.namePodr) = -1) Then
                            dt.Rows(i)(podr.Value.namePodr) = "ПРОБЛЕМ"
                        Else
                            If (arrOrdersAll(idZakaz).Rows(0)(podr.Value.namePodr) = arrOrdersReady(idZakaz).Rows(0)(podr.Value.namePodr)) Then
                                dt.Rows(i)(podr.Value.namePodr) = "ЗО (" + calculateWorkDays(arrOrdersReady(idZakaz).Rows(0)(podr.Value.namePodr + "dateUpdate")) + ")"
                            Else
                                dt.Rows(i)(podr.Value.namePodr) = "ЗОЧ"
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        Next
        dt.Columns.Add("-//-")
        xggcOrders.DataSource = Nothing
        xggcOrders.DataSource = dt.Copy()
        gxOrdersView1.RefreshData()
        gxOrdersView1.PopulateColumns()

        Me.Text = "Готовность комплектации | Количество заказов: " + dt.Rows.Count.ToString()

        ' Разрешаем слияние ячеек для определенных колонок и скрываем не нужные столбцы
        For Each column As DevExpress.XtraGrid.Columns.GridColumn In gxOrdersView1.Columns
            'If (Array.IndexOf({"-//-"}, column.FieldName) < 0) Then
            'column.OptionsColumn.FixedWidth = True
            'End If
            If (Array.IndexOf({"Дата отгрузки", "Погрузочный", "Дилер", "Дата комплектации"}, column.FieldName) < 0) Then
                column.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False
            End If
            If (Array.IndexOf({"idZakaz", "StatusBrak", "Status", "Binary"}, column.FieldName) >= 0) Then
                column.Visible = False
            End If
        Next


        gxOrdersView1.Columns("Дата отгрузки").Width = 90
        gxOrdersView1.Columns("Погрузочный").Width = 80
        gxOrdersView1.Columns("Номер заказа").Width = 90
        gxOrdersView1.Columns("Дата заказа").Width = 80

        'запись столбов по умолчанию
        query = "SELECT * FROM WidthColumnsForUser WHERE idUser = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND form_name = '" + Me.Name + "'"
        dt = ClassDbWorkBase.FillDataTable(query)
        If dt.Rows.Count = 0 Then
            For Each column As DevExpress.XtraGrid.Columns.GridColumn In gxOrdersView1.Columns
                query = "INSERT INTO WidthColumnsForUser (idUser, form_name, column_name, width_column, number_column) VALUES (" + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + ", '" + Me.Name + "', '" + column.FieldName + "', " + column.Width.ToString() + ", " + column.VisibleIndex.ToString() + "  )"
                ClassDbWorkBase.ExecuteNonQuery(query)
            Next
        End If

        Me.WindowState = FormWindowState.Maximized

        'ширина столбцов для конкретного пользователя
        For i As Integer = 0 To dt.Rows.Count - 1
            For Each column As DevExpress.XtraGrid.Columns.GridColumn In gxOrdersView1.Columns
                If column.FieldName = dt.Rows(i)("column_name") Then
                    column.Width = dt.Rows(i)("width_column")
                    column.VisibleIndex = dt.Rows(i)("number_column")
                    'column.OptionsColumn.FixedWidth = False
                End If
            Next
        Next

        gxOrdersView1.Tag = 0
        readyComplProgressPanel.Visible = False
    End Sub

    Private Function calculateWorkDays(ByVal fromDate As Date) As String
        Dim days As Integer = DateDiff(DateInterval.Day, fromDate, Date.Now)
        Dim dayOfWeek As Date = fromDate
        Dim i As Integer = 0

        While dayOfWeek <= Date.Now And i < 365
            If Weekday(dayOfWeek) = vbSaturday OrElse Weekday(dayOfWeek) = vbSunday Then
                days -= 1
            End If

            dayOfWeek = dayOfWeek.AddDays(1)
            i += 1
        End While

        days = IIf(days = 0, 0, days + 1)
        Return days.ToString()
    End Function

    Private Sub listBoxLoadComponents()
        query = "SELECT id, NamePodr FROM KatPodr WHERE Cex <> 0 AND Cex <> 999 AND NamePodr not like '%Сборщик%' AND Cex <= 5 ORDER BY NomRes"
        lbAreaAll.DataSource = ClassDbWorkBase.FillDataTable(query)
        lbAreaAll.DisplayMember = "NamePodr"
        lbAreaAll.ValueMember = "id"

        query = "SELECT UsersParam.id, UsersParam.idValue, KatPodr.NamePodr " + _
                "FROM UsersParam LEFT JOIN KatPodr ON UsersParam.idValue = KatPodr.id " + _
                "WHERE UsersParam.idUsers = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND UsersParam.isPodr = 1 ORDER BY NomRes"
        lbAreaAddedUser.DataSource = ClassDbWorkBase.FillDataTable(query)
        lbAreaAddedUser.DisplayMember = "NamePodr"
        lbAreaAddedUser.ValueMember = "idValue"

    End Sub

    Private Sub btPanelClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPanelClose.Click
        pnSetShowingArea.Visible = False
        refreshGrid()
    End Sub

    Private Sub lbAreaAll_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbAreaAll.DoubleClick
        tsmiAdd_Click(Nothing, Nothing)
    End Sub

    Private Sub lbAreaAddedUser_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbAreaAddedUser.DoubleClick
        tsmiDelete_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmiSetShowingArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSetShowingArea.Click
        pnSetShowingArea.Visible = True
        listBoxLoadComponents()

    End Sub

    Private Sub xggcOrders_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles xggcOrders.MouseDoubleClick
        If e.Button <> Windows.Forms.MouseButtons.Left Then Return

        Dim item = (From podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr Where podr.Value.namePodr = gxOrdersView1.FocusedColumn.ToString()).ToList()

        If (item.Count <> 0) Then
            Dim ProductionAreas As ProductionAreas
            ProductionAreas = New ProductionAreas(gxOrdersView1.GetRowCellValue(gxOrdersView1.FocusedRowHandle, "idZakaz").ToString(), item(0).Key, CType(Me.ParentForm, MainForm).protect.idUser)
            ProductionAreas.MdiParent = Me.ParentForm
            ProductionAreas.Show()
        End If

        If gxOrdersView1.FocusedColumn.FieldName = "Номер заказа" Then
            openDocumentForOrder(gxOrdersView1.GetRowCellValue(gxOrdersView1.FocusedRowHandle, "Номер заказа").ToString())
            pnShowDocumentForOrder.Visible = True
            PdfViewerDocumentForOrder.CloseDocument()
        End If
    End Sub

    Private Sub tsmiRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiRefresh.Click
        refreshGrid()
    End Sub

    Private Sub gxOrdersView1_RowCellStyle(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs) Handles gxOrdersView1.RowCellStyle
        If String.IsNullOrEmpty(e.CellValue.ToString()) Then
        'e.Appearance.BackColor = Color.LightGray
        Return
        End If

        Dim item = (From podr As KeyValuePair(Of Integer, TPodrInfo) In arrPodr Where podr.Value.namePodr = e.Column.FieldName).ToList()

        If (item.Count > 0) Then
            If (Array.IndexOf({"НЗ", "ЗОЧ"}, e.CellValue.ToString()) >= 0) Then
                e.Appearance.BackColor = Color.LightPink
            ElseIf e.CellValue.ToString().IndexOf("ЗО (") <> -1 Then
                Dim open, close As Integer
                open = e.CellValue.ToString().IndexOf("(") + 1
                close = e.CellValue.ToString().IndexOf(")")
                If e.CellValue.ToString().Substring(open, close - open) > 5 Then
                    e.Appearance.BackColor = Color.OrangeRed
                Else : e.Appearance.BackColor = Color.LightGreen
                End If

            ElseIf e.CellValue.ToString() = "ПРОБЛЕМ" Then
                e.Appearance.BackColor = Color.Crimson
            ElseIf e.CellValue.ToString() = "КОМПЛ" Then
                e.Appearance.BackColor = Color.LimeGreen
            End If
        End If

    End Sub

    Private Sub openDocumentForOrder(ByVal numberOrder As String)
        Dim pach As String = Setting.Xml.GetXmlValue("DestinationTo")
        Dim listOfFiles As String()

        lbDocumentsForOrder.Items.Clear()

        listOfFiles = Directory.GetFiles(pach, "*" + numberOrder.ToString().Substring(0, 6) + "*.*").Where(Function(s) Not s.ToLower().Contains("price")).ToArray()
        For Each fileName As String In listOfFiles
            lbDocumentsForOrder.Items.Add(Path.GetFileName(fileName))
        Next
    End Sub

    Private Sub lbDocumentsForOrder_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lbDocumentsForOrder.MouseClick
        If (lbDocumentsForOrder.SelectedItem Is Nothing) Then Return

        PdfViewerDocumentForOrder.CloseDocument()
        If lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".pdf") <> -1 Then
            PdfViewerDocumentForOrder.LoadDocument(Setting.Xml.GetXmlValue("DestinationTo") + lbDocumentsForOrder.SelectedItem.ToString())
        ElseIf lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".xls") <> -1 Then
            ClassCommon.openExcelReadOnly(Setting.Xml.GetXmlValue("DestinationTo") + lbDocumentsForOrder.SelectedItem.ToString())
        End If
    End Sub

    Private Sub btPrintDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPrintDocument.Click
        If lbDocumentsForOrder.SelectedIndex = -1 Then Return
        If lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".pdf") <> -1 Then
            PdfViewerDocumentForOrder.Print()
        End If
    End Sub

    Private Sub btClosePDFViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosePDFViewer.Click
        pnShowDocumentForOrder.Visible = False
    End Sub

    Private Sub tsmiMigration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiMigration.Click
        Dim listOrders As String = String.Empty
        Dim count As Integer
        Dim table As DataTable

        query = "SELECT idZakaz, ShipmentDate AS 'Дата отгрузки', idCargo AS 'Погрузочный', NameOrg AS 'Дилер', CASE WHEN Nomer > '800' THEN ('Пробный ' + Nomer) ELSE Nomer END as 'Номер заказа' , " + _
                 "StatusName AS 'Статус заказа', DataZak AS 'Дата заказа', OrderPrim AS 'Примечание', Binary " + _
                "FROM ( " + _
                    "SELECT Zakaz.id AS idZakaz, DateName(Year,DataOut)+'-'+LTRIM(STR(Month(DataOut)))+'-'+Right('0'+DateName(Day,DataOut),2) AS ShipmentDate, " + _
                        "Zakaz.idCargo, KatOrg.NameOrg + ' (' + KatOrg.Remark + ')' AS NameOrg, Zakaz.Nomer, Status.StatusName, Zakaz.DataZak, " + _
                        "Zakaz.PrimPar + ' ' + Zakaz.Prim AS OrderPrim, Zakaz.Binary " + _
                    "FROM KatOrg INNER JOIN  Zakaz ON KatOrg.id = Zakaz.idOrg INNER JOIN Status ON Zakaz.Status = Status.id LEFT OUTER JOIN Brak ON Zakaz.id = Brak.idZakaz " + _
                    "WHERE(Zakaz.Status < 90 And Zakaz.Status >= 3) AND (Zakaz.Status < 6)" + _
                 ") as z " + _
                "GROUP BY idZakaz, ShipmentDate, idCargo, NameOrg, Nomer, StatusName, DataZak, OrderPrim, Binary " + _
                "ORDER BY ShipmentDate, idCargo, NameOrg, Binary"
        table = ClassDbWorkBase.FillDataTable(query)

        For i As Integer = 0 To table.Rows.Count - 1
            listOrders += IIf(String.IsNullOrEmpty(listOrders), "", ",") + table.Rows(i)("idZakaz").ToString()
        Next

        query = "SELECT SpZak.idZakaz, SteepSpZak.idSpZak,  SteepSpZak.Steep3,  SteepSpZak.dataS3,  SteepSpZak.kol3 " + _
                "FROM SpZak INNER JOIN " + _
                 "SteepSpZak ON SteepSpZak.idSpZak = SpZak.id " + _
                "WHERE SpZak.idZakaz IN (" + listOrders + ") AND SteepSpZak.Steep3 IS NOT NULL AND SteepSpZak.kol3 IS NOT NULL AND SteepSpZak.dataS3 IS NOT NULL"
        table = ClassDbWorkBase.FillDataTable(query)
        For i As Integer = 0 To table.Rows.Count - 1
            If (Convert.ToInt32(table.Rows(i)("kol3")) <= 0) Then Continue For

            query = "SELECT count(idSpZak) " + _
                    "FROM OrderMark " + _
                    "WHERE idZakaz = " + table.Rows(i)("idZakaz").ToString() + " AND idSpZak = " + table.Rows(i)("idSpZak").ToString() + " AND idPodr = " + ClassCommon.PODR_COMPL.ToString()
            count = Convert.ToInt32(ClassDbWorkBase.ExecuteScalar(query))

            If count <> 0 Then
                query = "UPDATE OrderMark SET idStatus = " + ClassCommon.STATUS_COMPL.ToString() + ", isCompl = 1, dateUpdate = Convert(DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "') " + _
                        "WHERE idZakaz = " + table.Rows(i)("idZakaz").ToString() + " AND idSpZak = " + table.Rows(i)("idSpZak").ToString() + " AND idPodr = " + ClassCommon.PODR_COMPL.ToString()
            Else
                query = "INSERT INTO OrderMark (idZakaz, idSpZak, idPodr, idStatus, isCompl, dateCreate, dateUpdate) " + _
                        "VALUES(" + table.Rows(i)("idZakaz").ToString() + "," + table.Rows(i)("idSpZak").ToString() + "," + ClassCommon.PODR_COMPL.ToString() + "," + ClassCommon.STATUS_COMPL.ToString() + ", 1" + _
                            ", Convert(DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "'), Convert (DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "') )"
            End If
            ClassDbWorkBase.ExecuteNonQuery(query)

            query = "UPDATE OrderMark SET isCompl = 1 WHERE idSpZak = " + table.Rows(i)("idSpZak").ToString()
            ClassDbWorkBase.ExecuteNonQuery(query)
        Next
        MessageBox.Show("Готово", "Миграция комплектации", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub gxOrdersView1_ColumnPositionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gxOrdersView1.ColumnPositionChanged
        If gxOrdersView1.Tag <> 0 Then Return
        query = "SELECT * FROM WidthColumnsForUser WHERE idUser = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND form_name = '" + Me.Name + "'"
        dt = ClassDbWorkBase.FillDataTable(query)

        If dt.Rows.Count <> 0 Then
            For Each column As DevExpress.XtraGrid.Columns.GridColumn In gxOrdersView1.Columns
                For i As Integer = 0 To dt.Rows.Count - 1
                    If column.FieldName = dt.Rows(i)("column_name") Then
                        query = "UPDATE WidthColumnsForUser SET number_column = " + column.VisibleIndex.ToString() + " WHERE idUser = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND form_name = '" + Me.Name + _
                                "' AND column_name = '" + column.FieldName + "'"
                        ClassDbWorkBase.ExecuteNonQuery(query)
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub gxOrdersView1_ColumnWidthChanged(ByVal sender As System.Object, ByVal e As DevExpress.XtraGrid.Views.Base.ColumnEventArgs) Handles gxOrdersView1.ColumnWidthChanged
        query = "SELECT * FROM WidthColumnsForUser WHERE idUser = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND form_name = '" + Me.Name + "'"
        dt = ClassDbWorkBase.FillDataTable(query)
        If dt.Rows.Count <> 0 Then
            For Each column As DevExpress.XtraGrid.Columns.GridColumn In gxOrdersView1.Columns
                For i As Integer = 0 To dt.Rows.Count - 1
                    If column.FieldName = dt.Rows(i)("column_name") Then
                        query = "UPDATE WidthColumnsForUser SET width_column = " + column.Width.ToString() + _
                                " WHERE idUser = " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND form_name = '" + Me.Name + _
                                    "' AND column_name = '" + column.FieldName + "'"
                        ClassDbWorkBase.ExecuteNonQuery(query)
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer.Tick
        refreshGrid()
    End Sub

    Private Sub tsmiPrintTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiPrintTable.Click
        Dim result = MessageBox.Show("Раскрашивать ячейки Excel, аналогично таблице?", "Раскрашивание ячеек Excel", MessageBoxButtons.YesNoCancel)
        Dim options As XlsxExportOptionsEx = New XlsxExportOptionsEx()

        If (result = DialogResult.Yes) Then
            options.ExportType = DevExpress.Export.ExportType.WYSIWYG

        ElseIf (result = DialogResult.No) Then
            options.ExportType = DevExpress.Export.ExportType.DataAware
        ElseIf (result = DialogResult.Cancel) Then
            Return
        End If

        Using saveDialog As SaveFileDialog = New SaveFileDialog()
            saveDialog.Filter = "Excel (2010) (.xlsx)|*.xlsx"
            saveDialog.FileName = "Готовность комплектации " + Now.Day.ToString("00") + "_" + Now.Month.ToString("00") + "_" + Now.Year.ToString() + " " + _
                Now.Hour.ToString() + "ч" + Now.Minute.ToString() + "м" + Now.Second.ToString() + "c"

            If saveDialog.ShowDialog() <> DialogResult.Cancel Then
                Dim exportFilePath As String = saveDialog.FileName
                Dim fileExtenstion As String = New FileInfo(exportFilePath).Extension

                Select Case fileExtenstion
                    Case ".xlsx"
                        xggcOrders.ExportToXlsx(exportFilePath, options)
                    Case Else
                End Select

                If File.Exists(exportFilePath) Then

                    Try
                        System.Diagnostics.Process.Start(exportFilePath)
                    Catch
                        Dim msg As String = "Файл не может быть открыт." & Environment.NewLine + Environment.NewLine & "Path: " & exportFilePath
                        MessageBox.Show(msg, "Error!", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    End Try
                Else
                    Dim msg As String = "Файл не может быть сохранен." & Environment.NewLine + Environment.NewLine & "Path: " & exportFilePath
                    MessageBox.Show(msg, "Error!", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End If
            End If
        End Using
    End Sub

    Private Sub tsmiAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiAdd.Click
        If (lbAreaAll.SelectedIndex < 0) Then Return
        For i As Integer = 0 To lbAreaAddedUser.Items.Count - 1
            If (CType(lbAreaAddedUser.DataSource, DataTable).Rows(i)("idValue") = lbAreaAll.SelectedValue) Then
                Return
            End If
        Next
        query = "INSERT INTO UsersParam (idUsers, idValue, isPodr) VALUES (" + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + ", " + lbAreaAll.SelectedValue.ToString() + ", 1)"
        ClassDbWorkBase.ExecuteNonQuery(query)
        listBoxLoadComponents()
    End Sub

    Private Sub tsmiDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiDelete.Click
        If (lbAreaAddedUser.SelectedIndex < 0) Then Return
        query = "DELETE FROM UsersParam WHERE idUsers =" + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + " AND idValue = " + lbAreaAddedUser.SelectedValue.ToString() + " AND isPodr = 1"
        ClassDbWorkBase.ExecuteNonQuery(query)
        listBoxLoadComponents()
    End Sub

End Class
