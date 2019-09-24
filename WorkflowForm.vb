Imports SergeyDll
Imports System.IO
Imports System.Windows.Forms

Public Class WorkflowForm

    Private Const HEIGHT_GB As Integer = 28

    Private query As String
    Private filterStr As String
    Private excel As ExcelDocument


    Private Sub WorkflowForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CType(Me.ParentForm, MainForm).protect.Protect(Me)
        Dim dt As DataTable

        query = "SELECT * FROM pCompany Order by CompanyName"
        dt = ClassDbYavid.FillDataTable(query)
        cbClient.DataSource = dt.Copy()
        cbClient.DisplayMember = "CompanyName"
        cbClient.SelectedIndex = -1
        cbFilterClient.Tag = 0
        cbFilterClient.DataSource = dt.Copy()
        cbFilterClient.DisplayMember = "CompanyName"
        cbFilterClient.SelectedIndex = -1
        cbFilterClient.Tag = 1


        query = "SELECT KatMC.id as 'id', KatMC.NameMC as 'NameMC' FROM SpCompl INNER JOIN KatMC ON SpCompl.idSecond = KatMC.id WHERE (SpCompl.idMain = 2141978920) ORDER BY KatMC.Indx"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbType.DataSource = dt.Copy()
        cbType.DisplayMember = "NameMC"
        cbType.SelectedIndex = -1
        cbEditType.Tag = 0
        cbEditType.DataSource = dt.Copy()
        cbEditType.DisplayMember = "NameMC"
        cbEditType.SelectedIndex = -1
        cbEditType.Tag = 1
        cbFilterType.Tag = 0
        cbFilterType.DataSource = dt.Copy()
        cbFilterType.DisplayMember = "NameMC"
        cbFilterType.SelectedIndex = -1
        cbFilterType.Tag = 1

        cbStatus.Tag = 0
        query = "SELECT id, (CONVERT(varchar(10), id) + ' - ' + status) as 'status' FROM orderStatus ORDER BY status"
        cbStatus.DataSource = ClassDbEcadmaster.FillDataTable(query)
        cbStatus.DisplayMember = "status"
        cbStatus.Tag = 1

        dtpFilterDateTo.Tag = 0
        dtpFilterDateTo.Value = Now
        dtpFilterDateTo.Tag = 1
        dtpFilterDateFrom.Tag = 0
        dtpFilterDateFrom.Value = New Date(Now.Year - 1, Now.Month, Now.Day)
        dtpFilterDateFrom.Tag = 1

        Me.Tag = 0
        filterStr = " AND orderWorkflow.created > CONVERT(DATETIME,'" + dtpFilterDateFrom.Value.ToString().Split(" ")(0) + " 00:00:00', 104) AND " + _
                "orderWorkflow.created < CONVERT(DATETIME,'" + dtpFilterDateTo.Value.ToString().Split(" ")(0) + " 23:59:59', 104) "
        refreshGridWorkflow()
        Me.Tag = 1
        cbSettingVisibleOrders.Tag = 1
    End Sub

    Private Sub paintGridWorkflow()
        Dim periodNSD As Integer = Convert.ToInt32(tbEditPeriodTypeNSD.Text)
        Dim periodR As Integer = Convert.ToInt32(tbEditPeriodTypeR.Text)
        Dim lastDate As Date

        For i As Integer = 0 To dgvWorkflow.Rows.Count - 1
            lastDate = dgvWorkflow.Rows(i).Cells("Принят").Value

            If (Now.Date > lastDate.AddDays(+periodNSD)) And (dgvWorkflow.Rows(i).Cells("orderStatusID").Value.ToString() <> 10) Then
                dgvWorkflow.Rows(i).Cells("Номер заказа").Style.BackColor = Color.LightCoral
            End If

            If (Now.Date > lastDate.AddDays(+periodR)) And (dgvWorkflow.Rows(i).Cells("Тип").Value.ToString() = "R") Then
                dgvWorkflow.Rows(i).Cells("Номер заказа").Style.BackColor = Color.IndianRed
            End If
        Next
    End Sub

    Private Sub refreshGridWorkflow()
        If (cbSettingVisibleOrders.Checked) Then
            filterStr += " AND (orderStatus.id <> 10) "
        End If

        query = "SELECT orderWorkflow.id, orderWorkflow.numOrder as 'Номер заказа', orderWorkflow.numOrderClient as 'Номер диллера', " + _
                     "Yavid.dbo.pCompany.CompanyName as 'Диллер',orderWorkflow.numOrder3Cad as 'Номер 3CAD', KatMC.NameMC as 'Тип', orderWorkflow.draw1 as 'Сл.1', " + _
                     "orderWorkflow.draw2 as 'Сл.2', orderWorkflow.draw3 as 'Сл.3', orderStatus.id AS 'orderStatusID', " + _
                     "orderStatus.status as 'Статус', orderWorkflow.created as 'Принят',orderWorkflow.dateConverted as 'Конвертирован', Users.FIO AS 'Оформитель', " + _
                     "Users_1.FIO AS 'Конструктор', MAX([LOG].datetime)  AS 'Передан в производство' " + _
                "FROM Users AS Users_1 RIGHT OUTER JOIN " + _
                     "orderWorkflow INNER JOIN orderStatus ON orderWorkflow.status = orderStatus.id INNER JOIN " + _
                     "KatMC ON orderWorkflow.type = KatMC.id INNER JOIN " + _
                     "Yavid.dbo.pCompany ON orderWorkflow.client = Yavid.dbo.pCompany.CompanyID ON Users_1.idUsers = orderWorkflow.constructor LEFT OUTER JOIN " + _
                     "Users ON orderWorkflow.registrator = Users.idUsers LEFT OUTER JOIN " + _
                     "[LOG] ON orderWorkflow.id = [LOG].pr_key_id " + _
                "WHERE (([LOG].table_name = 'orderWorkflow' and [LOG].col_name = 'status' and [LOG].new_value = 10) or [LOG].new_value is Null) " + filterStr + _
                "GROUP BY orderWorkflow.id, orderWorkflow.numOrder, orderWorkflow.numOrderClient, orderWorkflow.numOrder3Cad, " + _
                            "Yavid.dbo.pCompany.CompanyName, KatMC.NameMC, orderWorkflow.draw1, " + _
                            "orderWorkflow.draw2, orderWorkflow.draw3, orderStatus.id, " + _
                            "orderStatus.status, orderWorkflow.created, orderWorkflow.dateConverted, Users.FIO, " + _
                            "Users_1.FIO " + _
                "ORDER BY orderWorkflow.numOrder"

        dgvWorkflow.DataSource = ClassDbEcadmaster.FillDataTable(query)
        dgvWorkflow.Columns("id").Visible = False
        dgvWorkflow.Columns("orderStatusID").Visible = False
        dgvWorkflow.Columns("Передан в производство").Visible = False
        dgvWorkflow.Columns("Номер заказа").Width = 55
        dgvWorkflow.Columns("Тип").Width = 30
        dgvWorkflow.Columns("Сл.1").Width = 30
        dgvWorkflow.Columns("Сл.2").Width = 30
        dgvWorkflow.Columns("Сл.3").Width = 30
        dgvWorkflow.Columns("Статус").Width = 150

        If Me.Tag <> 0 Then
            paintGridWorkflow()
        End If

        Me.Text = "Журнал заказов (" + dgvWorkflow.RowCount.ToString() + ")"

    End Sub

    Private Sub btCreateOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCreateOrder.Click
        Dim len As Integer = mtbOrderNumber.Text.Replace(" ", "").Length
        If len < 6 Then
            MessageBox.Show("Укажите номер заказа", "Workflow", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim month As Integer = Convert.ToInt32(mtbOrderNumber.Text.Substring(4, 2))
        If month > 12 Or month < 1 Then
            MessageBox.Show("Не правильно введен номер заказа", "Конвертор заказов", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrEmpty(tbClientNumber.Text) Then
            MessageBox.Show("Укажите номер диллера", "Workflow", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If cbClient.SelectedIndex = -1 Then
            MessageBox.Show("Выберите диллера", "Workflow", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If cbType.SelectedIndex = -1 Then
            MessageBox.Show("Выберите тип заказа", "Workflow", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        query = "INSERT INTO orderWorkflow(numOrder, numOrderClient, client, type, draw1, draw2, draw3, status, created) " + _
            "values('" + _
                mtbOrderNumber.Text + "','" + _
                tbClientNumber.Text + "', " + _
                DirectCast(cbClient.Items(cbClient.SelectedIndex), DataRowView)("CompanyID").ToString() + ", " + _
                DirectCast(cbType.Items(cbType.SelectedIndex), DataRowView)("id").ToString() + ", " + _
                IIf(String.IsNullOrEmpty(mtbDraw1.Text.Trim()), "0", mtbDraw1.Text.Trim()) + ", " + _
                IIf(String.IsNullOrEmpty(mtbDraw2.Text.Trim()), "0", mtbDraw2.Text.Trim()) + ", " + _
                IIf(String.IsNullOrEmpty(mtbDraw3.Text.Trim()), "0", mtbDraw3.Text.Trim()) + ", " + _
                "1, " + _
                "CONVERT(DATETIME, '" + DateTime.Now() + "', 104)" + _
            ")"

        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            mtbOrderNumber.Text = String.Empty
            tbClientNumber.Text = String.Empty
            mtbDraw1.Text = String.Empty
            mtbDraw2.Text = String.Empty
            mtbDraw3.Text = String.Empty
            cbClient.SelectedIndex = -1
            cbType.SelectedIndex = -1
        End If
    End Sub

    Private Sub pbChangeStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbChangeStatus.Click
        cbStatus.DroppedDown = True
    End Sub

    Private Sub pbChangeStatus_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbChangeStatus.MouseLeave, pbChangeDraw.MouseLeave
        CType(sender, PictureBox).BorderStyle = BorderStyle.None
    End Sub

    Private Sub pbChangeStatus_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbChangeStatus.MouseMove, pbChangeDraw.MouseMove
        CType(sender, PictureBox).BorderStyle = BorderStyle.Fixed3D
    End Sub

    Private Sub tpWorkflow_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tpWorkflow.MouseMove, panelChangeWorkflow.MouseMove
        pbChangeStatus.BorderStyle = BorderStyle.None
        pbChangeDraw.BorderStyle = BorderStyle.None
    End Sub

    Private Sub dgvWorkflow_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkflow.CellDoubleClick
        If dgvWorkflow.SelectedRows.Count = 0 Or btEditRight.Tag = 0 Then Return

        mtbChangeDraw1.Text = IIf(dgvWorkflow.SelectedRows(0).Cells("Сл.1").Value <> 0, dgvWorkflow.SelectedRows(0).Cells("Сл.1").Value, "")
        mtbChangeDraw2.Text = IIf(dgvWorkflow.SelectedRows(0).Cells("Сл.2").Value <> 0, dgvWorkflow.SelectedRows(0).Cells("Сл.2").Value, "")
        mtbChangeDraw3.Text = IIf(dgvWorkflow.SelectedRows(0).Cells("Сл.3").Value <> 0, dgvWorkflow.SelectedRows(0).Cells("Сл.3").Value, "")
        cbEditType.Text = dgvWorkflow.SelectedRows(0).Cells("Тип").Value.ToString()

        panelChangeWorkflow.Enabled = True
    End Sub

    Private Sub dgvWorkflow_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvWorkflow.SelectionChanged
        panelChangeWorkflow.Enabled = False
    End Sub

    Private Sub btChangeRegistrator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btChangeRegistrator.Click
        Dim user As String = CType(Me.ParentForm, MainForm).protect.idUser.ToString()

        query = "UPDATE orderWorkflow SET registrator=" + user + _
            "WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()
        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            dgvWorkflow.SelectedRows(0).Cells("Оформитель").Value = ClassDbEcadmaster.ExecuteScalar("SELECT FIO FROM Users WHERE idUsers=" + user)
        End If
    End Sub

    Private Sub btChangeConstructor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btChangeConstructor.Click
        Dim user As String = CType(Me.ParentForm, MainForm).protect.idUser.ToString()

        query = "UPDATE orderWorkflow SET constructor=" + user + _
            "WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()
        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            dgvWorkflow.SelectedRows(0).Cells("Конструктор").Value = ClassDbEcadmaster.ExecuteScalar("SELECT FIO FROM Users WHERE idUsers=" + user)
        End If
    End Sub

    Private Sub pbChangeDraw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbChangeDraw.Click
        Dim draw1, draw2, draw3 As String

        draw1 = mtbChangeDraw1.Text.Trim().Trim("_")
        draw2 = mtbChangeDraw2.Text.Trim().Trim("_")
        draw3 = mtbChangeDraw3.Text.Trim().Trim("_")
        draw1 = IIf(String.IsNullOrEmpty(draw1), "0", draw1)
        draw2 = IIf(String.IsNullOrEmpty(draw2), "0", draw2)
        draw3 = IIf(String.IsNullOrEmpty(draw3), "0", draw3)

        query = "UPDATE orderWorkflow SET " + _
                    "draw1=" + draw1 + ", " + _
                    "draw2=" + draw2 + ", " + _
                    "draw3=" + draw3 + " " + _
                "WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()
        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            dgvWorkflow.SelectedRows(0).Cells("Сл.1").Value = draw1
            dgvWorkflow.SelectedRows(0).Cells("Сл.2").Value = draw2
            dgvWorkflow.SelectedRows(0).Cells("Сл.3").Value = draw3
        End If
    End Sub

    Private Sub btEditRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEditRight.Click
        If btEditRight.ImageKey = "off.png" Then
            btEditRight.ImageKey = "on.png"
            btEditRight.Tag = 1
        Else
            btEditRight.ImageKey = "off.png"
            btEditRight.Tag = 0
        End If
    End Sub

    Private Sub tsmiHistoryShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btHistoryShow.Click
        If dgvWorkflow.SelectedRows.Count = 0 Then Return

        Dim logForm As New LogForm("orderWorkflow", Convert.ToInt32(dgvWorkflow.SelectedRows(0).Cells("id").Value))
        logForm.MdiParent = Me.MdiParent
        logForm.Show()
    End Sub

    Private Sub cbStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStatus.SelectedIndexChanged
        If dgvWorkflow.SelectedRows.Count = 0 Or cbStatus.Tag = 0 Then Return

        query = "UPDATE orderWorkflow SET " + _
                    "status=" + DirectCast(cbStatus.Items(cbStatus.SelectedIndex), DataRowView)("id").ToString() + " " + _
                "WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()

        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            dgvWorkflow.SelectedRows(0).Cells("Статус").Value = DirectCast(cbStatus.Items(cbStatus.SelectedIndex), DataRowView)("status").ToString().Split("-")(1).Trim()
        End If
    End Sub

    Private Sub btDelWorkflow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btDelWorkflow.Click
        If dgvWorkflow.SelectedRows.Count = 0 Then Return

        If MessageBox.Show("Удалить заказ?", "Workflow", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = DialogResult.Yes Then
            query = "DELETE FROM orderWorkflow WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()
            If (ClassDbEcadmaster.ExecuteNonQuery(query)) Then
                query = "DELETE FROM LOG WHERE table_name='orderWorkflow' and pr_key_id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()
                ClassDbEcadmaster.ExecuteNonQuery(query)
                dgvWorkflow.Rows.Remove(dgvWorkflow.SelectedRows(0))
            End If
        End If
    End Sub

    Private Sub cbEditType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbEditType.SelectedIndexChanged
        If cbEditType.Tag = 0 Then Return

        query = "UPDATE orderWorkflow SET " + _
                    "type=" + DirectCast(cbEditType.Items(cbEditType.SelectedIndex), DataRowView)("id").ToString() + " " + _
                    "WHERE id=" + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString()

        If ClassDbEcadmaster.ExecuteNonQuery(query) Then
            dgvWorkflow.SelectedRows(0).Cells("Тип").Value = DirectCast(cbEditType.Items(cbEditType.SelectedIndex), DataRowView)("NameMC").ToString()
        End If
    End Sub

    Private Sub btShowHideEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btShowHideEdit.Click
        If (gbEdit.Height = HEIGHT_GB) Then
            gbEdit.Height = 300
            btShowHideEdit.ImageKey = "arrow-up-grey.png"
        Else
            gbEdit.Height = HEIGHT_GB
            btShowHideEdit.ImageKey = "arrow-down-grey.png"
        End If
    End Sub

    Private Sub brShowHideFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btShowHideFilter.Click
        Dim mahHeight As Integer = 175

        If (gbFilter.Height = HEIGHT_GB) Then
            gbFilter.Height = mahHeight
            gbEdit.Location = New Point(gbEdit.Location.X, gbEdit.Location.Y + (mahHeight - HEIGHT_GB))
            btShowHideFilter.ImageKey = "arrow-up-grey.png"
        Else
            gbFilter.Height = HEIGHT_GB
            gbEdit.Location = New Point(gbEdit.Location.X, gbEdit.Location.Y - (mahHeight - HEIGHT_GB))
            btShowHideFilter.ImageKey = "arrow-down-grey.png"
        End If
    End Sub

    Private Sub cbFilterType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFilterType.SelectedIndexChanged
        If (cbFilterType.Tag = 0) Then Return

        createFilter()
    End Sub

    Private Sub cbFilterClient_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFilterClient.SelectedIndexChanged
        If (cbFilterClient.Tag = 0) Then Return

        createFilter()
    End Sub

    Private Sub createFilter()
        Dim hasFilter As Boolean = False

        filterStr = String.Empty

        If (cbFilterClient.SelectedIndex <> -1) Then
            filterStr = "AND orderWorkflow.client=" + DirectCast(cbFilterClient.Items(cbFilterClient.SelectedIndex), DataRowView)("CompanyID").ToString() + " "
            hasFilter = True
        End If
        If (cbFilterType.SelectedIndex <> -1) Then
            filterStr += " AND orderWorkflow.type=" + DirectCast(cbFilterType.Items(cbFilterType.SelectedIndex), DataRowView)("id").ToString() + " "
        End If
        If (Not String.IsNullOrEmpty(mtbFilterOrderNumber.Text.Trim().Trim("_"))) Then
            Dim arr As String()
            Dim number As String

            arr = mtbFilterOrderNumber.Text.Split("_")
            arr(0) = arr(0).Replace(" ", "_")
            arr(1) = arr(1).Replace(" ", "_")
            arr(1) = IIf(arr(1).Length = 0, "__", arr(1))
            arr(1) = IIf(arr(1).Length = 1, arr(1) + "_", arr(1))
            number = IIf(String.IsNullOrEmpty(arr(1)), arr(0) + "_", arr(0) + "_" + arr(1))

            filterStr += " AND orderWorkflow.numOrder LIKE '" + number + "' "
        End If

        If (Not String.IsNullOrEmpty(tbFilterDealerNumber.Text)) Then
            filterStr += " AND orderWorkflow.numOrderClient " + IIf(cbFilterDealerNumberMask.Checked, "like '%" + tbFilterDealerNumber.Text + "%' ", "= '" + tbFilterDealerNumber.Text + "' ")
        End If

        filterStr += " AND orderWorkflow.created > CONVERT(DATETIME,'" + dtpFilterDateFrom.Value.ToString().Split(" ")(0) + " 00:00:00', 104) AND " + _
                "orderWorkflow.created < CONVERT(DATETIME,'" + dtpFilterDateTo.Value.ToString().Split(" ")(0) + " 23:59:59', 104) "

        refreshGridWorkflow()
    End Sub

    Private Sub btClearFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClearFilter.Click
        mtbFilterOrderNumber.Text = String.Empty
        cbFilterType.SelectedIndex = -1
        cbFilterClient.SelectedIndex = -1
        tbFilterDealerNumber.Text = String.Empty
        btnCleareFilterDate_Click(Nothing, Nothing)

        createFilter()
    End Sub

    Private Sub btClearFilterOrderNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClearFilterOrderNumber.Click
        mtbFilterOrderNumber.Text = String.Empty

        createFilter()
    End Sub

    Private Sub btClearFilterType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClearFilterType.Click
        cbFilterType.SelectedIndex = -1
    End Sub

    Private Sub btClearFilterClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClearFilterClient.Click
        cbFilterClient.SelectedIndex = -1
    End Sub

    Private Sub mtbFilterOrderNumber_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mtbFilterOrderNumber.KeyUp

        If (e.KeyValue >= 96 And e.KeyValue <= 105) Or (e.KeyValue >= 48 And e.KeyValue <= 57) Or e.KeyValue = 8 Or e.KeyValue = 46 Or e.KeyValue = 13 Or e.KeyValue = 32 Then
            createFilter()
        End If
    End Sub

    Private Sub dgvWorkflow_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWorkflow.CellClick
        If (dgvWorkflow.SelectedRows.Count = 0) Then Return

        Dim value As String

        query = "SELECT MIN(datetime) FROM LOG " + _
            "WHERE pr_key_id = " + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString() + " " + _
            " AND table_name = 'orderWorkflow' AND col_name = 'status' AND new_value = 4;"
        value = ClassDbEcadmaster.ExecuteScalar(query)
        lbDateConstructors.Text = value

        query = "SELECT MIN(datetime) FROM LOG " + _
            "WHERE pr_key_id = " + dgvWorkflow.SelectedRows(0).Cells("id").Value.ToString() + " " + _
            " AND table_name = 'orderWorkflow' AND col_name = 'status' AND new_value = 7;"
        value = ClassDbEcadmaster.ExecuteScalar(query)
        lbDateOformlinie.Text = value
    End Sub

    Private Sub tbFilterDealerNumber_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbFilterDealerNumber.KeyUp
        createFilter()
    End Sub

    Private Sub btnClearFilterDealerNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearFilterDealerNumber.Click
        tbFilterDealerNumber.Text = String.Empty
        createFilter()
    End Sub


    Private Sub dtpFilterDateFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFilterDateFrom.ValueChanged, dtpFilterDateTo.ValueChanged
        If CType(sender, DateTimePicker).Tag = 0 Then Return

        createFilter()
    End Sub

    Private Sub btnCleareFilterDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCleareFilterDate.Click
        dtpFilterDateTo.Tag = 0
        dtpFilterDateTo.Value = Now
        dtpFilterDateTo.Tag = 1
        dtpFilterDateFrom.Value = New Date(Now.Year - 1, Now.Month, Now.Day)
    End Sub

    Private Sub btnOutputExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOutputExcel.Click
        Dim paintExcel As Boolean = False
        Dim dialogResult As DialogResult
        Dim templatesReportsDir As String = Setting.Xml.GetXmlValue("TemplatesReports")

        dialogResult = MessageBox.Show("Раскрасить ячейки Excel, аналогично таблице?" + Environment.NewLine + "Раскрашивание может занять некоторое время.", "Раскраска заказа", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information)

        If (dialogResult = dialogResult.Yes) Then
            paintExcel = True
        ElseIf (dialogResult = dialogResult.Cancel) Then
            Return
        End If

        If (dgvWorkflow.ColumnCount = 0) Or (dgvWorkflow.Rows.Count = 0) Then Return

        If (Not excel Is Nothing) Then
            excel.Close()
        End If

        If (Not File.Exists(templatesReportsDir + "workflowReport.xlsx")) Then
            MessageBox.Show("Не найден файл шаблона workflowReport.xlsx!", "Workflow", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            Return
        End If

        excel = New ExcelDocument(templatesReportsDir + "workflowReport.xlsx")
        excel.SelectSheet("Лист1")
        ClassCommon.exportFromDgvToExcel(dgvWorkflow, excel, "A5", 5, "|id|Сл.1|Сл.2|Сл.3|orderStatusID|Оформитель|Конструктор|Передан в производство|", paintExcel)
        excel.SetRangeCellValue(dtpFilterDateFrom, "F2", "F2")
        excel.SetRangeCellValue(dtpFilterDateTo, "F3", "F3")
        excel.SetRangeCellValue(dgvWorkflow.RowCount.ToString(), "F4", "F4")

        '===========Скрытый DGV========= Таблица "Дилер - Кол-во заказов" =========== 
        query = "SELECT pCompany.CompanyName As 'Название', count(*) as 'Количество заказов' " + _
                      "FROM [Ecadmaster].[dbo].[orderWorkflow] JOIN yavid.dbo.pCompany ON pCompany.CompanyID =  orderWorkflow.client " + _
                      "WHERE orderWorkflow.created > CONVERT(DATETIME,'" + dtpFilterDateFrom.Value.ToString().Split(" ")(0) + " 00:00:00', 104) AND orderWorkflow.created < CONVERT(DATETIME,'" + dtpFilterDateTo.Value.ToString().Split(" ")(0) + " 23:59:59', 104) " + _
                      "GROUP BY orderWorkflow.client,pCompany.CompanyName "

        dgvNotVisible.DataSource = ClassDbEcadmaster.FillDataTable(query)

        excel.SelectSheet("Лист2")
        ClassCommon.exportFromDgvToExcel(dgvNotVisible, excel, "A5", 5)
        excel.SetRangeCellValue(dtpFilterDateFrom, "F2", "F2")
        excel.SetRangeCellValue(dtpFilterDateTo, "F3", "F3")

        refreshGridWorkflow()

        excel.Visible = True
    End Sub

    Private Sub WorkflowForm_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown, dgvWorkflow.Sorted
        paintGridWorkflow()
    End Sub

    Private Sub btRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRefresh.Click
        refreshGridWorkflow()
    End Sub

    Private Sub WorkflowForm_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If (Not excel Is Nothing) Then
            excel.Close()
        End If
    End Sub

    Private Sub cbFilterDealerNumberMask_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFilterDealerNumberMask.Click, btEditPeriodType.Click
        createFilter()
    End Sub

    Private Sub cbSettingVisibleOrders_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSettingVisibleOrders.CheckedChanged
        If cbSettingVisibleOrders.Tag = 0 Then Return
        createFilter()
    End Sub

    Private Sub tbEditPeriodTypeNSD_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbEditPeriodTypeNSD.KeyPress, tbEditPeriodTypeR.KeyPress
        If (e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = ChrW(Keys.Back) Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbEditPeriodTypeNSD_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbEditPeriodTypeNSD.KeyUp, tbEditPeriodTypeR.KeyUp
        If (e.KeyCode = Keys.Enter) Then
            createFilter()
        End If
    End Sub
End Class