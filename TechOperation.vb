Imports SergeyDll

Public Class TechOperation

    Private query As String
    Private dt As DataTable
    Private arrNumberCode As New Dictionary(Of Integer, String)
    Private codTP As String = String.Empty

    Public Sub New(ByVal numberTehProc As String)
        codTP = numberTehProc
    End Sub


    'Этот метод загружает все необходимые компоненты формы. 
    Private Sub TechOperation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        arrNumberCode.Add(851, "Первичная")
        arrNumberCode.Add(852, "Вторичная")
        arrNumberCode.Add(853, "Отделка")
        arrNumberCode.Add(854, "Сборка")
        arrNumberCode.Add(855, "Контроль")

        cbNumberCode.DataSource = New BindingSource(arrNumberCode, Nothing)
        cbNumberCode.ValueMember = "Key"
        cbNumberCode.DisplayMember = "Value"

        LoadTehOperatin()
        loadTepOperationAll()
        loadChronometrage()

    End Sub

    ''' <summary>
    ''' Загрузка всех тех операций (ТО)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub loadTepOperationAll()
        query = "SELECT TehOperation.id, (TehOperation.CodIdLeft + TehOperation.CodIdRight) as 'Номер', " + _
                    "TehOperation.Name as 'Название', " + _
                    "TehOperation.DateCreate as 'Создан', " + _
                    "TehOperation.DateUpdate as 'Обновлен', " + _
                    "TypeTransaction.Name AS 'Тип', " + _
                    "KatPodr.NamePodr as 'Участок', " + _
                    "Equipment.Name as 'Оборудование', " + _
                    "TehOperation.DateEnd as 'Завершен', " + _
                    "TehOperation.Remark as 'Примечание' " + _
                "FROM TehOperation INNER JOIN  [Work_Base].[dbo].[Status] ON TehOperation.Status = Status.id " + _
                     "INNER JOIN [Work_Base].[dbo].[KatPodr] ON TehOperation.Area = KatPodr.id " + _
                     "INNER JOIN TypeTransaction ON TehOperation.TypeTehOper = TypeTransaction.id " + _
                     "INNER JOIN Equipment ON TehOperation.Equipment = Equipment.id"

        gcTehOperationAll.DataSource = ClassDbEcadmaster.FillDataTable(query)
    End Sub

    ''' <summary>
    '''  Загрузка всех ТО, входящих ТП CardTehProcess
    ''' </summary>
    Private Sub LoadTehOperatin()
        If gcTehOperationView.SelectedRowsCount = 0 Then
            Return
        End If

        query = "SELECT TehOperation.id, TehOperation.CodIdLeft + TehOperation.CodIdRight AS 'Код', TehOperation.Name AS 'Наименование', " +
                      "TehOperation.DateCreate AS 'Создан', TehOperation.DateUpdate AS 'Обновлен', TehOperation.DateEnd AS 'Завершен', " +
                      "TehOperation.Remark AS 'Примечание', Equipment.Name as 'Оборудование', Status.StatusName as 'Статус', KatPodr.NamePodr as 'Участок' " +
                "FROM GroupOperation INNER JOIN " +
                      "TehOperation ON GroupOperation.idTehOper = TehOperation.id INNER JOIN " +
                      "Equipment ON TehOperation.Equipment = Equipment.id INNER JOIN " +
                      "[Work_Base].[dbo].[Status] ON TehOperation.Status = Status.id INNER JOIN " +
                      "[Work_Base].[dbo].[KatPodr] ON TehOperation.Area = KatPodr.id " +
                "WHERE GroupOperation.idTehProc = " '+ gcTehProcessView.GetRowCellValue(gcTehProcessView.FocusedRowHandle, "id").ToString()
        gcTehOperation.DataSource = ClassDbEcadmaster.FillDataTable(query)
    End Sub

    'Метод изменения видимости панели, в которой происходит добавление ТО
    Private Sub btClosedPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosedPanel.Click
        pnAddTehOperation.Visible = False
    End Sub

    'Метод Добавления ТО
    Private Sub btAddTehOperation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAddTehOperation.Click
        Dim result As String = String.Empty

        If tbNameTO.Text = String.Empty Then
            MessageBox.Show("Присутствуют пустые поля!", "Ошибка!", MessageBoxButtons.OK)
            Return
        End If

        query = "SELECT id FROM TehOperation WHERE Name like'%" + tbNameTO.Text + "%'"
        result = ClassDbEcadmaster.ExecuteScalar(query)

        If result = String.Empty Then
            query = "INSERT INTO TehOperation (CodIdLeft, CodIdRight, Name, DateCreate, DateUpdate, TypeTehOper, Equipment, Status, Area, DateEnd, Remark) " + _
                "VALUES ('" + cbNumberCode.SelectedValue.ToString() + "', '" + tbIncrementCode.Text + "', '" + tbNameTO.Text + _
                      "', CONVERT (DATETIME, '" + dtpDateCreateTO.Value.ToString() + "', 104), CONVERT (DATETIME, '" + dtpDateCreateTO.Value.ToString() + "',104), " + cbTypeAddOP.SelectedValue.ToString() + _
                      ", " + cbEquipmentAddTO.SelectedValue.ToString() + ", " + cbStatusAddTO.SelectedValue.ToString() + ", " + cbAreaAddTO.SelectedValue.ToString() + _
                      ", CONVERT(DATETIME, '" + dtpDateEndTO.Value.ToString() + "',104), '" + tbRemarkTO.Text + "')"
            ClassDbEcadmaster.ExecuteScalar(query)
            MessageBox.Show("OK")
            pnAddTehOperation.Visible = False
            loadTepOperationAll()
        Else
            MessageBox.Show("Данное наименование уже существует!", "Ошибка!", MessageBoxButtons.OK)
        End If

    End Sub
    ''' <summary>
    ''' Метод формирования части кода (00000000Х)
    ''' </summary>
    ''' <param name="number"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MakingCod(ByVal number As Integer)
        Dim cod As String = String.Empty
        If number <> 0 Then
            number += 1
            Select Case number
                Case Is <= 9 : cod = "00000000" + Convert.ToString(number)
                Case Is <= 99 : cod = "0000000" + Convert.ToString(number)
                Case Is <= 999 : cod = "000000" + Convert.ToString(number)
                Case Is <= 9999 : cod = "00000" + Convert.ToString(number)
                Case Is <= 99999 : cod = "0000" + Convert.ToString(number)
                Case Is <= 999999 : cod = "000" + Convert.ToString(number)
                Case Is <= 9999999 : cod = "00" + Convert.ToString(number)
                Case Is <= 99999999 : cod = "0" + Convert.ToString(number)
                Case Is <= 999999999 : cod = Convert.ToString(number)
            End Select
        End If
        Return cod
    End Function

    'Private Sub tsmiAddTO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiAddTO.Click
    '    Dim maxNumber As Integer
    '    Dim nextCod As String = String.Empty

    '    tbNameTO.Text = String.Empty
    '    tbRemarkTO.Text = String.Empty

    '    'Формирование кода ТО
    '    query = "SELECT MAX(CodIdRight) as 'Cod' FROM TehOperation"
    '    If ClassDbEcadmaster.ExecuteScalar(query) <> "" Then
    '        maxNumber = ClassDbEcadmaster.ExecuteScalar(query)
    '        tbIncrementCode.Text = Convert.ToString(maxNumber)
    '    Else
    '        tbIncrementCode.Text = "000000001"
    '    End If


    '    'query = "SELECT id, StatusName FROM Status WHERE Type = 2"
    '    'dt = ClassDbEcadmaster.FillDataTable(query)
    '    'cbStatusAddTO.DataSource = dt.Copy()
    '    'cbStatusAddTO.DisplayMember = "StatusName"
    '    'cbStatusAddTO.ValueMember = "id"

    '    query = "SELECT id, NamePodr FROM KatPodr WHERE Cex <> 0 AND Cex <> 999 AND NamePodr not like '%Сборщик%' AND Cex <= 5 ORDER BY NomRes"
    '    dt = ClassDbEcadmaster.FillDataTable(query)
    '    cbAreaAddTO.DataSource = dt.Copy()
    '    cbAreaAddTO.DisplayMember = "NamePodr"
    '    cbAreaAddTO.ValueMember = "id"

    '    query = "SELECT id, name FROM Equipment"
    '    dt = ClassDbEcadmaster.FillDataTable(query)
    '    cbEquipmentAddTO.DataSource = dt.Copy()
    '    cbEquipmentAddTO.DisplayMember = "Name"
    '    cbEquipmentAddTO.ValueMember = "id"

    '    query = "SELECT id, Name FROM TypeTransaction"
    '    dt = ClassDbEcadmaster.FillDataTable(query)
    '    cbTypeAddOP.DataSource = dt.Copy()
    '    cbTypeAddOP.DisplayMember = "Name"
    '    cbTypeAddOP.ValueMember = "id"

    '    dtpDateCreateTO.Value = New Date(Now.Year, Now.Month, Now.Day)
    '    dtpDateEndTO.Value = New Date(Now.Year, Now.Month, Now.Day)

    '    pnAddTehOperation.Visible = True
    'End Sub

    Private Sub loadComboBoxStatus()
        query = "SELECT id, StatusName FROM Status WHERE Type = 2"
        dt = ClassDbWorkBase.FillDataTable(query)
        cbStatusChronometrage.DataSource = dt.Copy()
        cbStatusChronometrage.DisplayMember = "StatusName"
        cbStatusChronometrage.ValueMember = "id"
    End Sub

    'Метод изменения видимости панели, в которой происходит привязка ТО к ТП
    Private Sub tsmiBondTO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiBondTO.Click

        query = "SELECT id, Name FROM TehOperation WHERE id = " + gcTehOperationViewAll.GetRowCellValue(gcTehOperationViewAll.FocusedRowHandle, "id").ToString()
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbToBondTP.DataSource = dt.Copy()
        cbToBondTP.DisplayMember = "Name"
        cbToBondTP.ValueMember = "id"

        query = "SELECT id, Name FROM TehProcess"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbTpBondTP.DataSource = dt.Copy()
        cbTpBondTP.DisplayMember = "Name"
        cbTpBondTP.ValueMember = "id"

        pnBondTO.Visible = True
    End Sub

    'Метод привязки ТО к ТП
    Private Sub btBondTO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btBondTO.Click
        query = "INSERT INTO GroupOperation (idTehProc, idTehOper, Variation, SerialNumber) " + _
                    "VALUES(" + cbTpBondTP.SelectedValue.ToString() + ", " + cbToBondTP.SelectedValue.ToString() + ", 1, 1)"
        ClassDbEcadmaster.ExecuteScalar(query)
        MessageBox.Show("Тех операция привязана")
        'gcTehProcess_MouseClick(Nothing, Nothing)
    End Sub

    'Метод изменения видимости панели, в которой происходит привязка ТО к ТП
    Private Sub btClosedBondTO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosedBondTO.Click
        pnBondTO.Visible = False
    End Sub

    'Метод изменения видимости панели, в которой происходит добавление новых записей в comboBox
    Private Sub btAddNodeNotVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAddNodeNotVisible.Click
        Dim maxNumberType As Integer = 0
        Dim maxNumberEquipment As Integer = 0

        tbNameTypeTO.Text = String.Empty
        tbNameEquipmentTO.Text = String.Empty

        pnAddNodeNotVisible.Visible = True
    End Sub

    Private Sub btClosePnAddNodeNotVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosePnAddNodeNotVisible.Click
        tbNameTypeTO.Text = String.Empty
        tbNameEquipmentTO.Text = String.Empty
        tsTypeTO.EditValue = False
        tsEquipmentTO.EditValue = False
        pnAddNodeNotVisible.Visible = False
    End Sub

    Private Sub tsTypeTO_Toggled(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsTypeTO.Toggled
        If tsTypeTO.EditValue Then
            gbTypeTO.Enabled = True
            tbNameTypeTO.Text = String.Empty
            tsEquipmentTO.EditValue = False
        Else : gbTypeTO.Enabled = False
        End If
    End Sub

    Private Sub tsEquipmentTO_Toggled(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsEquipmentTO.Toggled
        If tsEquipmentTO.EditValue Then
            gbEquipmentTO.Enabled = True
            tbNameEquipmentTO.Text = String.Empty
            tsTypeTO.EditValue = False
        Else : gbEquipmentTO.Enabled = False
        End If
    End Sub

    Private Sub btAddPnTypeOrEquipment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAddPnTypeOrEquipment.Click
        Dim result As String = String.Empty

        'If tbNameTypeTO.Text = "" OrElse tbNameEquipmentTO.Text = "" Then
        '    MessageBox.Show("Присутствуют пустые поля!", "Ошибка!", MessageBoxButtons.OK)
        '    Return
        'End If

        If tsTypeTO.EditValue = True Then
            If tbNameTypeTO.Text <> String.Empty Then
                query = "SELECT id FROM TypeTransaction WHERE NAME like'%" + tbNameTypeTO.Text + "%'"
                result = ClassDbEcadmaster.ExecuteScalar(query)
                If result = String.Empty Then
                    query = "INSERT INTO TypeTransaction (Name) VALUES('" + tbNameTypeTO.Text + "')"
                    ClassDbEcadmaster.ExecuteScalar(query)
                Else : MessageBox.Show("Данное наименование уже существует!", "Ошибка!", MessageBoxButtons.OK)
                End If
            Else : MessageBox.Show("Присутствуют пустые поля!", "Ошибка!", MessageBoxButtons.OK)
            End If
        End If

        If tsEquipmentTO.EditValue = True Then
            If tbNameEquipmentTO.Text <> String.Empty Then
                query = "SELECT id FROM Equipment WHERE NAME like'%" + tbNameEquipmentTO.Text + "%'"
                result = ClassDbEcadmaster.ExecuteScalar(query)
                If result = String.Empty Then
                    query = "INSERT INTO Equipment (Name) VALUES('" + tbNameEquipmentTO.Text + "')"
                    ClassDbEcadmaster.ExecuteScalar(query)
                Else : MessageBox.Show("Данное наименование уже существует!", "Ошибка!", MessageBoxButtons.OK)
                End If
            Else : MessageBox.Show("Присутствуют пустые поля!", "Ошибка!", MessageBoxButtons.OK)
            End If
        End If

        tsmiAddTO1_Click(Nothing, Nothing) 'обновление данных комбобоксов
    End Sub

    Private Sub gcTehOperation_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles gcTehOperation.MouseClick
        If gcTehOperationView.SelectedRowsCount = 0 Then
            Return
        End If

        query = "SELECT Chronometrage.id,Chronometrage.Cod as 'Номер', Chronometrage.CodeOperation as 'Номер ТО', Chronometrage.DateApplying as 'Применен', Chronometrage.DateValidBy as 'Завершен', " + _
                    "Chronometrage.NextCodeChronometrage as 'Сл. номер хронометража', Status.StatusName as 'Статус', Users.FIO as 'Обработал', Chronometrage.Remark as 'Примечание' " + _
                "FROM Chronometrage INNER JOIN " + _
                      "Status ON Chronometrage.Status = Status.id INNER JOIN " + _
                      "Users ON Chronometrage.Working = Users.idUsers " + _
                "WHERE Chronometrage.CodeOperation = '" + gcTehOperationView.GetRowCellValue(gcTehOperationView.FocusedRowHandle, "Код").ToString() + "'"
        gcChronometrage.DataSource = ClassDbEcadmaster.FillDataTable(query)
    End Sub

    Private Sub tsmiDisplayChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiDisplayChronometrage.Click

        If gcChronometrageView.SelectedRowsCount = 0 Then
            Return
        End If

        DisplayChronometrage(gcChronometrageView.GetRowCellValue(gcChronometrageView.FocusedRowHandle, "id"))
        pnDisplay.Visible = True
    End Sub
    ''' <summary>
    ''' Метод для отображения хронометража
    ''' </summary>
    ''' <param name="idChronometrage">dbo.Chronometrage.id</param>
    ''' <remarks>В методе загрузка данных в gcDisplay</remarks>
    Private Sub DisplayChronometrage(ByVal idChronometrage As Integer)
        query = "SELECT id ,Cod ,CodeOperation,Status, " + _
                    "DateCreate, DateApplying, " + _
                    "DateValidBy ,Working, Remark,C ,TL, " + _
                    "CL, TW, CW, TH, CH, " + _
                    "TSP, CSP, TSK, CSK, " + _
                    "TST, CST, TV, CV, TPP, CPP, TPK, CPK, " + _
                    "TPT, CPT, NextCodeChronometrage " + _
                "FROM Chronometrage where id= " + idChronometrage.ToString()
        gcDisplay.DataSource = ClassDbEcadmaster.FillDataTable(query)

    End Sub

    Private Sub btCloseDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCloseDisplay.Click
        pnDisplay.Visible = False
    End Sub

    Private Sub btClosePnAddCronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosePnAddCronometrage.Click
        btUpdateChronometrage.Visible = False
        pnAddCronometrage.Visible = False
    End Sub

    Private Sub tsmiAddChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiAddChronometrage.Click
        Dim result As String = String.Empty

        
        ClearPnChronometrage()
        loadComboBoxStatus()

        'query = "SELECT id, StatusName FROM Status WHERE Type = 2"
        'dt = ClassDbWorkBase.FillDataTable(query)
        'cbStatusChronometrage.DataSource = dt.Copy()
        'cbStatusChronometrage.DisplayMember = "StatusName"
        'cbStatusChronometrage.ValueMember = "id"

        query = "SELECT MAX(Cod) FROM Chronometrage"
        result = ClassDbEcadmaster.ExecuteScalar(query)
        result += 1
        tbNumberChrono.Text = result

        pnAddCronometrage.Visible = True
    End Sub

    ''' <summary>
    ''' Очистка панели добавления / изменения хронометража
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearPnChronometrage()
        tbNumberChrono.Text = String.Empty
        mtbNextCodChrono.Text = String.Empty
        dtpDateApplyingChrono.Value = Date.Now
        dtpDateEndChrono.Value = Date.Now
        tbRemarkChrono.Text = String.Empty
        For Each ctrl As Control In GroupBox5.Controls
            If TypeOf ctrl Is TextBox Then
                If ctrl.Text <> String.Empty Then
                    ctrl.Text = String.Empty
                End If
            End If
        Next
    End Sub

    Private Sub btAddChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAddChronometrage.Click
        Dim arrayCV As New Dictionary(Of String, Integer)

        'If mtbNextCodChrono.Text = String.Empty Then
        '    Return
        'End If

        'Если одно из полей для коэффицентов не равно Empty
        For Each ctrl As Control In GroupBox5.Controls
            If TypeOf ctrl Is TextBox Then
                If ctrl.Tag = "tb_coefficient" Then
                    If ctrl.Text <> "" Then
                        arrayCV.Add(ctrl.Name, ctrl.Text)
                    End If
                End If
            End If
        Next

        Dim seet As String = String.Empty
        Dim seetValue As String = String.Empty

        For i As Integer = 0 To arrayCV.Keys.Count - 1
            seet += IIf(String.IsNullOrEmpty(seet), "", ",") + arrayCV.Keys(i).ToString()
            seetValue += IIf(String.IsNullOrEmpty(seetValue), "", ",") + arrayCV.Values(i).ToString()
        Next

        'Если длина равно 0, то добавляем в БД без коэффицентов
        If gcTehOperationView.SelectedRowsCount <> 0 Then
            If seet.Length <> 0 AndAlso seetValue.Length <> 0 Then
                query = "INSERT INTO Chronometrage (Cod, CodeOperation, NextCodeChronometrage, Status, DateCreate, DateApplying,DateValidBy, Working, Remark, " + seet + ") " + _
                        "VALUES(" + tbNumberChrono.Text + ", " + gcTehOperationView.GetRowCellValue(gcTehOperationView.FocusedRowHandle, "Код").ToString() + ", '" + mtbNextCodChrono.Text + "'" + _
                        ", " + cbStatusChronometrage.SelectedValue.ToString() + ", Convert (DateTime,'" + Date.Now + "',104), Convert(DateTime, '" + dtpDateApplyingChrono.Value.ToString() + "',104), Convert(DateTime, '" + dtpDateEndChrono.Value.ToString() + "', 104), " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + ", '" + tbRemarkChrono.Text + "', " + seetValue + " )"
            Else
                query = "INSERT INTO Chronometrage (Cod, CodeOperation, NextCodeChronometrage, Status, DateCreate, DateApplying,DateValidBy, Working, Remark) " + _
                            "VALUES(" + tbNumberChrono.Text + ", " + gcTehOperationView.GetRowCellValue(gcTehOperationView.FocusedRowHandle, "Код").ToString() + ", '" + mtbNextCodChrono.Text + "'" + _
                            ", " + cbStatusChronometrage.SelectedValue.ToString() + ", Convert (DateTime,'" + Date.Now + "',104), Convert(DateTime, '" + dtpDateApplyingChrono.Value.ToString() + "',104), Convert(DateTime, '" + dtpDateEndChrono.Value.ToString() + "', 104), " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + ", '" + tbRemarkChrono.Text + "' )"
            End If
            ClassDbEcadmaster.ExecuteNonQuery(query)
            loadChronometrage()
            gcTehOperation_MouseClick(Nothing, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' Загрузка всех хронометражей
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub loadChronometrage()
        query = "SELECT id, Cod, CodeOperation, NextCodeChronometrage, Status, " + _
                    "DateCreate, DateApplying, DateValidBy, Working, Remark " + _
                "FROM Chronometrage WHERE CodeOperation = ''"
        gcChronometrageAll.DataSource = ClassDbEcadmaster.FillDataTable(query)
    End Sub

    Private Sub cbToBondTP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbToBondTP.Click
        query = "SELECT id, Name FROM TehOperation"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbToBondTP.DataSource = dt.Copy()
        cbToBondTP.DisplayMember = "Name"
        cbToBondTP.ValueMember = "id"
    End Sub

    Private Sub tsmiUpdateChrono_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiUpdateChrono.Click
        If gcChronometrageView.SelectedRowsCount = 0 Then
            Return
        End If

        query = "SELECT Chronometrage.id, Chronometrage.Cod, Chronometrage.CodeOperation, Chronometrage.NextCodeChronometrage, Status.StatusName, Chronometrage.Status, " + _
                        "Chronometrage.DateCreate, Chronometrage.DateApplying, Chronometrage.DateValidBy, Chronometrage.Working, Chronometrage.Remark, " + _
                        "Chronometrage.C, Chronometrage.TL, Chronometrage.CL, Chronometrage.TW, Chronometrage.CW, Chronometrage.TH, " + _
                        "Chronometrage.CH, Chronometrage.TSP, Chronometrage.CSP, Chronometrage.TSK, Chronometrage.CSK, Chronometrage.TST, " + _
                        "Chronometrage.CST, Chronometrage.TV, Chronometrage.CV, Chronometrage.TPP, Chronometrage.CPP, Chronometrage.TPK, " + _
                        "Chronometrage.CPK, Chronometrage.TPT, Chronometrage.CPT " + _
                "FROM Chronometrage INNER JOIN Status ON Chronometrage.Status = Status.id " + _
                "WHERE Chronometrage.id = " + gcChronometrageView.GetRowCellValue(gcChronometrageView.FocusedRowHandle, "id").ToString()
        dt = ClassDbEcadmaster.FillDataTable(query)

        tbNumberChrono.Text = dt.Rows(0)("Cod")
        If IsDBNull(dt.Rows(0)("NextCodeChronometrage")) Then
            mtbNextCodChrono.Text = ""
        Else
            mtbNextCodChrono.Text = dt.Rows(0)("NextCodeChronometrage")
        End If

        dtpDateApplyingChrono.Value = dt.Rows(0)("DateApplying")
        dtpDateEndChrono.Value = dt.Rows(0)("DateValidBy")
        tbRemarkChrono.Text = dt.Rows(0)("Remark")

        'заносим значения коэффицентов в textbox если,
        'имена texbox совпадают с названиями колонок dt, то заносим значения
        For Each ctrl As Control In GroupBox5.Controls
            If TypeOf ctrl Is TextBox Then
                For i As Integer = 0 To dt.Columns.Count - 1
                    If ctrl.Name = dt.Columns(i).ColumnName Then
                        ctrl.Text = dt.Rows(0)(i).ToString()
                    End If
                Next
            End If
        Next

        cbStatusChronometrage.DataSource = dt.Copy()
        cbStatusChronometrage.DisplayMember = "StatusName"
        cbStatusChronometrage.ValueMember = "Status"

        btUpdateChronometrage.Visible = True
        btAddChronometrage.Visible = False
        pnAddCronometrage.Visible = True
    End Sub

    Private Sub cbStatusChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStatusChronometrage.Click
        loadComboBoxStatus()
    End Sub

    Private Sub btUpdateChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btUpdateChronometrage.Click
        query = "UPDATE Chronometrage " + _
                "SET Cod = " + tbNumberChrono.Text + _
                  " ,NextCodeChronometrage = '" + mtbNextCodChrono.Text + "'" + _
                  " ,Status =  " + cbStatusChronometrage.SelectedValue.ToString() + _
                  " ,DateApplying = Convert(Datetime, '" + dtpDateApplyingChrono.Value.ToString() + "', 104)" + _
                  " ,DateValidBy = Convert(Datetime, '" + dtpDateEndChrono.Value.ToString() + "', 104)" + _
                  " ,Working =  " + CType(Me.ParentForm, MainForm).protect.idUser.ToString() + _
                  " ,Remark = '" + tbRemarkChrono.Text + "'" + _
                  " ,C = " + C.Text + _
                  " ,TL = " + TL.Text + _
                  " ,CL = " + CL.Text + _
                  " ,TW = " + TW.Text + _
                  " ,CW = " + CW.Text + _
                  " ,TH = " + TH.Text + _
                  " ,CH = " + CH.Text + _
                  " ,TSP = " + TSP.Text + _
                  " ,CSP = " + CSP.Text + _
                  " ,TSK = " + TSK.Text + _
                  " ,CSK = " + CSK.Text + _
                  " ,TST = " + TST.Text + _
                  " ,CST = " + CST.Text + _
                  " ,TV = " + TV.Text + _
                  " ,CV= " + CV.Text + _
                  " ,TPP = " + TPP.Text + _
                  " ,CPP = " + CPP.Text + _
                  " ,TPK = " + TPK.Text + _
                  " ,CPK = " + CPK.Text + _
                  " ,TPT = " + TPT.Text + _
                  " ,CPT = " + CPT.Text + _
              " WHERE id = " + gcChronometrageView.GetRowCellValue(gcChronometrageView.FocusedRowHandle, "id").ToString()

        If MessageBox.Show("Вы уверены что хотите изменить данные о хронометраже ?!", "Изменение данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            ClassDbEcadmaster.ExecuteNonQuery(query)
            MessageBox.Show("Запись успешно изменена !", "Изменение данных", MessageBoxButtons.OK)
            loadChronometrage()
            gcTehOperation_MouseClick(Nothing, Nothing)
        Else : Return
        End If
    End Sub
    ' ''' <summary>
    ' ''' Редактирование данных хронометража
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub UpdateChronometrage(ByVal id)

    'End Sub

    Private Sub tsmiBondChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiBondChronometrage.Click

        If gcChronometrageViewAll.SelectedRowsCount = 0 Then
            Return
        End If

        query = "SELECT id, Name FROM TehOperation "
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbBondNameTOandChrono.DataSource = dt.Copy()
        cbBondNameTOandChrono.DisplayMember = "Name"
        cbBondNameTOandChrono.ValueMember = "id"

        query = "SELECT id,Cod,CodeOperation FROM Chronometrage WHERE id = " + gcChronometrageViewAll.GetRowCellValue(gcChronometrageViewAll.FocusedRowHandle, "id").ToString()
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbBondCodeChrono.DataSource = dt.Copy()
        cbBondCodeChrono.DisplayMember = "Cod"
        cbBondCodeChrono.ValueMember = "id"

        pnBondChronometrage.Visible = True
    End Sub


    Private Sub btClosePnBondChrono_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosePnBondChrono.Click
        pnBondChronometrage.Visible = False
    End Sub

    Private Sub cbBondCodeChrono_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBondCodeChrono.Click
        query = "SELECT id,Cod FROM Chronometrage WHERE CodeOperation = ''"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbBondCodeChrono.DataSource = dt.Copy()
        cbBondCodeChrono.DisplayMember = "Cod"
        cbBondCodeChrono.ValueMember = "id"
    End Sub


    Private Sub btBondToOfChronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btBondToOfChronometrage.Click
        Dim numberTO As String = String.Empty

        query = "SELECT (CodIdLeft + CodIdRight) as 'Cod', Name FROM TehOperation WHERE id = " + cbBondNameTOandChrono.SelectedValue.ToString()
        dt = ClassDbEcadmaster.FillDataTable(query)
        'numberTO = ClassDbWorkBase.ExecuteScalar(query)
        If MessageBox.Show("Вы уверены что хотите привязать " + dt.Rows(0)("Name"), "Изменение данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            query = "UPDATE Chronometrage SET CodeOperation = " + dt.Rows(0)("Cod").ToString() + " WHERE id = " + cbBondCodeChrono.SelectedValue.ToString()
            ClassDbEcadmaster.ExecuteScalar(query)
            MessageBox.Show("Хронометраж успешно привязан!", "", MessageBoxButtons.OK)
            loadChronometrage()
            gcTehOperation_MouseClick(Nothing, Nothing)
        End If
    End Sub

    Private Sub tsmiToUntieCronometrage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiToUntieCronometrage.Click
        If MessageBox.Show("Хотите отвязать хронометраж №" + gcChronometrageView.GetRowCellValue(gcChronometrageView.FocusedRowHandle, "Номер").ToString(), "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            query = "UPDATE Chronometrage SET CodeOperation = '' WHERE id = " + gcChronometrageView.GetRowCellValue(gcChronometrageView.FocusedRowHandle, "id").ToString()
            ClassDbEcadmaster.ExecuteScalar(query)
            loadChronometrage()
            gcTehOperation_MouseClick(Nothing, Nothing)
            MessageBox.Show("Операция успешно завершена!", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Else : MessageBox.Show("Операция НЕ завершена!", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End If
    End Sub

    Private Sub tsmiDisplay_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiDisplay.Click
        If gcChronometrageView.SelectedRowsCount = 0 Then
            Return
        End If

        DisplayChronometrage(gcChronometrageViewAll.GetRowCellValue(gcChronometrageViewAll.FocusedRowHandle, "id"))
        pnDisplay.Visible = True
    End Sub

    Private Sub tsmiAddTO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiAddTO1.Click
        Dim maxNumber As Integer
        Dim nextCod As String = String.Empty

        tbNameTO.Text = String.Empty
        tbRemarkTO.Text = String.Empty

        'Формирование кода ТО
        query = "SELECT MAX(CodIdRight) as 'Cod' FROM TehOperation"
        If ClassDbEcadmaster.ExecuteScalar(query) <> "" Then
            maxNumber = ClassDbEcadmaster.ExecuteScalar(query)
            tbIncrementCode.Text = MakingCod(maxNumber)
        Else
            tbIncrementCode.Text = "000000001"
        End If

        query = "SELECT id, StatusName FROM Status WHERE Type = 2"
        dt = ClassDbWorkBase.FillDataTable(query)
        cbStatusAddTO.DataSource = dt.Copy()
        cbStatusAddTO.DisplayMember = "StatusName"
        cbStatusAddTO.ValueMember = "id"

        query = "SELECT id, NamePodr FROM KatPodr WHERE Cex <> 0 AND Cex <> 999 AND NamePodr not like '%Сборщик%' AND Cex <= 5 ORDER BY NomRes"
        dt = ClassDbWorkBase.FillDataTable(query)
        cbAreaAddTO.DataSource = dt.Copy()
        cbAreaAddTO.DisplayMember = "NamePodr"
        cbAreaAddTO.ValueMember = "id"

        query = "SELECT id, name FROM Equipment"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbEquipmentAddTO.DataSource = dt.Copy()
        cbEquipmentAddTO.DisplayMember = "Name"
        cbEquipmentAddTO.ValueMember = "id"

        query = "SELECT id, Name FROM TypeTransaction"
        dt = ClassDbEcadmaster.FillDataTable(query)
        cbTypeAddOP.DataSource = dt.Copy()
        cbTypeAddOP.DisplayMember = "Name"
        cbTypeAddOP.ValueMember = "id"

        dtpDateCreateTO.Value = New Date(Now.Year, Now.Month, Now.Day)
        dtpDateEndTO.Value = New Date(Now.Year, Now.Month, Now.Day)

        pnAddTehOperation.Visible = True
    End Sub

    'Private Sub tsmiToUntieTO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiToUntieTO.Click
    '    If gcChronometrageView.SelectedRowsCount = 0 Then
    '        Return
    '    End If

    '    If MessageBox.Show("Хотите отвязать ТО: " + gcTehOperationView.GetRowCellValue(gcTehOperationView.FocusedRowHandle, "Наименование").ToString(), "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
    '        query = "" + gcTehProcessView.GetRowCellValue(gcTehProcessView.FocusedRowHandle, "id").ToString() + " ___ " + gcTehOperationView.GetRowCellValue(gcTehOperationView.FocusedRowHandle, "id").ToString()
    '        MessageBox.Show(query)
    '    End If

    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form_ As New CardTehProcess
        form_.Show()
    End Sub
End Class