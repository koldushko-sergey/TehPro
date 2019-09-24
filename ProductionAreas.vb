Imports SergeyDll
Imports System.IO
Imports DevExpress.XtraTreeList.Nodes
Imports System.Collections
Imports System.Drawing
Imports System.Windows.Forms.ListView
Imports System.Threading


Public Class ProductionAreas
    Private query As String
    Private idZakaz, idPodr, idUser As Integer
    Private dt As DataTable
    Private lvSelectedRowColor As Color = Color.LimeGreen
    Private lvwColumnSorter As ListViewColumnSorter

    Public Sub New(ByVal _idZakaz As Integer, ByVal _idPodr As Integer, ByVal _idUser As Integer)
        idZakaz = _idZakaz
        idPodr = _idPodr
        idUser = _idUser
        InitializeComponent()

        lvwColumnSorter = New ListViewColumnSorter()
        lvwColumnSorter.SortColumn = 1
        lvwColumnSorter.Order = SortOrder.Ascending
        lvUserGroup.ListViewItemSorter = lvwColumnSorter
    End Sub

    Private Sub ProductionAreas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CType(Me.ParentForm, MainForm).protect.Protect(Me)

        For Each btn As Control In pnRegistration.Controls
            If (Not TypeOf btn Is Button) Then Continue For
            btn.AutoSize = (btn.Enabled And btn.Visible)
        Next
        For Each btn As Control In pnComplButtons.Controls
            If (Not TypeOf btn Is Button) Then Continue For
            btn.AutoSize = (btn.Enabled And btn.Visible)
        Next
        btRegistratoinStatus.AutoSize = (btRegistratoinStatus.Enabled And btRegistratoinStatus.Visible)

        processLoadForm()
        RedefinitionCbFilterGroup()
    End Sub

    Private Sub processLoadForm()
        'участок
        query = "SELECT KatPodr.id, KatPodr.NamePodr " + _
                "FROM KatPodr INNER JOIN UsersParam ON KatPodr.id = UsersParam.idValue " + _
                "WHERE KatPodr.Cex <> 0 AND KatPodr.Cex <> 999 AND KatPodr.NamePodr not like '%Сборщик%' AND KatPodr.Cex <= 5 AND UsersParam.isPodr = 1 " + _
                    "AND UsersParam.idUsers = " + idUser.ToString() + " " + _
                "ORDER BY KatPodr.NomRes "
        dt = ClassDbWorkBase.FillDataTable(query)
        cbDevision.Tag = 1
        cbDevision.DataSource = dt.Copy()
        cbDevision.DisplayMember = "NamePodr"
        cbDevision.ValueMember = "id"
        cbDevision.SelectedValue = idPodr.ToString()
        cbDevision.Tag = 0

        'статус
        query = "SELECT id, StatusName FROM Status Where Type = 1"
        dt = ClassDbWorkBase.FillDataTable(query)
        cbStatus.DataSource = dt.Copy()
        cbStatus.DisplayMember = "StatusName"
        cbStatus.SelectedIndex = -1
        cbStatus.ValueMember = "id"

        'Dim thr As New Thread(
        'Sub()
        listBoxLoadComponents()
        'End Sub)
        'thr.Start()

        query = "SELECT id As idGroupMC, NameGr FROM GroupMC WHERE KodGr <> 999 AND cMainGr = 0 ORDER BY NameGr "
        dt = ClassDbWorkBase.FillDataTable(query)
        cbFilterGroup.DisplayMember = "NameGr"
        cbFilterGroup.ValueMember = "idGroupMC"
        cbFilterGroup.DataSource = dt.Copy()

        TreeViewLoad()
        RedefinitionCbFilterGroup()
    End Sub

    Private Sub TreeViewLoad() 'загрузка дерева деталей заказа
        Dim orderBranch, groupBranch, detailBranch As TreeNode
        Dim listGroups As String = "-1"
        Dim numberOrder, listOfKatMcIds As String
        Dim idGroup As Integer
        Dim filter = String.Empty, descriptionNode, imageForNode As String
        Dim dtIsCompl, dtBuf, dtBuf2 As DataTable
        Dim listOfFiles As String()
        Dim isWork, isCompl As Boolean

        tvListOfDetails.Nodes.Clear()
        lbStatus.Visible = False
        groupBranch = Nothing

        dt = ClassDbWorkBase.FillDataTable("SELECT * FROM BondKatPodrAndGroupMC WHERE idKatPodr = " + idPodr.ToString())

        If dt.Rows.Count <> 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                listGroups += IIf(String.IsNullOrEmpty(listGroups), "", ",") + dt.Rows(i)("idGroupMC").ToString()
            Next
        End If

        listOfKatMcIds = ClassDbWorkBase.ExecuteScalar("SELECT listOfKatMcId From KatPodr WHERE id = " + idPodr.ToString())
        listOfKatMcIds = IIf(String.IsNullOrEmpty(listOfKatMcIds), "-1", listOfKatMcIds)

        'Выбираем изделия привязанные к участку через группу
        filter = " AND KatMC.idGrMC IN (" + listGroups + ")"
        query = "SELECT SpZak.id AS idSpZak, SpZak.Kol, SpZak.SizeX, SpZak.SizeY, SpZak.SizeZ, SpZak.Rem, SpZak.RemFakt, SpZak.idGroup, KatMC.NameMC,GroupMC.NameGr, OrderMark.idStatus, OrderMark.isCompl " + _
                "FROM KatMC INNER JOIN " + _
                    "SpZak ON SpZak.idMC = KatMC.id INNER JOIN " + _
                    "GroupMC ON KatMC.idGrMC = GroupMC.id INNER JOIN " + _
                    "BondKatPodrAndGroupMC ON BondKatPodrAndGroupMC.idGroupMC = KatMC.idGrMC LEFT JOIN " + _
                    "OrderMark ON OrderMark.idSpZak = SpZak.id AND BondKatPodrAndGroupMC.idKatPodr = OrderMark.idPodr " + _
                "WHERE SpZak.idZakaz = " + idZakaz.ToString() + " AND " + _
                    "BondKatPodrAndGroupMC.idKatPodr = " + idPodr.ToString() + filter + " " + _
                    "GROUP BY SpZak.id, SpZak.Kol, SpZak.SizeX, SpZak.SizeY, SpZak.SizeZ, SpZak.Rem, SpZak.RemFakt, SpZak.idGroup, KatMC.NameMC,GroupMC.NameGr, OrderMark.idStatus, OrderMark.isCompl " + _
                "ORDER BY GroupMC.NameGr, KatMC.NameMC, SpZak.SizeX, SpZak.SizeY "
        dt = ClassDbWorkBase.FillDataTable(query)

        'Выбираем изделия привязанные к участку напрямую
       
        query = "SELECT SpZak.id AS idSpZak, SpZak.Kol, SpZak.SizeX, SpZak.SizeY, SpZak.SizeZ, SpZak.Rem, SpZak.RemFakt, SpZak.idGroup, KatMC.NameMC,GroupMC.NameGr, null as idStatus, null as isCompl " + _
                "FROM KatMC INNER JOIN " + _
                    "SpZak ON SpZak.idMC = KatMC.id INNER JOIN " + _
                    "GroupMC ON KatMC.idGrMC = GroupMC.id " + _
                "WHERE SpZak.idZakaz = " + idZakaz.ToString() + " AND KatMC.id IN (" + listOfKatMcIds + ") " + _
                "ORDER BY SpZak.id, GroupMC.NameGr, KatMC.NameMC, SpZak.SizeX, SpZak.SizeY"
        dtBuf = ClassDbWorkBase.FillDataTable(query)

        Dim listOfSpZakIds As String = "-1"
        For i As Integer = 0 To dtBuf.Rows.Count - 1
            listOfSpZakIds += IIf(String.IsNullOrEmpty(listOfSpZakIds), "", ",") + dtBuf.Rows(i)("idSpZak").ToString()
        Next

        query = "SELECT * FROM OrderMark WHERE idZakaz = " + idZakaz.ToString() + " AND idPodr = " + idPodr.ToString() + " AND idSpZak IN (" + listOfSpZakIds + ")"
        dtBuf2 = ClassDbWorkBase.FillDataTable(query)

        dtBuf.Columns("idStatus").ReadOnly = False
        dtBuf.Columns("isCompl").ReadOnly = False

        For i As Integer = 0 To dtBuf2.Rows.Count - 1
            For j As Integer = 0 To dtBuf.Rows.Count - 1
                If dtBuf2.Rows(i)("idSpZak") = dtBuf.Rows(j)("idSpZak") Then
                    dtBuf.Rows(j)("idStatus") = dtBuf2.Rows(i)("idStatus")
                    dtBuf.Rows(j)("isCompl") = (dtBuf2.Rows(i)("isCompl"))
                End If
            Next
        Next

        dt.Merge(dtBuf, True, MissingSchemaAction.Ignore)

        query = "SELECT idSpZak FROM OrderMark WHERE idZakaz = " + idZakaz.ToString() + " AND idStatus = " + ClassCommon.STATUS_COMPL.ToString() + " AND isCompl = 1 GROUP BY idSpZak"
        dtIsCompl = ClassDbWorkBase.FillDataTable(query)
        Dim arrIsCompl(dtIsCompl.Rows.Count - 1) As Integer
        For i As Integer = 0 To dtIsCompl.Rows.Count - 1
            arrIsCompl(i) = dtIsCompl.Rows(i)("idSpZak")
        Next

        numberOrder = ClassDbWorkBase.ExecuteScalar("SELECT Nomer FROM Zakaz WHERE id = " + idZakaz.ToString())
        orderBranch = tvListOfDetails.Nodes.Add("nodeOrder", numberOrder, "ilPngBoxOpened", "ilPngBoxOpened")

        For i As Integer = 0 To dt.Rows.Count - 1

            If dt.Rows(i)("idGroup") <> idGroup Then
                If Not IsNothing(groupBranch) Then
                    If isCompl Then
                        groupBranch.ImageKey = "ilPngBoxGreen"
                        groupBranch.SelectedImageKey = "ilPngBoxGreen"
                    ElseIf isWork Then
                        groupBranch.ImageKey = "ilPngBox"
                        groupBranch.SelectedImageKey = "ilPngBox"
                    End If
                End If

                isWork = True
                isCompl = True

                groupBranch = orderBranch.Nodes.Add("nodeGroup" + i.ToString(), dt.Rows(i)("NameGr"), "ilPngBoxRed", "ilPngBoxRed")
                groupBranch.Tag = dt.Rows(i)("idGroup").ToString()
            End If

            isWork = (isWork And IIf(Not IsDBNull(dt.Rows(i)("idStatus")) AndAlso dt.Rows(i)("idStatus") = ClassCommon.STATUS_READY, True, False))
            isCompl = (isCompl And IIf((Not IsDBNull(dt.Rows(i)("idStatus")) AndAlso dt.Rows(i)("isCompl") = True) OrElse Array.IndexOf(arrIsCompl, dt.Rows(i)("idSpZak")) >= 0, True, False))

            descriptionNode = dt.Rows(i)("NameMC") + "   " + dt.Rows(i)("SizeX").ToString() + "x" + dt.Rows(i)("SizeY").ToString() + "x" + dt.Rows(i)("SizeZ").ToString() + "   " + _
                dt.Rows(i)("RemFakt") + dt.Rows(i)("Rem") + dt.Rows(i)("Kol").ToString() + " Шт."
            imageForNode = IIf(idPodr = ClassCommon.PODR_COMPL, "ilPngNotCompl", "ilPngCancel")

            If Not IsDBNull(dt.Rows(i)("idStatus")) Then
                Select Case dt.Rows(i)("idStatus")
                    Case ClassCommon.STATUS_COMPL
                        imageForNode = "ilPngCompl"
                    Case ClassCommon.STATUS_READY
                        If dt.Rows(i)("isCompl") = True Then
                            imageForNode = "ilPngCompl"
                        Else
                            imageForNode = "ilPngOk"
                        End If
                    Case ClassCommon.STATUS_PROBLEM
                        imageForNode = "ilPngWarning"
                    Case ClassCommon.STATUS_NOT_READY
                        If dt.Rows(i)("isCompl") = True Then
                            imageForNode = "ilPngComplNotWork"
                        Else
                            imageForNode = "ilPngCancel"
                        End If
                    Case ClassCommon.STATUS_NOT_COMPL
                        imageForNode = "ilPngNotCompl"
                End Select
            ElseIf (Array.IndexOf(arrIsCompl, dt.Rows(i)("idSpZak")) >= 0) Then
                imageForNode = "ilPngComplNotWork"
            End If

            detailBranch = groupBranch.Nodes.Add("nodeDetail" + i.ToString(), descriptionNode, imageForNode, imageForNode)
            detailBranch.Tag = dt.Rows(i)("idSpZak").ToString()
            idGroup = dt.Rows(i)("idGroup")
        Next


        If isCompl Then
            groupBranch.ImageKey = "ilPngBoxGreen"
            groupBranch.SelectedImageKey = "ilPngBoxGreen"
        ElseIf isWork Then
            groupBranch.ImageKey = "ilPngBox"
            groupBranch.SelectedImageKey = "ilPngBox"
        End If

        Me.Text = "Регистрация заказа | Всего деталей: " + dt.Rows.Count.ToString()
        orderBranch.Expand()

        ' поиск документов для заказа 
        lbDocumentsForOrder.Items.Clear()
        listOfFiles = Directory.GetFiles(Setting.Xml.GetXmlValue("DestinationTo"), "*" + numberOrder.ToString().Substring(0, 6) + "*.*").Where(Function(s) Not s.ToLower().Contains("price")).ToArray()
        For Each fileName As String In listOfFiles
            lbDocumentsForOrder.Items.Add(Path.GetFileName(fileName))
        Next

        buttonRightForPodr()
        lblInfoDate.Text = String.Empty
    End Sub

    'Private Function ReturnGreenOrRedIcoTV(ByVal idGroup As Integer)

    '    Try
    '        Dim colRowsGood As Integer = 0
    '        query = "SELECT SpZak.id AS idSpZak, SpZak.Kol, SpZak.SizeX, SpZak.SizeY, SpZak.SizeZ, SpZak.Rem, SpZak.RemFakt, SpZak.idGroup, KatMC.NameMC,GroupMC.NameGr, OrderMark.idStatus, OrderMark.isCompl " + _
    '          "FROM KatMC INNER JOIN " + _
    '              "SpZak ON SpZak.idMC = KatMC.id INNER JOIN " + _
    '              "GroupMC ON KatMC.idGrMC = GroupMC.id INNER JOIN " + _
    '              "BondKatPodrAndGroupMC ON BondKatPodrAndGroupMC.idGroupMC = KatMC.idGrMC LEFT JOIN " + _
    '              "OrderMark ON OrderMark.idSpZak = SpZak.id AND BondKatPodrAndGroupMC.idKatPodr = OrderMark.idPodr " + _
    '          "WHERE SpZak.idZakaz = " + idZakaz.ToString() + " AND " + _
    '              "BondKatPodrAndGroupMC.idKatPodr = " + idPodr.ToString() + " AND KatMC.idGrMC = " + idGroup.ToString() + _
    '          "ORDER BY GroupMC.NameGr, KatMC.NameMC, SpZak.SizeX, SpZak.SizeY "

    '        Dim dtGroup = ClassDbWorkBase.FillDataTable(query)

    '        For k As Integer = 0 To dtGroup.Rows.Count - 1
    '            If Not IsDBNull(dt.Rows(k)("idStatus")) OrElse Not IsNothing(dt.Rows(k)("idStatus")) Then
    '                If dtGroup.Rows(k)("idStatus").ToString() = ClassCommon.STATUS_READY.ToString() OrElse dtGroup.Rows(k)("idStatus").ToString() = ClassCommon.STATUS_COMPL.ToString() Then
    '                    colRowsGood += 1
    '                End If
    '            End If
    '        Next

    '        If colRowsGood = dtGroup.Rows.Count Then
    '            Return "Green"
    '        Else : Return "Red"
    '        End If

    '    Catch ex As Exception
    '        Console.WriteLine(ex.Message)
    '    End Try
    '    Return "NULL"
    'End Function

    Private Sub buttonRightForPodr()
        If (idPodr = ClassCommon.PODR_COMPL) Then
            For Each btn As Control In pnComplButtons.Controls
                If (Not TypeOf btn Is Button) Then Continue For
                btn.Enabled = btn.AutoSize
            Next
            For Each btn As Control In pnRegistration.Controls
                If (Not TypeOf btn Is Button) Then Continue For
                btn.Enabled = False
            Next
            btRegistratoinStatus.Enabled = False
        Else
            For Each btn As Control In pnComplButtons.Controls
                If (Not TypeOf btn Is Button) Then Continue For
                btn.Enabled = False
            Next
            For Each btn As Control In pnRegistration.Controls
                If (Not TypeOf btn Is Button) Then Continue For
                btn.Enabled = btn.AutoSize
            Next
            btRegistratoinStatus.Enabled = btRegistratoinStatus.AutoSize
        End If
    End Sub

    Private Sub btRegistratoinStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRegistratoinStatus.Click
        If (cbStatus.SelectedIndex = -1 And String.IsNullOrEmpty(CType(sender, Control).Tag)) OrElse IsNothing(tvListOfDetails.SelectedNode) Then Return
        lbStatus.Visible = False

        Dim status As String
        If (String.IsNullOrEmpty(CType(sender, Control).Tag)) Then
            status = cbStatus.SelectedValue.ToString()
        Else
            status = CType(sender, Control).Tag.ToString()
        End If

        If tvListOfDetails.SelectedNode.Name.IndexOf("Order") <> -1 Then
            For Each node As TreeNode In tvListOfDetails.SelectedNode.Nodes
                If node.Name.IndexOf("Group") Then
                    For Each nodeDetail As TreeNode In node.Nodes
                        registerDetail(nodeDetail, status)
                    Next
                End If
            Next
        ElseIf tvListOfDetails.SelectedNode.Name.IndexOf("Group") <> -1 Then
            For Each node As TreeNode In tvListOfDetails.SelectedNode.Nodes
                If node.Name.IndexOf("Detail") >= 0 Then
                    registerDetail(node, status)
                End If
            Next
        ElseIf tvListOfDetails.SelectedNode.Name.IndexOf("Detail") <> -1 Then
            registerDetail(tvListOfDetails.SelectedNode, status)
        End If

        lbStatus.Visible = True
        setIconGroupTv()
    End Sub

    Private Sub registerDetail(ByRef nodeDetail As TreeNode, ByVal status As String)
        Dim count As Integer
        Dim dtAlreadyCompl As DataTable

        query = "SELECT * FROM OrderMark " + _
                "WHERE idZakaz = " + idZakaz.ToString() + " And idSpZak = " + nodeDetail.Tag.ToString() + _
                    " And idPodr = " + ClassCommon.PODR_COMPL.ToString() + " And idStatus = " + ClassCommon.STATUS_COMPL.ToString() + " And isCompl = 1"
        dtAlreadyCompl = ClassDbWorkBase.FillDataTable(query)

        Dim isCompl As Integer = IIf(dtAlreadyCompl.Rows.Count > 0, 1, 0)

        query = "SELECT count(idSpZak) FROM OrderMark WHERE idZakaz = " + idZakaz.ToString() + " AND idSpZak = " + nodeDetail.Tag.ToString() + " AND idPodr = " + idPodr.ToString()
        count = Convert.ToInt32(ClassDbWorkBase.ExecuteScalar(query))

        If count <> 0 Then
            query = "UPDATE OrderMark SET idStatus = " + status + ", isCompl = " + isCompl.ToString() + ", dateUpdate = Convert (DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "') " + _
                    "WHERE idZakaz = " + idZakaz.ToString() + " AND idSpZak = " + nodeDetail.Tag.ToString() + " AND idPodr = " + idPodr.ToString()
        Else
            query = "INSERT INTO OrderMark (idZakaz, idSpZak, idPodr, idStatus, isCompl, dateCreate, dateUpdate) " + _
                    "VALUES(" + idZakaz.ToString() + "," + nodeDetail.Tag.ToString() + "," + idPodr.ToString() + "," + status + "," + isCompl.ToString() + _
                        ", Convert(DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "'), Convert (DateTime, '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "') )"
        End If

        ClassDbWorkBase.ExecuteNonQuery(query)

        If isCompl = 1 And idPodr <> ClassCommon.PODR_COMPL Then
            If status = ClassCommon.STATUS_READY Then
                nodeDetail.ImageKey = "ilPngCompl"
                nodeDetail.SelectedImageKey = "ilPngCompl"
            ElseIf status = ClassCommon.STATUS_NOT_READY Then
                nodeDetail.ImageKey = "ilPngComplNotWork"
                nodeDetail.SelectedImageKey = "ilPngComplNotWork"
            ElseIf status = ClassCommon.STATUS_PROBLEM Then
                nodeDetail.ImageKey = "ilPngWarning"
                nodeDetail.SelectedImageKey = "ilPngWarning"
            End If
        ElseIf status = ClassCommon.STATUS_NOT_READY Then
            nodeDetail.ImageKey = "ilPngCancel"
            nodeDetail.SelectedImageKey = "ilPngCancel"
        ElseIf status = ClassCommon.STATUS_READY Then
            nodeDetail.ImageKey = "ilPngOk"
            nodeDetail.SelectedImageKey = "ilPngOk"
        ElseIf status = ClassCommon.STATUS_COMPL Then
            nodeDetail.ImageKey = "ilPngCompl"
            nodeDetail.SelectedImageKey = "ilPngCompl"
        ElseIf status = ClassCommon.STATUS_NOT_COMPL Then
            nodeDetail.ImageKey = "ilPngNotCompl"
            nodeDetail.SelectedImageKey = "ilPngNotCompl"
        ElseIf status = ClassCommon.STATUS_PROBLEM Then
            nodeDetail.ImageKey = "ilPngWarning"
            nodeDetail.SelectedImageKey = "ilPngWarning"
        End If
    End Sub

    Private Sub cbDevision_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDevision.SelectedIndexChanged
        If cbDevision.Tag = 1 Then Return
        idPodr = cbDevision.SelectedValue
        cbDateAddWord.Visible = Not (idPodr = ClassCommon.PODR_COMPL)

        TreeViewLoad()
        'Dim thr As New Thread(
        '    Sub()
        listBoxLoadComponents()
        '    End Sub)
        'thr.Start()
    End Sub

    Private Sub btSetStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSetStatusGiveWork.Click, btSetStatusProblem.Click
        btRegistratoinStatus_Click(sender, Nothing)
    End Sub

    Private Sub btSetStatusNotWork_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSetStatusNotWork.Click
        If MessageBox.Show("Вы действительно хотите удалить отметку о передаче в работу?", "Отмена", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            btRegistratoinStatus_Click(sender, Nothing)
        End If
    End Sub

    Private Sub setIsComplToDetail(ByVal node As TreeNode, ByVal isCompl As Integer)
        If (node.Nodes.Count = 0) Then
            If (node.Name.IndexOf("Detail") >= 0) Then
                Dim count As Integer

                query = "UPDATE OrderMark SET isCompl = " + isCompl.ToString() + " WHERE idSpZak = " + node.Tag.ToString()
                ClassDbWorkBase.ExecuteNonQuery(query)

                query = "SELECT count(id) As cnt FROM SteepSpZak WHERE idSpZak = " + node.Tag.ToString()
                count = CType(ClassDbWorkBase.ExecuteScalar(query), Integer)
                If (count = 0) Then
                    If (isCompl <> 0) Then
                        query = "INSERT Into SteepSpZak(idSpZak, Steep3, dataS3, kol3) values(" + _
                                    node.Tag.ToString() + ", " + _
                                    isCompl.ToString() + ", '" + _
                                    Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "', " + _
                                    isCompl.ToString() + ")"
                        ClassDbWorkBase.ExecuteNonQuery(query)
                    End If
                Else
                    query = "UPDATE SteepSpZak SET " + _
                                "Steep3 = " + isCompl.ToString() + ", " + _
                                "dataS3 = '" + Format(Date.Now, "yyyy-MM-dd HH:mm:ss") + "', " + _
                                "kol3 = " + isCompl.ToString() + " " + _
                             "WHERE idSpZak = " + node.Tag.ToString()
                    ClassDbWorkBase.ExecuteNonQuery(query)
                End If
            End If
        Else
            For Each subNode As TreeNode In node.Nodes
                setIsComplToDetail(subNode, isCompl)
            Next
        End If
    End Sub

    Private Sub btSetStatusCompl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSetStatusCompl.Click
        If (IsNothing(tvListOfDetails.SelectedNode)) Then Return

        Dim node As TreeNode = tvListOfDetails.SelectedNode

        btRegistratoinStatus_Click(sender, Nothing)
        setIsComplToDetail(node, 1)
    End Sub

    Private Sub btSetStatusNotCompl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSetStatusNotCompl.Click
        If (IsNothing(tvListOfDetails.SelectedNode)) Then Return

        If (CType(sender, Button).Name = "btSetStatusNotCompl" AndAlso _
            MessageBox.Show("Вы действительно хотите удалить отметку о комплектации?", "Отмена", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> Windows.Forms.DialogResult.Yes) Then
            Return
        End If

        Dim node As TreeNode = tvListOfDetails.SelectedNode

        btRegistratoinStatus_Click(sender, Nothing)
        setIsComplToDetail(node, 0)
    End Sub

    Private Sub btRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRefresh.Click
        TreeViewLoad()
    End Sub

    Private Sub cmsTreeView_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsTreeView.Opening
        If (IsNothing(tvListOfDetails.SelectedNode)) Then Return

        If tvListOfDetails.SelectedNode.Name.IndexOf("Group") <> -1 Then
            tsmiDeleteGroup.Enabled = True
        Else
            tsmiDeleteGroup.Enabled = False
        End If

    End Sub

    Private Sub listBoxLoadComponents()
        Dim tableAll, tableSelected As DataTable
        Dim item As ListViewItem

        'Группы
        query = "SELECT BondKatPodrAndGroupMC.idGroupMC, GroupMC.NameGr " + _
                "FROM BondKatPodrAndGroupMC INNER JOIN GroupMC ON BondKatPodrAndGroupMC.idGroupMC = GroupMC.id " + _
                "WHERE BondKatPodrAndGroupMC.idKatPodr = " + idPodr.ToString() + " AND cMainGr = 0 " + _
                "ORDER BY GroupMC.NameGr"
        tableSelected = ClassDbWorkBase.FillDataTable(query)
        query = "SELECT id As idGroupMC, NameGr FROM GroupMC WHERE KodGr <> 999 AND cMainGr = 0 ORDER BY NameGr"
        tableAll = ClassDbWorkBase.FillDataTable(query)

        Dim existGroup(tableSelected.Rows.Count - 1) As Integer

        Invoke(CType(Me.ParentForm, MainForm).delegat.ClearListView, lvUserGroup)

        For i As Integer = 0 To tableSelected.Rows.Count - 1
            existGroup(i) = tableSelected.Rows(i)("idGroupMC")

            item = New ListViewItem(tableSelected.Rows(i)("idGroupMC").ToString())
            item.SubItems.Add(tableSelected.Rows(i)("NameGr"))
            Invoke(CType(Me.ParentForm, MainForm).delegat.AddListViewItems, lvUserGroup, item)
        Next

        Invoke(CType(Me.ParentForm, MainForm).delegat.ClearListView, lvAllGroup)

        For i As Integer = 0 To tableAll.Rows.Count - 1
            item = New ListViewItem(tableAll.Rows(i)("idGroupMC").ToString())
            item.SubItems.Add(tableAll.Rows(i)("NameGr"))
            If (Array.IndexOf(existGroup, tableAll.Rows(i)("idGroupMC")) >= 0) Then
                item.BackColor = lvSelectedRowColor
            End If
            Invoke(CType(Me.ParentForm, MainForm).delegat.AddListViewItems, lvAllGroup, item)
        Next

        'изделия 
        Invoke(CType(Me.ParentForm, MainForm).delegat.ClearListView, lvUserDetails)

        query = "SELECT ListOfKatMCid FROM KatPodr WHERE id = " + idPodr.ToString()
        Dim listIdDetails As String = ClassDbWorkBase.ExecuteScalar(query)

        Dim existDetails() As String = {}
        If Not String.IsNullOrEmpty(listIdDetails) Then
            existDetails = listIdDetails.Split(",")
            query = "SELECT id, NameMC FROM KatMC WHERE id IN (" + listIdDetails + ")"
            dt = ClassDbWorkBase.FillDataTable(query)

            For i As Integer = 0 To dt.Rows.Count - 1
                item = New ListViewItem(dt.Rows(i)("id").ToString())
                item.SubItems.Add(dt.Rows(i)("NameMC"))
                Invoke(CType(Me.ParentForm, MainForm).delegat.AddListViewItems, lvUserDetails, item)
            Next
        End If
        setColorGroupByDetails()
    End Sub

    Private Sub setColorGroupByDetails()
        query = "SELECT ListOfKatMCid FROM KatPodr WHERE id = " + Invoke(CType(Me.ParentForm, MainForm).delegat.ComboBoxGetSelectedValue, cbDevision).ToString()
        Dim listIdDetails As String = ClassDbWorkBase.ExecuteScalar(query)
        If listIdDetails <> "" Then
            query = "SELECT GroupMC.id, dbo.GroupMC.NameGr " + _
                 "FROM GroupMC INNER JOIN KatMC ON GroupMC.id = KatMC.idGrMC " + _
                 "WHERE KatMC.id in (" + listIdDetails + ")"
            dt = ClassDbWorkBase.FillDataTable(query)

            For j As Integer = 0 To lvAllGroup.Items.Count - 1
                For i As Integer = 0 To dt.Rows.Count - 1
                    If lvAllGroup.Items(j).BackColor <> Color.LimeGreen Then
                        lvAllGroup.Items(j).BackColor = Color.White
                        If lvAllGroup.Items(j).Text = dt.Rows(i)("id") Then
                            lvAllGroup.Items(j).BackColor = Color.Yellow
                            Exit For
                        End If
                    End If
                Next
            Next
        Else
            For j As Integer = 0 To lvAllGroup.Items.Count - 1
                If lvAllGroup.Items(j).BackColor <> Color.LimeGreen Then
                    lvAllGroup.Items(j).BackColor = Color.White
                End If
            Next
        End If

    End Sub

    Private Sub lbDocumentsForOrder_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lbDocumentsForOrder.MouseClick
        If (lbDocumentsForOrder.SelectedItem Is Nothing) Then Return

        PdfViewerDocumentsForOrder.Visible = True
        btClosePDFViewer.Visible = True
        tvListOfDetails.Visible = False
        If lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".pdf") <> -1 Then
            PdfViewerDocumentsForOrder.LoadDocument(Setting.Xml.GetXmlValue("DestinationTo") + lbDocumentsForOrder.SelectedItem.ToString())
        ElseIf lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".xls") <> -1 Then
            ClassCommon.openExcelReadOnly(Setting.Xml.GetXmlValue("DestinationTo") + lbDocumentsForOrder.SelectedItem.ToString())
        End If
    End Sub

    Private Sub btPrintDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPrintDocument.Click
        If lbDocumentsForOrder.SelectedIndex = -1 Then Return
        If lbDocumentsForOrder.SelectedItem.ToString().IndexOf(".pdf") <> -1 Then
            PdfViewerDocumentsForOrder.Print()
        End If
    End Sub

    Private Sub btClosePDFViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClosePDFViewer.Click
        PdfViewerDocumentsForOrder.Visible = False
        btClosePDFViewer.Visible = False
        tvListOfDetails.Visible = True
    End Sub

    Private Sub tvListOfDetails_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvListOfDetails.MouseClick
        lbStatus.Visible = False
    End Sub

    Private Sub setOpenCloseIconForNode()
        tvListOfDetails.Nodes("nodeOrder").ImageKey = IIf(tvListOfDetails.Nodes("nodeOrder").IsExpanded, "ilPngBoxOpened", "ilPngBox")
        tvListOfDetails.Nodes("nodeOrder").SelectedImageKey = IIf(tvListOfDetails.Nodes("nodeOrder").IsExpanded, "ilPngBoxOpened", "ilPngBox")

        For Each node As TreeNode In tvListOfDetails.Nodes("nodeOrder").Nodes
            If (node.Name.IndexOf("Group") >= 0) Then
                Select Case node.ImageKey
                    Case "ilPngBoxRed", "ilPngBoxOpenedRed"
                        node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
                        node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
                    Case "ilPngBoxGreen", "ilPngBoxOpenedGreen"
                        node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
                        node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
                    Case "ilPngBox", "ilPngBoxOpened"
                        node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpened", "ilPngBox")
                        node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpened", "ilPngBox")
                End Select
            End If
        Next
    End Sub

    Private Sub tvListOfDetails_AfterCollapse(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvListOfDetails.AfterCollapse
        setOpenCloseIconForNode()
    End Sub

    Private Sub tvListOfDetails_AfterExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvListOfDetails.AfterExpand
        setOpenCloseIconForNode()
    End Sub

    Private Sub tvListOfDetails_NodeMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvListOfDetails.NodeMouseClick
        Dim dateAddWork = String.Empty, dateCompl As String = String.Empty

        cbDateAddWord.Checked = False
        cbDateCompl.Checked = False
        cbDateCompl.Text = "Комплектация: "
        cbDateAddWord.Text = "Задание отдано: "
        lblInfoDate.Text = String.Empty

        If e.Node.Name.IndexOf("Detail") = -1 Then Return

        query = "SELECT OrderMark.idZakaz, OrderMark.idSpZak,OrderMark.idStatus, OrderMark.isCompl, OrderMark.idPodr, KatPodr.NamePodr, OrderMark.dateUpdate " + _
                "FROM OrderMark INNER JOIN KatPodr ON OrderMark.idPodr = KatPodr.id  " + _
                "WHERE idZakaz =" + idZakaz.ToString() + " AND idSpZak = " + e.Node.Tag.ToString()


        dt = ClassDbWorkBase.FillDataTable(query)

        For i As Integer = 0 To dt.Rows.Count - 1
            If dt.Rows(i)("idStatus") = ClassCommon.STATUS_READY And dt.Rows(i)("idPodr") = idPodr Then
                cbDateAddWord.Checked = True
                cbDateAddWord.Text += dt.Rows(i)("dateUpdate").ToString()
                dateAddWork = dt.Rows(i)("dateUpdate").ToString()
            ElseIf dt.Rows(i)("idStatus") = ClassCommon.STATUS_COMPL And dt.Rows(i)("idPodr") = ClassCommon.PODR_COMPL Then
                cbDateCompl.Checked = True
                cbDateCompl.Text += dt.Rows(i)("dateUpdate").ToString()
                dateCompl = dt.Rows(i)("dateUpdate").ToString()
            End If
        Next
        If dateCompl <> "" And dateAddWork <> "" Then
            If dateCompl < dateAddWork Then
                lblInfoDate.Text = "Дата выдачи задания позже чем дата комплектации!"
            End If
        End If

    End Sub

    Private Sub lvAllGroup_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvAllGroup.MouseDoubleClick
        If (IsNothing(lvAllGroup.SelectedItems)) Then Return

        If lvAllGroup.SelectedItems.Item(0).BackColor = Color.Yellow Then
            Return
        End If

        For i As Integer = 0 To lvUserGroup.Items.Count - 1
            If (lvUserGroup.Items(i).Text = lvAllGroup.SelectedItems(0).Text) Then
                Return
            End If
        Next
        query = "INSERT INTO BondKatPodrAndGroupMC (idKatPodr, idGroupMC) VALUES (" + cbDevision.SelectedValue.ToString() + ", " + lvAllGroup.SelectedItems(0).Text + ")"

        If (ClassDbWorkBase.ExecuteNonQuery(query)) Then
            Dim addItem As New ListViewItem(lvAllGroup.SelectedItems(0).Text)
            addItem.SubItems.Add(lvAllGroup.SelectedItems(0).SubItems(1).Text)
            lvUserGroup.Items.Add(addItem)

            lvUserGroup.Sort()

            lvAllGroup.SelectedItems(0).BackColor = lvSelectedRowColor
            lvAllGroup.SelectedItems(0).Selected = False
            TreeViewLoad()

            RedefinitionCbFilterGroup()
        End If
    End Sub

    Private Sub RedefinitionCbFilterGroup()
        'переопределние cbFilterGroup, показываем все группы, без учета групп выбранных пользователем
        Dim userGroup As String = String.Empty
        For i As Integer = 0 To lvUserGroup.Items.Count - 1
            userGroup += IIf(String.IsNullOrEmpty(userGroup), "", ",") + lvUserGroup.Items(i).Text
        Next
        If Not String.IsNullOrEmpty(userGroup) Then
            query = "SELECT id As idGroupMC, NameGr FROM GroupMC WHERE KodGr <> 999 AND cMainGr = 0 AND NOT id IN (" + userGroup + ") ORDER BY NameGr"
            dt = ClassDbWorkBase.FillDataTable(query)
            cbFilterGroup.DataSource = dt.Copy()
        End If
        If userGroup = "" Then
            query = "SELECT id As idGroupMC, NameGr FROM GroupMC WHERE KodGr <> 999 AND cMainGr = 0 ORDER BY NameGr"
            dt = ClassDbWorkBase.FillDataTable(query)
            cbFilterGroup.DataSource = dt.Copy()
        End If
    End Sub

    Private Sub lvUserGroup_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvUserGroup.MouseDoubleClick
        If (IsNothing(lvUserGroup.SelectedItems)) Then Return
        query = "DELETE FROM BondKatPodrAndGroupMC WHERE idKatPodr = " + cbDevision.SelectedValue.ToString() + " AND idGroupMC = " + lvUserGroup.SelectedItems(0).Text

        If (ClassDbWorkBase.ExecuteNonQuery(query)) Then
            Dim listDeleteGroup As String = String.Empty

            'получаем список idSpZak для удаления записей из OrderMark
            query = "SELECT SpZak.id " + _
                    "FROM KatMC INNER JOIN GroupMC ON KatMC.idGrMC = GroupMC.id INNER JOIN " + _
                        "SpZak ON KatMC.id = SpZak.idMC " + _
                    "WHERE SpZak.idZakaz = " + idZakaz.ToString() + " And GroupMC.id = " + sender.SelectedItems(0).Text
            dt = ClassDbWorkBase.FillDataTable(query)

            For i As Integer = 0 To dt.Rows.Count - 1
                listDeleteGroup += IIf(String.IsNullOrEmpty(listDeleteGroup), "", ",") + dt.Rows(i)("id").ToString()
            Next

            If Not String.IsNullOrEmpty(listDeleteGroup) Then
                query = "DELETE FROM OrderMark WHERE idZakaz = " + idZakaz.ToString() + " AND idPodr = " + idPodr.ToString() + " AND idSpZak IN (" + listDeleteGroup + ")"
                ClassDbWorkBase.ExecuteScalar(query)
            End If

            For Each item As ListViewItem In lvAllGroup.Items
                If (item.Text = lvUserGroup.SelectedItems(0).Text) Then
                    item.BackColor = Color.Transparent
                    Exit For
                End If
            Next
            lvUserGroup.SelectedItems(0).Remove()
            TreeViewLoad()

            RedefinitionCbFilterGroup()
            

        End If
    End Sub

    Private Sub setIconGroupTv()
        If tvListOfDetails.SelectedNode.Name.IndexOf("Order") <> -1 Then
            For Each node As TreeNode In tvListOfDetails.SelectedNode.Nodes
                setIconForNone(node)
            Next
        Else
            If (tvListOfDetails.SelectedNode.Name.IndexOf("Group") <> -1) Then
                setIconForNone(tvListOfDetails.SelectedNode)
            Else
                setIconForNone(tvListOfDetails.SelectedNode.Parent)
            End If
        End If
    End Sub

    Private Sub setIconForNone(ByVal node As TreeNode)
        Dim isWork As Boolean = True
        Dim isCompl As Boolean = True

        For Each nodeDetail As TreeNode In node.Nodes
            isWork = (isWork And IIf(nodeDetail.ImageKey = "ilPngOk", True, False))
            isCompl = (isCompl And IIf(nodeDetail.ImageKey = "ilPngCompl" OrElse nodeDetail.ImageKey = "ilPngComplNotWork", True, False))
        Next
        If idPodr = ClassCommon.PODR_COMPL Then
            If isCompl Then
                node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
                node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
            Else
                node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
                node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
            End If
        Else
            If isCompl Then
                node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
                node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedGreen", "ilPngBoxGreen")
            ElseIf isWork Then
                node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpened", "ilPngBox")
                node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpened", "ilPngBox")
            Else
                node.ImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
                node.SelectedImageKey = IIf(node.IsExpanded, "ilPngBoxOpenedRed", "ilPngBoxRed")
            End If
        End If
    End Sub


    Private Sub tsListGroupsOrDetails_Toggled(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsListGroupsOrDetails.Toggled
        If tsListGroupsOrDetails.EditValue Then
            pnDetails.Visible = True
            tpSetingsGroupsDetails.Text = "Настройка изделий"
        Else
            pnDetails.Visible = False
            tpSetingsGroupsDetails.Text = "Настройка групп"
            setColorGroupByDetails()
        End If
        RedefinitionCbFilterGroup()
    End Sub


    Private Sub colorChangeListViewDetails()
        For i As Integer = 0 To lvUserDetails.Items.Count - 1
            For j As Integer = 0 To lvAllDetails.Items.Count - 1
                If lvUserDetails.Items(i).Text = lvAllDetails.Items(j).Text Then
                    lvAllDetails.Items(j).BackColor = lvSelectedRowColor
                    lvAllDetails.Items(j).Selected = False
                End If
            Next
        Next
    End Sub

    Private Sub lvAllDetails_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvAllDetails.MouseDoubleClick
        If (IsNothing(lvAllDetails.SelectedItems)) Then Return

        For i As Integer = 0 To lvUserDetails.Items.Count - 1
            If (lvUserDetails.Items(i).Text = lvAllDetails.SelectedItems(0).Text) Then
                Return
            End If
        Next

        Dim addItem As New ListViewItem(lvAllDetails.SelectedItems(0).Text)
        addItem.SubItems.Add(lvAllDetails.SelectedItems(0).SubItems(1).Text)
        lvUserDetails.Items.Add(addItem)

        lvUserDetails.Sort()

        Dim listOfKatMCid As String = String.Empty
        For i As Integer = 0 To lvUserDetails.Items.Count - 1
            listOfKatMCid += IIf(String.IsNullOrEmpty(listOfKatMCid), "", ",") + lvUserDetails.Items(i).Text
        Next
        query = "UPDATE KatPodr Set ListOfKatMCid = '" + listOfKatMCid + "' WHERE id = " + cbDevision.SelectedValue.ToString()

        If (ClassDbWorkBase.ExecuteNonQuery(query)) Then
            lvAllDetails.SelectedItems(0).BackColor = lvSelectedRowColor
            lvAllDetails.SelectedItems(0).Selected = False
            TreeViewLoad()
        End If
    End Sub

    Private Sub lvUserDetails_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvUserDetails.MouseDoubleClick
        If (IsNothing(lvUserDetails.SelectedItems)) Then Return

        Dim listOfKatMCid As String = String.Empty
        For i As Integer = 0 To lvUserDetails.Items.Count - 1
            If (lvUserDetails.SelectedItems(0).Text <> lvUserDetails.Items(i).Text) Then
                listOfKatMCid += IIf(String.IsNullOrEmpty(listOfKatMCid), "", ",") + lvUserDetails.Items(i).Text
            End If
        Next

        query = "UPDATE KatPodr Set ListOfKatMCid = '" + listOfKatMCid + "' WHERE id = " + cbDevision.SelectedValue.ToString()

        If (ClassDbWorkBase.ExecuteNonQuery(query)) Then
            Dim listDeleteKatMC As String = String.Empty
            'получаем список idSpZak для удаления записей из OrderMark
            query = "SELECT SpZak.id " + _
                    "FROM KatMC INNER JOIN SpZak ON KatMC.id = SpZak.idMC " + _
                    "WHERE SpZak.idZakaz = " + idZakaz.ToString() + " And KatMC.id = " + lvUserDetails.SelectedItems(0).Text
            dt = ClassDbWorkBase.FillDataTable(query)

            If dt.Rows.Count <> 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    listDeleteKatMC += IIf(String.IsNullOrEmpty(listDeleteKatMC), "", ",") + dt.Rows(i)("id").ToString()
                Next

                query = "DELETE FROM OrderMark WHERE idZakaz = " + idZakaz.ToString() + " AND idPodr = " + idPodr.ToString() + " AND idSpZak IN (" + listDeleteKatMC + ")"
                ClassDbWorkBase.ExecuteScalar(query)
            End If

            For Each item As ListViewItem In lvAllDetails.Items
                If (item.Text = lvUserDetails.SelectedItems(0).Text) Then
                    item.BackColor = Color.Transparent
                    Exit For
                End If
            Next
            lvUserDetails.SelectedItems(0).Remove()
            TreeViewLoad()
        End If
    End Sub

    Private Sub ProductionAreas_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub

    Private Sub cbFilterGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFilterGroup.SelectedIndexChanged
        If cbFilterGroup.Tag = 1 OrElse cbFilterGroup.Items.Count = 0 Then Return

        Dim item As ListViewItem
        query = "SELECT id, NameMC FROM KatMC WHERE idGrMC = " + cbFilterGroup.SelectedValue.ToString()
        dt = ClassDbWorkBase.FillDataTable(query)
        lvAllDetails.Items.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            item = New ListViewItem(dt.Rows(i)("id").ToString())
            item.SubItems.Add(dt.Rows(i)("NameMC"))
            lvAllDetails.Items.Add(item)
        Next

        For Each items As ListViewItem In lvAllDetails.Items
            If (lvUserDetails.Items.Count > 0 AndAlso items.Text = lvUserDetails.Items(0).Text) Then
                items.BackColor = lvSelectedRowColor
                Exit For
            End If
        Next
        colorChangeListViewDetails()
    End Sub
End Class