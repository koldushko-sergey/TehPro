Public Class ClassDelegateForThread
    'Устанавливаем Tag в Control
    Public Delegate Sub TControlTagSet(ByVal isControl As Control, ByVal tag As Object)
    Public ControlTagSet As New TControlTagSet(AddressOf subControlTagSet)
    'Изменение Text в Control
    Public Delegate Sub TControlSetText(ByVal _control As Control, ByVal text As String)
    Public ControlSetText As New TControlSetText(AddressOf subControlSetText)
    'Добавление контрола на форму (по центру и по верх всех контролов)
    Public Delegate Sub TAddControl(ByVal frmParent As Form, ByVal newControl As Control)
    Public AddControl As TAddControl = New TAddControl(AddressOf subAddControl)
    'Изменение видимости контрола 
    Public Delegate Sub TControlVisible(ByVal control As Control, ByVal isVisible As Boolean)
    Public ControlVisible As New TControlVisible(AddressOf subControlVisible)
    'Изменяем свойство ReadOnly для колонки
    Public Delegate Sub TDataGridViewColumnReadOnly(ByVal Column As DataGridViewColumn, ByVal isReadOnly As Boolean)
    Public SetDataGridViewColumnReadOnly As New TDataGridViewColumnReadOnly(AddressOf subDataGridViewColumnReadOnly)
    'Изменяем свойство ReadOnly для ячейки
    Public Delegate Sub TDataGridViewCellReadOnly(ByVal Cell As DataGridViewCell, ByVal isReadOnly As Boolean)
    Public SetDataGridViewCellReadOnly As New TDataGridViewCellReadOnly(AddressOf subDataGridViewCellReadOnly)
    'Очистка столбцов и строк в DataGridView 
    Public Delegate Sub TDataGridViewClear(ByVal grid As DataGridView)
    Public DataGridViewClear As New TDataGridViewClear(AddressOf subDataGridViewClear)
    'Очистка строк в DataGridView 
    Public Delegate Sub TDataGridViewRowClear(ByVal grid As DataGridView)
    Public DataGridViewRowClear As New TDataGridViewRowClear(AddressOf subDataGridViewRowClear)
    'Добавление колонки в DataGridView
    Public Delegate Sub TAddColumnToDataGridView(ByVal grid As DataGridView, ByVal column As DataGridViewColumn)
    Public AddColumnToDataGridView As New TAddColumnToDataGridView(AddressOf subAddColumnToDataGridView)
    'Добавление строки в DataGridView
    Public Delegate Sub TAddRowToDataGridView(ByVal grid As DataGridView, ByVal row As DataGridViewRow)
    Public AddRowToDataGridView As New TAddRowToDataGridView(AddressOf subAddRowToDataGridView)
    'Изменение шрифта в ячейке DataGridView
    Public Delegate Sub TSetFontCell(ByVal _cell As DataGridViewCell, ByVal _font As Font)
    Public SetFontCell As New TSetFontCell(AddressOf subSetFontCell)
    'Изменение ширины колонки DataGridView
    Public Delegate Sub TSetWidthColumn(ByVal column As DataGridViewColumn, ByVal _width As Integer)
    Public SetWidthColumn As New TSetWidthColumn(AddressOf subSetWidthColumn)
    'Изменение цвета ячейки DataGridView
    Public Delegate Sub TSetColorCell(ByVal cell As DataGridViewCell, ByVal _color As Color)
    Public SetColorCell As New TSetColorCell(AddressOf subSetColorCell)
    'Изменение цвета строки в DataGridView
    Public Delegate Sub TSetColorRow(ByVal row As DataGridViewRow, ByVal _color As Color)
    Public SetColorRow As New TSetColorRow(AddressOf subSetColorRow)
    'Присваиваем DataGridView.DataSource значение
    Public Delegate Sub TSetDataSource(ByVal grid As DataGridView, ByVal table As DataTable)
    Public SetDataSource As New TSetDataSource(AddressOf subSetDataSource)
    'Изменяем видимость колонки в DataGridView
    Public Delegate Sub TSetVisibleColumn(ByVal column As DataGridViewColumn, ByVal value As Boolean)
    Public SetVisibleColumn As New TSetVisibleColumn(AddressOf subSetVisibleColumn)
    'Установить значение в DataGridViewCell
    Public Delegate Sub TSetValueDataGridViewCell(ByVal cell As DataGridViewCell, ByVal value As Object)
    Public SetValueDataGridViewCell As New TSetValueDataGridViewCell(AddressOf subSetValueDataGridViewCell)
    'Изменение текста в ToolStripStatusLabel
    Public Delegate Sub TSetToolStripStatusLabelText(ByVal StatusLabel As ToolStripStatusLabel, ByVal text As String)
    Public SetToolStripStatusLabelText As New TSetToolStripStatusLabelText(AddressOf subSetToolStripStatusLabelText)
    'Выделение ячейки в DataGridView
    Public Delegate Sub TSelectCellInDataGridView(ByVal grid As DataGridView, ByVal cell As DataGridViewCell)
    Public SelectCellInDataGridView As New TSelectCellInDataGridView(AddressOf subSelectCellInDataGridView)
    'Устанавливаем активный контрол на форме
    Public Delegate Sub TSetActiveControl(ByVal form As Form, ByVal _control As Control)
    Public SetActiveControl As New TSetActiveControl(AddressOf subSetActiveControl)
    'Получение текста из ListView
    Public Delegate Function TGetTextFromListView(ByVal list As ListView, ByVal itemNumber As Integer)
    Public GetTextFromListView As New TGetTextFromListView(AddressOf subGetTextFromListView)
    'Получет индекс первого выделенного item-а в ListView
    Public Delegate Function TGetFirstSelectedItemFromListView(ByVal list As ListView)
    Public GetFirstSelectedItemFromListView As New TGetFirstSelectedItemFromListView(AddressOf subGetFirstSelectedItemFromListView)
    'Выделение item-а в ListView
    Public Delegate Sub TSetActiveItemInListView(ByVal list As ListView, ByVal itemNumber As Integer)
    Public SetActiveItemInListView As New TSetActiveItemInListView(AddressOf subSetActiveItemInListView)
    'Устанавливаем максимальное значение для ToolStripProgressBar
    Public Delegate Sub TSetMaximumToolStripProgressBar(ByVal ProgresBar As ToolStripProgressBar, ByVal value As Integer)
    Public SetMaximumToolStripProgressBar As New TSetMaximumToolStripProgressBar(AddressOf subSetMaximumToolStripProgressBar)
    'Устанавливаем значение для ToolStripProgressBar
    Public Delegate Sub TSetValueToolStripProgressBar(ByVal ProgresBar As ToolStripProgressBar, ByVal value As Integer)
    Public SetValueToolStripProgressBar As New TSetValueToolStripProgressBar(AddressOf subSetValueToolStripProgressBar)
    'Добавляем Node в TreeView
    Public Delegate Function TTreeViewAddNode(ByVal tree As TreeView, ByVal keyName As String, ByVal textNode As String, ByVal imgIndex As Integer, ByVal selImgIndex As Integer)
    Public TreeViewAddNode As New TTreeViewAddNode(AddressOf subTreeViewAddNode)
    'Добавляем дочерний Node 
    Public Delegate Function TTreeViewAddNodeInNode(ByVal node As TreeNode, ByVal keyName As String, ByVal textNode As String, ByVal imgIndex As Integer, ByVal selImgIndex As Integer)
    Public TreeViewAddNodeInNode As New TTreeViewAddNodeInNode(AddressOf subTreeViewAddNodeInNode)
    'Очищаем TreeView
    Public Delegate Sub TTreeViewNodesClear(ByVal tree As TreeView)
    Public TreeViewNodesClear As New TTreeViewNodesClear(AddressOf subTreeViewNodesClear)
    'Получаем текст Node-a
    Public Delegate Function TTreeViewGetTextNode(ByVal node As TreeNode)
    Public TreeViewGetTextNode As New TTreeViewGetTextNode(AddressOf subTreeViewGetTextNode)
    'Возвращает SelectedNode для TreeView
    Public Delegate Function TTreeViewGetSelectedNode(ByVal tree As TreeView)
    Public TreeViewGetSelectedNode As New TTreeViewGetSelectedNode(AddressOf subTreeViewGetSelectedNode)
    'Очищаем items-ы в ComboBox-е
    Public Delegate Sub TComboBoxItemsClear(ByVal Box As ComboBox)
    Public ComboBoxItemsClear As New TComboBoxItemsClear(AddressOf subComboBoxItemsClear)
    'Устанавливаем DateSource в ComboBox
    Public Delegate Sub TComboBoxSetDataSource(ByVal Box As ComboBox, ByVal table As DataTable)
    Public ComboBoxSetDataSource As New TComboBoxSetDataSource(AddressOf subComboBoxSetDataSource)
    'Устанавливаем SelectIndex в ComboBox
    Public Delegate Sub TComboBoxSetSelectIndex(ByVal box As ComboBox, ByVal index As Integer)
    Public ComboBoxSetSelectIndex As New TComboBoxSetSelectIndex(AddressOf subComboBoxSetSelectIndex)
    'Устанавливаем DisplayMember в ComboBox
    Public Delegate Sub TComboBoxSetDisplayMember(ByVal box As ComboBox, ByVal member As String)
    Public ComboBoxSetDisplayMember As New TComboBoxSetDisplayMember(AddressOf subComboBoxSetDisplayMember)
    'Получаем SelectIndex из ComboBox
    Public Delegate Function TComboBoxGetSelectedIndex(ByVal box As ComboBox)
    Public ComboBoxGetSelectedIndex As New TComboBoxGetSelectedIndex(AddressOf subComboBoxGetSelectedIndex)
    'Получаем DataSourceValue из ComboBox
    Public Delegate Function TComboBoxGetDataSourceValue(ByVal box As ComboBox, ByVal column As String)
    Public ComboBoxGetDataSourceValue As New TComboBoxGetDataSourceValue(AddressOf subComboBoxGetDataSourceValue)
    'Получает SelectedItem из ComboBox
    Public Delegate Function TComboBoxGetSelectedItem(ByVal box As ComboBox)
    Public ComboBoxGetSelectedItem As New TComboBoxGetSelectedItem(AddressOf subComboBoxGetSelectedItem)
    'Получает SelectedValue из ComboBox
    Public Delegate Function TComboBoxGetSelectedText(ByVal box As ComboBox)
    Public ComboBoxGetSelectedText As New TComboBoxGetSelectedItem(AddressOf subComboBoxGetSelectedText) 'subComboBoxGetSelectedValue AND Return box.Text.ToString()  30.05.2019

    Public Delegate Function TComboBoxGetSelectedValue(ByVal box As ComboBox)
    Public ComboBoxGetSelectedValue As New TComboBoxGetSelectedItem(AddressOf subComboBoxGetSelectedValue)

    'Устанавливаем Tag в TreeView Node
    Public Delegate Sub TTreeViewSetTagInNode(ByVal node As TreeNode, ByVal tag As Object)
    Public TreeViewSetTagInNode As New TTreeViewSetTagInNode(AddressOf subTreeViewSetTagInNode)

    'Очищение ListView
    Public Delegate Sub TClearListView(ByVal lv As ListView)
    Public ClearListView As New TClearListView(AddressOf subClearListView)
    'Добавление Item ListView
    Public Delegate Sub TAddListViewItems(ByVal lv As ListView, ByVal item As ListViewItem)
    Public AddListViewItems As New TAddListViewItems(AddressOf subAddListViewItems)
    'Получение Item из ListView
    Public Delegate Function TGetListViewItem(ByVal lv As ListView, ByVal index As Integer)
    Public GetListViewItem As New TGetListViewItem(AddressOf subGetListViewItem)

    Private Sub subAddControl(ByVal frmParent As Form, ByVal newControl As Control)
        newControl.Parent = frmParent
        newControl.Location = New Point(Math.Round(frmParent.Size.Width / 2 - newControl.Size.Width / 2), _
                                        Math.Round(frmParent.Size.Height / 2 - newControl.Size.Height / 2))
        newControl.Anchor = AnchorStyles.Bottom And AnchorStyles.Left
        frmParent.Controls.Add(newControl)
        newControl.BringToFront()
    End Sub

    Private Sub subControlVisible(ByVal _control As Control, ByVal isVisible As Boolean)
        _control.Visible = isVisible
    End Sub

    Private Sub subDataGridViewClear(ByVal grid As DataGridView)
        grid.Rows.Clear()
        grid.Columns.Clear()
    End Sub

    Private Sub subAddColumnToDataGridView(ByVal grid As DataGridView, ByVal column As DataGridViewColumn)
        grid.Columns.Add(column)
    End Sub

    Private Sub subAddRowToDataGridView(ByVal grid As DataGridView, ByVal row As DataGridViewRow)
        grid.Rows.Add(row)
    End Sub

    Private Sub subSetFontCell(ByVal _cell As DataGridViewCell, ByVal _font As Font)
        _cell.Style.Font = _font
    End Sub

    Private Sub subSetWidthColumn(ByVal column As DataGridViewColumn, ByVal _width As Integer)
        column.Width = _width
    End Sub

    Private Sub subSetColorCell(ByVal cell As DataGridViewCell, ByVal _color As Color)
        cell.Style.BackColor = _color
    End Sub
    Private Sub subSetColorRow(ByVal row As DataGridViewRow, ByVal _color As Color)
        row.DefaultCellStyle.BackColor = _color
    End Sub

    Private Sub subControlSetText(ByVal _control As Control, ByVal text As String)
        _control.Text = text
    End Sub

    Private Sub subSelectCellInDataGridView(ByVal grid As DataGridView, ByVal cell As DataGridViewCell)
        grid.CurrentCell = cell
    End Sub

    Private Sub subSetActiveControl(ByVal form As Form, ByVal _control As Control)
        form.ActiveControl = _control
    End Sub

    Private Function subGetTextFromListView(ByVal list As ListView, ByVal itemNumber As Integer)
        Return list.Items(itemNumber).Text
    End Function

    Private Function subGetFirstSelectedItemFromListView(ByVal list As ListView)
        If list.SelectedItems.Count > 0 Then
            Return list.SelectedItems(0).Index
        Else
            Return 0
        End If
    End Function

    Private Sub subSetActiveItemInListView(ByVal list As ListView, ByVal itemNumber As Integer)
        list.Items(itemNumber).Selected = True
        list.Items(itemNumber).EnsureVisible()
    End Sub

    Private Sub subSetVisibleColumn(ByVal column As DataGridViewColumn, ByVal value As Boolean)
        column.Visible = value
    End Sub

    Private Sub subDataGridViewRowClear(ByVal grid As DataGridView)
        grid.Rows.Clear()
    End Sub

    Private Sub subSetToolStripStatusLabelText(ByVal StatusLabel As ToolStripStatusLabel, ByVal text As String)
        StatusLabel.Text = text
    End Sub

    Private Sub subSetMaximumToolStripProgressBar(ByVal ProgresBar As ToolStripProgressBar, ByVal value As Integer)
        ProgresBar.Maximum = value
    End Sub

    Private Sub subSetValueToolStripProgressBar(ByVal ProgresBar As ToolStripProgressBar, ByVal value As Integer)
        ProgresBar.Value = value
    End Sub

    Private Sub subDataGridViewColumnReadOnly(ByVal Column As DataGridViewColumn, ByVal isReadOnly As Boolean)
        Column.ReadOnly = isReadOnly
    End Sub

    Private Sub subDataGridViewCellReadOnly(ByVal Cell As DataGridViewCell, ByVal isReadOnly As Boolean)
        Cell.ReadOnly = isReadOnly
    End Sub

    Private Function subTreeViewAddNode(ByVal tree As TreeView, ByVal keyName As String, ByVal textNode As String, ByVal imgIndex As Integer, ByVal selImgIndex As Integer)
        Dim n As TreeNode
        If String.IsNullOrEmpty(keyName) Then
            n = tree.Nodes.Add(textNode)
        Else
            n = tree.Nodes.Add(keyName, textNode, imgIndex, selImgIndex)
        End If
        Return n
    End Function

    Private Function subTreeViewAddNodeInNode(ByVal node As TreeNode, ByVal keyName As String, ByVal textNode As String, ByVal imgIndex As Integer, ByVal selImgIndex As Integer)
        Dim n As TreeNode
        If String.IsNullOrEmpty(keyName) Then
            n = node.Nodes.Add(textNode)
        Else
            n = node.Nodes.Add(keyName, textNode, imgIndex, selImgIndex)
        End If
        Return n
    End Function

    Private Sub subTreeViewNodesClear(ByVal tree As TreeView)
        tree.Nodes.Clear()
    End Sub

    Private Sub subComboBoxItemsClear(ByVal Box As ComboBox)
        Box.Items.Clear()
    End Sub

    Private Sub subComboBoxSetDataSource(ByVal Box As ComboBox, ByVal table As DataTable)
        Box.DataSource = table
    End Sub

    Private Sub subSetDataSource(ByVal grid As DataGridView, ByVal table As DataTable)
        grid.DataSource = table
    End Sub

    Private Sub subSetValueDataGridViewCell(ByVal cell As DataGridViewCell, ByVal value As Object)
        cell.Value = value
    End Sub

    Private Function subTreeViewGetTextNode(ByVal node As TreeNode)
        Return node.Text
    End Function

    Private Function subTreeViewGetSelectedNode(ByVal tree As TreeView)
        If (tree.SelectedNode Is Nothing) Then
            Return Nothing
        Else
            Return tree.SelectedNode
        End If
    End Function

    Private Sub subComboBoxSetSelectIndex(ByVal box As ComboBox, ByVal index As Integer)
        box.SelectedIndex = index
    End Sub

    Private Sub subComboBoxSetDisplayMember(ByVal box As ComboBox, ByVal member As String)
        box.DisplayMember = member
    End Sub

    Private Sub subControlTagSet(ByVal isControl As Control, ByVal tag As Object)
        isControl.Tag = tag
    End Sub

    Private Sub subTreeViewSetTagInNode(ByVal node As TreeNode, ByVal tag As Object)
        node.Tag = tag
    End Sub

    Private Function subComboBoxGetSelectedIndex(ByVal box As ComboBox)
        Return box.SelectedIndex
    End Function

    Private Function subComboBoxGetDataSourceValue(ByVal box As ComboBox, ByVal column As String)
        Return DirectCast(box.Items(box.SelectedIndex), DataRowView)(column).ToString()
    End Function

    Private Function subComboBoxGetSelectedItem(ByVal box As ComboBox)
        Return box.SelectedItem.ToString()
    End Function

    Private Function subComboBoxGetSelectedText(ByVal box As ComboBox)
        Return box.Text.ToString()
    End Function

    Private Function subComboBoxGetSelectedValue(ByVal box As ComboBox)
        Return box.SelectedValue
    End Function

    Private Function subGetListViewItem(ByVal lv As ListView, ByVal index As Integer)
        Return lv.Items(index)
    End Function

    Private Sub subClearListView(ByVal lv As ListView)
        lv.Items.Clear()
    End Sub

    Private Sub subAddListViewItems(ByVal lv As ListView, ByVal item As ListViewItem)
        lv.Items.Add(item)
    End Sub

End Class
