Imports SergeyDll
Imports System.IO
Imports System.Windows.Forms.ListView

Public Class DrawingsStandard

    Private query As String
    Private dt As DataTable

    Private Sub DrawingsStandard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mainBranch, parentNode, babyNode As New TreeNode
        Dim cod, cod_prev As Integer
        Dim des As String = String.Empty
        Dim codArTtneutri As String = String.Empty

        query = "SELECT TIPOLOGIE.COD, TIPOLOGIE.DES, (case when ARTNEUTRI.COD is not NULL then ARTNEUTRI.COD else '' end) as 'codARTNEUTRI' " + _
                "FROM TIPOLOGIE left JOIN ARTNEUTRI ON TIPOLOGIE.COD = ARTNEUTRI.TIP " + _
                "ORDER BY TIPOLOGIE.COD"
        dt = ClassDbYavid.FillDataTable(query)

        cod = -1
        cod_prev = -1

        For i As Integer = 0 To dt.Rows.Count - 1 'загрузка дерева, вкладка Стандарт
            If (Not Integer.TryParse(dt.Rows(i)("COD").ToString(), cod)) Then Continue For 'проверка на число

            If (dt.Rows(i)("COD") <> cod_prev) Then
                If (cod Mod 100 = 0) Then 'узел 0 - главное изделие
                    mainBranch = tvDS.Nodes.Add(dt.Rows(i)("DES"))
                    parentNode = mainBranch
                Else
                    parentNode = mainBranch.Nodes.Add(dt.Rows(i)("DES"))
                End If
            End If

            If (Not String.IsNullOrEmpty(dt.Rows(i)("codARTNEUTRI"))) Then
                babyNode = parentNode.Nodes.Add(dt.Rows(i)("codARTNEUTRI"))
                babyNode.Tag = 1
            End If

            cod_prev = dt.Rows(i)("COD")
        Next

        tpStandart.Tag = Setting.Xml.GetXmlValue("DrawningKDStandart")
        tpNotice.Tag = Setting.Xml.GetXmlValue("SezamNoticeFiles")
        tpAssemblyDiagram.Tag = Setting.Xml.GetXmlValue("SezamAssemblyDiagramFiles")

        statusStandartLabel.Text = String.Empty
    End Sub

    Private Sub tvDS_NodeMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvDS.NodeMouseClick
        If Not IsNothing(e.Node) AndAlso e.Node.Tag = 1 Then
            Dim buf = e.Node.Text.Split("-")
            Dim listOfFiles As String()
            Dim fileError = "\\sezamdell\KD\STANDART\File_not_found.pdf"
            listOfFiles = Directory.GetFiles(tcDS.SelectedTab.Tag.ToString(), buf(0) + "*.pdf", System.IO.SearchOption.TopDirectoryOnly)
            If listOfFiles.Length > 0 AndAlso File.Exists(listOfFiles(0)) Then
                PdfViewer.LoadDocument(listOfFiles(0))
                statusStandartLabel.Text = "Загружено"
            Else
                statusStandartLabel.Text = "Файл не найден"
                If File.Exists(fileError) Then
                    PdfViewer.LoadDocument(fileError)
                End If
            End If
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim SearchText As String = tbFolderStandartSearch.Text
        Dim buf As Object
        Dim listOfFiles As String()
        Dim fileError = "\\sezamdell\KD\STANDART\File_not_found.pdf"

        If String.IsNullOrEmpty(SearchText) Then Return

        Select Case (tcDS.SelectedTab.Name)
            Case "tpStandart"
                Dim SelectedNode As TreeNode = SearchNode(SearchText, tvDS.Nodes(0))

                If SelectedNode IsNot Nothing Then
                    tvDS.SelectedNode = SelectedNode
                    tvDS.SelectedNode.Expand()
                    tvDS.[Select]()
                    buf = tvDS.SelectedNode.Text.Split("-")
                    listOfFiles = Directory.GetFiles(tcDS.SelectedTab.Tag.ToString(), buf(0) + "*.pdf", System.IO.SearchOption.TopDirectoryOnly)
                    If listOfFiles.Length > 0 AndAlso File.Exists(listOfFiles(0)) Then
                        PdfViewer.LoadDocument(listOfFiles(0))
                        statusStandartLabel.Text = "Загружено"
                    Else
                        statusStandartLabel.Text = "Файл не найден"
                        If File.Exists(fileError) Then
                            PdfViewer.LoadDocument(fileError)
                        End If
                    End If
                End If
            Case "tpNotice", "tpAssemblyDiagram"
                Dim j As Integer = findWrd(SearchText, lvAssemblyDiagramFiles)
                If j < 0 Then Exit Sub
                lvAssemblyDiagramFiles.Focus()
                lvAssemblyDiagramFiles.Items.Item(j).Selected = True
                lvAssemblyDiagramFiles.EnsureVisible(j)

                'buf = tvDS.SelectedNode.Text.Split("-") 
                listOfFiles = Directory.GetFiles(tcDS.SelectedTab.Tag.ToString(), lvAssemblyDiagramFiles.Items(j).Text)
                If listOfFiles.Length > 0 AndAlso File.Exists(listOfFiles(0)) Then
                    PdfViewer.LoadDocument(listOfFiles(0))
                    statusStandartLabel.Text = "Загружено"
                Else
                    statusStandartLabel.Text = "Файл не найден"
                    If File.Exists(fileError) Then
                        PdfViewer.LoadDocument(fileError)
                    End If
                End If
        End Select
    End Sub

    'поиск по вкладке Стандарт
    Private Function SearchNode(ByVal SearchText As String, ByVal StartNode As TreeNode) As TreeNode 'поиск по treeView
        Dim node As TreeNode = Nothing
        While StartNode IsNot Nothing
            If StartNode.Text.ToLower().Contains(SearchText.ToLower()) Then
                node = StartNode 'что-то нашли - выходим
                Exit While
            End If
            If StartNode.Nodes.Count <> 0 Then 'у узла есть дочерние элементы
                node = SearchNode(SearchText, StartNode.Nodes(0)) 'ищем рекурсивно в дочерних
                If node IsNot Nothing Then
                    Exit While 'что-то нашли
                End If
            End If
            StartNode = StartNode.NextNode
        End While
        Return node
    End Function

    'поиск во вкладке извещения и схема сборки
    Private Function findWrd(ByVal sFind As String, ByVal lv As ListView) As Integer
        Dim lvc As New ListViewItemCollection(lv)
        Dim s As String = 0
        Dim jj As Integer = -1
        For Each lvi As ListViewItem In lvc
            s = lvi.Text

            If s.IndexOf(sFind, StringComparison.OrdinalIgnoreCase) <> -1 Then
                jj = lvc.IndexOf(lvi)
                Exit For  'до первого совпадения
            End If
        Next
        Return jj
    End Function

    'по нажатию на Enter срабатывает свойство Button_Click для поиск записи
    Private Sub tbFolderStandart_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbFolderStandartSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnSearch_Click(Nothing, Nothing)
        End If
    End Sub

    Sub Search(ByVal Fol As String, ByVal Node As TreeNode)
        Dim TmpNode As TreeNode

        For Each S As String In IO.Directory.GetDirectories(Fol, "*.*", SearchOption.TopDirectoryOnly)
            TmpNode = New TreeNode(IO.Path.GetFileName(S))
            TmpNode.ImageIndex = 0
            Node.Nodes.Add(TmpNode)

            Search(S, TmpNode)
        Next
    End Sub

    'загрузка файлов в ListBox из папки Схемы Сборки
    Private Sub tvAssemblyDiagramFiles_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvAssemblyDiagramFiles.AfterSelect
        Dim Files() As String = IO.Directory.GetFiles(IO.Path.GetDirectoryName(tcDS.SelectedTab.Tag.ToString()) & "\" & e.Node.FullPath, "*.pdf", SearchOption.TopDirectoryOnly)
        lvAssemblyDiagramFiles.Items.Clear()
        For Each File As String In Files
            lvAssemblyDiagramFiles.Items.Add(IO.Path.GetFileName(File)).Tag = File
            lvAssemblyDiagramFiles.Items(lvAssemblyDiagramFiles.Items.Count - 1).ImageIndex = 1
        Next
    End Sub

    'открытие pdf файла из папки Извещения в окне
    Private Sub lvAssemblyDiagramFiles_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvAssemblyDiagramFiles.MouseClick
        If (lvAssemblyDiagramFiles.SelectedItems.Count = 0) Then Return

        Dim fileError As String = "\\sezamdell\KD\STANDART\File_not_found.pdf"
        Dim fileName As String = tcDS.SelectedTab.Tag.ToString() + "\" + lvAssemblyDiagramFiles.SelectedItems(0).Text

        If File.Exists(fileName) Then
            Try
                PdfViewer.LoadDocument(fileName)
                statusStandartLabel.Text = "Загружено"
            Catch ex As Exception
                statusStandartLabel.Text = "Файл не найден"
                If File.Exists(fileError) Then
                    PdfViewer.LoadDocument(fileError)
                End If
            End Try
        Else
            statusStandartLabel.Text = "Файл не найден"
            If File.Exists(fileError) Then
                PdfViewer.LoadDocument(fileError)
            End If
        End If
    End Sub

    'при переключении вкладок
    Private Sub tcDS_Selecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles tcDS.Selecting
        Select Case (tcDS.SelectedTab.Name)
            Case "tpStandart"
                pnListView.Visible = False
                pnTreeView.Visible = True
                statusStandartLabel.Text = String.Empty
            Case "tpNotice", "tpAssemblyDiagram"
                pnListView.Visible = True
                pnTreeView.Visible = False
                tvAssemblyDiagramFiles.Nodes.Clear()
                lvAssemblyDiagramFiles.Items.Clear()
                statusStandartLabel.Text = String.Empty
                tvAssemblyDiagramFiles.Nodes.Add(IO.Path.GetFileName(tcDS.SelectedTab.Tag.ToString()))
                tvAssemblyDiagramFiles.SelectedNode = tvAssemblyDiagramFiles.Nodes(0)
        End Select
        tbFolderStandartSearch.Text = String.Empty
        statusStandartLabel.Text = String.Empty
        tvDS.CollapseAll()
    End Sub
End Class



