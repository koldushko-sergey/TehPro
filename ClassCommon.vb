Imports System.Xml
Imports System.IO
Imports SergeyDll
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop

Public Class ClassCommon

    Public Const STATUS_NOT_READY = 302
    Public Const STATUS_READY = 304
    Public Const STATUS_COMPL = 305
    Public Const STATUS_NOT_COMPL = 306
    Public Const STATUS_PROBLEM = 307
    Public Const PODR_COMPL = 78

    ''' <summary>
    ''' Копирует файлы в указанную директорию
    ''' </summary>
    ''' <param name="listOfFiles">Список файлов, котрые надо скопировать</param>
    ''' <param name="toDir">Куда копировать</param>
    ''' <remarks></remarks>
    Public Shared Sub CopyFiles(ByVal listOfFiles As String(), ByVal toDir As String, ByVal prefix As String)
        For Each fileName As String In listOfFiles
            Try
                File.Copy(fileName, toDir + prefix + Path.GetFileName(fileName), True)
            Catch ex As Exception
                MessageBox.Show("Ошибка копирования файла " + fileName + ". Код ошибки:" + ex.Message, "Копирование файлов", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next
    End Sub

    '''  <summary>
    ''' Создание и инициализация файла настроек
    ''' </summary>
    Public Shared Function CreateSettingFile()
        Try
            Dim Xml As New XmlTextWriter("Setting.xml", Nothing)
            Xml.Formatting = Formatting.Indented
            Xml.WriteStartDocument(False)
            Xml.WriteStartElement("YavidSetting")
            Xml.WriteComment("Строка подключения")
            Xml.WriteElementString("Yavid", My.Settings.Yavid)
            Xml.WriteElementString("Ecadmaster", My.Settings.Ecadmaster)
            Xml.WriteElementString("WorkBase", My.Settings.WorkBase)
            Xml.Flush()
            Xml.Close()
            Return True
        Catch
            Return False
        End Try
    End Function

    '''  <summary>
    ''' Записывает значение в ini-файл
    ''' </summary>
    ''' <param name="FileName">Полное имя ini-файла файла</param>
    ''' <param name="Section">Секция</param>
    ''' <param name="Key">Параметр, значение которого нужно поменять</param>
    ''' <param name="Value">Новое значение параметра</param>
    Public Shared Function SetIniValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
        Try
            Dim ini As New ClassINI()
            ini.Filename = FileName
            ini.Section = Section
            ini.Key = Key
            ini.Value = Value
            Return (True)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Синхронизирует таблицы на SezamBuh c таблицами на SezamDell
    ''' </summary>
    ''' <param name="tableName">Имя таблицы, которую надо синхронизировать</param>
    ''' <param name="ColumnSearch">Имя колонки уникальной, по которой будут сопоставляться таблицы</param>
    Public Shared Function UpdateTableOn3CadDb(ByVal tableName As String, ByVal ColumnSearch As String, ByVal label As Control) As Boolean
        Dim tblWorkBase As DataTable
        Dim newRow As DataRow
        Dim dr() As DataRow
        Dim iLamda As Integer

        Try
            tblWorkBase = ClassDbWorkBase.FillDataTable("Select * From " + tableName)
            ClassDbEcadmaster.DynamicDataTable("Select * From " + tableName, tableName)
        Catch ex As Exception
            MessageBox.Show("Ошибка выборки таблиц. Код ошибки: " + ex.Message, _
                                "Синхронизация", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try

        Try
            For i As Integer = 0 To tblWorkBase.Rows.Count - 1
                iLamda = i + 1
                label.Invoke(New Action(Sub() label.Text = "Добавление: " + iLamda.ToString() + " из " + tblWorkBase.Rows.Count.ToString()))

                dr = ClassDbEcadmaster.DataSet.Tables(tableName).Select(ColumnSearch + "=" + tblWorkBase.Rows(i)(ColumnSearch).ToString())
                If dr.Count = 1 Then
                    ' Обновление строки
                    For j As Integer = 0 To tblWorkBase.Columns.Count - 1
                        dr(0)(j) = tblWorkBase.Rows(i)(j)
                    Next
                ElseIf dr.Count = 0 Then
                    ' Добавление новой строки
                    newRow = ClassDbEcadmaster.DataSet.Tables(tableName).NewRow()
                    For j As Integer = 0 To tblWorkBase.Columns.Count - 1
                        newRow(j) = tblWorkBase.Rows(i)(j)
                    Next
                    ClassDbEcadmaster.DataSet.Tables(tableName).Rows.Add(newRow)
                Else
                    MessageBox.Show("Неоднозначное соответствие таблиц один ко многим. Значение поиска: " + tblWorkBase.Rows(i)(0).ToString(), _
                                    "Синхронизация", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Next

            For i As Integer = 0 To ClassDbEcadmaster.DataSet.Tables(tableName).Rows.Count - 1
                iLamda = i + 1
                label.Invoke(New Action(Sub() label.Text = "Удаление: " + iLamda.ToString() + " из " + tblWorkBase.Rows.Count.ToString()))

                dr = tblWorkBase.Select(ColumnSearch + "=" + ClassDbEcadmaster.DataSet.Tables(tableName).Rows(i)(ColumnSearch).ToString())
                If dr.Count = 0 Then
                    ClassDbEcadmaster.DataSet.Tables(tableName).Rows(i).Delete()
                End If
            Next

            ClassDbEcadmaster.UpdateDataTable(tableName)

            ClassDbEcadmaster.ExecuteNonQuery("DELETE FROM [Ecadmaster].[dbo].[Sootv] " + _
                                                "WHERE [Ecadmaster].[dbo].[Sootv].idmc in (SELECT [Ecadmaster].[dbo].[Sootv].idmc as 'ID' " + _
                                                              "FROM [Ecadmaster].[dbo].[Sootv] left join [Ecadmaster].[dbo].[KatMC] on " + _
                                                                  "[Ecadmaster].[dbo].[Sootv].idmc = [Ecadmaster].[dbo].[KatMC].id " + _
                                                              "WHERE [Ecadmaster].[dbo].[KatMC].NameMC is null)")
            Return True
        Catch ex As Exception
            MessageBox.Show("Ошибка синхронизации. Код ошибки: " + ex.Message, _
                                    "Синхронизация", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Возвращает соответствующее значение из программы TehPro для пеереданного типа
    ''' </summary>
    ''' <param name="Var">Строка со всеми вариантами</param>
    ''' <param name="NameVariant">Название варианта</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSootvVariant(ByVal Var As String, ByVal NameVariant As String)
        Dim n, lenstr As Integer
        Dim query, value As String

        n = Var.ToUpper().IndexOf(NameVariant + "=") + NameVariant.Length + 1
        If n = (NameVariant.Length) Then
            Return "0"
        Else
            lenstr = Var.IndexOf(";", n) - n
            If lenstr < 0 Then lenstr = Var.Length - n
            value = Var.Substring(n, lenstr)
            query = "Select idmc From Sootv Where Kod='" + NameVariant + ":" + value + "'"
            value = ClassDbEcadmaster.ExecuteScalar(query)
            If String.IsNullOrEmpty(value) Then
                value = Var.Substring(n, lenstr)
                MessageBox.Show("Не задано соответствие для " + NameVariant + ":" + value + ".", "Конвертор заказов", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                value = "-1"
            End If
        End If

        Return value
    End Function


    Public Shared Function GetSootvByValue(ByVal value As String)
        Dim query, result As String

        query = "Select idmc From Sootv Where Kod='" + value + "'"
        result = ClassDbEcadmaster.ExecuteScalar(query)

        Return result
    End Function

    ''' <summary>
    ''' Возвращает значение варианта
    ''' </summary>
    ''' <param name="Var">Строка со всеми вариантами</param>
    ''' <param name="NameVariant">Название варианта</param>
    ''' <param name="DefValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetValueVariant(ByVal var As String, ByVal NameVariant As String, ByVal DefValue As String)
        Dim n, lenstr As Integer
        Dim value As String

        n = var.ToUpper().IndexOf(NameVariant + "=") + NameVariant.Length + 1
        If n = (NameVariant.Length) Then
            Return DefValue
        Else
            lenstr = var.IndexOf(";", n) - n
            If lenstr < 0 Then lenstr = var.Length - n
            value = var.Substring(n, lenstr)
        End If

        Return value
    End Function

    ''' <summary>
    ''' Определение столбца в Excel
    ''' </summary>
    ''' <param name="num">Количество столбцов</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NumToLate(ByVal num As Integer) As String
        Select Case (num)
            Case 0 : Return "Z"
            Case 1 : Return "A"
            Case 2 : Return "B"
            Case 3 : Return "C"
            Case 4 : Return "D"
            Case 5 : Return "E"
            Case 6 : Return "F"
            Case 7 : Return "G"
            Case 8 : Return "H"
            Case 9 : Return "I"
            Case 10 : Return "J"
            Case 11 : Return "K"
            Case 12 : Return "L"
            Case 13 : Return "M"
            Case 14 : Return "N"
            Case 15 : Return "O"
            Case 16 : Return "P"
            Case 17 : Return "Q"
            Case 18 : Return "R"
            Case 19 : Return "S"
            Case 20 : Return "T"
            Case 21 : Return "U"
            Case 22 : Return "V"
            Case 23 : Return "W"
            Case 24 : Return "X"
            Case 25 : Return "Y"
            Case 26 : Return "Z"
            Case Else : Return NumToLate(Math.Truncate(num / 26)) + NumToLate(num - Math.Truncate(num / 26) * 26)
        End Select
    End Function

    ''' <summary>
    ''' Вывод значений из DataGridView в Excel
    ''' </summary>
    ''' <param name="dgv">DataGridView, из которого будут скопированны данные в Excel</param>
    ''' <param name="excel">Экземпляр Excel</param>
    ''' <param name="topLeft">Верхняя левая граница диапазона в Excel, куда будут вставленны данные из DataGridView</param>
    ''' <param name="numberVerticalIndent">Отступ по вертикали</param>
    ''' <param name="skipColumns">Имена колонок, которые не нужно выгружать в Excel. В формате: |name1|name2|name|</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function exportFromDgvToExcel(ByVal dgv As DataGridView, ByVal excel As ExcelDocument, ByVal topLeft As String, ByVal numberVerticalIndent As Integer, ByVal skipColumns As String, ByVal paintExcel As Boolean)
        Try
            Dim colSymbol As Integer = New Regex("[|]").Matches(skipColumns).Count - 1 'количество пропущенных столбцов 
            Dim array(dgv.Rows.Count, dgv.ColumnCount - colSymbol - 1) As Object
            Dim index As Integer = 0

            For i As Integer = 0 To dgv.ColumnCount - 1
                If (skipColumns.IndexOf("|" + dgv.Columns(i).HeaderText + "|") = -1) Then
                    array(0, index) = dgv.Columns(i).HeaderText
                    index += 1
                End If
            Next

            For i As Integer = 0 To dgv.Rows.Count - 1
                index = 0
                For j As Integer = 0 To dgv.ColumnCount - 1
                    If (skipColumns.IndexOf("|" + dgv.Columns(j).HeaderText + "|") = -1) Then
                        If (TypeOf dgv.Rows(i).Cells(j).Value Is Date) Then
                            array(i + 1, index) = dgv.Rows(i).Cells(j).Value.ToString()
                        Else
                            array(i + 1, index) = dgv.Rows(i).Cells(j).Value
                        End If

                        index += 1
                    End If
                Next
            Next

            If (Not excel Is Nothing) Then
                excel.SetRangeCellValue(array, topLeft, NumToLate(dgv.ColumnCount - colSymbol) + (dgv.Rows.Count + numberVerticalIndent).ToString())

                If (paintExcel) Then
                    Dim clr As Color
                    For i As Integer = 0 To dgv.Rows.Count - 1
                        For j As Integer = 0 To dgv.ColumnCount - 1
                            clr = dgv.Rows(i).Cells(j).Style.BackColor
                            If Not clr.IsEmpty Then
                                excel.SetColor(NumToLate(j) + (i + numberVerticalIndent + 1).ToString(), NumToLate(j) + (i + numberVerticalIndent + 1).ToString(), clr)
                            End If
                        Next
                    Next
                End If
            Else
                MessageBox.Show("Excel не инициализирован.", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("Код ошибки: " + ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        Return True
    End Function

    Public Shared Function exportFromDgvToExcel(ByVal dgv As DataGridView, ByVal excel As Object, ByVal topLeft As String, ByVal numberVerticalIndent As Integer)
        Return ClassCommon.exportFromDgvToExcel(dgv, excel, topLeft, numberVerticalIndent, "", False)
    End Function

    Public Shared Function exportFromDgvToExcel(ByVal dgv As DataGridView, ByVal excel As Object, ByVal topLeft As String, ByVal numberVerticalIndent As Integer, ByVal skipColumns As String)
        Return ClassCommon.exportFromDgvToExcel(dgv, excel, topLeft, numberVerticalIndent, skipColumns, False)
    End Function

    Public Shared Function exportFromDgvToExcel(ByVal dgv As DataGridView, ByVal excel As Object, ByVal topLeft As String, ByVal numberVerticalIndent As Integer, ByVal paintExcel As Boolean)
        Return ClassCommon.exportFromDgvToExcel(dgv, excel, topLeft, numberVerticalIndent, "", paintExcel)
    End Function

    Public Shared Function openExcelReadOnly(ByVal name As String) As ExcelDocument
        Dim excel As New ExcelDocument(name, True)
        excel.Visible = True
        Return excel
    End Function
End Class
