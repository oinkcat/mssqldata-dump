Imports System.IO
Imports System.Data

''' <summary>
''' Представляет таблицу базы данных
''' </summary>
Public Class SQLTable

    ' Тип столбца таблицы
    Private Enum ColumnType
        Number
        Logical
        Text
        Binary
    End Enum

    Private Const SelectSql As String = "SELECT * FROM {0}"

    Private columns As Dictionary(Of String, ColumnType)

    ''' <summary>
    ''' Имя таблицы данных
    ''' </summary>
    Public Property Name As String

    ''' <summary>
    ''' Добавить в таблицу определение столюца
    ''' </summary>
    ''' <param name="definition">Массив определений столбца</param>
    Public Sub AppendDefinition(definition() As String)
        Const NameIdx As Integer = 1
        Const TypeIdx As Integer = 2

        Dim asBool = {"bit"}
        Dim asString = {"varchar", "nvarchar", "char", "nchar", "text",
                        "datetime", "datetime2", "smalldatetime",
                        "date", "time"}
        Dim asBinary = {"binary", "varbinary"}
        Dim ignorable = {"sysname"}

        Dim defType = definition(TypeIdx)
        If Not ignorable.Contains(defType) Then
            Dim colType As ColumnType = ColumnType.Number
            If asString.Contains(defType) Then
                colType = ColumnType.Text
            ElseIf asBool.Contains(defType) Then
                colType = ColumnType.Logical
            ElseIf asBinary.Contains(defType) Then
                colType = ColumnType.Binary
            End If

            columns.Add(definition(NameIdx), colType)
        End If
    End Sub

    ''' <summary>
    ''' Экспортировать данные таблицы в поток
    ''' </summary>
    ''' <param name="outStream">Поток для экспорта</param>
    Public Sub ExportData(connection As IDbConnection,
                          outStream As Stream,
                          packetSize As Integer)

        Dim writer As New StreamWriter(outStream)
        writer.WriteLine("-- Таблица {0}", Name)
        writer.WriteLine("SET IDENTITY_INSERT [{0}] ON", Name)
        writer.WriteLine()

        Dim fields = String.Join(", ", columns.Keys.Select(
                                 Function(Name) _
                                     String.Concat("[", Name, "]")).ToArray())
        Dim insertStmt = String.Format("INSERT INTO [{0}] ({1}) VALUES",
                                       Name, fields)

        Dim packetStarted As Boolean = False
        Dim rowIdx As Integer = 0

        For Each row In FetchData(connection)
            If Not packetStarted Then
                ' Начало пакета - определение
                writer.WriteLine(insertStmt)
                packetStarted = True
            Else
                ' Далее - выводить запятую для разделения записей
                writer.WriteLine(",")
            End If

            Dim formattedRow = FormatRow(row)
            Dim values = String.Concat("(", String.Join(", ", formattedRow), ")")
            writer.Write(vbTab & values)

            Dim lastInPacket = packetSize > 0 AndAlso rowIdx Mod packetSize = 0
            If lastInPacket Then
                writer.WriteLine()
                If packetSize > 1 Then
                    writer.WriteLine("GO")
                End If
                packetStarted = False
            End If

            rowIdx += 1
        Next

        writer.WriteLine()
        writer.WriteLine()
        writer.WriteLine("SET IDENTITY_INSERT [{0}] OFF", Name)
        writer.WriteLine("GO")
        writer.WriteLine()
        writer.Flush()
    End Sub

    ' Построчно извлечь данные таблицы 
    Private Iterator Function FetchData(conn As IDbConnection) _
        As IEnumerable(Of Object())

        Dim dataCmd = conn.CreateCommand()
        dataCmd.CommandType = CommandType.Text
        dataCmd.CommandText = String.Format(SelectSql, Name)

        Dim reader As IDataReader = dataCmd.ExecuteReader()
        Dim row(reader.FieldCount - 1) As Object
        While reader.Read()
            reader.GetValues(row)
            Yield row
        End While

        reader.Close()
    End Function

    ' Форматировать данные столбца для экспорта
    Private Function FormatRow(row() As Object) As String()
        Dim exportRow(row.Length - 1) As String

        Dim types = columns.Values.ToArray()
        For i = 0 To row.Length - 1
            If Not IsDBNull(row(i)) Then
                If types(i) = ColumnType.Number Then
                    exportRow(i) = row(i)
                ElseIf types(i) = ColumnType.Text Then
                    Dim escaped = row(i).ToString().Replace("'", "''")
                    exportRow(i) = String.Concat("'", escaped, "'")
                ElseIf types(i) = ColumnType.Binary Then
                    Dim byteAttay As Byte() = row(i)
                    Dim hexString = BitConverter.ToString(byteAttay) _
                        .Replace("-", String.Empty)
                    exportRow(i) = String.Concat("0x", hexString)
                Else
                    exportRow(i) = If(row(i) = True, 1, 0)
                End If
            Else
                exportRow(i) = "NULL"
            End If
        Next

        Return exportRow
    End Function

    Public Sub New(name As String)
        Me.Name = name
        columns = New Dictionary(Of String, ColumnType)
    End Sub

End Class
