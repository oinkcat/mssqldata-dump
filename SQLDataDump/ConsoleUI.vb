Imports System.IO

''' <summary>
''' Пользовательский интерфейс CLI
''' </summary>
Module ConsoleUI

    Private Function GetCmdLineParams() As CmdLineParams
        Dim cmdParams As New CmdLineParams()
        Dim cmdLine = Environment.GetCommandLineArgs()

        Dim argIdx = 0
        While argIdx < cmdLine.Length
            Dim param = cmdLine(argIdx).ToLower()
            If param = "-s" Then
                cmdParams.ServerName = cmdLine(argIdx + 1)
                argIdx += 1
            ElseIf param = "-d" Then
                cmdParams.DBName = cmdLine(argIdx + 1)
                argIdx += 1
            ElseIf param = "-o" Then
                cmdParams.OutputFile = cmdLine(argIdx + 1)
                argIdx += 1
            ElseIf param = "-p" Then
                Dim size As Integer
                If Integer.TryParse(cmdLine(argIdx + 1), size) Then
                    cmdParams.PacketSize = If(size > 0, size, 0)
                End If
            ElseIf param = "-t" Then
                cmdParams.TablesToInclude = cmdLine(argIdx + 1).Split(","c)
            End If
            argIdx += 1
        End While

        Return cmdParams
    End Function

    Private Sub TableExporting(sender As Object, e As TableExportEventArgs)
        Console.WriteLine("Таблица {0}...", e.TableName)
    End Sub

    Sub Main()
        Dim cmdLineParams = GetCmdLineParams()

        Console.WriteLine("Соединение с {0}\{1}...",
                          cmdLineParams.ServerName,
                          cmdLineParams.DBName)
        Using database As New SQLDatabase(cmdLineParams.ServerName,
                                          cmdLineParams.DBName),
            outStream As New FileStream(cmdLineParams.OutputFile,
                                        FileMode.Create)
            database.IncludeTables = cmdLineParams.TablesToInclude

            database.DataPacketSize = cmdLineParams.PacketSize
            AddHandler database.TableExporting, AddressOf TableExporting
            database.Connect()
            Console.WriteLine("OK")
            Console.WriteLine("Анализ таблиц базы данных...")
            database.AnalyzeTables()
            Console.WriteLine("Экспорт данных")
            database.Export(outStream)
            Console.WriteLine("Завершено")
        End Using
    End Sub

End Module
