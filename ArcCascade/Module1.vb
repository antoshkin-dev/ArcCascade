Imports System.IO
Imports System.Net.Mime.MediaTypeNames

Module Module1
    Private ParamList As New List(Of Params) From {
            New Params("pf ", False),
            New Params("sf ", False, "y"),
            New Params("fm ", False),
            New Params("delete", False),
            New Params("view", False),
            New Params("?", False),
            New Params("help", False)
        }
    Dim WorkPath As String
    Dim FileMask As String
    Dim SubFolders As Boolean = True
    Dim ViewOnly As Boolean
    Dim DeleteAfterArc As Boolean = False
    Dim FoundCount As Integer = 0
    Dim FilesSize As Long = 0
    Dim FilesSizeAfter As Long = 0
    Dim StartupPath As String = AppDomain.CurrentDomain.BaseDirectory
    Sub Main()
        Dim msg As String = "
****************************************
       Каскадный архиватор файлов
****************************************
"
        AddLog(msg)

        Dim TMass() As String = Split(Command.ToString, "/")

        ' Получаем значения заданных параметров из адресной строки
        For Each CMD As String In TMass
            For Each Param As Params In ParamList
                If (LCase(Mid(CMD, 1, Len(Param.Name))) = Param.Name) Then
                    Param.Value = Trim(Mid(CMD, Len(Param.Name) + 1))
                End If
            Next
        Next

        'Отображаем помощь, если у нас есть в этом потребность
        Dim FND As Params
        FND = ParamList.Find(Function(x) x.Name = "?")
        If FND.IsSet = True Then ShowHelpAndEnd()

        FND = ParamList.Find(Function(x) x.Name = "help")
        If FND.IsSet = True Then ShowHelpAndEnd()

        'Проверяем полученные параметры и что все параметры указаны в нужном формате
        Dim WasErr As Boolean = False
        For Each Result As Params In ParamList
            If Result.Reqired = True And Result.IsSet = False Then
                WasErr = True
                AddLog("Отсутствует обязательный параметр '" & RTrim(Result.Name) & "'")
            End If
        Next
        If WasErr = True Then ShowHelpAndEnd()

        REM Определяем параметры
        'Каталог, в котором будем работать
        FND = ParamList.Find(Function(x) x.Name = "pf ")
        If FND.IsSet = False Then 'используем путь запуска приложения
            WorkPath = StartupPath
        Else
            WorkPath = FND.Value.Replace(Chr(34), "") 'Убираем ковычки, если они были
            If IO.Directory.Exists(WorkPath) = False Then
                AddLog("Не найден каталог для обработки: " & WorkPath)
                End
            End If
        End If

        'Маска файлов
        FND = ParamList.Find(Function(x) x.Name = "fm ")
        If FND.IsSet = False Or FND.Value = "" Then
            'используем путь запуска приложения
            FileMask = "*.*"
        Else
            FileMask = FND.Value
        End If

        'Обработка подкаталогов
        FND = ParamList.Find(Function(x) x.Name = "sf ")
        If LCase(FND.Value) = "n" Then SubFolders = False

        'Режим удаления файлов после архивации
        FND = ParamList.Find(Function(x) x.Name = "delete")
        If FND.IsSet Then
            DeleteAfterArc = True
        End If

        'Режим просмотра списка
        FND = ParamList.Find(Function(x) x.Name = "view")
        If FND.IsSet Then
            ViewOnly = True
        End If

        'Выводим сводную информацию по параметрам
        msg = "
***************
   Параметры
***************
Исходный каталог: {0}
Путь для архивации: {1}
Маска поиска: {2}
Поиск в подкаталогах: {3}
Удалять файлы после упаковки: {4}
Режим просмотра: {5}
***************
"
        AddLog(String.Format(msg, StartupPath, WorkPath, FileMask, SubFolders.ToString, DeleteAfterArc.ToString, ViewOnly.ToString))

        'Всё ок, уходим в процедуру архивации
        ProcessFolder(WorkPath)

        'Итог операции
        Dim ProfitSize As Long = Fix((FilesSize - FilesSizeAfter) / 1048576)
        Dim ProfitText As String = IIf(DeleteAfterArc, "Освобождено: " & ProfitSize & "Mb", "")

        msg = "
Процесс архивации завершен.
Найдено файлов, соответствующих фильтру: {0}
Исходный размер файлов: {1}Mb
Размер файлов после упаковки: {2}Mb
{3}"
        AddLog(String.Format(msg, Fix(FoundCount), Fix(FilesSize / 1048576), Fix(FilesSizeAfter / 1048576), ProfitText))
    End Sub
    Private Sub ShowHelpAndEnd()
        Dim msg As String = "
Утилита ArcCascade предназначена для упаковки файлов в индивидуальные архивы.
Параметры запуска: 
/pf <путь> - process folder, путь к папке, в которой будет производиться операция архивирования.
    Если параметр не указывать, обработка будет проводиться в каталоге с утилитой.
/fm *.* - file mask, маска отбора по типу файлов. По умолчанию фильтр файлов: *.*
/sf [y|n] - subfolders, производить обработку в подкаталогах. По умолчанию - поиск в подкаталогах включен.
/delete - Ключ удаления исходный файлов после упаковки.
/view - только отображение информации без выполнения команд архивации. Для отладки параметров запуска."
        AddLog(msg)
        Console.ReadKey()
        End
    End Sub

    Private Sub ProcessFolder(ByVal Path As String)
        ' Перебирваем файлы в данной папке
        Dim Files() As String
        Try
            Files = Directory.GetFiles(Path, FileMask)
        Catch ex As Exception
            AddLog("Ошибка открытия каталога " & Path)
            Exit Sub
        End Try
        For Each F As String In Files
            'Расширения файла не должно соответствовать файлу архива
            If Right(F, 3) = "rar" Then Continue For

            'Получаем информация о файле
            Dim Info As New FileInfo(F)
            FoundCount += 1
            FilesSize += Info.Length

            'Процесс упаковки файла
            ArcFile(F)
        Next


        If SubFolders = False Then Exit Sub ' Обрабатываем файлы только в данной папке 
        Dim Folders() As String = Directory.GetDirectories(Path)
        For Each Folder As String In Folders
            ProcessFolder(Folder)
        Next
    End Sub

    Private Sub ArcFile(ByVal FileForArc As String)
        Dim RunRar As String
        '-md4g
        Dim DestinationFile As String = FileForArc & ".rar"
        Dim DeleteKey As String = IIf(DeleteAfterArc = True, " -df ", "") 'ключ удаления исходного файла
        RunRar = StartupPath & "\rar.exe  a -m3 -ep -y -inul " & DeleteKey & Chr(34) & DestinationFile & Chr(34) & " " & Chr(34) & FileForArc & Chr(34)
        AddLog("RARing " & FileForArc)
        If ViewOnly = False Then
            Call Shell(RunRar, AppWinStyle.Hide, True)
            'Прибавляем размер запакованого файла к итоговой сумме
            Try
                If (File.Exists(DestinationFile)) Then
                    Dim Info As New FileInfo(DestinationFile)
                    FilesSizeAfter += Info.Length
                End If
            Catch
            End Try
        End If
    End Sub
    Private Sub AddLog(TX As String)
        Dim MSG As String = Now & ": " & TX
        Console.WriteLine(MSG)
        Debug.Print(MSG)
    End Sub

End Module
