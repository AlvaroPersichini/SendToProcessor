Option Explicit On
Option Strict On
Module Program

    Sub Main()

        ' CATIA Session
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            MsgBox(session.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False


        ' Directorios
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
            Console.WriteLine("Carpeta creada: " & folderPath)
        Else
            Console.WriteLine("La carpeta ya existe: " & folderPath)
        End If


        ' Ejecutar eSendTo
        Dim stProcessor As New SendToProcessor(session.Application)
        stProcessor.Execute(session.RootProduct, folderPath)


        ' Limpieza COM
        Dim cleaner As New ComCleaner()
        cleaner.Release(oProduct, session.Application)


    End Sub

End Module
