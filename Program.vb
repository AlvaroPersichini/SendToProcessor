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
        Dim timestamp As String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = System.IO.Path.Combine(baseDir, "Export_" & timestamp)
        ' Verificamos si la carpeta existe, y si no, la creamos
        If Not IO.Directory.Exists(folderPath) Then
            ' CreateDirectory crea toda la ruta necesaria (incluyendo carpetas padre si no existen)
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
