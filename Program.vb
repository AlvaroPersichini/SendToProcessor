Option Explicit On
Option Strict On
Module Program

    Sub Main()

        ' --- 1. CONEXIÓN Y VALIDACIÓN ---
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            MsgBox(session.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False


        ' --- GESTIÓN DE DIRECTORIOS ---
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



        ' -------------------
        ' SendToWithPartN
        ' -------------------
        ' 3. Ejecutar el Manager
        Dim stProcessor As New SendToProcessor(session.Application)
        stProcessor.Execute(session.RootProduct, folderPath)


    End Sub


End Module
