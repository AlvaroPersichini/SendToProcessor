Option Explicit On
Option Strict On


Public Class SendToProcessor
    Private ReadOnly _app As INFITF.Application

    Public Sub New(catiaApp As INFITF.Application)
        _app = catiaApp
    End Sub

    '
    Public Sub Execute(oProduct As ProductStructureTypeLib.Product, strDir As String)

        Dim oDic1 As Specialized.StringDictionary = GetMap(oProduct)
        Dim oProductDocument = CType(oProduct.ReferenceProduct.Parent, ProductStructureTypeLib.ProductDocument)
        SendTOWPN(oProductDocument, oDic1, strDir)

    End Sub


    Private Function GetMap(objRoot As ProductStructureTypeLib.Product) As Specialized.StringDictionary
        Dim dicc As New Specialized.StringDictionary()
        FillMap(objRoot, dicc)
        Return dicc
    End Function

    Private Sub FillMap(current As ProductStructureTypeLib.Product, ByRef dicc As Specialized.StringDictionary)
        Try
            Dim parentObj As Object = current.ReferenceProduct.Parent
            Dim docName As String = CType(parentObj, INFITF.Document).Name
            Dim pn As String = current.PartNumber

            If Not dicc.ContainsKey(docName) Then
                dicc.Add(docName, pn)
            End If

            For Each child As ProductStructureTypeLib.Product In current.Products
                FillMap(child, dicc)
            Next
        Catch

        End Try
    End Sub


    Private Sub SendTOWPN(oProductDocument As ProductStructureTypeLib.ProductDocument, oDic1 As Specialized.StringDictionary, strDir As String)

        Dim SendTo As INFITF.SendToService = _app.CreateSendTo()
        SendTo.SetInitialFile(oProductDocument.FullName)


        Dim oWillBeCopied(oDic1.Count - 1) As Object
        SendTo.GetListOfToBeCopiedFiles(oWillBeCopied)
        SendTo.SetDirectoryFile(strDir)

        Dim oDicPendientes As New Specialized.StringDictionary()

        ' RENOMBRADO (PRIMERA PASADA)
        For i As Integer = 0 To UBound(oWillBeCopied)
            Dim strFullPath As String = oWillBeCopied(i).ToString()
            Dim lastSlash As Integer = strFullPath.LastIndexOf("\")
            Dim strFileName As String = If(lastSlash > -1, strFullPath.Substring(lastSlash + 1), strFullPath)

            If oDic1.ContainsKey(strFileName) Then
                Dim strNewName As String = oDic1(strFileName)
                Dim dotIdx As Integer = strFileName.LastIndexOf(".")
                Dim currentNameNoExt As String = If(dotIdx > 0, strFileName.Substring(0, dotIdx), strFileName)

                If strNewName <> currentNameNoExt Then
                    Dim yaExisteEnConjunto As Boolean = False
                    For Each objPath In oWillBeCopied
                        If objPath IsNot Nothing AndAlso objPath.ToString().Contains(strNewName & ".") Then
                            yaExisteEnConjunto = True
                            Exit For
                        End If
                    Next

                    If Not yaExisteEnConjunto Then
                        SendTo.SetRenameFile(strFileName, strNewName)
                    Else
                        If Not oDicPendientes.ContainsKey(strFileName) Then
                            oDicPendientes.Add(strFileName, strNewName)
                        End If
                    End If
                End If
            End If
        Next

        ' Renombrado Segunda pasada
        If oDicPendientes.Count > 0 Then
            Dim llavesPendientes(oDicPendientes.Count - 1) As String
            oDicPendientes.Keys.CopyTo(llavesPendientes, 0)
            For Each strFileKey As String In llavesPendientes
                If strFileKey Is Nothing Then Continue For
                Try
                    SendTo.SetRenameFile(strFileKey, oDicPendientes(strFileKey))
                Catch
                End Try
            Next
        End If
        SendTo.Run()
        MsgBox("SendTo finalizado con éxito." & vbCrLf & "Pendientes intentados: " & oDicPendientes.Count)
    End Sub

End Class







' IMPORTANTE:
' "oWillBeCopied" es una lista que puede o no contener todos los archivos que se van a copiar en funcion de su tamaño (lenght)
' Si "oWillBeCopied" se declara de un tamaño menor a la cantidad de archivos que se van a copiar,
' no da error, pero el renombrado solo se lleva a cabo con la cantidad que esa lista contiene. Si "oWillBeCopied" es mas grande, da error.

' Cauando se utilizan referencias externas (por ejemplo utilizando el módulo "Structure Design" hay archivos de extension .CATMaterial o también
' Parts que estan en memoria pero no cargados (unloaded), esto no es tenido en cuenta por el diccionario de PartNumbers.
' Es por esto que, si el product raíz "arrastra archivos que son usados como referencia externas como es el caso de los croquiz de los perfiles,
' el oDic no los computa y el oDic.Count va a dar diferente a la cantidad "oWillBeCopied.lenght")

' También, en el procedimiento de verificar si los archivos ya existen en el directorio destino,
' no se tienen en cuenta las referencias externas (.CATMaterial, croquiza de CATParts, etc,)
' entonces, al querer pisar nuevamente todos los arhivos, el número de "n archivos ya existen" puede diferir de lo que contiene "oWillBeCopied"

' (*) Me fijo si el dicionario ya contiene un nombre de los nuevos,
' porque lo que estaría pasando es que quiera asignar un nombre nuevo que es identico a uno que ya existe
' Es decir quiere dar el nombre "A" a una pieza, pero ese nombre "A" ya es el nombre de otro archivo de mas abajo.
' Si ese es el caso, entonces no puedo renombrar en este momento.
' Lo que hace es, guarda ese par en el diccionario "oDicNoRenamed" y lo procesa luego cuando la pieza de mas abajo, ya no es mas "A"
' Utilizar un Segundo cilco de renombrado: NO FUNCIONA SIEMPRE!

' Conclusión:
' Es preferible utilizar el servicio "SendTo" sin referencias externas, es decir, los product que forma el Structure Design,
' cambiarlos a "allCatPart" o eliminar las referencias externas, para que solo queden archivos del tipo "CATProdcut" y "CATPart".

' Al realizar el SendTo el comando no tiene en cuanta si el product tiene propiedades como ser "Description", "Source", "Definition", etc.
' Entonces al hacer el SendTo, esas propiedades no se copian al nuevo archivo. Hay que realizar un proceso aparte para copiar esas propiedades.


' Una advertencia sobre el tamaño del Array
' Como estás manteniendo la línea: Dim oWillBeCopied(oDic1.Count - 1) As Object
' Si el producto raíz tiene referencias externas (como un .CATMaterial que no está en tu diccionario), el SendTo querrá meterlo en el array.
' Como tu array tiene el tamaño exacto de tu diccionario, y el raíz ya ocupa un lugar,
' si hay elementos "extra" que CATIA detecta, la línea GetListOfToBeCopiedFiles podría darte el Error de rango esperado que vimos antes.
' El único riesgo técnico sigue siendo que CATIA encuentre más archivos de los que tu diccionario tiene contabilizados



' Desajuste de Conteo: El diccionario cuenta elementos del árbol, pero SendTo cuenta archivos físicos en disco;
' basta con que exista un solo archivo "extra" (como el Producto Raíz o un .CATMaterial) para que la lista
' supere el tamaño del array.Error de Rango Crítico: Al ser un objeto COM, SendTo no puede redimensionar un array de .NET;
' si intenta escribir el archivo $n+1$ en un espacio de $n$, el programa se detiene inmediatamente con una excepción de rango.
' Invisibilidad de Dependencias:
' El método asume que tu estructura lógica es idéntica a la estructura de archivos, ignorando que CATIA arrastra
' vínculos ocultos que no aparecen en el árbol de productos pero que el servicio de copia está obligado a procesar.







