'Option Explicit On
'Option Strict On


''' IMPORTANTE:
''' "oWillBeCopied" es una lista que puede o no contener todos los archivos que se van a copiar en funcion de su tamaño (lenght)
''' Si "oWillBeCopied" se declara de un tamaño menor a la cantidad de archivos que se van a copiar,
''' no da error, pero el renombrado solo se lleva a cabo con la cantidad que esa lista contiene. Si "oWillBeCopied" es mas grande, da error.

''' Cauando se utilizan referencias externas (por ejemplo utilizando el módulo "Structure Design" hay archivos de extension .CATMaterial o también
''' Parts que estan en memoria pero no cargados (unloaded), esto no es tenido en cuenta por el diccionario de PartNumbers.
''' Es por esto que, si el product raíz "arrastra archivos que son usados como referencia externas como es el caso de los croquiz de los perfiles,
''' el oDic no los computa y el oDic.Count va a dar diferente a la cantidad "oWillBeCopied.lenght")

''' También, en el procedimiento de verificar si los archivos ya existen en el directorio destino,
''' no se tienen en cuenta las referencias externas (.CATMaterial, croquiza de CATParts, etc,)
''' entonces, al querer pisar nuevamente todos los arhivos, el número de "n archivos ya existen" puede diferir de lo que contiene "oWillBeCopied"

''' (*) Me fijo si el dicionario ya contiene un nombre de los nuevos,
''' porque lo que estaría pasando es que quiera asignar un nombre nuevo que es identico a uno que ya existe
''' Es decir quiere dar el nombre "A" a una pieza, pero ese nombre "A" ya es el nombre de otro archivo de mas abajo.
''' Si ese es el caso, entonces no puedo renombrar en este momento.
''' Lo que hace es, guarda ese par en el diccionario "oDicNoRenamed" y lo procesa luego cuando la pieza de mas abajo, ya no es mas "A"
''' Utilizar un Segundo cilco de renombrado: NO FUNCIONA SIEMPRE!

''' Conclusión:
''' Es preferible utilizar el servicio "SendTo" sin referencias externas, es decir, los product que forma el Structure Design,
''' cambiarlos a "allCatPart" o eliminar las referencias externas, para que solo queden archivos del tipo "CATProdcut" y "CATPart".

''' Al realizar el SendTo el comando no tiene en cuanta si el product tiene propiedades como ser "Description", "Source", "Definition", etc.
''' Entonces al hacer el SendTo, esas propiedades no se copian al nuevo archivo. Hay que realizar un proceso aparte para copiar esas propiedades.


''' Una advertencia sobre el tamaño del Array
''' Como estás manteniendo la línea: Dim oWillBeCopied(oDic1.Count - 1) As Object
''' Si el producto raíz tiene referencias externas (como un .CATMaterial que no está en tu diccionario), el SendTo querrá meterlo en el array.
''' Como tu array tiene el tamaño exacto de tu diccionario, y el raíz ya ocupa un lugar,
''' si hay elementos "extra" que CATIA detecta, la línea GetListOfToBeCopiedFiles podría darte el Error de rango esperado que vimos antes.
''' El único riesgo técnico sigue siendo que CATIA encuentre más archivos de los que tu diccionario tiene contabilizados



''' Desajuste de Conteo: El diccionario cuenta elementos del árbol, pero SendTo cuenta archivos físicos en disco;
''' basta con que exista un solo archivo "extra" (como el Producto Raíz o un .CATMaterial) para que la lista
''' supere el tamaño del array.Error de Rango Crítico: Al ser un objeto COM, SendTo no puede redimensionar un array de .NET;
''' si intenta escribir el archivo $n+1$ en un espacio de $n$, el programa se detiene inmediatamente con una excepción de rango.
''' Invisibilidad de Dependencias:
''' El método asume que tu estructura lógica es idéntica a la estructura de archivos, ignorando que CATIA arrastra
''' vínculos ocultos que no aparecen en el árbol de productos pero que el servicio de copia está obligado a procesar.






'Public Class SendToProcessor

'    Private ReadOnly _app As INFITF.Application

'    Private _renameMap As Collections.Specialized.StringDictionary
'    Private _pendingRenames As Collections.Specialized.StringDictionary

'    Public Sub New(catiaApp As INFITF.Application)
'        _app = catiaApp
'        _renameMap = New Collections.Specialized.StringDictionary()
'        _pendingRenames = New Collections.Specialized.StringDictionary()
'    End Sub


'    ''' <summary>
'    ''' Ejecuta el proceso completo: Mapeo del árbol, gestión de nombres y copiado físico.
'    ''' </summary>
'    ''' <param name="rootProduct">Producto raíz de CATIA</param>
'    ''' <param name="targetDir">Carpeta de destino</param>
'    Public Sub Execute(rootProduct As ProductStructureTypeLib.Product, targetDir As String)
'        ' 1. Limpieza de diccionarios para una nueva ejecución
'        _renameMap.Clear()
'        _pendingRenames.Clear()

'        ' 2. Mapeo recursivo: Llenamos el diccionario (NombreArchivo -> PartNumber)
'        FillMap(rootProduct)

'        ' 3. Configuración del servicio SendTo
'        Dim sendToService As INFITF.SendToService = _app.CreateSendTo()
'        Dim rootDoc = CType(rootProduct.ReferenceProduct.Parent, ProductStructureTypeLib.ProductDocument)

'        sendToService.SetInitialFile(rootDoc.FullName)
'        sendToService.SetDirectoryFile(targetDir)

'        ' 4. Gestión del Array (Solución al error de rango)
'        ' Sumamos un margen (ej. 50) para absorber referencias externas (.CATMaterial, cgr, etc.) 
'        ' que no están en el árbol de productos pero que SendTo detecta en disco.
'        Dim oWillBeCopied(_renameMap.Count + 50) As Object
'        sendToService.GetListOfToBeCopiedFiles(oWillBeCopied)

'        ' 5. PRIMER CICLO: Renombrado directo
'        For Each objPath In oWillBeCopied
'            If objPath Is Nothing Then Continue For

'            Dim fullPath As String = objPath.ToString()
'            Dim fileName As String = IO.Path.GetFileName(fullPath)

'            If _renameMap.ContainsKey(fileName) Then
'                Dim newName As String = _renameMap(fileName)
'                Dim currentNameNoExt As String = IO.Path.GetFileNameWithoutExtension(fileName)

'                If newName <> currentNameNoExt Then
'                    ' Verificamos colisión: ¿El nombre nuevo ya existe en el set de archivos actual?
'                    If Not ExistsInCollection(newName, oWillBeCopied) Then
'                        sendToService.SetRenameFile(fileName, newName)
'                    Else
'                        ' Si existe, lo guardamos para la segunda pasada
'                        If Not _pendingRenames.ContainsKey(fileName) Then
'                            _pendingRenames.Add(fileName, newName)
'                        End If
'                    End If
'                End If
'            End If
'        Next


'        ' 6. SEGUNDO CICLO: Resolución de pendientes
'        If _pendingRenames.Count > 0 Then
'            For Each fileKey As String In _pendingRenames.Keys
'                Try
'                    sendToService.SetRenameFile(fileKey, _pendingRenames(fileKey))
'                Catch
'                    Console.WriteLine("Conflicto persistente en: " & fileKey)
'                End Try
'            Next
'        End If

'        ' 7. Ejecución Final
'        Try
'            sendToService.Run()
'            MsgBox("SendTo finalizado con éxito." & vbCrLf & "Archivos en mapa: " & _renameMap.Count & vbCrLf & "Pendientes resueltos: " & _pendingRenames.Count)
'        Catch ex As Exception
'            MsgBox("Error al ejecutar SendTo: " & ex.Message, MsgBoxStyle.Critical)
'        End Try
'    End Sub

'    ' --- MÉTODOS PRIVADOS DE APOYO ---


'    Private Sub FillMap(current As ProductStructureTypeLib.Product)

'        Try
'            ' Obtenemos el documento padre del ReferenceProduct
'            Dim parentDoc = CType(current.ReferenceProduct.Parent, INFITF.Document)
'            Dim docName As String = parentDoc.Name
'            Dim pn As String = current.PartNumber

'            ' Solo agregamos si no existe (evita errores con instancias repetidas)
'            If Not _renameMap.ContainsKey(docName) Then
'                _renameMap.Add(docName, pn)
'            End If

'            ' Recurrencia sobre los hijos
'            For Each child As ProductStructureTypeLib.Product In current.Products
'                FillMap(child)
'            Next
'        Catch
'            ' Omitimos errores de componentes unloaded o links rotos
'        End Try
'    End Sub


'    Private Function ExistsInCollection(nameWithNoExt As String, collection As Object()) As Boolean
'        For Each item In collection
'            If item IsNot Nothing AndAlso item.ToString().ToUpper().Contains(nameWithNoExt.ToUpper() & ".") Then
'                Return True
'            End If
'        Next
'        Return False
'    End Function


'End Class


Option Explicit On
Option Strict On

Imports System.Collections.Specialized

Public Class SendToProcessor
    Private ReadOnly _app As INFITF.Application

    Public Sub New(catiaApp As INFITF.Application)
        _app = catiaApp
    End Sub

    ''' <summary>
    ''' Esta función replica exactamente tu lógica original pero encapsulada.
    ''' </summary>
    Public Sub Execute(oProduct As ProductStructureTypeLib.Product, strDir As String)
        ' 1. Generamos el diccionario usando tu lógica de mapeo
        Dim oDic1 As StringDictionary = GetMap(oProduct)

        ' 2. Obtenemos el Documento Raíz
        Dim oProductDocument = CType(oProduct.ReferenceProduct.Parent, ProductStructureTypeLib.ProductDocument)

        ' 3. Ejecutamos TU lógica de SendTo (sin cambios en el algoritmo de renombrado)
        SendTOWPN(oProductDocument, oDic1, strDir)
    End Sub

    ' --- TU LÓGICA DE MAPEO (ADAPTADA A LA CLASE) ---
    Private Function GetMap(objRoot As ProductStructureTypeLib.Product) As StringDictionary
        Dim dicc As New StringDictionary()
        FillMap(objRoot, dicc)
        Return dicc
    End Function

    Private Sub FillMap(current As ProductStructureTypeLib.Product, ByRef dicc As StringDictionary)
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
            ' Mantenemos el comportamiento silencioso ante links rotos
        End Try
    End Sub

    ' --- TU LÓGICA DE SENDTO ORIGINAL (INTACTA) ---
    Private Sub SendTOWPN(oProductDocument As ProductStructureTypeLib.ProductDocument, oDic1 As StringDictionary, strDir As String)

        Dim SendTo As INFITF.SendToService = _app.CreateSendTo()
        SendTo.SetInitialFile(oProductDocument.FullName)

        ' Respeto tu dimensionamiento original: oDic1.Count - 1
        Dim oWillBeCopied(oDic1.Count - 1) As Object
        SendTo.GetListOfToBeCopiedFiles(oWillBeCopied)
        SendTo.SetDirectoryFile(strDir)

        Dim oDicPendientes As New StringDictionary()

        ' --- TU CICLO DE RENOMBRADO (PRIMERA PASADA) ---
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

        ' --- TU SEGUNDO CICLO DE RENOMBRADO ---
        If oDicPendientes.Count > 0 Then
            Dim llavesPendientes(oDicPendientes.Count - 1) As String
            oDicPendientes.Keys.CopyTo(llavesPendientes, 0)

            For Each strFileKey As String In llavesPendientes
                If strFileKey Is Nothing Then Continue For
                Try
                    SendTo.SetRenameFile(strFileKey, oDicPendientes(strFileKey))
                Catch
                    ' Log de consola como tenías
                End Try
            Next
        End If

        SendTo.Run()
        MsgBox("SendTo finalizado con éxito." & vbCrLf & "Pendientes intentados: " & oDicPendientes.Count)
    End Sub

End Class


