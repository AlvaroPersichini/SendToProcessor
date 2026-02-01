Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Public Class ComCleaner
    ''' <summary>
    ''' Libera una lista de objetos COM de CATIA (Product, Document, Application, etc.)
    ''' </summary>
    Public Sub Release(ParamArray objects As Object())
        If objects Is Nothing Then Exit Sub

        For Each obj In objects
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Try
                    ' Mata la referencia COM inmediatamente
                    Marshal.FinalReleaseComObject(obj)
                Catch
                    ' Si el objeto ya se liberó, ignoramos el error
                End Try
            End If
        Next

        ' Forzamos la limpieza de memoria de .NET
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class