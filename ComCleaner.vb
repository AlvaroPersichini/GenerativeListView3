Imports System.Runtime.InteropServices

Public Class ComCleaner

    ''' <summary>
    ''' Limpia toda la suite de Excel: Hoja, Libro, Colecciones y Aplicación.
    ''' </summary>
    Public Sub CleanExcel(ByRef app As Microsoft.Office.Interop.Excel.Application,
                         ByRef workbook As Microsoft.Office.Interop.Excel.Workbook,
                         ByRef sheets As Microsoft.Office.Interop.Excel.Sheets,
                         ByRef sheet As Microsoft.Office.Interop.Excel.Worksheet)

        ' 1. Cerrar el libro sin guardar (ya se debió guardar antes)
        If workbook IsNot Nothing Then
            Try : workbook.Close(False) : Catch : End Try
        End If

        ' 2. Salir de la aplicación
        If app IsNot Nothing Then
            Try : app.Quit() : Catch : End Try
        End If

        ' 3. Liberación atómica de objetos COM (Hijos -> Padres)
        ReleaseObject(sheet)
        ReleaseObject(sheets)
        ReleaseObject(workbook)
        ReleaseObject(app)
    End Sub

    ''' <summary>
    ''' Limpia toda la suite de CATIA: Diccionario de datos, Producto, Documento y Aplicación.
    ''' </summary>
    Public Sub CleanCatia(ByRef app As INFITF.Application,
                         ByRef doc As INFITF.Document,
                         ByRef prod As ProductStructureTypeLib.Product,
                         ByRef data As Dictionary(Of String, PwrProduct))

        ' 1. Restaurar interfaz
        If app IsNot Nothing Then
            app.Interactive = True
            app.DisplayFileAlerts = True
        End If

        ' 2. Liberar objetos del diccionario (Referencias internas de CATIA)
        If data IsNot Nothing Then
            For Each kvp In data
                If kvp.Value?.Product IsNot Nothing Then
                    ReleaseObject(kvp.Value.Product)
                End If
            Next
            data.Clear()
        End If

        ' 3. Liberar objetos principales
        ReleaseObject(prod)
        ReleaseObject(doc)
        ReleaseObject(app)

        ' 4. Forzar al Garbage Collector (Indispensable para CATIA)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    ' Método privado auxiliar para evitar repetir código
    Private Sub ReleaseObject(ByRef obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.FinalReleaseComObject(obj)
            End If
        Catch
        Finally
            obj = Nothing
        End Try
    End Sub

End Class