Public Class ComCleaner


    Public Sub CleanExcel(ByRef app As Microsoft.Office.Interop.Excel.Application,
                     ByRef workbooks As Microsoft.Office.Interop.Excel.Workbooks,
                     ByRef workbook As Microsoft.Office.Interop.Excel.Workbook,
                     ByRef sheets As Microsoft.Office.Interop.Excel.Sheets,
                     ByRef sheet As Microsoft.Office.Interop.Excel.Worksheet)

        ' 1. La orden de cierre
        Try
            If workbook IsNot Nothing Then workbook.Close(SaveChanges:=False)
            If app IsNot Nothing Then app.Quit()
        Catch ex As Exception
            ' Por si ya estaba cerrado
        End Try

        ' 2. Liberar TODO en orden inverso
        ' No dejes ni un solo objeto COM vivo
        ReleaseObject(sheet)
        ReleaseObject(sheets)
        ReleaseObject(workbook)
        ReleaseObject(workbooks) ' Esta colección es clave
        ReleaseObject(app)

        ' 3. El golpe final
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub



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
            If obj IsNot Nothing AndAlso Runtime.InteropServices.Marshal.IsComObject(obj) Then
                Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            End If
        Catch
        Finally
            obj = Nothing
        End Try
    End Sub

End Class