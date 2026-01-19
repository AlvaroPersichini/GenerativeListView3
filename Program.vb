Option Explicit On
Option Strict On

Module Program

    Sub Main()


        ' --- 1. INICIALIZACIÓN DE CATIA ---
        Dim oCATIA As New CATIA
        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = CType(oCATIA.AppCATIA.ActiveDocument, ProductStructureTypeLib.ProductDocument)
        Dim oProduct As ProductStructureTypeLib.Product = oProductDocument.Product


        ' --- 2. INICIALIZACIÓN DE EXCEL ---
        Dim myExcel As New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False,
            .ScreenUpdating = False,
            .DisplayAlerts = False
        }
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.Workbooks.Add()
        Dim oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(oWorkbook.Worksheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)



        ' --- 3. GESTIÓN DE DIRECTORIOS ---
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)

        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If


        ' --- 4. EXTRACCIÓN Y FORMATEO ---
        oCATIA.AppCATIA.DisplayFileAlerts = False
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim oExcelFormater As New ExcelFormatter
        Dim oCatiaData As Dictionary(Of String, PwrProduct) = oCatiaDataextractor.ExtractData(oProduct, folderPath, True)
        To_Excel.CompletaListView2(oProduct, oWorkSheet, folderPath, oCatiaData, True)
        oExcelFormater.FormatoListView2(oWorkSheet, oCatiaData.Count)



        ' --- 5. GUARDADO y CIERRE ---
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
        oWorkbook.SaveAs(excelFileName)



        ' --- 6. LIMPIEZA ATÓMICA DE OBJETOS COM ---

        ' A. Cerrar archivos y aplicaciones primero
        oWorkbook.Close()
        myExcel.Quit()

        ' B. Liberar objetos del diccionario (Referencias a Products de CATIA)
        For Each kvp In oCatiaData
            If kvp.Value.Product IsNot Nothing Then
                Runtime.InteropServices.Marshal.FinalReleaseComObject(kvp.Value.Product)
            End If
        Next
        oCatiaData.Clear()

        ' C. Liberar variables de contenido (Excel)
        Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkSheet)
        Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkbook)
        Runtime.InteropServices.Marshal.FinalReleaseComObject(myExcel)

        ' D. Liberar variables de contenido (CATIA)
        Runtime.InteropServices.Marshal.FinalReleaseComObject(oProduct)
        Runtime.InteropServices.Marshal.FinalReleaseComObject(oProductDocument)

        ' E. Restaurar interfaz de CATIA y liberar la APP
        ' No usamos FinalReleaseComObject sobre "oCATIA" porque es tu clase.
        ' Si tu clase oCATIA tiene una propiedad que es la App de CATIA, liberas esa:
        If oCATIA.AppCATIA IsNot Nothing Then
            oCATIA.AppCATIA.Interactive = True
            oCATIA.AppCATIA.DisplayFileAlerts = True
            ' Solo liberamos la propiedad interna si es un objeto COM puro
            Runtime.InteropServices.Marshal.FinalReleaseComObject(oCATIA.AppCATIA)
        End If

        ' F. Forzar al Garbage Collector (Doble pasada)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

End Module