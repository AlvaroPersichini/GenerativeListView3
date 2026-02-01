Option Explicit On
Option Strict On

Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine("------------------------------------------------")


        ' --- 1. Conexión con CATIA ---
        ' Este programa funciona con un Product, no con un Part.
        ' Validación estricta del estado
        ' Comprobamos si el Status es exactamente ProductDocument
        Dim session As New CatiaSession()
        If session.Status <> CatiaSession.CatiaSessionStatus.ProductDocument Then
            MsgBox("Error: Se requiere un Product activo." & vbCrLf &
           "Estado actual: " & session.Description, MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim oAppCatia As INFITF.Application = session.Application
        oAppCatia.DisplayFileAlerts = False
        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = CType(oAppCatia.ActiveDocument, ProductStructureTypeLib.ProductDocument)
        ' Comprobamos si el documento está guardado
        If Not CheckSaveStatus(oProductDocument) Then
            MessageBox.Show("El documento actual no ha sido guardado. Guárdelo antes de continuar.", "Aviso")
            Exit Sub
        End If
        ' En este punto el oProduct ya esta validado
        Dim oProduct As ProductStructureTypeLib.Product = oProductDocument.Product





        ' --- 2. INICIALIZACIÓN DE EXCEL ---
        Dim myExcel As New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False,
            .ScreenUpdating = False,
            .DisplayAlerts = False
        }
        Dim oWorkbooks As Microsoft.Office.Interop.Excel.Workbooks = myExcel.Workbooks
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = oWorkbooks.Add()
        Dim oWorkSheets As Microsoft.Office.Interop.Excel.Sheets = oWorkbook.Worksheets
        Dim oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(oWorkSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)





        ' --- 3. GESTIÓN DE DIRECTORIOS y nombre del nuevo archivo .xlsx ---
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If







        '**************************************************************************************
        ' Funcionalidades principales: EXTRACCIÓN, ESCRITURA Y FORMATEO
        '**************************************************************************************

        Dim oCatiaDataextractor As New CatiaDataExtractor

        Dim oExcelDataInjector As New ExcelDataInjector

        Dim oExcelFormater As New ExcelFormatter

        Dim oCatiaData As Dictionary(Of String, PwrProduct)


        ' Extraer datos
        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, True)

        ' Escribir en Excel
        oExcelDataInjector.InjectData(oWorkSheet, oCatiaData)

        ' Formatear 
        oExcelFormater.FormatoListView2(oWorkSheet)

        ' Guardar
        oWorkbook.SaveAs(excelFileName)





        ' --- 6. LIMPIEZA ATÓMICA DE OBJETOS COM ---
        ' A. Liberar objetos del diccionario (Referencias internas de CATIA)
        If oCatiaData IsNot Nothing Then
            For Each kvp In oCatiaData
                If kvp.Value?.Product IsNot Nothing Then
                    If Runtime.InteropServices.Marshal.IsComObject(kvp.Value.Product) Then
                        Runtime.InteropServices.Marshal.FinalReleaseComObject(kvp.Value.Product)
                    End If
                End If
            Next
            oCatiaData.Clear()
        End If
        ' B. Liberar variables de contenido de Excel (Hijos a Padres)
        ' Cerramos el libro pero NO la aplicación todavía
        If oWorkbook IsNot Nothing Then oWorkbook.Close(False)
        If oWorkSheet IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkSheet)
        If oWorkSheets IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkSheets)
        If oWorkbook IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkbook)
        If oWorkbooks IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oWorkbooks)
        ' Ahora que todo lo demás de Excel se liberó, cerramos la App y la liberamos
        If myExcel IsNot Nothing Then
            myExcel.Quit()
            Runtime.InteropServices.Marshal.FinalReleaseComObject(myExcel)
        End If
        ' C. Liberar variables de contenido de CATIA
        If oProduct IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oProduct)
        If oProductDocument IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oProductDocument)
        ' D. Restaurar interfaz y liberar la App de CATIA
        If oAppCatia IsNot Nothing Then
            oAppCatia.Interactive = True
            oAppCatia.DisplayFileAlerts = True
            Runtime.InteropServices.Marshal.FinalReleaseComObject(oAppCatia)
        End If
        ' E. Forzar al Garbage Collector
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()






        Console.WriteLine("-----------------------------------------------------------------")
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))





    End Sub



    Private Function CheckSaveStatus(oDoc As INFITF.Document) As Boolean
        ' Un documento sin ruta (Path vacío) nunca ha sido guardado
        If String.IsNullOrEmpty(oDoc.Path) Then
            Return False
        End If
        ' Si Saved es False, tiene cambios pendientes
        If Not oDoc.Saved Then
            Return False
        End If
        Return True
    End Function




End Module