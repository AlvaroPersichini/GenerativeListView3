'Option Explicit On
'Option Strict On

'Module Program

'    Sub Main()


'        Console.WriteLine(">>> Starting Export Process...")
'        Console.WriteLine("------------------------------------------------")


'        ' --- 1. Conexión con CATIA ---
'        ' Este programa funciona con un Product, no con un Part.
'        ' Validación estricta del estado
'        ' Comprobamos si el Status es exactamente ProductDocument
'        Dim session As New CatiaSession()
'        If session.Status <> CatiaSession.CatiaSessionStatus.ProductDocument Then
'            MsgBox("Error: Se requiere un Product activo." & vbCrLf &
'           "Estado actual: " & session.Description, MsgBoxStyle.Critical)
'            Exit Sub
'        End If
'        Dim oAppCatia As INFITF.Application = session.Application
'        oAppCatia.DisplayFileAlerts = False
'        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = CType(oAppCatia.ActiveDocument, ProductStructureTypeLib.ProductDocument)
'        ' Comprobamos si el documento está guardado
'        If Not CheckSaveStatus(oProductDocument) Then
'            MessageBox.Show("El documento actual no ha sido guardado. Guárdelo antes de continuar.", "Aviso")
'            Exit Sub
'        End If
'        ' En este punto el oProduct ya esta validado
'        Dim oProduct As ProductStructureTypeLib.Product = oProductDocument.Product





'        ' --- 2. INICIALIZACIÓN DE EXCEL ---
'        Dim myExcel As New Microsoft.Office.Interop.Excel.Application With {
'            .Visible = False,
'            .ScreenUpdating = False,
'            .DisplayAlerts = False
'        }
'        Dim oWorkbooks As Microsoft.Office.Interop.Excel.Workbooks = myExcel.Workbooks
'        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = oWorkbooks.Add()
'        Dim oWorkSheets As Microsoft.Office.Interop.Excel.Sheets = oWorkbook.Worksheets
'        Dim oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(oWorkSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)





'        ' --- 3. GESTIÓN DE DIRECTORIOS y nombre del nuevo archivo .xlsx ---
'        Dim baseDir As String = "C:\Temp"
'        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
'        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
'        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
'        If Not IO.Directory.Exists(folderPath) Then
'            IO.Directory.CreateDirectory(folderPath)
'        End If







'        '**************************************************************************************
'        ' Funcionalidades principales: EXTRACCIÓN, ESCRITURA Y FORMATEO
'        '**************************************************************************************

'        Dim oCatiaDataextractor As New CatiaDataExtractor

'        Dim oExcelDataInjector As New ExcelDataInjector

'        Dim oExcelFormater As New ExcelFormatter

'        Dim oCatiaData As Dictionary(Of String, PwrProduct)

'        ' Extraer datos
'        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, True)

'        ' Escribir en Excel
'        oExcelDataInjector.InjectData(oWorkSheet, oCatiaData)

'        ' Formatear 
'        oExcelFormater.FormatoListView2(oWorkSheet)

'        ' Guardar
'        oWorkbook.SaveAs(excelFileName)






'        ' --- SECCIÓN DE LIMPIEZA ---
'        Dim oCleaner As New ComCleaner()

'        oCleaner.CleanExcel(myExcel, oWorkbook, oWorkSheets, oWorkSheet)

'        oCleaner.CleanCatia(oAppCatia, CType(oProductDocument, INFITF.Document), oProduct, oCatiaData)

'        Console.WriteLine(">>> Cleanup Complete.")



'    End Sub



'    Private Function CheckSaveStatus(oDoc As INFITF.Document) As Boolean
'        ' Un documento sin ruta (Path vacío) nunca ha sido guardado
'        If String.IsNullOrEmpty(oDoc.Path) Then
'            Return False
'        End If
'        ' Si Saved es False, tiene cambios pendientes
'        If Not oDoc.Saved Then
'            Return False
'        End If
'        Return True
'    End Function




'End Module




Option Explicit On
Option Strict On

Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine("------------------------------------------------")



        ' --- 1. CONEXIÓN Y VALIDACIÓN CON CATIA ---
        Dim session As New CatiaSession()
        If session.Status <> CatiaSession.CatiaSessionStatus.ProductDocument Then
            MsgBox(session.Description, MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim oAppCatia As INFITF.Application = session.Application : oAppCatia.DisplayFileAlerts = False
        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = CType(oAppCatia.ActiveDocument, ProductStructureTypeLib.ProductDocument)
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




        ' --- 3. GESTIÓN DE DIRECTORIOS ---
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")

        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If



        ' --- 4. EXTRACCIÓN, ESCRITURA Y FORMATEO ---
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim oExcelDataInjector As New ExcelDataInjector
        Dim oExcelFormater As New ExcelFormatter
        Dim oCatiaData As Dictionary(Of String, PwrProduct)




        ' Extraer
        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, True)

        ' Inyectar
        oExcelDataInjector.InjectData(oWorkSheet, oCatiaData)

        ' Formatear (ahora detecta filas automáticamente internamente)
        oExcelFormater.FormatoListView2(oWorkSheet)

        ' Guardar
        oWorkbook.SaveAs(excelFileName)




        ' --- 5. LIMPIEZA ATÓMICA ---
        Dim oCleaner As New ComCleaner()
        oCleaner.CleanExcel(myExcel, oWorkbook, oWorkSheets, oWorkSheet)
        oCleaner.CleanCatia(oAppCatia, CType(oProductDocument, INFITF.Document), oProduct, oCatiaData)




        Console.WriteLine("-----------------------------------------------------------------")
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))
        Console.WriteLine(">>> Cleanup Complete.")

    End Sub

End Module