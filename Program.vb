Option Explicit On
Option Strict On

Module Program

    Sub Main()

        Console.WriteLine(">>> Starting Export Process...")
        Console.WriteLine("------------------------------------------------")


        ' --- 1. CONEXIÓN Y VALIDACIÓN CON CATIA ---
        Dim session As New CatiaSession()
        If Not session.IsReady Then
            MsgBox(session.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = session.RootProduct
        session.Application.DisplayFileAlerts = False



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

        ' Formatear 
        oExcelFormater.FormatoListView2(oWorkSheet)

        ' Guardar
        oWorkbook.SaveAs(excelFileName)

        ' Limpieza
        Dim oCleaner As New ComCleaner()
        oCleaner.CleanExcel(myExcel, oWorkbooks, oWorkbook, oWorkSheets, oWorkSheet)
        oCleaner.CleanCatia(session.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)



        Console.WriteLine("-----------------------------------------------------------------")
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))
        Console.WriteLine(">>> Cleanup Complete.")

    End Sub

End Module