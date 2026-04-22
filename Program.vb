
Option Explicit On
Option Strict On
Imports CATIAClassLibrary

Module Program

    Sub Main()

        ' Inicio
        Console.WriteLine(">>> Starting Process...")



        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False




        ' Excel
        Dim oExcelSession As New ExcelSession()
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = oExcelSession.CreateNewWorkbook()
        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets = oWorkbook.Sheets
        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(oSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)




        ' Directorios y nombres
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If


        ' Extraer de CATIA
        Dim oCatiaData As Dictionary(Of String, PwrProduct)
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim addImages As Boolean = True
        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, addImages)


        ' Inyectar a EXCEL
        Dim oExcelDataInjector As New ExcelDataInjector
        oExcelDataInjector.InjectData(oSheet, oCatiaData)


        ' Formatear EXCEL
        Dim oExcelFormater As New ExcelFormatter
        oExcelFormater.FormatoListView2(oSheet)


        ' Guardar EXCEL
        oExcelSession.Workbook.SaveAs(excelFileName)


        ' Limpieza
        Dim oCleaner As New ComCleaner()
        oCleaner.CleanExcel(oExcelSession.Application, oExcelSession.Workbooks, oExcelSession.Workbook, oSheets, oSheet)
        oCleaner.CleanCatia(CATIAsession.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)



        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))



    End Sub



End Module