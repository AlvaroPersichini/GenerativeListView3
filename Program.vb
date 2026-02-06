Option Explicit On
Option Strict On

Module Program

    Sub Main()

        ' Inicio
        Console.WriteLine(">>> Starting Process...")
        Console.WriteLine("------------------------------------------------")


        ' Catia
        Dim CATIAsession As New CatiaSession()
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False


        ' Excel
        Dim oExcelSession As New ExcelSession()
        oExcelSession.CreateNewWorkbook()
        If Not oExcelSession.IsReady Then
            Console.WriteLine(oExcelSession.ErrorMessage)
            Return
        End If




        ' Directorios y nombres
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")

        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If



        ' Extraer
        Dim oCatiaData As Dictionary(Of String, PwrProduct)
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim addImages As Boolean = True
        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, addImages)



        ' Inyectar
        Dim oExcelDataInjector As New ExcelDataInjector
        oExcelDataInjector.InjectData(oExcelSession.Worksheet, oCatiaData)



        ' Formatear 
        Dim oExcelFormater As New ExcelFormatter
        oExcelFormater.FormatoListView2(oExcelSession.Worksheet)



        ' Guardar
        oExcelSession.Workbook.SaveAs(excelFileName)



        ' Limpieza
        Dim oCleaner As New ComCleaner()
        oCleaner.CleanExcel(oExcelSession.Application, oExcelSession.Workbooks, oExcelSession.Workbook, oExcelSession.Worksheets, oExcelSession.ActiveSheet)
        oCleaner.CleanCatia(CATIAsession.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)



        Console.WriteLine("-----------------------------------------------------------------")
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))
        Console.WriteLine(">>> Cleanup Complete.")


    End Sub



End Module