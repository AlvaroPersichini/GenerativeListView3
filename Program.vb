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
        Dim xlSession As New ExcelSession()
        xlSession.CreateNewWorkbook()
        If Not xlSession.IsReady Then
            Console.WriteLine(xlSession.ErrorMessage)
            Return
        End If



        ' --- 3. GESTIÓN DE DIRECTORIOS ---
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
        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, True)



        ' Inyectar
        Dim oExcelDataInjector As New ExcelDataInjector
        oExcelDataInjector.InjectData(xlSession.Worksheet, oCatiaData)



        ' Formatear 
        Dim oExcelFormater As New ExcelFormatter
        oExcelFormater.FormatoListView2(xlSession.Worksheet)



        ' Guardar
        xlSession.Workbook.SaveAs(excelFileName)



        ' Limpieza
        Dim oCleaner As New ComCleaner()
        oCleaner.CleanExcel(xlSession.Application, xlSession.Workbooks, xlSession.Workbook, xlSession.Worksheets, xlSession.ActiveSheet)
        oCleaner.CleanCatia(CATIAsession.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)



        Console.WriteLine("-----------------------------------------------------------------")
        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))
        Console.WriteLine(">>> Cleanup Complete.")


    End Sub



End Module