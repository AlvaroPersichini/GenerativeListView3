Option Explicit On
Option Strict On
Imports CATIAClassLibrary
Imports EXCELClassLibrary

Module Program

    Sub Main()

        Console.WriteLine("> Starting Process...")

        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False


        ' Excel
        Dim oExcelSession As New ExcelSession
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


        ' Extracción CATIA
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim oDataTable As New DataTable
        oDataTable = oCatiaDataextractor.ExtractData(oProduct, folderPath)


        ' Inyectar a EXCEL
        Dim oExcelDataInjector As New ExcelDataInjector
        ' oExcelDataInjector.InjectData(oSheet, oDataTable)


        ' Formatear EXCEL
        Dim oExcelFormater As New ExcelFormatter
        oExcelFormater.FormatoListView2(oSheet)


        ' Guardar EXCEL
        oExcelSession.Workbook.SaveAs(excelFileName)


        ' Limpieza
        'Dim oCleaner As New ComCleaner()
        'oCleaner.CleanExcel(oExcelSession.Application, oExcelSession.Workbooks, oExcelSession.Workbook, oSheets, oSheet)
        'oCleaner.CleanCatia(CATIAsession.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)


        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))

    End Sub

End Module


























'Option Explicit On
'Option Strict On
'Imports CATIAClassLibrary
'Imports EXCELClassLibrary



'Module Program

'    Sub Main()

'        Console.WriteLine(">>> Starting Process...")

'        ' Catia
'        Dim CATIAsession As New CatiaSession
'        If Not CATIAsession.IsReady Then
'            MsgBox(CATIAsession.Description)
'            Exit Sub
'        End If
'        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
'        CATIAsession.Application.DisplayFileAlerts = False


'        ' Excel
'        Dim oExcelSession As New ExcelSession
'        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = oExcelSession.CreateNewWorkbook()
'        Dim oSheets As Microsoft.Office.Interop.Excel.Sheets = oWorkbook.Sheets
'        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(oSheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)


'        ' Directorios y nombres
'        Dim baseDir As String = "C:\Temp"
'        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
'        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
'        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
'        If Not IO.Directory.Exists(folderPath) Then
'            IO.Directory.CreateDirectory(folderPath)
'        End If


'        ' Extraer de CATIA
'        Dim oCatiaData As Dictionary(Of String, (FullPath As String,
'                                         FileName As String,
'                                         ImageFilePath As String,
'                                         PartNumber As String,
'                                         DescriptionRef As String,
'                                         Nomenclature As String,
'                                         Definition As String,
'                                         Quantity As Integer,
'                                         Level As Integer,
'                                         ProductType As String,
'                                         Source As Integer))





'        ' Extraccion

'        Dim oCatiaDataextractor As New CatiaDataExtractor

'        Dim addImages As Boolean = True

'        oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, addImages)


'        'para usar el SubTracer, se debe envolver el bloque de código que se desea medir con la instrucción Using New SubTracer("Nombre del Proceso")
'        'Using New SubTracer("Extracción CATIA (Datos + Imágenes)")
'        '    Dim oCatiaDataextractor As New CatiaDataExtractor
'        '    Dim addImages As Boolean = True
'        '    oCatiaData = oCatiaDataextractor.ExtractData(oProduct, folderPath, addImages)
'        'End Using



'        ' Inyectar a EXCEL
'        Dim oExcelDataInjector As New ExcelDataInjector
'        oExcelDataInjector.InjectData(oSheet, oCatiaData)


'        ' Formatear EXCEL
'        Dim oExcelFormater As New ExcelFormatter
'        oExcelFormater.FormatoListView2(oSheet)


'        ' Guardar EXCEL
'        oExcelSession.Workbook.SaveAs(excelFileName)


'        ' Limpieza
'        'Dim oCleaner As New ComCleaner()
'        'oCleaner.CleanExcel(oExcelSession.Application, oExcelSession.Workbooks, oExcelSession.Workbook, oSheets, oSheet)
'        'oCleaner.CleanCatia(CATIAsession.Application, CType(oProduct.ReferenceProduct.Parent, INFITF.Document), oProduct, oCatiaData)


'        Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))


'    End Sub



'End Module
