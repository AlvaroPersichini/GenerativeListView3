Option Explicit On
Option Strict On
Imports CATIAClassLibrary


Module Program2

    Sub Main()

        Console.WriteLine(">>> Starting Process...")

        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False


        Dim oCatiaApp As INFITF.Application = CATIAsession.Application

        oCatiaApp.StartCommand("New Window")


        'Dim extractor As New CatiaDataExtractor()

        'Dim dims As (DimX As Double, DimY As Double, DimZ As Double) = extractor.GetBoundingBoxDimensions(oProduct)

        'MsgBox($"X: {dims.DimX:F2} mm | Y: {dims.DimY:F2} mm | Z: {dims.DimZ:F2} mm", MsgBoxStyle.Information, "Dimensiones Bounding Box")

        'Console.WriteLine(">>> Finished Successfully at " & DateTime.Now.ToString("HH:mm:ss"))


    End Sub



End Module