


Module To_Excel
    Sub CompletaListView2(oProduct As ProductStructureTypeLib.Product,
                      oSheetListView As Microsoft.Office.Interop.Excel.Worksheet,
                      strDir As String,
                      oDiccType3 As Dictionary(Of String, PwrProduct))

        ' MENSAJE DE PROGRESO
        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Step 2/3: Filling Excel with extracted data...")

        Dim i As Integer = 3
        Dim oShape As Microsoft.Office.Interop.Excel.Shape

        ' --- MEJORA: Formateo masivo ANTES del bucle ---
        Dim ultimoFila As Integer = oDiccType3.Count + 2
        oSheetListView.Range("A3:L" & ultimoFila).NumberFormat = "@"

        For Each kvp As KeyValuePair(Of String, PwrProduct) In oDiccType3

            ' Usamos la ruta que ya procesamos en la extracción
            Dim sImgPath As String = kvp.Value.ImageFilePath

            With oSheetListView
                .Cells(i, "A").Value2 = i - 2
                .Cells(i, "B").Value2 = kvp.Value.Product.PartNumber
                .Cells(i, "C").Value2 = kvp.Value.ProductType
                .Cells(i, "D").Value2 = kvp.Value.Product.ReferenceProduct.Parent.Name
                .Cells(i, "E").Value2 = kvp.Value.FullPath
                .Cells(i, "F").Value2 = kvp.Value.Product.DescriptionRef
                .Cells(i, "G").Value2 = kvp.Value.Quantity
                .Cells(i, "H").Value2 = kvp.Value.Source
                .Cells(i, "I").Value2 = kvp.Value.Level
                .Cells(i, "J").Value2 = kvp.Value.Product.Nomenclature

                ' Inserción de imagen
                If IO.File.Exists(sImgPath) Then
                    Dim cl As Microsoft.Office.Interop.Excel.Range = .Cells(i, "K")
                    oShape = .Shapes.AddPicture(sImgPath, False, True, cl.Left + 5.5, cl.Top + 5, 90, 90)
                End If
            End With
            i += 1
        Next
        oSheetListView.Application.ActiveWindow.DisplayVerticalScrollBar = True
        oSheetListView.Application.ActiveWindow.DisplayHorizontalScrollBar = True
    End Sub


End Module