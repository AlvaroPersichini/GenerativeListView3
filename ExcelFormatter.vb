Option Explicit On
Option Strict On

Public Class ExcelFormatter

    Sub FormatoListView2(oWorkSheetListView As Microsoft.Office.Interop.Excel.Worksheet, rows As Integer)

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Step 3/3: Formatting Excel...")

        oWorkSheetListView.Activate() : oWorkSheetListView.Name = "ListView"

        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = CType(oWorkSheetListView.Parent, Microsoft.Office.Interop.Excel.Workbook)

        'Está asignando el item 1 del total de todas las ventanas.
        Dim viewListView As Microsoft.Office.Interop.Excel.WorksheetView = CType(oWorkBook.Windows.Item(1).SheetViews.Item(1), Microsoft.Office.Interop.Excel.WorksheetView)
        viewListView.DisplayGridlines = False
        Dim oRangoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A1", "L2")
        Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A3", "L3")
        Dim strColumnLetter As String
        Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
        Dim a As String
        Dim b As String


        ' // Arma el diccionario con los textos del encabezado. Si a futuro se requieren otras columnas
        ' // hay que modificar esto. Se pueden armar diccionarios o listas aparte y luego pasarlas como argumentos
        Dim oDicListViewColumnText As New Dictionary(Of String, String) From {
            {"A1", "#"},
            {"B1", "CurrentDirectory"},
            {"C1", "FileName"},
            {"D1", "CurrentPartNumber"},
            {"E1", "NewPartNumber"},
            {"F1", "DescriptionRef"},
            {"G1", "Quantity"},
            {"H1", "Source"},
            {"I1", "Level"},
            {"J1", "Nomenclature"},
            {"K1", "Definition"},
            {"L1", "Image"}
        }
        For Each kvp As KeyValuePair(Of String, String) In oDicListViewColumnText
            oWorkSheetListView.Range(kvp.Key).Value = kvp.Value
        Next

        ' Bordes del encabezado
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado.Cells
            With c
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            End With
        Next

        ' Fuente, tamaño y alineado de todo el documento
        With oWorkSheetListView.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        ' Fuente, tamaño y alineado del encabezado
        With oWorkSheetListView
            .Range("A1", "D1").Interior.ColorIndex = 15
            .Range("E1", "K1").Interior.Color = RGB(204, 255, 255)
            .Range("L1").Interior.Color = RGB(255, 165, 0)
            .Range("B1", "L1").Orientation = 90
            .Range("A1", "L1").Font.Bold = True
        End With


        ' Hace AutoFit pero a la columna de imagenes no.
        ' Aca hay que incluir la opcion de que si la planilla va a tener imagenes entonces que no haga AutoFit,
        ' pero si son incluidas las imagenes, no debería hacer autofit.
        For Each C As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado
            C.EntireColumn.AutoFit()
        Next


        ' Formato aplicado a todo el cuerpo
        With oWorkSheetListView
            .Range("A3", "L" & rows + 2).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            .Range("A3", "L" & rows + 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .Range("L3", "L" & rows + 2).RowHeight = 100
            .Range("L3", "L" & rows + 2).ColumnWidth = 18
            .Range("E3", "E" & rows + 2).ColumnWidth = 18
        End With


        ' Para aplicar los bordes a cada columna hasta la última fila de datos,
        ' hay que hacer estos pasos para armar el rango
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpo
            strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
            b = strColumnLetter & rows + 2
            oCurrentRange = oWorkSheetListView.Range(a, b)
            With oCurrentRange
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            End With
        Next

    End Sub

End Class