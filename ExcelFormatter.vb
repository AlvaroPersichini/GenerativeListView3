Option Explicit On
Option Strict On

Public Class ExcelFormatter

    Sub FormatoListView2(oWorkSheetListView As Microsoft.Office.Interop.Excel.Worksheet)

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Step 3/3: Formatting Excel...")


        ' Desproteger hoja
        oWorkSheetListView.Unprotect()


        ' Última fila con datos
        Dim lastRow As Integer = CType(oWorkSheetListView.Cells(oWorkSheetListView.Rows.Count, 1), Microsoft.Office.Interop.Excel.Range).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        If lastRow < 3 Then lastRow = 3 ' Mínimo hasta la fila 3 para no romper el formato


        ' Nombre de la hoja
        oWorkSheetListView.Name = "ListView"


        ' Sacar gridlines
        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = CType(oWorkSheetListView.Parent, Microsoft.Office.Interop.Excel.Workbook)
        Dim viewListView As Microsoft.Office.Interop.Excel.WorksheetView = CType(oWorkBook.Windows.Item(1).SheetViews.Item(1), Microsoft.Office.Interop.Excel.WorksheetView)
        viewListView.DisplayGridlines = False


        ' Rango de encabezados y cuerpo
        Dim oRangoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A1", "L2")
        Dim oRangoCuerpoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A3", "L3")



        ' Encabezados
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


        ' Bordes de encabezados
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado.Cells
            With c.Borders
                .Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            End With
        Next

        ' Formato general de la hoja
        With oWorkSheetListView.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        With oWorkSheetListView
            .Range("A1", "D1").Interior.ColorIndex = 15
            .Range("E1", "K1").Interior.Color = RGB(204, 255, 255)
            .Range("L1").Interior.Color = RGB(255, 165, 0)
            .Range("B1", "L1").Orientation = 90
            .Range("A1", "L1").Font.Bold = True
        End With

        ' Autofit de columnas
        For i As Integer = 1 To 11
            CType(oWorkSheetListView.Columns(i), Microsoft.Office.Interop.Excel.Range).EntireColumn.AutoFit()
        Next


        ' Formato cuerpo
        With oWorkSheetListView
            Dim rangeCuerpoCompleto As String = "A3:L" & lastRow
            .Range(rangeCuerpoCompleto).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            .Range(rangeCuerpoCompleto).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .Range("L3:L" & lastRow).RowHeight = 100
            .Range("L3:L" & lastRow).ColumnWidth = 18
            .Range("E3:E" & lastRow).ColumnWidth = 18
        End With


        'Bordes cuerpo 
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpoEncabezado.Cells
            Dim colLetter As String = Split(c.Address(True, False), "$")(0)
            Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range(colLetter & "3", colLetter & lastRow)
            With oCurrentRange.Borders
                .Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            End With
        Next


    End Sub

End Class