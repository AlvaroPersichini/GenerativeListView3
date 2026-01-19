Option Explicit On
Option Strict On

Imports System.IO
Imports INFITF
Imports ProductStructureTypeLib

'El manejo de los "Components" 
'Los detecta y los salta: Si encuentra un "Component"
'(no tienen un archivo propio y solo sirven para organizar),
'el programa se da cuenta y no los pone en la lista.
'Aunque salte el Component, entra a mirar qué tiene dentro. Si adentro hay piezas reales, las trata normalmente.

Public Class CatiaDataExtractor

    ''' <summary>
    ''' Extrae los datos de la estructura de CATIA y genera capturas de pantalla.
    ''' </summary>
    Public Function ExtractData(oRootProduct As Product,
                                folderPath As String,
                                takeSnaps As Boolean) As Dictionary(Of String, PwrProduct)

        If takeSnaps AndAlso Not String.IsNullOrEmpty(folderPath) Then
            If Not Directory.Exists(folderPath) Then Directory.CreateDirectory(folderPath)
        End If

        Dim oDictionary As New Dictionary(Of String, PwrProduct)
        Dim PPRoot As New PwrProduct

        ' Documento del Root
        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)

        With PPRoot
            .Product = oRootProduct
            .Quantity = 1
            .ProductType = TypeName(rootDoc)
            .Source = oRootProduct.Source
            .Level = 0
            .FileName = rootDoc.Name
            .FullPath = rootDoc.FullName
            .ImageFilePath = If(takeSnaps, TakeSnapshot(oRootProduct, folderPath, True), "")
        End With

        oDictionary.Add(oRootProduct.PartNumber, PPRoot)

        ' Iniciamos recursividad pasando el documento del padre (en este caso el root)
        ProcesarHijosRecursivo(oRootProduct, oDictionary, 1, folderPath, takeSnaps, rootDoc)

        Return oDictionary
    End Function

    Private Sub ProcesarHijosRecursivo(oParent As Product,
                                      ByRef oDictionary As Dictionary(Of String, PwrProduct),
                                      ByVal currentLevel As Integer,
                                      folderPath As String,
                                      takeSnaps As Boolean,
                                      oParentDoc As INFITF.Document) ' <-- Recibe el Doc del Padre

        For Each oChild As Product In oParent.Products
            ' Obtenemos el documento al que pertenece la referencia del hijo
            Dim oChildDoc As INFITF.Document = CType(oChild.ReferenceProduct.Parent, INFITF.Document)

            ' COMPARACIÓN DINÁMICA: Si el hijo vive en el mismo archivo que el padre
            If oChildDoc.FullName = oParentDoc.FullName Then

                ' ES UN COMPONENT: Atravesamos sin registrar y mantenemos el nivel
                ' Pasamos el mismo oParentDoc porque el componente no genera un archivo nuevo
                ProcesarHijosRecursivo(oChild, oDictionary, currentLevel, folderPath, takeSnaps, oParentDoc)
            Else
                ' ES UN ARCHIVO REAL (Part o Product)
                Dim pNumber As String = oChild.PartNumber

                If oDictionary.ContainsKey(pNumber) Then
                    oDictionary.Item(pNumber).Quantity += 1
                Else
                    Dim PP As New PwrProduct
                    With PP
                        .Product = oChild
                        .Quantity = 1
                        .ProductType = TypeName(oChildDoc)
                        .Source = oChild.Source
                        .Level = currentLevel
                        .FileName = oChildDoc.Name
                        .FullPath = oChildDoc.FullName
                        .ImageFilePath = If(takeSnaps, TakeSnapshot(oChild, folderPath, False), "")
                    End With
                    oDictionary.Add(pNumber, PP)
                End If

                ' Si es un ensamble real, profundizamos pasando el documento del HIJO como nuevo padre
                If TypeOf oChildDoc Is ProductDocument Then
                    ProcesarHijosRecursivo(oChild, oDictionary, currentLevel + 1, folderPath, takeSnaps, oChildDoc)
                End If
            End If
        Next
    End Sub

    Private Function TakeSnapshot(oProd As Product, folder As String, isRoot As Boolean) As String
        ' Limpiar el nombre para evitar errores con caracteres como / o *
        Dim safePartNumber As String = CleanFileName(oProd.PartNumber)
        Dim finalFileName As String = Path.Combine(folder, safePartNumber & ".jpg")

        ' Filtro para piezas auxiliares
        If Left(oProd.PartNumber, 3) = "AUX" Then Return ""

        Dim oApp As INFITF.Application = oProd.Application
        Dim docPrincipal As INFITF.Document = oApp.ActiveDocument

        ' --- 1. GESTIÓN DE VENTANAS ---
        If Not isRoot Then
            Dim oSelection As Selection = docPrincipal.Selection
            oSelection.Clear()
            oSelection.Add(oProd)
            oApp.StartCommand("Open in New Window")
            oApp.RefreshDisplay = True

            ' Control de seguridad: si no cambió la ventana, abortar para no cerrar el principal
            If oApp.ActiveDocument Is docPrincipal Then
                oSelection.Clear()
                Return ""
            End If
        End If

        ' --- 2. CONFIGURACIÓN VISUAL ---
        Dim oCurrentWindow As INFITF.Window = oApp.ActiveWindow
        Dim oSpecsWin As SpecsAndGeomWindow = CType(oCurrentWindow, SpecsAndGeomWindow)
        Dim oViewer As Viewer3D = CType(oSpecsWin.Viewers.Item(1), Viewer3D)

        ' Fondo Blanco
        Dim oldColor(2), white(2) As Object
        white(0) = 1 : white(1) = 1 : white(2) = 1
        oViewer.GetBackgroundColor(oldColor)
        oViewer.PutBackgroundColor(white)

        ' Interfaz y Cámara
        oSpecsWin.Layout = CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        oApp.StartCommand("Compass") ' Ocultar

        oCurrentWindow.Height = 300
        oCurrentWindow.Width = 300

        ' Aplicar Vista Isométrica (Cámara 1)
        oViewer.Viewpoint3D = CType(oApp.ActiveDocument.Cameras.Item(1), Camera3D).Viewpoint3D
        oViewer.Reframe()
        oViewer.Update()
        oApp.RefreshDisplay = True

        ' --- 3. CAPTURA ---
        oViewer.CaptureToFile(CatCaptureFormat.catCaptureFormatJPEG, finalFileName)

        ' --- 4. RESTAURACIÓN ---
        oViewer.PutBackgroundColor(oldColor)
        oApp.StartCommand("Compass") ' Mostrar
        oSpecsWin.Layout = CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom

        If Not isRoot Then
            oApp.ActiveDocument.Close()
            docPrincipal.Activate()
        Else
            oCurrentWindow.WindowState = CatWindowState.catWindowStateMaximized
        End If

        Return finalFileName
    End Function

    ''' <summary>
    ''' Reemplaza caracteres inválidos del PartNumber para poder guardar el archivo en Windows.
    ''' </summary>
    Private Function CleanFileName(name As String) As String
        Dim invalidChars As String = New String(Path.GetInvalidFileNameChars())
        Dim cleaned As String = name
        For Each c As Char In invalidChars
            cleaned = cleaned.Replace(c, "_"c)
        Next
        Return cleaned
    End Function

End Class