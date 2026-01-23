
Module To_Excel
    Sub CompletaListView2(oProduct As ProductStructureTypeLib.Product,
                      oSheetListView As Microsoft.Office.Interop.Excel.Worksheet,
                      strDir As String,
                      oDiccType3 As Dictionary(Of String, PwrProduct))

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
















'Dim ListViewPartNumCol As String = "C"        ' PartNumber está en C1
'Dim ListViewDescriptionCol As String = "D"    ' Description está en D1
'Dim ListViewCantidad As String = "E"           ' Cantidad está en E1
'Dim ListViewConjuntoParteCol As String = "F"  ' FileType (o Conjunto Parte) está en F1
'Dim ListViewMadeOrBoughtCol As String = "G"   ' Source está en G1
'Dim ListViewVendor_Code_IDCol As String = "H" ' Vendor_Code_ID está en H1
'Dim ListViewMaterialCol As String = "I"       ' Material está en I1
'Dim ListViewTrataTermicoCol As String = "N"   ' (Si quieres usar la columna N)
'Dim ListViewImageCol As String = "M"          ' <--- ESTA ES LA QUE BUSCABAS




' **************************************************************************************************************************
' IMPORTANTE !!!!!!!!!!!!!!!!!!!!!!!!
' Habria que hacer un procedimiento para comprobar que no hay "\" (backlash) en los nombres de los archivos!
' Importante: Naming Rules:
' Create a New product and give it a name with a backslash: "Cube2\Elementary Source".
' Save this product And you will see that all the words before the backslash, And the backslash, are Not taken into account.






' **************************************************************************************************************************
' IMPORTANTE !!!!!!!!!!!!!!!!!!!!!!!!
' Habria que hacer un procedimiento para comprobar que no hay "backlash" en los nombres de los archivos!
' "Valvula 5_2 24V.CATPart": este archivo tenía backlash y daba error
' Importante: Naming Rules:
' Create a New product and give it a name with a backslash: "Cube2\Elementary Source".
' Save this product And you will see that all the words before the backslash, And the backslash, are Not taken into account.











' Se referencia la aplicación a partir del oProduct que se recibe por parámetro.
' Otra opción puede ser, recibir por parámetro la nueva ventana que fue abierta.
' Dim oAppCATIA As INFITF.Application = oProduct.Application


' Bloquea la interacccion para evitar cambio de "ActiveDocument" por un click sobre CATIA.
' Si bien ya hay que desactivar la interaccion antes de entrar a este método, lo vuelvo a hacer acá.
' Tambien deshabilita los mensajes emergentes.
' Hay un problema con la función "Interactive", porque al terminar la ejecución, CATIA queda con menu greyed-out
' oAppCATIA.Interactive = False
' oAppCATIA.DisplayFileAlerts = False


' Este objeto selection va a servir para hacer un "OpenInNewWindow" por cada product para poder sacar las capturas de pantalla
' Ver que apenas creo el objeto, luego hago "clear"
'  Dim oSelection As INFITF.Selection = oAppCATIA.ActiveDocument.Selection : oSelection.Clear()





' Esta parte de parametros no esta bien pulida
'  Dim oUserRefParameters As KnowledgewareTypeLib.Parameters




'Para completar la pestaña "ListView" se necesita un diccionario de tipo 3
' Dim oDiccType3 As Dictionary(Of String, PwrProduct) = Diccionarios.DiccT3_Rev2(oProduct)






' Reset de la ventana
'oViewer3D.PutBackgroundColor(oldColor)
'oAppCATIA.StartCommand("Compass")
'oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom
'oViewer3D.Reframe()
'oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateNormal
'oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateMaximized





' Una vez que ya he tomado la captura de pantalla cierro la ventana activa
' Hay un condicional para no cerrar el producto raíz
' No se si es la mejor manera, hay que ver si existe otra forma mejor.
' lo que hace esta forma es, ver que tipo de padre tiene el product.
' El product raíz va a tener de padre un objeto de tipo collection (products)
' Creo que a lo que me refería es a que el Raiz no va a tener un padre de tipo "collection".
' El padre del root que es?

'If TypeName(kvp.Value.Product.Parent) = "Products" Then
'    kvp.Value.Product.Application.ActiveDocument.Close()
'End If







'Sub CapturarImagenProducto(oProd As ProductStructureTypeLib.Product, path As String)
'    Dim oApp As INFITF.Application = oProd.Application

'    ' Abrir en ventana nueva
'    Dim oSel As INFITF.Selection = oApp.ActiveDocument.Selection
'    oSel.Clear()
'    oSel.Add(oProd)
'    oApp.StartCommand("Open in New Window")
'    oSel.Clear()

'    ' Configurar Visor
'    Dim oWin As INFITF.Window = oApp.ActiveWindow
'    Dim oSpecsWin As INFITF.SpecsAndGeomWindow = CType(oWin, INFITF.SpecsAndGeomWindow)
'    Dim oViewer As INFITF.Viewer3D = CType(oSpecsWin.Viewers.Item(1), INFITF.Viewer3D)

'    ' Vista Isométrica y Fondo Blanco
'    Dim white(2) As Object : white(0) = 1 : white(1) = 1 : white(2) = 1
'    oViewer.PutBackgroundColor(white)
'    oSpecsWin.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly

'    oWin.Height = 300 : oWin.Width = 300
'    oViewer.Viewpoint3D = CType(oApp.ActiveDocument.Cameras.Item(1), INFITF.Camera3D).Viewpoint3D
'    oViewer.Reframe()
'    oViewer.Update()

'    ' Capturar y Cerrar
'    oViewer.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, path)
'    oApp.ActiveDocument.Close()
'End Sub







' *********************************************************
' Seleccion y Apertura en Nueva Ventana y Captura de Imagen
' *********************************************************

' Toma la selección del product actual
'oSelection.Add(kvp.Value.Product)
'oAppCATIA.StartCommand("Open in New Window")
'oAppCATIA.RefreshDisplay = True
'oSelection.Clear()

'' --- Configuración de Ventana y Visor ---
'Dim oCurrentWindow As INFITF.Window = oAppCATIA.ActiveWindow
'Dim oSpecsAndGeomWindow As INFITF.SpecsAndGeomWindow = CType(oCurrentWindow, INFITF.SpecsAndGeomWindow)
'Dim oViewer3D As INFITF.Viewer3D = CType(oSpecsAndGeomWindow.Viewers.Item(1), INFITF.Viewer3D)

'' --- Configuración de Cámara (Vista Isométrica) ---
'Dim oCameras As INFITF.Cameras = oAppCATIA.ActiveDocument.Cameras
'Dim oCamera3D As INFITF.Camera3D = CType(oCameras.Item(1), INFITF.Camera3D) ' Camara 1 = isometrica

'' --- Preparación del Fondo (Blanco) ---
'Dim oldColor(2), whiteColor(2) As Object
'whiteColor(0) = 1 : whiteColor(1) = 1 : whiteColor(2) = 1
'oViewer3D.GetBackgroundColor(oldColor)
'oViewer3D.PutBackgroundColor(whiteColor)

'' --- Ajustes de Interfaz para Captura ---
'oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly
'oAppCATIA.StartCommand("Compass") ' Alterna visibilidad del compass

'' --- Aplicación de Vista y Redimensionado ---
'oCurrentWindow.Height = 300
'oCurrentWindow.Width = 300
'oViewer3D.Viewpoint3D = oCamera3D.Viewpoint3D ' Aplica orientación isométrica
'oViewer3D.Reframe() ' Centra el objeto en la nueva vista

'' --- Refresco Final ---
'oViewer3D.Update()
'oAppCATIA.RefreshDisplay = True


' Toma la captura y la amlmacena en el directorio indicado
' oViewer3D.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, sFullPathFileName)


' ****************************************************************************
' Propiedades de Usuario
' (1) Material en Bruto
' (2) Material
' NOTA: Si se agregan propiedades de usuario, entonces hay que agregarlas acá
' ****************************************************************************
'oUserRefParameters = kvp.Value.Product.ReferenceProduct.UserRefProperties
'If oUserRefParameters.Count <> 0 Then 'Primero ver si hay propiedades de usuario
'    For Each Parametro As KnowledgewareTypeLib.Parameter In oUserRefParameters  'Luego, buscar cuales hay
'        If Parametro.Name = "Material" Then
'            oSheetListView.Cells(i, ListViewMaterialCol) = oUserRefParameters.Item("Material").ValueAsString()
'        End If
'        If Parametro.Name = "Material en Bruto" Then
'            oSheetListView.Cells(i, ListViewMaterialenBrutoCol) = oUserRefParameters.Item("Material en Bruto").ValueAsString()
'        End If
'    Next
'End If

'Dim oCurrentWindow As INFITF.Window = oAppCATIA.ActiveWindow
'Dim oSpecsAndGeomWindow As INFITF.SpecsAndGeomWindow = CType(oCurrentWindow, INFITF.SpecsAndGeomWindow)
' Dim oViewer3D As INFITF.Viewer3D = CType(oSpecsAndGeomWindow.Viewers.Item(1), INFITF.Viewer3D)
'Dim oViewPoint3D As INFITF.Viewpoint3D
'Dim oCameras As INFITF.Cameras = oAppCATIA.ActiveDocument.Cameras
'Dim oCamera3D As INFITF.Camera3D = oCameras.Item(1)
'Dim arrBackgroundColor(2) As Object
'Dim arrWhiteColor(2) As Object
'arrWhiteColor(0) = 1
'arrWhiteColor(1) = 1
'arrWhiteColor(2) = 1
'oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly
'oViewPoint3D = oCamera3D.Viewpoint3D
'oViewer3D.Viewpoint3D = oCamera3D.Viewpoint3D
'oAppCATIA.StartCommand("Compass")
'oViewer3D.GetBackgroundColor(arrBackgroundColor)
'oViewer3D.PutBackgroundColor(arrWhiteColor)
'oCurrentWindow.Height = 300
'oCurrentWindow.Width = 300
'oViewer3D.Update()
'oAppCATIA.RefreshDisplay = True
'oViewer3D.Reframe()








'' *****************************************************
'' Sobre la Ventana abierta se realiza todo lo siguiente
'' *****************************************************
'' oCurrentWindow: esto esta asegurado porque al ingresar a este metodo antes se ejecutó: " oAppCATIA.Interactive = False"
'' Pero creo que se podría mejorar si se utiliza Windows.Item("Nombre de la ventana que corresponde al product que ingresó por parámetro"):
'' y luego se activa esa ventana. Aunque para hacer esto, se necesita ver que tipo de doc se va a abrir en la nueva ventana.
'Dim oCurrentWindow As INFITF.Window = oAppCATIA.ActiveWindow

'Dim oSpecsAndGeomWindow As INFITF.SpecsAndGeomWindow = CType(oCurrentWindow, INFITF.SpecsAndGeomWindow) '(QueryInterface) (Pag.235)

'' El oViewer3D antes estaba resuelto con la línea: "Dim objViewer3D As INFITF.Viewer3D = objAppCATIA.ActiveWindow.ActiveViewer"
'Dim oViewer3D As INFITF.Viewer3D = CType(oSpecsAndGeomWindow.Viewers.Item(1), INFITF.Viewer3D)  '(QueryInterface) (Pag.235)

'Dim oViewPoint3D As INFITF.Viewpoint3D

'' En las cámaras pasa lo mismo que en "oCurrentWindow", ya que para poder referenciar las camaras se necesita el documento activo
'' Trabajar con la sentencia "ActiveDocuemnt" o "ActiveWindow" en un método que recibe por parámetro el oProduct tiene cierta inconsistencia,
'' ya que debería estár todo referenciado a el parámetro que ha ingresado como argumento. De todas formas, se utiliza el " oAppCATIA.Interactive = False"
'' para que el usuario no cambie el documento o ventana activa.
'Dim oCameras As INFITF.Cameras = oAppCATIA.ActiveDocument.Cameras

'Dim oCamera3D As INFITF.Camera3D = oCameras.Item(1)  ' las primeras 7 camaras son de tipo "Camera3D". Camara 1 = isometrica  '(QueryInterface) (Pag.235)

'' Estos arrays son para el color de fondo de la ventana. Estos arreglos contiene tipo genérico "Object", ya que en la documentación de CATIA
'' se indica que deben ser del tipo Variant
'Dim arrBackgroundColor(2) As Object
'Dim arrWhiteColor(2) As Object
'arrWhiteColor(0) = 1
'arrWhiteColor(1) = 1
'arrWhiteColor(2) = 1

'' ***************************************************************************
'' Seteo de la ventana para poder tomar la captura de pantalla
'' ***************************************************************************

'oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly  ' Apaga el arbol de especificaciones

'oViewPoint3D = oCamera3D.Viewpoint3D   ' Setteo de la vista en la que se va a tomar la captura

'' Esto es lo que realmente mueve la "cámara" del visor
'oViewer3D.Viewpoint3D = oCamera3D.Viewpoint3D

'oAppCATIA.StartCommand("Compass")  ' Oculta el compass

'oViewer3D.GetBackgroundColor(arrBackgroundColor)  ' Toma el color actual de fondo y lo almacena en "arrBackgroundColor" para luego reestablecerlo.
'oViewer3D.PutBackgroundColor(arrWhiteColor)  ' Luego, setea el fondo de color a blanco para tomar la captura

'oCurrentWindow.Height = 300  ' Altura de la pantalla para la captura
'oCurrentWindow.Width = 300   ' Ancho de la pantalla para la captura

'oViewer3D.Update()
'oAppCATIA.RefreshDisplay = True
'oViewer3D.Reframe()

