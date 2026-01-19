Option Explicit On
Option Strict On   ' (Boxing - Unboxing pag.116) Fuerza la conversión de tipos

' Clase: "CATIA"
' Finalidad: Enlazar la aplicación con "GetObject" y asignar la propiedades Status y AppCatia


Public Class CATIA

    Implements ICATIA
    Private oAppCATIA As INFITF.Application
    Private intCATStatus As Integer

    Public Sub New()

        ICATIA_CATIAStatus()

    End Sub

    Public ReadOnly Property CATStatus As Integer Implements ICATIA.CATStatus
        Get
            Return intCATStatus
        End Get
    End Property

    Public ReadOnly Property AppCATIA As INFITF.Application Implements ICATIA.AppCATIA
        Get
            Return oAppCATIA
        End Get
    End Property


    ' ***************************************************************************************************************
    ' Procedimiento Sub "CATIAStatus"
    '
    ' Finalidad: Comprobar el estado de la aplicacion CATIA
    '
    ' Entradas: 
    '
    ' Resultados:    Devuelve un entero, siendo 5 los casos posibles:
    '                0: No hay sesion de CATIA abierta
    '                1: Hay sesion de CATIA abierta, pero no hay ninguna ventana abierta (oWindows.Count = 0 )
    '                2: Hay sesion de CATIA y al menos una ventana activa con un Conjunto
    '                3: Hay sesion de CATIA y al menos una ventana activa con un Part 
    '                4: Hay sesion de CATIA y al menos una ventana activa con un Drawing
    '                5: Hay sesion de CATIA y al menos una ventana activa con un Catalogo
    '            
    ' Observaciones: Al querer observar si hay alguna ventana abierta en la aplicacion, se observa
    '                si hay algun editor (window) abierto. Ésto se hace así porque si se fuera a evaluar
    '                a través de los documentos, puede ocurrir que no haya ninguna ventana abierta sin embargo
    '                puede quedar un archivo (ej.CATStructureDesignSample.CATfct) cargado en sesion y no ser visible.
    '
    '                Usar "isObject" o "IsNothing" (interceptacion de errores para variable objeto)     
    ' ****************************************************************************************************************

    Private Sub ICATIA_CATIAStatus() Implements ICATIA.CATIAStatus

        On Error Resume Next

        oAppCATIA = CType(GetObject(, "CATIA.Application"), INFITF.Application) ' Enlazar un objeto ya existente.(Boxing/Unboxing)

        ' 0: No hay sesion de CATIA abierta
        If IsNothing(oAppCATIA) Then
            intCATStatus = 0
            Exit Sub
        End If

        ' 1: Hay sesion de CATIA abierta, pero no hay ninguna ventana abierta
        If oAppCATIA.Windows.Count = 0 Then
            intCATStatus = 1
            Exit Sub
        End If

        Select Case TypeName(oAppCATIA.ActiveDocument)
            Case "ProductDocument"  ' 2: Hay sesion de CATIA y al menos una ventana activa con un Conjunto
                intCATStatus = 2
                Exit Sub
            Case "PartDocument"     ' 3: Hay sesion de CATIA y al menos una ventana activa con un Part
                intCATStatus = 3
                Exit Sub
            Case "DrawingDocument"  ' 4: Hay sesion de CATIA y al menos una ventana activa con un Drawing
                intCATStatus = 4
                Exit Sub
            Case "CatalogDocument"  ' 5: Hay sesion de CATIA y al menos una ventana activa con un Catalogo
                intCATStatus = 5
                Exit Sub
            Case Else
                MsgBox(Err.Number & " - " & Err.Description & " - " & Err.Source & vbCrLf)
                'poner aca un cartel que indique que no hay ninguno de los anteriores archivos abierto(verlo)
                Exit Sub
        End Select

    End Sub

End Class