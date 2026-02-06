Option Explicit On

Public Class ExcelSession
    ' Declaración de la API de Windows para obtener el PID
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    Public Property Application As Microsoft.Office.Interop.Excel.Application
    Public Property Workbooks As Microsoft.Office.Interop.Excel.Workbooks
    Public Property Workbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property Worksheets As Microsoft.Office.Interop.Excel.Sheets
    Public Property Worksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property IsReady As Boolean = False
    Public Property ErrorMessage As String = ""

    Public Sub New()

        Me.Application = New Microsoft.Office.Interop.Excel.Application
        If Me.Application Is Nothing Then Throw New Exception("No se pudo iniciar Excel.")

        Me.Application.Visible = False
        Me.Application.ScreenUpdating = False
        Me.Application.DisplayAlerts = False

    End Sub



    Sub CreateNewWorkbook()

        Try
            ' 1. Inicializar la aplicación

            Me.Application = New Microsoft.Office.Interop.Excel.Application
            If Me.Application Is Nothing Then Throw New Exception("No se pudo iniciar Excel.")

            Me.Application.Visible = False
            Me.Application.ScreenUpdating = False
            Me.Application.DisplayAlerts = False


            ' 2. Obtener la colección de libros y AGREGAR uno nuevo
            Me.Workbooks = Me.Application.Workbooks
            Me.Workbook = Me.Workbooks.Add() ' <-- Crucial: Crea el archivo

            ' 3. Asignar la hoja activa o la primera hoja del libro creado
            ' Es más seguro referenciarla desde el objeto Workbook recién creado
            Me.Worksheet = CType(Me.Workbook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
            Me.ActiveSheet = Me.Worksheet

            Me.IsReady = True


        Catch ex As System.Runtime.InteropServices.COMException
            Me.ErrorMessage = ">>> No se pudo crear nuevo archivo Excel. Error COM: " & ex.Message
        Catch ex As Exception
            Me.ErrorMessage = ">>> [ERROR] " & ex.Message
        End Try

    End Sub


    Sub GetActiveWorkbook()

        Try

            ' Excel existente
            Me.Application = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)


            ' Keystroke ESC y cerrar posibles celda en edición
            Dim excelHwnd As New IntPtr(Me.Application.Hwnd)
            Dim excelPid As Integer
            GetWindowThreadProcessId(excelHwnd, excelPid)
            AppActivate(excelPid)
            SendKeys.SendWait("{ESC}")



            If Me.Application.ActiveWorkbook Is Nothing Then
                Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
                Return
            End If


            Me.Workbooks = Me.Application.Workbooks
            Me.Workbook = Me.Application.ActiveWorkbook
            Me.ActiveSheet = CType(Me.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Me.IsReady = True


        Catch ex As System.Runtime.InteropServices.COMException
            Me.ErrorMessage = ">>> [ERROR] No se detectó ninguna instancia de Excel abierta."
        Catch ex As System.Exception
            Me.ErrorMessage = ">>> [ERROR] " & ex.Message
        End Try

    End Sub



End Class