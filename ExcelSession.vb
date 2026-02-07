Option Explicit On
Option Strict On

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

    End Sub



    Sub CreateNewWorkbook()
        Try
            Me.Application = New Microsoft.Office.Interop.Excel.Application
            With Me
                .Application.Visible = False
                .Application.ScreenUpdating = False
                .Application.DisplayAlerts = False
                .Workbooks = Me.Application.Workbooks
                .Workbook = Me.Workbooks.Add()
                .Worksheets = Me.Worksheets
                .Worksheet = CType(Me.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                .ActiveSheet = Me.Worksheet
                .IsReady = True
            End With
        Catch ex As Exception
            MsgBox("Error al iniciar Excel: " & ex.Message, MsgBoxStyle.Critical)
            Me.IsReady = False
        End Try
    End Sub



    Sub GetActiveWorkbook()

        ' Excel existente
        Try
            Me.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.Runtime.InteropServices.COMException
            Me.ErrorMessage = ">>> [ERROR] GetActiveObject(Excel.Application)"
            Exit Sub
        End Try
        If Me.Application.ActiveWorkbook Is Nothing Then
            Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
            Return
        End If

        ' Keystroke ESC y cerrar posibles celda en edición
        Dim excelHwnd As New IntPtr(Me.Application.Hwnd)
        Dim excelPid As Integer
        GetWindowThreadProcessId(excelHwnd, excelPid)
        AppActivate(excelPid)
        SendKeys.SendWait("{ESC}")



        Me.Workbooks = Me.Application.Workbooks
        Me.Workbook = Me.Application.ActiveWorkbook
        Me.ActiveSheet = CType(Me.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

        ' bandera
        Me.IsReady = True

    End Sub

End Class