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
    Public Property Sheets As Microsoft.Office.Interop.Excel.Sheets
    Public Property Worksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property NCUWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property NCUSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property IsReady As Boolean = False
    Public Property ErrorMessage As String = ""

    Sub New()


    End Sub



    Function CreateNewWorkbook() As Microsoft.Office.Interop.Excel.Workbook
        Try
            With Me
                .Application = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False,
                .ScreenUpdating = False,
                .DisplayAlerts = False
            }
                .Workbook = .Application.Workbooks.Add()
                .IsReady = True
                Return .Workbook
            End With
        Catch ex As Exception
            Me.ErrorMessage = "Error al iniciar Excel: " & ex.Message
            MsgBox(Me.ErrorMessage, MsgBoxStyle.Critical)
            Me.IsReady = False
            Return Nothing
        End Try
    End Function

    Sub GetActiveWorkbook()
        Try
            Me.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            If Me.Application.ActiveWorkbook Is Nothing Then
                Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
                Return
            End If
        Catch ex As Exception
            Me.ErrorMessage = ">>> [ERROR] No se pudo conectar con Excel: " & ex.Message
            Return
        End Try
        ' Desbloqueo de celda en edición
        Dim pid As Integer
        GetWindowThreadProcessId(New IntPtr(Me.Application.Hwnd), pid)
        AppActivate(pid)
        SendKeys.SendWait("{ESC}")
        With Me
            .Workbooks = .Application.Workbooks
            .Workbook = .Application.ActiveWorkbook
            .ActiveSheet = CType(.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            .IsReady = True
        End With
    End Sub


    Function GetNCUSheet(ByVal ncuPath As String) As Microsoft.Office.Interop.Excel.Worksheet
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = Me.NCUWorkbook.Worksheets
        Try
            With Me
                .Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                .Application.ScreenUpdating = False
                .Application.DisplayAlerts = False

                .NCUWorkbook = .Application.Workbooks.Open(ncuPath, ReadOnly:=True)
                .NCUWorkbook.Windows(1).Visible = False
                NCUSheet = CType(.NCUWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
            End With
            Return NCUSheet
        Catch ex As Exception
            Me.ErrorMessage = "No se pudo abrir el archivo NCU: " & ex.Message
            MsgBox(Me.ErrorMessage)
            Return Nothing ' Asegura un retorno aunque falle
        Finally
            If Me.Application IsNot Nothing Then
                Me.Application.ScreenUpdating = True
                Me.Application.DisplayAlerts = True
            End If
        End Try
    End Function

    Public Sub CloseNCU()
        If Me.NCUWorkbook IsNot Nothing Then
            Me.NCUWorkbook.Close(SaveChanges:=False)
            Me.NCUWorkbook = Nothing
        End If
    End Sub



End Class