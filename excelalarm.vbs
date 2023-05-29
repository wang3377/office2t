Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim excelApp As Excel.Application
    Dim excelBook As Excel.Workbook
    Dim excelSheet As Excel.Worksheet

    Private Sub btnCreateExcel_Click(sender As Object, e As EventArgs) Handles btnCreateExcel.Click
        ' 创建Excel应用程序对象
        excelApp = New Excel.Application()
        excelBook = excelApp.Workbooks.Add()
        excelSheet = excelBook.Sheets(1)

        ' 设置表头
        excelSheet.Cells(1, 1).Value = "闹钟时间"

        ' 添加闹钟时间
        For i As Integer = 2 To 6
            Dim alarmTime As DateTime = DateTime.Now.AddMinutes(i * 10)
            excelSheet.Cells(i, 1).Value = alarmTime.ToString("yyyy-MM-dd HH:mm:ss")
        Next

        ' 保存Excel文件并关闭应用程序
        Dim savePath As String = "C:\Path\To\Save\Excel\File.xlsx"
        excelBook.SaveAs(savePath)
        excelBook.Close()
        excelApp.Quit()

        ' 释放资源
        ReleaseObject(excelSheet)
        ReleaseObject(excelBook)
        ReleaseObject(excelApp)

        MessageBox.Show("Excel文件创建成功！")
    End Sub

    ' 释放COM对象
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
