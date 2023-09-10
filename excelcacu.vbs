Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        ' 从文本框中获取输入的两个数字
        Dim num1 As Double
        Dim num2 As Double

        If Double.TryParse(txtNum1.Text, num1) AndAlso Double.TryParse(txtNum2.Text, num2) Then
            ' 执行计算
            Dim result As Double = num1 + num2

            ' 将结果显示在标签上
            lblResult.Text = "结果: " & result

            ' 将结果写入到 Excel
            Dim excelApp As New Excel.Application()
            Dim excelBook As Excel.Workbook = excelApp.Workbooks.Add()
            Dim excelSheet As Excel.Worksheet = excelBook.Sheets(1)

            excelSheet.Cells(1, 1).Value = "数字1"
            excelSheet.Cells(1, 2).Value = "数字2"
            excelSheet.Cells(1, 3).Value = "结果"
            excelSheet.Cells(2, 1).Value = num1
            excelSheet.Cells(2, 2).Value = num2
            excelSheet.Cells(2, 3).Value = result

            ' 保存 Excel 文件
            Dim savePath As String = "C:\Path\To\Save\Excel\File.xlsx"
            excelBook.SaveAs(savePath)
            excelBook.Close()
            excelApp.Quit()

            ' 释放资源
            ReleaseObject(excelSheet)
            ReleaseObject(excelBook)
            ReleaseObject(excelApp)
        Else
            MessageBox.Show("请输入有效的数字。")
        End If
    End Sub

    ' 释放 COM 对象
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
