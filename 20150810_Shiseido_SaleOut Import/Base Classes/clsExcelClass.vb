Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class clsExcelClass
    Private oRecordSet As SAPbobsCOM.Recordset
    Dim strQuery As String = String.Empty

    Public Function updateExcelTemplate(ByVal strPath As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim excel As Application = New Application
            Dim w As Workbook = excel.Workbooks.Open(strPath)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' Loop over all sheets.
            For i As Integer = 1 To w.Sheets.Count
                Dim sheet As Worksheet = w.Sheets(i)
                If sheet.Name = "Sheet1" Or sheet.Name = "INPUT" Then
                    Dim r As Range = sheet.UsedRange
                    Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)
                    If array IsNot Nothing Then
                        Dim bound0 As Integer = array.GetUpperBound(0)
                        Dim bound1 As Integer = array.GetUpperBound(1)
                        Dim intStyleCol, intParentCol, intDescription As Integer
                        Dim intItemCodeCol As Integer = 9999
                        For intCol As Integer = 1 To bound1
                            Dim xRng As Excel.Range = CType(sheet.Cells(1, intCol), Excel.Range)
                            Dim strValue As String = xRng.Value()
                            If strValue = "parent_item_code" Then
                                intParentCol = intCol
                            ElseIf strValue = "Description" Then
                                intDescription = intCol
                            ElseIf strValue = "U_Style" Then
                                intStyleCol = intCol
                            ElseIf strValue = "ItemCode" Then
                                intItemCodeCol = intCol
                            End If
                        Next

                        If intParentCol = 0 Or intDescription = 0 Or intStyleCol = 0 Then
                            excel.Quit()
                            ReleaseComObject(w)
                            ReleaseComObject(excel)
                            Return False
                        End If

                        For j As Integer = 2 To bound0
                            Dim xRng As Excel.Range = CType(sheet.Cells(j, intStyleCol), Excel.Range)
                            Dim strStyle As String = xRng.Value()
                            strQuery = "Select ItemCode,ItemName From OITM Where U_Style = '" + strStyle + "'"
                            oRecordSet.DoQuery(strQuery)
                            If Not oRecordSet.EoF Then
                                sheet.Cells(j, intParentCol) = oRecordSet.Fields.Item("ItemCode").Value
                                sheet.Cells(j, intDescription) = oRecordSet.Fields.Item("ItemName").Value
                                If intItemCodeCol <> 9999 Then
                                    sheet.Cells(j, intItemCodeCol) = oRecordSet.Fields.Item("ItemCode").Value
                                End If
                            Else
                                sheet.Cells(j, intParentCol) = "N/A"
                                sheet.Cells(j, intDescription) = "N/A"
                                If intItemCodeCol <> 9999 Then
                                    sheet.Cells(j, intItemCodeCol) = "N/A"
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            w.Close(SaveChanges:=True)
            excel.Quit()
            ReleaseComObject(w)
            ReleaseComObject(excel)
            Return _retVal
        Catch ex As Exception
            Throw ex
        Finally
           
        End Try
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

End Class
