Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsImportWizard
    Inherits clsBase
    Private strQuery As String
    Private oGrid As SAPbouiCOM.Grid

    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_Price As SAPbouiCOM.DataTable
    Private oDt_Duplicate As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable

    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ImpWiz, frm_ImpWiz)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            '  CType(oForm.Items.Item("19").Specific, SAPbouiCOM.EditText).Value = objFrmID
            oForm.Items.Item("6").TextStyle = 7
            oForm.Items.Item("20").TextStyle = 7
            oForm.Items.Item("21").TextStyle = 7
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_SellOut
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ImpWiz Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "9" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "17")
                                ElseIf pVal.ItemUID = "24" Then
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "22")
                                ElseIf (pVal.ItemUID = "7") Then 'Next
                                    'CreateItem(oForm)
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "17") Then
                                            If (oApplication.Utilities.GetExcelData(oForm, "17")) Then
                                                loadData(oForm)
                                                oForm.Items.Item("4").Enabled = True
                                                oApplication.Utilities.Message("SellOut Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "23") Then 'Next
                                    If CType(oForm.Items.Item("22").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateIEFile(oForm, "22") Then
                                            Dim strDFNPath As String = String.Empty
                                            If oApplication.Utilities.CopyFile(oForm, "22", strDFNPath) Then
                                                'If oApplication.Excel.updateExcelTemplate(strDFNPath) Then
                                                '    oApplication.Utilities.strSFilePath = strDFNPath
                                                '    oApplication.Utilities.ShowSaveFile(oForm, strDFNPath)
                                                '    Dim intVal As Integer = oApplication.SBO_Application.MessageBox("Do You Want to Open File", 2, "Yes", "No", "")
                                                '    If intVal = 1 Then
                                                '        System.Diagnostics.Process.Start(oApplication.Utilities.strDFilePath)
                                                '    End If
                                                'Else
                                                '    oApplication.Utilities.Message("Cannot Find Style,ParentItem and Description....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                'End If
                                                'If strDFNPath <> "" Then
                                                '    If File.Exists(strDFNPath) Then
                                                '        File.Delete(strDFNPath)
                                                '    End If
                                                'End If
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "3") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 1
                                    oForm.Items.Item("3").Enabled = False
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "4") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Items.Item("3").Enabled = True
                                    oForm.Items.Item("5").Enabled = True
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "5") Then
                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Do you Want to Import successful Data to Form?", 2, "Yes", "No", "")
                                    If _retVal = 1 Then
                                        If AddtoUDT1(oForm) = True Then
                                            oForm.Close()
                                        End If
                                    Else
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "13") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "16") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 3
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "18") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 4
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "26") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 5
                                    oForm.Freeze(False)
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ImpWiz Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from Z_SODC where IsError='N'")
        Dim oStatic As SAPbouiCOM.StaticText
        Dim struppercount As String = oRec.RecordCount
        oRec.DoQuery("Select * from Z_SODC where IsError='N'")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            oStatic = oForm.Items.Item("100").Specific
            oStatic.Caption = "Processing : " & intRow.ToString & " of " & strUpperCount
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strCode = oApplication.Utilities.getMaxCode("@Z_SODC", "Code")
            oUserTable = oApplication.Company.UserTables.Item("Z_SODC")
            If 1 = 1 Then
                'strCode = oApplication.Utilities.getMaxCode("@Z_PAY_OEAR", "Code")
                strECode = oRec.Fields.Item("Date").Value
                oUserTable.Code = strCode
                oUserTable.Name = strCode
                Dim dtDate As New Date(strECode.Substring(0, 4), strECode.Substring(5, 2), strECode.Substring(8, 2))
                oUserTable.UserFields.Fields.Item("U_Z_Date").Value = dtDate
                oUserTable.UserFields.Fields.Item("U_Z_Month").Value = oRec.Fields.Item("Month").Value
                oUserTable.UserFields.Fields.Item("U_Z_Year").Value = oRec.Fields.Item("Year").Value
                oUserTable.UserFields.Fields.Item("U_Z_BPCode").Value = oRec.Fields.Item("BPCode").Value
                oUserTable.UserFields.Fields.Item("U_Z_ItemCode").Value = oRec.Fields.Item("ItemCode").Value
                oUserTable.UserFields.Fields.Item("U_Z_ItemName").Value = oRec.Fields.Item("ItemName").Value
                oUserTable.UserFields.Fields.Item("U_Z_Qty").Value = oRec.Fields.Item("Quantity").Value
                oUserTable.UserFields.Fields.Item("U_Z_Currency").Value = oRec.Fields.Item("Currency").Value
                oUserTable.UserFields.Fields.Item("U_Z_ExRate").Value = oRec.Fields.Item("ExRate").Value
                oUserTable.UserFields.Fields.Item("U_Z_Value").Value = oRec.Fields.Item("Value").Value
                oUserTable.UserFields.Fields.Item("U_Z_ValueLC").Value = oRec.Fields.Item("ValueLC").Value
                oUserTable.UserFields.Fields.Item("U_Z_Stock").Value = oRec.Fields.Item("Stock").Value
                If oUserTable.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else

                End If
            End If
            oRec.MoveNext()
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True

    End Function
    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try

            oForm.DataSources.DataTables.Add("Dt_Import")
            oForm.DataSources.DataTables.Add("Dt_Price")
            oForm.DataSources.DataTables.Add("Dt_Duplicate")
            oForm.DataSources.DataTables.Add("Dt_ErrorLog")

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            oDt_Import.ExecuteQuery("Select *  From [@Z_SODC] Where 1 = 2 ")
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Price")
            oDt_Duplicate.ExecuteQuery("Select *  From [@Z_SODC] Where 1 = 2  ")
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_Duplicate

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Duplicate")
            oDt_Duplicate.ExecuteQuery("Select *  From [@Z_SODC] Where 1 = 2 ")
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable = oDt_Duplicate

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            oDt_ErrorLog.ExecuteQuery("Select Convert(VarChar(250),'') As 'Error'")
            oGrid = oForm.Items.Item("25").Specific
            oGrid.DataTable = oDt_ErrorLog

            formatGrid(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadData(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oStatic As SAPbouiCOM.StaticText
            oStatic = oForm.Items.Item("100").Specific
            oStatic.Caption = ""
            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
            strQuery += " T0.BarCode1,T0.BarCode2,SUM(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList,T0.Brand,T0.CountryOfOrigin From Z_POIM T0 "
            strQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
            strQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
            strQuery += " JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
            strQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
            strQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
            strQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode1,T0.BarCode2, "
            strQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList,T0.Brand,T0.CountryOfOrigin   "
            strQuery = "Select * from Z_SODC where IsError='N' "
            oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Price")
            strQuery = " Select T1.ItemCode As 'Item',ISNull(T2.Price,0) As 'Old Price',T0.SP_Price As 'New Price', "

            If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.CheckBox).Checked Then
                strQuery += " CONVERT(VarChar(1),'Y') As 'Check' "
            Else
                strQuery += " CONVERT(VarChar(1),'N') As 'Check' "
            End If

            strQuery += " ,SP_Currency,T0.PriceList From Z_POIM T0 JOIN OITM T1  "
            strQuery += " On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size))  "
            strQuery += " JOIN ITM1 T2 On T1.ItemCode = T2.ItemCode   "
            strQuery += " And T0.PriceList = T2.PriceList "
            strQuery += " Where T0.SP_Price <> ISNull(T2.Price,0) And ISNULL(T0.SP_Price,0) > 0 "

            strQuery = "Select * from Z_SODC where IsError='N'"
            oDt_Duplicate.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_Duplicate

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Duplicate")
            strQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
            strQuery += " T0.BarCode1,T0.BarCode2,(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList,Brand,CountryOfOrigin From Z_POIM T0 "
            strQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
            strQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
            strQuery += " JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
            strQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
            strQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
            strQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode1,T0.BarCode2, "
            strQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList,T0.Brand,T0.CountryOfOrigin   "
            strQuery += " Having Count(*) > 1  "

            strQuery = "Select * from Z_SODC where IsError='N'"
            oDt_Duplicate.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable = oDt_Duplicate

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            strQuery = " Select 'InValid Item : ' + (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As 'Error' From Z_POIM T0"
            strQuery += " LEFT OUTER JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size))"
            strQuery += " Where T1.ItemCode Is Null "
            strQuery += " Union All "
            strQuery += " Select 'InValid Purchase Currency : ' + T0.PO_Currency As 'Error' From Z_POIM T0 "
            strQuery += " LEFT OUTER JOIN OCRN T1 On T1.CurrCode = T0.PO_Currency "
            strQuery += " Where T1.CurrCode Is Null "
            strQuery += " Union All "
            strQuery += " Select 'InValid Sale Currency : ' + T0.SP_Currency As 'Error' From Z_POIM T0 "
            strQuery += " LEFT OUTER JOIN OCRN T1 On T1.CurrCode = T0.SP_Currency "
            strQuery += " Where T1.CurrCode Is Null "
            strQuery += " Union All "
            strQuery += " Select 'InValid WareHouse : ' + T0.WareHouse As 'Error' From Z_POIM T0"
            strQuery += " LEFT OUTER JOIN OWHS T1 ON T1.WhsCode = T0.WareHouse "
            strQuery += " Where T1.WhsCode Is Null "
            strQuery += " Union All "
            strQuery += " Select 'InValid Price List : ' + T0.PriceList As 'Error' From Z_POIM T0  "
            strQuery += " LEFT OUTER JOIN OPLN T1 On T1.ListNum = T0.PriceList "
            strQuery += " Where isnull(T0.PriceList,'')<>'' and T1.ListNum Is Null "

            strQuery = "Select * from Z_SODC where IsError='Y'"
            oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("25").Specific
            oGrid.DataTable = oDt_ErrorLog

            formatGrid(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("8").Specific
            '  formatAll(oForm, oGrid)
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oGrid = oForm.Items.Item("14").Specific
            formatPrice(oForm, oGrid)
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oGrid = oForm.Items.Item("15").Specific
            formatAll(oForm, oGrid)
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oGrid = oForm.Items.Item("25").Specific
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid)
        Try
            oGrid.Columns.Item("PCode").TitleObject.Caption = "Parent Code"
            oGrid.Columns.Item("Desc").TitleObject.Caption = "Description"
            oGrid.Columns.Item("Color").TitleObject.Caption = "Color"
            oGrid.Columns.Item("Size").TitleObject.Caption = "Size"
            oGrid.Columns.Item("Item").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("BarCode1").TitleObject.Caption = "Bar Code1"
            oGrid.Columns.Item("BarCode2").TitleObject.Caption = "Bar Code2"
            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("PO_Currency").TitleObject.Caption = "Purchase Currency"
            oGrid.Columns.Item("PO_Price").TitleObject.Caption = "Purchase Price"
            oGrid.Columns.Item("SP_Currency").TitleObject.Caption = "Sales Currency"
            oGrid.Columns.Item("SP_Price").TitleObject.Caption = "Sales Price"
            oGrid.Columns.Item("WareHouse").TitleObject.Caption = "Ware House"
            oGrid.Columns.Item("PriceList").TitleObject.Caption = "Price List"
            oGrid.Columns.Item("Brand").TitleObject.Caption = "Manufacture"
            oGrid.Columns.Item("CountryOfOrigin").TitleObject.Caption = "Country of Origin"

            oGrid.Columns.Item("Item").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("Item")
            oEditColumn.LinkedObjectType = "4"

            'Currency 
            oGrid.Columns.Item("PO_Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("PO_Currency")
            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            'Currency 
            oGrid.Columns.Item("SP_Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("SP_Currency")
            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            'Price List 
            oGrid.Columns.Item("PriceList").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("PriceList")
            strQuery = "Select ListNum,ListName From OPLN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("ListNum").Value, oRecordSet.Fields.Item("ListName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("Quantity").RightJustified = True
            oGrid.Columns.Item("PO_Price").RightJustified = True
            oGrid.Columns.Item("SP_Price").RightJustified = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub formatPrice(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid)
        Try
            oGrid.Columns.Item("Item").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("Old Price").TitleObject.Caption = "Old Price"
            oGrid.Columns.Item("New Price").TitleObject.Caption = "New Price"
            oGrid.Columns.Item("Check").TitleObject.Caption = "Check"


            oGrid.Columns.Item("Item").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("Item")
            oEditColumn.LinkedObjectType = "4"

            'Check Box to Add. 
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            'Currency 
            oGrid.Columns.Item("SP_Currency").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("SP_Currency")
            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            'Price List 
            oGrid.Columns.Item("PriceList").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("PriceList")
            strQuery = "Select ListNum,ListName From OPLN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("ListNum").Value, oRecordSet.Fields.Item("ListName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("Item").Editable = False
            oGrid.Columns.Item("Old Price").Editable = False
            oGrid.Columns.Item("New Price").Editable = False
            oGrid.Columns.Item("Check").Editable = True
            oGrid.Columns.Item("SP_Currency").Editable = False
            oGrid.Columns.Item("PriceList").Editable = False

            oGrid.Columns.Item("Old Price").RightJustified = True
            oGrid.Columns.Item("New Price").RightJustified = True
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub CreateItem(ByVal oForm As SAPbouiCOM.Form)
    '    Dim oItem As SAPbobsCOM.Items
    '    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '    Dim intStart As Double = 5011413610
    '    For index As Integer = 1 To 1000
    '        Dim strItemCode As String = (intStart + index).ToString() + "-" + "1234" + "-" + "36"
    '        oItem.ItemCode = strItemCode
    '        oItem.ItemCode = strItemCode
    '        oItem.ForeignName = strItemCode
    '        oItem.Add()
    '    Next
    'End Sub
#End Region

End Class
