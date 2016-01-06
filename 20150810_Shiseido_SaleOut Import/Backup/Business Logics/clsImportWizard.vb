Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsImportWizard
    Inherits clsBase
    Private strQuery As String
    Private oGrid As SAPbouiCOM.Grid

    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_Duplicate As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable

    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal objFrmID As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ImpWiz, frm_ImpWiz)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            CType(oForm.Items.Item("19").Specific, SAPbouiCOM.EditText).Value = objFrmID
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
                                ElseIf (pVal.ItemUID = "7") Then 'Next
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "17") Then
                                            If (oApplication.Utilities.GetExcelData(oForm, "17")) Then
                                                loadData(oForm)
                                                oForm.Items.Item("4").Enabled = True
                                                oApplication.Utilities.Message("Purchase Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
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
                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you Want to Import Data to Form?", 2, "Yes", "No", "")
                                    If _retVal = 1 Then
                                        oApplication.Utilities.setData(oForm, CType(oForm.Items.Item("19").Specific, SAPbouiCOM.EditText).Value)
                                        oApplication.Utilities.Message("Purchase Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        oForm.Close()
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
    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try

            oForm.DataSources.DataTables.Add("Dt_Import")
            oForm.DataSources.DataTables.Add("Dt_Duplicate")
            oForm.DataSources.DataTables.Add("Dt_ErrorLog")

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            oDt_Import.ExecuteQuery("Select PCode,[Desc],Color,Size,(Convert(VarChar,PCode) + '-' + Convert(VarChar,Color) + '-' + Convert(Varchar,Size)) As Item ,Barcode,Quantity,PO_Currency,PO_Price,SP_Currency,SP_Price,WareHouse,PriceList From Z_POIM Where 1 = 2 ")
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Duplicate")
            oDt_Duplicate.ExecuteQuery("Select PCode,[Desc],Color,Size,(Convert(VarChar,PCode) + '-' + Convert(VarChar,Color) + '-' + Convert(Varchar,Size)) As Item ,Barcode,Quantity,PO_Currency,PO_Price,SP_Currency,SP_Price,WareHouse,PriceList From Z_POIM Where 1 = 2 ")
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_Duplicate

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            oDt_ErrorLog.ExecuteQuery("Select Convert(VarChar(250),'') As 'Error'")
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable = oDt_ErrorLog

            formatGrid(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadData(ByVal oForm As SAPbouiCOM.Form)
        Try

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
            strQuery += " T0.BarCode,SUM(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList From Z_POIM T0 "
            strQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
            strQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
            strQuery += " JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
            strQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
            strQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
            strQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode, "
            strQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList   "
            oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import
          

            oDt_Duplicate = oForm.DataSources.DataTables.Item("Dt_Duplicate")
            strQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
            strQuery += " T0.BarCode,(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList From Z_POIM T0 "
            strQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
            strQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
            strQuery += " JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
            strQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
            strQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
            strQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode, "
            strQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList   "
            strQuery += " Having Count(*) > 1  "
            oDt_Duplicate.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("14").Specific
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
            oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("15").Specific
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
            formatAll(oForm, oGrid)
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oGrid = oForm.Items.Item("14").Specific
            formatAll(oForm, oGrid)
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next
            oGrid = oForm.Items.Item("15").Specific
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
            oGrid.Columns.Item("BarCode").TitleObject.Caption = "Bar Code"
            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("PO_Currency").TitleObject.Caption = "Purchase Currency"
            oGrid.Columns.Item("PO_Price").TitleObject.Caption = "Purchase Price"
            oGrid.Columns.Item("SP_Currency").TitleObject.Caption = "Sales Currency"
            oGrid.Columns.Item("SP_Price").TitleObject.Caption = "Sales Price"
            oGrid.Columns.Item("WareHouse").TitleObject.Caption = "Ware House"
            oGrid.Columns.Item("PriceList").TitleObject.Caption = "Price List"

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
#End Region

End Class
