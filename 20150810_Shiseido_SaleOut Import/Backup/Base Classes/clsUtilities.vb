Imports System.Xml
Imports System.Collections.Specialized
Imports System.IO


Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Private strFilepath As String = String.Empty
    Private oRecordSet As SAPbobsCOM.Recordset
    Private sQuery As String = String.Empty

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function getLocalCurrency(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select Maincurrncy from OADM")
        Return oTemp.Fields.Item(0).Value
    End Function

#Region "Get ExchangeRate"
    Public Function getExchangeRate(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select isNull(Rate,0) from ORTT where convert(nvarchar(10),RateDate,101)=Convert(nvarchar(10),getdate(),101) and currency='" & strCurrency & "'")
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function

    Public Function getExchangeRate(ByVal strCurrency As String, ByVal dtdate As Date) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSql As String
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSql = "Select isNull(Rate,0) from ORTT where ratedate='" & dtdate.ToString("yyyy-MM-dd") & "' and currency='" & strCurrency & "'"
            oTemp.DoQuery(strSql)
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#Region "Get DocCurrency"
    Public Function GetDocCurrency(ByVal aDocEntry As Integer) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select DocCur from OINV where docentry=" & aDocEntry)
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region

#Region "GetEditTextValues"
    Public Function getEditTextvalue(ByVal aForm As SAPbouiCOM.Form, ByVal strUID As String) As String
        Dim oEditText As SAPbouiCOM.EditText
        oEditText = aForm.Items.Item(strUID).Specific
        Return oEditText.Value
    End Function
#End Region

#Region "Get Currency"
    Public Function GetCurrency(ByVal strChoice As String, Optional ByVal aCardCode As String = "") As String
        Dim strCurrQuery, Currency As String
        Dim oTempCurrency As SAPbobsCOM.Recordset
        If strChoice = "Local" Then
            strCurrQuery = "Select MainCurncy from OADM"
        Else
            strCurrQuery = "Select Currency from OCRD where CardCode='" & aCardCode & "'"
        End If
        oTempCurrency = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempCurrency.DoQuery(strCurrQuery)
        Currency = oTempCurrency.Fields.Item(0).Value
        Return Currency
    End Function

#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0, Optional ByVal Enable As Boolean = True)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim oEditText As SAPbouiCOM.EditText
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 3

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "COPY" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left
                    .Height = objOldItem.Height
                    .Width = objOldItem.Width
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Add Condition CFL"
    Public Sub AddConditionCFL(ByVal FormUID As String, ByVal strQuery As String, ByVal strQueryField As String, ByVal sCFL As String)
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim Conditions As SAPbouiCOM.Conditions
        Dim oCond As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim sDocEntry As New ArrayList()
        Dim sDocNum As ArrayList
        Dim MatrixItem As ArrayList
        sDocEntry = New ArrayList()
        sDocNum = New ArrayList()
        MatrixItem = New ArrayList()

        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFL = oCFLs.Item(sCFL)

            Dim oRec As SAPbobsCOM.Recordset
            oRec = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(strQuery)
            oRec.MoveFirst()

            Try
                If oRec.EoF Then
                    sDocEntry.Add("")
                Else
                    While Not oRec.EoF
                        Dim DocNum As String = oRec.Fields.Item(strQueryField).Value.ToString()
                        If DocNum <> "" Then
                            sDocEntry.Add(DocNum)
                        End If
                        oRec.MoveNext()
                    End While
                End If
            Catch generatedExceptionName As Exception
                Throw
            End Try

            'If IsMatrixCondition = True Then
            '    Dim oMatrix As SAPbouiCOM.Matrix
            '    oMatrix = DirectCast(oForm.Items.Item(Matrixname).Specific, SAPbouiCOM.Matrix)

            '    For a As Integer = 1 To oMatrix.RowCount
            '        If a <> pVal.Row Then
            '            MatrixItem.Add(DirectCast(oMatrix.Columns.Item(columnname).Cells.Item(a).Specific, SAPbouiCOM.EditText).Value)
            '        End If
            '    Next
            '    If removelist = True Then
            '        For xx As Integer = 0 To MatrixItem.Count - 1
            '            Dim zz As String = MatrixItem(xx).ToString()
            '            If sDocEntry.Contains(zz) Then
            '                sDocEntry.Remove(zz)
            '            End If
            '        Next
            '    End If
            'End If

            'oCFLs = oForm.ChooseFromLists
            'oCFLCreationParams = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'If systemMatrix = True Then
            '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
            '    oCFLEvento = DirectCast(pVal, SAPbouiCOM.IChooseFromListEvent)
            '    Dim sCFL_ID As String = Nothing
            '    sCFL_ID = oCFLEvento.ChooseFromListUID
            '    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
            'Else
            '    oCFL = oForm.ChooseFromLists.Item(sCHUD)
            'End If

            Conditions = New SAPbouiCOM.Conditions()
            oCFL.SetConditions(Conditions)
            Conditions = oCFL.GetConditions()
            oCond = Conditions.Add()
            oCond.BracketOpenNum = 2
            For i As Integer = 0 To sDocEntry.Count - 1
                If i > 0 Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCond = Conditions.Add()
                    oCond.BracketOpenNum = 1
                End If

                oCond.[Alias] = strQueryField
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = sDocEntry(i).ToString()
                If i + 1 = sDocEntry.Count Then
                    oCond.BracketCloseNum = 2
                Else
                    oCond.BracketCloseNum = 1
                End If
            Next

            oCFL.SetConditions(Conditions)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function getFreightName(ByVal strExpCode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Try
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select ExpnsName From OEXD Where ExpnsCode = '" + strExpCode + "'")
            Return oTemp.Fields.Item(0).Value
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTemp1 As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select CurrCode  from OCRN")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strQuantity = strQuantity.Replace(oRec.Fields.Item(0).Value, "")
            oRec.MoveNext()
        Next
        strTemp1 = strQuantity
        strTemp = CompanyDecimalSeprator
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        If strQuantity = "" Then
            Return 0
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTemp1)
        End Try

        Return dblQuant
    End Function

    Public Sub OpenFileDialogBox(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String)
        Dim _retVal As String = String.Empty
        Try
            FileOpen()
            CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption = strFilepath
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Try
            Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
            mythr.SetApartmentState(Threading.ApartmentState.STA)
            mythr.Start()
            mythr.Join()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowFileDialog()
        Try
            Dim oDialogBox As New OpenFileDialog
            Dim strMdbFilePath As String
            Dim oProcesses() As Process
            Try
                oProcesses = Process.GetProcessesByName("SAP Business One")
                If oProcesses.Length <> 0 Then
                    For i As Integer = 0 To oProcesses.Length - 1
                        Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                        If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                            strMdbFilePath = oDialogBox.FileName
                            strFilepath = oDialogBox.FileName
                        Else
                        End If
                    Next
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function ValidateFile(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            If Path.GetExtension(strPath) <> ".txt" Then
                _retVal = False
                oApplication.Utilities.Message("In Valid File Format...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                _retVal = True
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetExcelData(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim _retVal As Boolean = False
        Dim _oDt As New DataTable
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            _oDt.TableName = "POIMPORT"
            _oDt.Columns.Add("PCode", GetType(String)).Caption = "Parent_Item_Code"
            _oDt.Columns.Add("Desc", GetType(String)).Caption = "Description"
            _oDt.Columns.Add("Color", GetType(String)).Caption = "Color"
            _oDt.Columns.Add("Size", GetType(String)).Caption = "Size"
            _oDt.Columns.Add("BarCode", GetType(String)).Caption = "BarCode"
            _oDt.Columns("BarCode").DefaultValue = ""
            _oDt.Columns.Add("Quantity", GetType(String)).Caption = "Quantity"
            _oDt.Columns.Add("PO_Currency", GetType(String)).Caption = "Cost_Curr"
            _oDt.Columns.Add("PO_Price", GetType(String)).Caption = "Cost"
            _oDt.Columns("PO_Price").DefaultValue = "0"
            _oDt.Columns.Add("SP_Currency", GetType(String)).Caption = "SP_Curr"
            _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
            _oDt.Columns("SP_Price").DefaultValue = "0"
            _oDt.Columns.Add("WareHouse", GetType(String)).Caption = "WareHouse"
            _oDt.Columns.Add("PriceList", GetType(String)).Caption = "PriceList"
            _oDt.Columns("PriceList").DefaultValue = "0"
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            If strPath.Length > 0 Then
                Dim txtRows() As String
                Dim fields() As String
                Dim oDr As DataRow
                txtRows = System.IO.File.ReadAllLines(strPath)
                Dim intRow As Integer = 0
                For Each txtrow As String In txtRows
                    If intRow = 0 Then
                        fields = txtrow.Split(vbTab)
                        For index As Integer = 0 To _oDt.Columns.Count - 1
                            If fields(index).ToLower() <> _oDt.Columns(index).Caption.ToLower() Then
                                Throw New Exception("In Valid Column : " + fields(index).ToLower() + " Should be " + _oDt.Columns(index).Caption.ToLower())
                                Exit Function
                            End If
                        Next
                    ElseIf intRow > 0 Then
                        fields = txtrow.Split(vbTab)
                        oDr = _oDt.NewRow()
                        oDr.Item(0) = fields(0) '  _oDt.Columns.Add("PCode", GetType(String)).Caption = "Parent_Item_Code"
                        oDr.Item(1) = fields(1) '  _oDt.Columns.Add("Desc", GetType(String)).Caption = "Description"
                        oDr.Item(2) = fields(2) ' _oDt.Columns.Add("Color", GetType(String)).Caption = "Color"
                        oDr.Item(3) = fields(3) '   _oDt.Columns.Add("Size", GetType(String)).Caption = "Size"
                        oDr.Item(4) = fields(4) '   _oDt.Columns.Add("BarCode", GetType(String)).Caption = "BarCode"
                        oDr.Item(5) = fields(5) '    _oDt.Columns.Add("Quantity", GetType(String)).Caption = "Quantity"
                        oDr.Item(6) = fields(6) '  _oDt.Columns.Add("PO_Currency", GetType(String)).Caption = "Cost_Curr"
                        ' oDr.Item(7) = fields(7) '     _oDt.Columns.Add("PO_Price", GetType(String)).Caption = "Cost"
                        If IsDBNull(fields(7)) Or fields(7) = "" Then
                            oDr.Item(7) = "0" '    _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
                        Else
                            oDr.Item(7) = fields(7) '    _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
                        End If
                        oDr.Item(8) = fields(8) '   _oDt.Columns.Add("SP_Currency", GetType(String)).Caption = "SP_Curr"
                        oDr.Item(9) = fields(9) '    _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
                        If IsDBNull(fields(9)) Or fields(9) = "" Then
                            oDr.Item(9) = "0" '    _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
                        Else
                            oDr.Item(9) = fields(9) '    _oDt.Columns.Add("SP_Price", GetType(String)).Caption = "Sales_Price"
                        End If

                        oDr.Item(10) = fields(10) '  _oDt.Columns.Add("WareHouse", GetType(String)).Caption = "WareHouse"
                        oDr.Item(11) = fields(11) '   _oDt.Columns.Add("PriceList", GetType(String)).Caption = "PriceList"
                        '    oDr.ItemArray = fields
                        _oDt.Rows.Add(oDr)
                    End If
                    intRow = intRow + 1
                Next
            End If
            Dim strDtXML As String = getXMLstring(_oDt)
            oRecordSet.DoQuery("Exec [Insert_POImport] '" + strDtXML + "'")
            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getXMLstring(ByVal oDt As System.Data.DataTable) As String
        Dim _retVal As String = String.Empty
        Try
            Dim sr As New System.IO.StringWriter()
            oDt.WriteXml(sr, False)
            _retVal = sr.ToString()
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub setData(ByVal oIMForm As SAPbouiCOM.Form, ByVal strFormID As String)
        Dim oForm As SAPbouiCOM.Form = oApplication.SBO_Application.Forms.Item(strFormID)
        If IsNothing(oForm) Then
            Exit Sub
        End If
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim intRow As Integer = 1
            If oForm.TypeEx = frm_GR_INVENTORY Then
                oMatrix = oForm.Items.Item("13").Specific
                sQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
                sQuery += " T0.BarCode,SUM(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList From Z_POIM T0 "
                sQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
                sQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
                sQuery += "  JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
                sQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
                sQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
                sQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode, "
                sQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList   "
                sQuery += " Having SUM(T0.Quantity) > 0 "
                oRecordSet.DoQuery(sQuery)

                If Not oRecordSet.EoF Then
                    If oMatrix.RowCount > 1 Then
                        oMatrix.Clear()
                        oMatrix.AddRow(1, -1)
                    End If
                    While Not oRecordSet.EoF
                        oForm.Freeze(True)
                        If CType(oIMForm.Items.Item("11").Specific, SAPbouiCOM.CheckBox).Checked Then
                            '   importBarCodePrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                            UpdateBarcode(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                        End If
                        If CType(oIMForm.Items.Item("12").Specific, SAPbouiCOM.CheckBox).Checked Then
                            '   importBarCodePrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                            UpdateSalesPrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                        End If
                        CType(oIMForm.Items.Item("6").Specific, SAPbouiCOM.StaticText).Caption = "Importing Item : " + oRecordSet.Fields.Item("Item").Value
                        CType(oMatrix.Columns.Item("1").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("Item").Value
                        Try
                            CType(oMatrix.Columns.Item("26").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("BarCode").Value
                        Catch ex As Exception

                        End Try

                        CType(oMatrix.Columns.Item("9").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("Quantity").Value
                        CType(oMatrix.Columns.Item("10").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("PO_Currency").Value.ToString() + " " + oRecordSet.Fields.Item("PO_Price").Value.ToString()
                        CType(oMatrix.Columns.Item("15").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("WareHouse").Value
                        oRecordSet.MoveNext()
                        intRow = intRow + 1
                        oForm.Freeze(False)
                    End While
                End If

            Else
                oMatrix = oForm.Items.Item("38").Specific
                sQuery = " Select T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) As Item, "
                sQuery += " T0.BarCode,SUM(T0.Quantity) As Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList From Z_POIM T0 "
                sQuery += " JOIN OITM T1 On T1.ItemCode = (Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)) "
                sQuery += " JOIN OCRN T2 On T2.CurrCode = T0.PO_Currency "
                sQuery += " JOIN OCRN T3 On T3.CurrCode = T0.SP_Currency "
                sQuery += " JOIN OWHS T4 ON T4.WhsCode = T0.WareHouse "
                sQuery += " Left Outer JOIN OPLN T5 On T5.ListNum = T0.PriceList "
                sQuery += " Group By T0.PCode,T0.[Desc],T0.Color,T0.Size,(Convert(VarChar,T0.PCode) + '-' + Convert(VarChar,T0.Color) + '-' + Convert(Varchar,T0.Size)),T0.BarCode, "
                sQuery += " T0.Quantity,T0.PO_Currency,T0.PO_Price,T0.SP_Currency,T0.SP_Price,T0.WareHouse,T0.PriceList   "
                sQuery += " Having SUM(T0.Quantity) > 0 "
                oRecordSet.DoQuery(sQuery)

                If Not oRecordSet.EoF Then
                    If oMatrix.RowCount > 1 Then
                        oMatrix.Clear()
                        oMatrix.AddRow(1, -1)
                    End If
                    While Not oRecordSet.EoF
                        oForm.Freeze(True)
                        If CType(oIMForm.Items.Item("11").Specific, SAPbouiCOM.CheckBox).Checked Then
                            '   importBarCodePrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                            UpdateBarcode(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                        End If
                        If CType(oIMForm.Items.Item("12").Specific, SAPbouiCOM.CheckBox).Checked Then
                            '   importBarCodePrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                            UpdateSalesPrice(oIMForm, oRecordSet.Fields.Item("Item").Value, oRecordSet.Fields.Item("BarCode").Value, oRecordSet.Fields.Item("PriceList").Value, oRecordSet.Fields.Item("SP_Currency").Value, oRecordSet.Fields.Item("SP_Price").Value)
                        End If
                        CType(oIMForm.Items.Item("6").Specific, SAPbouiCOM.StaticText).Caption = "Importing Item : " + oRecordSet.Fields.Item("Item").Value
                        CType(oMatrix.Columns.Item("1").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("Item").Value
                        CType(oMatrix.Columns.Item("4").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("BarCode").Value
                        CType(oMatrix.Columns.Item("11").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("Quantity").Value
                        CType(oMatrix.Columns.Item("14").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("PO_Currency").Value.ToString() + " " + oRecordSet.Fields.Item("PO_Price").Value.ToString()
                        CType(oMatrix.Columns.Item("24").Cells.Item(intRow).Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("WareHouse").Value
                        oRecordSet.MoveNext()
                        intRow = intRow + 1
                        oForm.Freeze(False)
                    End While
                End If

            End If


        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub AddBarCode(ByVal aItemCode As String, ByVal aBarCode As String, ByVal aUOMEntry As Integer)
        Dim lpCmpSer As SAPbobsCOM.ICompanyService
        Dim lpBCSer As SAPbobsCOM.IBarCodesService
        Dim lpBCPar As SAPbobsCOM.IBarCodeParams
        Dim lpBC As SAPbobsCOM.IBarCode
        Dim lRS As SAPbobsCOM.IRecordset
        Dim lUomEntry As Long, lBcdEntry As Long
        Try

            lRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            lUomEntry = aUOMEntry
            lpCmpSer = oApplication.Company.GetCompanyService
            lpBCSer = lpCmpSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BarCodesService)
            lpBC = lpBCSer.GetDataInterface(SAPbobsCOM.BarCodesServiceDataInterfaces.bsBarCode)
            lpBC.ItemNo = aItemCode
            lpBC.UoMEntry = aUOMEntry
            lpBC.BarCode = aBarCode
            lpBCPar = lpBCSer.Add(lpBC)
            '  MsgBox(lpBCPar.AbsEntry)

        Catch ex As Exception
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub importBarCodePrice(ByVal oForm As SAPbouiCOM.Form, ByVal strItemCode As String, ByVal strBarCode As String, ByVal strPriceList As String, ByVal strSPCurrency As String, ByVal strPrice As String)
        Dim oItem As SAPbobsCOM.Items = Nothing
        Try
            Dim oItemPrice As SAPbobsCOM.Items_Prices
            Dim oBarcodes As SAPbobsCOM.BarCodesService
            Dim intStatus, intBarcodes, intDefaultPOUOM As Integer
            Dim oRec, oRec1 As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If oItem.GetByKey(strItemCode) Then
                If CType(oForm.Items.Item("11").Specific, SAPbouiCOM.CheckBox).Checked Then
                    intDefaultPOUOM = oItem.DefaultPurchasingUoMEntry
                    intBarcodes = oItem.BarCodes.Count
                    oRec.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "'")
                    If oRec.RecordCount > 0 Then
                        oRec1.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "' and BcdCode='" & strBarCode & "'")
                        If oRec1.RecordCount <= 0 Then
                            AddBarCode(strItemCode, strBarCode, intDefaultPOUOM)
                        End If
                    Else
                        AddBarCode(strItemCode, strBarCode, intDefaultPOUOM)
                    End If
                End If
                If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.CheckBox).Checked Then
                    oItemPrice = oItem.PriceList
                    oItemPrice.SetCurrentLine(strPriceList - 1)
                    If strSPCurrency.Length > 0 Then
                        oItemPrice.Currency = strSPCurrency.Trim()
                    End If
                    If CDbl(strPrice) > 0 Then
                        oItemPrice.Price = CDbl(strPrice)
                    End If
                End If
                intStatus = oItem.Update()
                If intStatus <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
        End Try
    End Sub
    Public Sub UpdateSalesPrice(ByVal oForm As SAPbouiCOM.Form, ByVal strItemCode As String, ByVal strBarCode As String, ByVal strPriceList As String, ByVal strSPCurrency As String, ByVal strPrice As String)
        Dim oItem As SAPbobsCOM.Items = Nothing
        Try
            Dim oItemPrice As SAPbobsCOM.Items_Prices
            Dim oBarcodes As SAPbobsCOM.BarCodesService
            Dim intStatus, intBarcodes, intDefaultPOUOM As Integer
            Dim oRec, oRec1 As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If CType(oForm.Items.Item("12").Specific, SAPbouiCOM.CheckBox).Checked And getDocumentQuantity(strPrice) > 0 And getDocumentQuantity(strPriceList) <> 0 Then
                If oItem.GetByKey(strItemCode) Then
                    oItemPrice = oItem.PriceList
                    oItemPrice.SetCurrentLine(strPriceList - 1)
                    If strSPCurrency.Length > 0 Then
                        oItemPrice.Currency = strSPCurrency.Trim()
                    End If
                    If CDbl(strPrice) > 0 Then
                        oItemPrice.Price = CDbl(strPrice)
                    End If
                End If
                intStatus = oItem.Update()
                If intStatus <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
        End Try
    End Sub
    Public Sub UpdateBarcode(ByVal oForm As SAPbouiCOM.Form, ByVal strItemCode As String, ByVal strBarCode As String, ByVal strPriceList As String, ByVal strSPCurrency As String, ByVal strPrice As String)
        Dim oItem As SAPbobsCOM.Items = Nothing
        Try
            Dim oItemPrice As SAPbobsCOM.Items_Prices
            Dim oBarcodes As SAPbobsCOM.BarCodesService
            Dim intStatus, intBarcodes, intDefaultPOUOM As Integer
            Dim oRec, oRec1 As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If CType(oForm.Items.Item("11").Specific, SAPbouiCOM.CheckBox).Checked And strBarCode <> "" Then
                If oItem.GetByKey(strItemCode) Then
                    intDefaultPOUOM = oItem.DefaultPurchasingUoMEntry
                    intBarcodes = oItem.BarCodes.Count
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                    oRec.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "'")
                    If oRec.RecordCount > 0 Then
                        Dim s As String = "Select * from OBCD where ItemCode='" & strItemCode & "' and BcdCode='" & strBarCode & "'"
                        oRec1.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "' and BcdCode='" & strBarCode & "'")
                        If oRec1.RecordCount <= 0 Then
                            AddBarCode(strItemCode, strBarCode, intDefaultPOUOM)
                        End If
                    Else
                        AddBarCode(strItemCode, strBarCode, intDefaultPOUOM)
                    End If
                End If
                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If oItem.GetByKey(strItemCode) Then
                    oRec.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "'")
                    If oRec.RecordCount > 0 Then
                        Dim s As String = "Select * from OBCD where ItemCode='" & strItemCode & "' and BcdCode='" & strBarCode & "'"
                        oRec1.DoQuery("Select * from OBCD where ItemCode='" & strItemCode & "' and BcdCode='" & strBarCode & "'")
                        If oRec1.RecordCount > 0 Then
                            oItem.BarCode = strBarCode
                            oItem.Update()
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                        End If
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                    End If
                End If


                'intStatus = oItem.Update()
                'If intStatus <> 0 Then
                '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
        End Try
    End Sub
End Class

