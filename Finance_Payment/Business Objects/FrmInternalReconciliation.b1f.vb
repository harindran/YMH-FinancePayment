Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("FOITR", "Business Objects/FrmInternalReconciliation.b1f")>
    Friend Class FrmInternalReconciliation
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Public WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Dim FormCount As Integer = 0
        Dim objRs As SAPbobsCOM.Recordset
        Dim strSQL As String
        Public Shared objFDT As New DataTable
        Public Shared oSelectedDT As New DataTable
        Private WithEvents objCheck As SAPbouiCOM.CheckBox
        Public Shared objActualDT As New DataTable
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix1 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("ldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("ldocnum").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.EditText1 = CType(Me.GetItem("t_docnum").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldr2").Specific, SAPbouiCOM.Folder)
            Me.StaticText3 = CType(Me.GetItem("lremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tremark").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("ltotdue").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("ttotdue").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lpaydate").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tpaydate").Specific, SAPbouiCOM.EditText)
            Me.Button2 = CType(Me.GetItem("btnadjent").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler DataAddBefore, AddressOf Me.Form_DataAddBefore

        End Sub
        'June 27th 2022 - Removed Round off in the screen

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("FOITR", Me.FormCount)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                odbdsHeader = objform.DataSources.DBDataSources.Item(CType(0, Object))
                odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object))
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "MIOITR")
                oSelectedDT.Clear()
                If oSelectedDT.Columns.Count = 0 Then
                    oSelectedDT.Columns.Add("paytot", GetType(Double))
                    oSelectedDT.Columns.Add("#", GetType(String))
                End If
                objActualDT.Clear()
                If objActualDT.Columns.Count = 0 Then
                    For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                        If iCol <> 1 Then
                            If Matrix1.Columns.Item(iCol).UniqueID = "paytot" Then
                                objActualDT.Columns.Add(Matrix1.Columns.Item(iCol).UniqueID, GetType(Double))
                            Else
                                objActualDT.Columns.Add(Matrix1.Columns.Item(iCol).UniqueID)
                            End If

                        End If
                    Next
                End If
                objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objform.Items.Item("tpaydate").Specific.string = PayInitDate.ToString("yyyyMMdd")
                objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery("select distinct T0.""BnkChgAct"" as ""BCGAcct"",T0.""LinkAct_3"" as ""CahAcct"",""LinkAct_24"" as ""Rounding"",""GLGainXdif"",""GLLossXdif"",""ExDiffAct"" " &
                              ",(Select ""SumDec"" from OADM) as ""SumDec"",(Select ""RateDec"" from OADM) as ""RateDec""" &
                              "from OACP T0 left join OFPR T1 on T1.""Category""=T0.""PeriodCat"" where T0.""PeriodCat""=(Select ""Category"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"")")
                If objRs.RecordCount > 0 Then
                    If objRs.Fields.Item(5).Value.ToString <> "" Then ForexDiff = objRs.Fields.Item(5).Value.ToString
                    If objRs.Fields.Item(6).Value.ToString <> "" Then SumRound = objRs.Fields.Item(6).Value.ToString
                    If objRs.Fields.Item(7).Value.ToString <> "" Then RateRound = objRs.Fields.Item(7).Value.ToString

                End If
                objform.Settings.Enabled = True
                Field_Setup()
                Matrix1.Columns.Item("#").Visible = False
                Matrix1.CommonSetting.FixedColumnsCount = 2
                Folder0.Item.Click()
                If Not LoadData(Query) Then
                    objform.Close()
                End If
                'objform.Items.Item("btnadjent").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Button2 As SAPbouiCOM.Button

#End Region

#Region "Form Events"

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            'Try
            '    If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
            '    If EditText1.Value = "" Then objaddon.objapplication.StatusBar.SetText("Numbering Series Not Found. Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '    If Not CDbl(EditText4.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '    'objFDT.Clear()
            '    'objFDT = build_Matrix_DataTable("paytot")
            '    If objActualDT.Rows.Count > 0 Then objFDT = objActualDT Else objFDT = build_Matrix_DataTable("paytot")
            '    Dim Branch As String 'Amt
            '    Dim Amt As Double
            '    Dim Line As Integer = 0
            '    Dim ErrorFlag As Boolean = False
            '    Try
            '        If objFDT.Rows.Count = 0 Then objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '        Branch = objFDT.Rows(0)("branchc").ToString


            '        Dim otherBranchDT = From dr In objFDT.AsEnumerable()
            '                            Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
            '                            Select New With {                   'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0  'Ph <> Branch And
            '.branch = Ph,
            '.LengthSum = Math.Round(drg.Sum(Function(dr) dr.Field(Of Double)("paytot")), SumRound)
            '}
            '        '    Dim otherBranchDT = From dr In objFDT.AsEnumerable()
            '        '                        Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .DTLine = dr.Field(Of String)("#")} Into drg = Group
            '        '                        Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
            '        '.branch = Ph.branch,
            '        '.line = Ph.DTLine,
            '        '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
            '        '}
            '        objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
            '        For Each RowID In otherBranchDT
            '            'Line = CInt(RowID.line.ToString())
            '            'Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
            '            Amt = CDbl(RowID.LengthSum)
            '            If CDbl(Amt) = 0 Then
            '                If BranchReconciliation(objFDT, RowID.branch.ToString()) = False Then
            '                    ErrorFlag = True
            '                    ': BubbleEvent = False : Exit Sub
            '                End If
            '            Else
            '                If JournalEntry(objFDT, RowID.branch.ToString(), CDbl(Amt)) = False Then
            '                    ErrorFlag = True
            '                    objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
            '                End If
            '            End If
            '        Next

            '        If ErrorFlag = True Then
            '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '            Try
            '                objform.Freeze(True)
            '                Matrix1.FlushToDataSource()
            '                For rowNum As Integer = 0 To odbdsDetails.Size - 1
            '                    odbdsDetails.SetValue("U_JENo", rowNum, "")
            '                    odbdsDetails.SetValue("U_RecoNo", rowNum, "")
            '                Next
            '                Matrix1.LoadFromDataSource()
            '            Catch ex As Exception
            '            Finally
            '                objform.Freeze(False)
            '            End Try
            '            objform.Refresh()
            '            objaddon.objapplication.StatusBar.SetText("Error while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '        Else
            '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '            objaddon.objapplication.StatusBar.SetText("Internal Reconciliations Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '        End If
            '    Catch ex As Exception
            '    End Try

            'Catch ex As Exception
            'End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix1.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                If Button2.Item.Enabled = False Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objaddon.objapplication.MessageBox("Do you want to Create the adjustment JE ?", 2, "Yes", "No") <> 1 Then Exit Sub
                'Dim GetPayBranchDT = From dr In objActualDT.AsEnumerable()
                '                     Where (dr.Field(Of String)("object") = "24" Or dr.Field(Of String)("object") = "46")
                '                     Select New With {Key .branch = dr.Field(Of String)("branchc"), Key .BPCode = dr.Field(Of String)("cardc")}


                Dim Branch As String = "", BPCode As String = ""
                ''If GetPayBranchDT.Count = 0 Then Exit Sub
                'For Each RowID In GetPayBranchDT
                '    Branch = RowID.branch.ToString()
                '    BPCode = RowID.BPCode.ToString()
                '    Exit For
                'Next
                Create_Manual_JE(objActualDT, Branch, BPCode)
                'Create_Manual_JE(objActualDT, objaddon.objglobalmethods.getSingleValue("Select ""BPLId"" from OBPL where ""MainBPL""='Y'"))
            Catch ex As Exception

            End Try

        End Sub

#End Region

#Region "Functions"

        Private Function LoadData(ByVal Query As String) As Boolean
            Try
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(Query)
                Matrix1.Clear()
                odbdsDetails.Clear()
                If objRs.RecordCount > 0 Then
                    objaddon.objapplication.StatusBar.SetText("Loading data Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objform.Freeze(True)
                    While Not objRs.EoF
                        Matrix1.AddRow()
                        Matrix1.GetLineData(Matrix1.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                        odbdsDetails.SetValue("U_Select", 0, objRs.Fields.Item("Selected").Value.ToString)
                        odbdsDetails.SetValue("U_TransId", 0, objRs.Fields.Item("TransId").Value.ToString)
                        odbdsDetails.SetValue("U_TLine", 0, objRs.Fields.Item("Line_ID").Value.ToString)
                        odbdsDetails.SetValue("U_DebCred", 0, objRs.Fields.Item("DebCred").Value.ToString)
                        odbdsDetails.SetValue("U_CardType", 0, objRs.Fields.Item("CardType").Value.ToString)
                        odbdsDetails.SetValue("U_Origin", 0, objRs.Fields.Item("Origin").Value.ToString)
                        odbdsDetails.SetValue("U_OriginNo", 0, objRs.Fields.Item("DocNum").Value.ToString)
                        odbdsDetails.SetValue("U_DocEntry", 0, objRs.Fields.Item("DocEntry").Value.ToString)
                        odbdsDetails.SetValue("U_CardCode", 0, objRs.Fields.Item("CardCode").Value.ToString)
                        odbdsDetails.SetValue("U_CardName", 0, objRs.Fields.Item("CardName").Value.ToString)
                        odbdsDetails.SetValue("U_DocDate", 0, objRs.Fields.Item("DocDate").Value)
                        odbdsDetails.SetValue("U_DocCur", 0, objRs.Fields.Item("DocCur").Value.ToString)
                        odbdsDetails.SetValue("U_TotalFC", 0, CDbl(objRs.Fields.Item("DocTotalFC").Value.ToString))
                        odbdsDetails.SetValue("U_BalDueFC", 0, CDbl(objRs.Fields.Item("BalanceFC").Value.ToString))
                        odbdsDetails.SetValue("U_Total", 0, CDbl(objRs.Fields.Item("DocTotal").Value.ToString))
                        odbdsDetails.SetValue("U_BalDue", 0, CDbl(objRs.Fields.Item("Balance").Value.ToString))
                        odbdsDetails.SetValue("U_PayTotal", 0, objRs.Fields.Item("Balance").Value.ToString)
                        odbdsDetails.SetValue("U_Memo", 0, objRs.Fields.Item("LineMemo").Value.ToString)
                        odbdsDetails.SetValue("U_BranchId", 0, objRs.Fields.Item("BPLId").Value.ToString)
                        odbdsDetails.SetValue("U_BranchNam", 0, objRs.Fields.Item("BPLName").Value.ToString)
                        odbdsDetails.SetValue("U_Object", 0, objRs.Fields.Item("ObjType").Value.ToString)
                        odbdsDetails.SetValue("U_Pay", 0, objRs.Fields.Item("Balance").Value.ToString)
                        odbdsDetails.SetValue("U_Ref1", 0, objRs.Fields.Item("Ref1").Value.ToString)
                        odbdsDetails.SetValue("U_Ref2", 0, objRs.Fields.Item("Ref2").Value.ToString)
                        odbdsDetails.SetValue("U_Ref3", 0, objRs.Fields.Item("Ref3").Value.ToString)
                        'objform.DataSources.UserDataSources.Item("UD_0").Value = objRs.Fields.Item("Balance").Value.ToString
                        Matrix1.SetLineData(Matrix1.VisualRowCount)
                        objRs.MoveNext()
                    End While
                    Matrix1.AutoResizeColumns()
                    objaddon.objapplication.StatusBar.SetText("Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objform.Freeze(False)
                    Return True
                Else
                    objaddon.objapplication.StatusBar.SetText("No records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                objform.Freeze(False)
                Return False
            End Try
        End Function

        Private Function build_Matrix_DataTable(ByVal sKeyFieldID As String) As DataTable
            Dim objcheckbox As SAPbouiCOM.CheckBox
            Try
                Dim oDT As New DataTable
                'Add all of the columns by unique ID to the DataTable
                For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                    'Skip invisible columns
                    'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                    If iCol <> 1 Then
                        oDT.Columns.Add(Matrix1.Columns.Item(iCol).UniqueID)
                    End If
                Next
                'Now, add all of the data into the DataTable
                For iRow As Integer = 1 To Matrix1.VisualRowCount
                    objcheckbox = Matrix1.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = oDT.NewRow
                        For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                            'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                            If iCol <> 1 Then
                                oRow.Item(Matrix1.Columns.Item(iCol).UniqueID) = Matrix1.Columns.Item(iCol).Cells.Item(iRow).Specific.Value
                            End If
                        Next
                        'If the Key field has no value, then the row is empty, skip adding it.
                        If oRow(sKeyFieldID).ToString.Trim = 0 Then Continue For
                        oDT.Rows.Add(oRow)
                    End If
                Next

                Return oDT
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Private Function Matrix_DataTable(ByVal Row As Integer, ByVal ColName As String, Optional ByVal AdjJE As Boolean = False) As DataTable
            Try
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim DataFlag As Boolean
                objcheckbox = Matrix1.Columns.Item("select").Cells.Item(Row).Specific
                Dim oRow As DataRow = objActualDT.NewRow
                If objcheckbox.Checked = True Or AdjJE = True Then
                    If objActualDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                            If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                If ColName <> "" Then
                                    objActualDT.Rows(DTRow)(Matrix1.Columns.Item(ColName).UniqueID) = Matrix1.Columns.Item(ColName).Cells.Item(Row).Specific.Value
                                    DataFlag = True
                                    Exit For
                                End If
                            End If
                        Next
                        If DataFlag = False Then
                            For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                                If iCol <> 1 Then
                                    oRow.Item(Matrix1.Columns.Item(iCol).UniqueID) = Matrix1.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                                End If
                            Next
                            objActualDT.Rows.Add(oRow)
                        End If
                    Else
                        For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                            If iCol <> 1 Then
                                oRow.Item(Matrix1.Columns.Item(iCol).UniqueID) = Matrix1.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                            End If
                        Next
                        objActualDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                        If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            objActualDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                Return objActualDT
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Sub Calculate_Total()
            Try
                Dim objDT As New DataTable
                Dim value As Double
                objDT.Columns.Add("paytot", GetType(Double))

                Dim objcheckbox As SAPbouiCOM.CheckBox
                For iRow As Integer = 1 To Matrix1.VisualRowCount
                    objcheckbox = Matrix1.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = objDT.NewRow
                        oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(iRow).Specific.Value
                        objDT.Rows.Add(oRow)
                    End If
                Next
                For i As Integer = 0 To objDT.Rows.Count - 1
                    If objDT.Rows(i)("paytot").ToString <> "" Then
                        value += CDbl(objDT.Rows(i)("paytot").ToString)
                    End If
                Next
                odbdsHeader.SetValue("U_Total", 0, value) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Calc_Total_14092021(ByVal Row As Integer, Optional ByVal AdjJE As Boolean = False)
            Try
                Dim value As Double
                Dim DataFlag As Boolean
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim oRow As DataRow = oSelectedDT.NewRow
                objcheckbox = Matrix1.Columns.Item("select").Cells.Item(Row).Specific

                If objcheckbox.Checked = True Or AdjJE = True Then
                    If oSelectedDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                            If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                oSelectedDT.Rows(DTRow)("paytot") = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                                DataFlag = True
                                Exit For
                            End If
                        Next
                        If DataFlag = False Then
                            oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                            oRow.Item(Matrix1.Columns.Item("#").UniqueID) = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value
                            oSelectedDT.Rows.Add(oRow)
                        End If
                    Else
                        oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                        oRow.Item(Matrix1.Columns.Item("#").UniqueID) = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value
                        oSelectedDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                        If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            oSelectedDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                For i As Integer = 0 To oSelectedDT.Rows.Count - 1
                    If Val(oSelectedDT.Rows(i)("paytot").ToString) <> 0 Then
                        value = Math.Round(value + CDbl(oSelectedDT.Rows(i)("paytot")), SumRound)
                        'value = value + CDbl(oSelectedDT.Rows(i)("paytot"))
                        'value += CDbl(oSelectedDT.Rows(i)("paytot"))
                    End If
                Next
                odbdsHeader.SetValue("U_Total", 0, value) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Field_Setup()
            Try
                Matrix1.Columns.Item("branchc").Visible = False
                Matrix1.Columns.Item("pay").Visible = False
                Matrix1.Columns.Item("tranline").Visible = False
                Matrix1.Columns.Item("debcred").Visible = False
                Matrix1.Columns.Item("cardtype").Visible = False
                Matrix1.Columns.Item("object").Visible = False
            Catch ex As Exception

            End Try
        End Sub

        Private Function JournalEntry(ByVal InDT As DataTable, ByVal Branch As String, ByVal JEAmount As Double) As Boolean
            Try
                Dim TransId As String = "", GLCode As String, Series As String, CardCode As String = "", MatLine As String = "", CardType As String = ""
                Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim Amount As Double
                Dim DTLine As Integer = 0
                Try
                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                            TransId = InDT.Rows(DTRow)("jeno").ToString
                            MatLine = CInt(InDT.Rows(DTRow)("#").ToString)
                            DTLine = DTRow
                            CardCode = InDT.Rows(DTRow)("cardc").ToString()
                            CardType = InDT.Rows(DTRow)("cardtype").ToString()
                            Exit For
                        End If
                    Next
                    If TransId = "" Then
                        Try
                            objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                            objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                            Dim oEdit As SAPbouiCOM.EditText
                            oEdit = objform.Items.Item("tdocdate").Specific
                            Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            objjournalentry.ReferenceDate = DocDate  'Now.Date.ToString("yyyyMMdd") 
                            'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
                            objjournalentry.TaxDate = DocDate  ' ConvertDate.ToString("dd/MM/yy") 
                            objjournalentry.Reference = "Int Rec Payment JE"
                            objjournalentry.Memo = "Posted thro' recon On: " & Now.ToString
                            objjournalentry.UserFields.Fields.Item("U_IntRecNo").Value = EditText1.Value
                            If Localization = "IN" Then
                                If objaddon.HANA Then
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                                Else
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                                End If
                            Else
                                objjournalentry.AutoVAT = BoYesNoEnum.tNO
                                objjournalentry.AutomaticWT = BoYesNoEnum.tNO
                                If objaddon.HANA Then
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                                Else
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                                End If
                            End If
                            If Series <> "" Then objjournalentry.Series = Series
                            Amount = IIf(JEAmount < 0, -JEAmount, JEAmount)
                            If JEAmount < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                            GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                            objjournalentry.Lines.AccountCode = GLCode
                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()
                            objjournalentry.Lines.ShortName = CardCode
                            If JEAmount < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                            'If CardType = "C" Then
                            '    If JEAmount < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                            'Else
                            '    If JEAmount < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                            'End If
                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()

                            'For Row As Integer = 0 To InDT.Rows.Count - 1
                            '    If InDT.Rows(Row)("branchc").ToString = Branch And InDT.Rows(Row)("jeno").ToString = String.Empty Then
                            '        JEAmount = IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, -CDbl(InDT.Rows(Row)("paytot").ToString), CDbl(InDT.Rows(Row)("paytot").ToString))
                            '        GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                            '        objjournalentry.Lines.AccountCode = GLCode
                            '        If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 And InDT.Rows(Row)("cardtype").ToString = "C" Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
                            '        objjournalentry.Lines.BPLID = Branch
                            '        objjournalentry.Lines.Add()
                            '        If CardCode = "" Then CardCode = InDT.Rows(Row)("cardc").ToString()
                            '        If MatLine = "" Then MatLine = CInt(InDT.Rows(Row)("#").ToString())
                            '        objjournalentry.Lines.ShortName = InDT.Rows(Row)("cardc").ToString
                            '        If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 And InDT.Rows(Row)("cardtype").ToString = "C" Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
                            '        objjournalentry.Lines.BPLID = Branch
                            '        objjournalentry.Lines.Add()
                            '    End If
                            'Next
                            If objjournalentry.Add <> 0 Then
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objaddon.objapplication.SetStatusBarMessage("Journal: " & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                Return False
                            Else
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                TransId = objaddon.objcompany.GetNewObjectKey()
                                Matrix1.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String = TransId
                                If Matrix1.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String <> "" And Matrix1.Columns.Item("recono").Cells.Item(CInt(MatLine)).Specific.String = "" Then
                                    If MultiBranch_InternalReconciliation(InDT, MatLine, TransId, CardCode, Branch) = False Then
                                        objaddon.objapplication.StatusBar.SetText("MultiBranch_InternalReco" & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Return False
                                    End If
                                End If
                                objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                        Catch ex As Exception
                            'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    Else
                        Try
                            If Not InDT.Rows(DTLine)("jeno").ToString = String.Empty And (InDT.Rows(DTLine)("recono").ToString = String.Empty Or InDT.Rows(DTLine)("recono").ToString = "0") Then
                                If MultiBranch_InternalReconciliation(InDT, MatLine, TransId, InDT.Rows(DTLine)("cardc").ToString, Branch) = True Then
                                    objaddon.objapplication.StatusBar.SetText("MultiBranch_InternalReco" & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Return False
                                End If
                            End If
                        Catch ex As Exception
                            objaddon.objapplication.SetStatusBarMessage("JE Rec " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                    objrecset = Nothing
                    Return True
                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try
            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function MultiBranch_InternalReconciliation(ByVal InDT As DataTable, ByVal Line As Integer, ByVal transid As String, ByVal BPCode As String, ByVal Branch As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                'openTrans.ReconDate = DocumentDate

                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim RecAmount As Double
                Dim Row As Integer = 0
                objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"" from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where  T1.""TransId""='" & transid & "' and T1.""ShortName""='" & BPCode & "'")
                If objRs.RecordCount > 0 Then
                    Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    openTrans.ReconDate = DocDate
                    For Rec As Integer = 0 To objRs.RecordCount - 1
                        If Val(objRs.Fields.Item(0).Value.ToString) <> 0 Then
                            'RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                            RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString)
                            openTrans.InternalReconciliationOpenTransRows.Add()
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = transid
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString) 'InDT.Rows(Row)("tranline").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount 'CDbl(objRs.Fields.Item(0).Value.ToString)
                            Row += 1
                        End If
                        objRs.MoveNext()
                    Next

                End If
                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                        'If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = -CDbl(InDT.Rows(DTRow)("paytot").ToString) Else RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                        'RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                        openTrans.InternalReconciliationOpenTransRows.Add()
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(InDT.Rows(DTRow)("paytot").ToString) '
                        Row += 1
                    End If
                Next
                Dim Reconum As Integer = 0
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciled Error : " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                End Try

                Reconum = reconParams.ReconNum
                Matrix1.Columns.Item("recono").Cells.Item(Line).Specific.String = Reconum
                objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function BranchReconciliation(ByVal InDT As DataTable, ByVal Branch As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                Dim RecAmount As Double
                Dim Row As Integer = 0
                Dim Line As String = ""
                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = Branch And Not InDT.Rows(DTRow)("recono").ToString = String.Empty Then
                        Line = CInt(InDT.Rows(DTRow)("#").ToString)
                        Exit For
                    End If
                Next
                If Line = "" Then
                    Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    openTrans.ReconDate = DocDate
                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                            If Line = "" Then Line = InDT.Rows(DTRow)("#").ToString
                            'If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                            If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = -CDbl(InDT.Rows(DTRow)("paytot").ToString) Else RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                            openTrans.InternalReconciliationOpenTransRows.Add()
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                            Row += 1
                        End If
                    Next
                    Dim Reconum As Integer = 0
                    Try
                        reconParams = service.Add(openTrans)
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Reconciled Error..." & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                    End Try
                    Reconum = reconParams.ReconNum
                    Matrix1.Columns.Item("recono").Cells.Item(CInt(Line)).Specific.String = Reconum
                    objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                    GC.Collect()
                End If

                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function GetBranchName(ByVal BCode As String) As String
            Try
                Dim BName As String
                BName = objaddon.objglobalmethods.getSingleValue("Select ""BPLName"" from OBPL where ""BPLId""='" & BCode & "' ")

                Return BName
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Private Sub Create_Manual_JE(ByVal InDT As DataTable, ByVal Branch As String, ByVal BPCode As String)
            Dim objJEform, objtempform As SAPbouiCOM.Form
            Dim cmbBranch As SAPbouiCOM.ComboBox
            Dim Amt As Double
            Dim CardCode As String = ""
            Try
                '    Dim AdjDT = From dr In InDT.AsEnumerable()
                '                Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                '                Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) <> 0
                '                Select New With {
                '.branch = Ph,
                '.LengthSum = Math.Round(drg.Sum(Function(dr) dr.Field(Of Double)("paytot")), SumRound)
                '}

                objaddon.objapplication.Menus.Item("1540").Activate()
                objJEform = objaddon.objapplication.Forms.ActiveForm
                objJEform = objaddon.objapplication.Forms.Item(objJEform.UniqueID)
                'objJEform.Freeze(True)
                objJEform.Visible = True
                If objJEform.IsSystem = False Then
                    objaddon.objapplication.StatusBar.SetText("Error while loading the Journal Entry screen...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Exit Sub
                End If
                'objaddon.objapplication.StatusBar.SetText("Loading Journal Entry screen. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'objJEform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                'objJEform.State = SAPbouiCOM.BoFormStateEnum.fs_Restore
                objMatrix = objJEform.Items.Item("76").Specific
                'cmbBranch = objJEform.Items.Item("1320002034").Specific
                'If Branch <> "" Then cmbBranch.Select(Branch, SAPbouiCOM.BoSearchKey.psk_ByValue)
                'If AdjDT.Count = 0 And BPCode <> "" Then
                '    objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Click()
                '    objaddon.objapplication.SendKeys("^{TAB}")
                '    objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String = BPCode
                '    objtempform = objaddon.objapplication.Forms.ActiveForm
                '    objtempform.Items.Item("6").Specific.String = BPCode
                '    objtempform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'End If
                'For Each RowID In AdjDT
                '    Amt = CDbl(RowID.LengthSum)
                '    If CDbl(Amt) <> 0 Then
                '        For DTRow As Integer = 0 To InDT.Rows.Count - 1
                '            If InDT.Rows(DTRow)("branchc").ToString = RowID.branch Then
                '                CardCode = InDT.Rows(DTRow)("cardc").ToString()
                '                Exit For
                '            End If
                '        Next
                '        If objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String <> "" Then objMatrix.AddRow(1)
                '        objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Click()
                '        objaddon.objapplication.SendKeys("^{TAB}")
                '        objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String = CardCode
                '        objtempform = objaddon.objapplication.Forms.ActiveForm
                '        objtempform.Items.Item("6").Specific.String = CardCode
                '        objtempform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '        'objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String = CardCode
                '        Dim Adjamt As Double = IIf(Amt < 0, -Amt, Amt)
                '        If Amt < 0 Then objMatrix.Columns.Item("5").Cells.Item(objMatrix.VisualRowCount - 1).Specific.String = Adjamt Else objMatrix.Columns.Item("6").Cells.Item(objMatrix.VisualRowCount - 1).Specific.String = Adjamt
                '        objMatrix.Columns.Item("1320002030").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Select(RowID.branch, SAPbouiCOM.BoSearchKey.psk_ByValue)

                '        If objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String <> "" Then objMatrix.AddRow(1)
                '        objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.String = ForexDiff
                '        If Amt < 0 Then objMatrix.Columns.Item("6").Cells.Item(objMatrix.VisualRowCount - 1).Specific.String = Adjamt Else objMatrix.Columns.Item("5").Cells.Item(objMatrix.VisualRowCount - 1).Specific.String = Adjamt
                '        objMatrix.Columns.Item("1320002030").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Select(RowID.branch, SAPbouiCOM.BoSearchKey.psk_ByValue)
                '    End If
                'Next
                objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Click()
                pModal = True
                'objaddon.objapplication.StatusBar.SetText("Journal Entry details loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'objJEform.Freeze(False)
            Catch ex As Exception
                'objJEform.Freeze(False)
            End Try
        End Sub


#End Region

#Region "Matrix Events"

        Private Sub Matrix1_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.PressedAfter
            Try
                objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "select" Then
                    If objCheck.Checked = True Then
                        'Matrix1.SelectRow(pVal.Row, True, True)
                        Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    Else
                        'Matrix1.SelectRow(pVal.Row, False, True)
                        Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Matrix1.Item.BackColor)
                        Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                    End If
                    'Calculate_Total()
                    Calc_Total_14092021(pVal.Row)
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix1_LinkPressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.LinkPressedBefore
            Dim ColItem As SAPbouiCOM.Column = Matrix1.Columns.Item("origin")
            Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
            Dim oForm As SAPbouiCOM.Form
            Try
                Select Case pVal.ColUID
                    Case "origin"
                        If Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "13" Then
                            objaddon.objapplication.Menus.Item("2053").Activate()  'AR Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "14" Then
                            objaddon.objapplication.Menus.Item("2055").Activate()  'AR Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "18" Then
                            objaddon.objapplication.Menus.Item("2308").Activate()  'AP Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "19" Then
                            objaddon.objapplication.Menus.Item("2309").Activate()  'AP Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "203" Then
                            objaddon.objapplication.Menus.Item("2071").Activate()  'AR DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "204" Then
                            objaddon.objapplication.Menus.Item("2317").Activate()  'AP DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "24" Then
                            objaddon.objapplication.Menus.Item("2817").Activate()  'Incoming Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "46" Then
                            objaddon.objapplication.Menus.Item("2818").Activate()  'Outgoing Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        Else 'If Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "30" Then
                            'objlink.LinkedObjectType = "30" ' Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String
                            'objlink.Item.LinkTo = "trannum"
                            objaddon.objapplication.Menus.Item("1540").Activate()  'Outgoing Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("5").Specific.String = Matrix1.Columns.Item("trannum").Cells.Item(pVal.Row).Specific.String
                            'oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        End If
                End Select

            Catch ex As Exception
                oForm.Freeze(False)
                oForm = Nothing
            End Try

        End Sub

        Private Sub Matrix1_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.ValidateBefore
            Try
                'If pVal.ItemChanged = False Then Exit Sub
                If pVal.InnerEvent = True Then Exit Sub
                Dim Balance, PayTotal As Double
                'Dim ActTotal As Double
                If Val(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) <> 0 Then Balance = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) Else Balance = 0
                If Val(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) <> 0 Then PayTotal = CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) Else PayTotal = 0
                'If Val(Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) > 0 Then ActTotal = CDbl(Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) Else ActTotal = 0
                Select Case pVal.ColUID
                    Case "paytot"
                        If pVal.InnerEvent = False Then
                            If PayTotal <= 0 Then
                                PayTotal = -CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                                Balance = -CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                If PayTotal > Balance Or PayTotal = 0 Then
                                    'Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                End If
                            ElseIf PayTotal > 0 Then
                                If PayTotal > Balance Then
                                    'Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                End If
                            Else
                                Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                            End If
                            If pVal.ItemChanged = True Then
                                objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                                objCheck.Checked = True
                                Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                                'Matrix1.SelectRow(pVal.Row, True, True)
                                Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                            End If
                            Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String = Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String
                            'Calculate_Total()
                            Calc_Total_14092021(pVal.Row)
                        End If
                        Matrix1.SetCellWithoutValidation(pVal.Row, "details", Matrix1.Columns.Item("details").Cells.Item(pVal.Row).Specific.String)
                        'Case "select"
                        '    objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                        '    If odbdsDetails.GetValue("U_Select", 0) = "Y" Then
                        '        Calc_Total_14092021(pVal.Row, True)
                        '        Matrix_DataTable(pVal.Row, pVal.ColUID, True)
                        '    End If
                        'If objCheck.Checked = True Then
                        'Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                        'Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))

                        'End If
                End Select
                If pVal.ItemChanged = True Then
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub DeleteRow()
            Try
                Dim Flag As Boolean = False
                Dim objSelect As SAPbouiCOM.CheckBox
                'Matrix1.Columns.Item("select").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                For i As Integer = Matrix1.VisualRowCount To 1 Step -1
                    objSelect = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                    If objSelect.Checked = False Then
                        Matrix1.DeleteRow(i)
                        odbdsDetails.RemoveRecord(i - 1)
                        Flag = True
                    End If
                Next
                If Flag = True Then
                    For i As Integer = 1 To Matrix1.VisualRowCount
                        objSelect = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Matrix1.Columns.Item("#").Cells.Item(i).Specific.String = i
                        End If
                    Next
                    objform.Freeze(False)
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objform.EnableMenu("1282", False)
                DeleteRow()
                For i As Integer = 1 To Matrix1.VisualRowCount
                    objCheck = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        If objCheck.Checked = True Then
                            'Matrix1.SelectRow(i, True, True)
                            Matrix1.CommonSetting.SetRowBackColor(i, Color.PeachPuff.ToArgb)
                        End If
                    End If
                Next
                Matrix1.AutoResizeColumns()
                objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If EditText1.Value = "" Then objaddon.objapplication.StatusBar.SetText("Numbering Series Not Found. Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                If Not CDbl(EditText4.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                'objFDT.Clear()
                'objFDT = build_Matrix_DataTable("paytot")
                If objActualDT.Rows.Count > 0 Then objFDT = objActualDT Else objFDT = build_Matrix_DataTable("paytot")
                Dim Branch As String 'Amt
                Dim Amt As Double
                Dim Line As Integer = 0
                Dim ErrorFlag As Boolean = False
                Try
                    If objFDT.Rows.Count = 0 Then objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                    Branch = objFDT.Rows(0)("branchc").ToString


                    Dim otherBranchDT = From dr In objFDT.AsEnumerable()
                                        Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                                        Select New With {                   'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0  'Ph <> Branch And
            .branch = Ph,
            .LengthSum = Math.Round(drg.Sum(Function(dr) dr.Field(Of Double)("paytot")), SumRound)
            }
                    '    Dim otherBranchDT = From dr In objFDT.AsEnumerable()
                    '                        Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .DTLine = dr.Field(Of String)("#")} Into drg = Group
                    '                        Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
                    '.branch = Ph.branch,
                    '.line = Ph.DTLine,
                    '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
                    '}
                    If (objaddon.objapplication.MessageBox("You cannot change this document after you have added it. Continue?", 2, "Yes", "No") <> 1) Then BubbleEvent = False : Return
                    objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                    For Each RowID In otherBranchDT
                        'Line = CInt(RowID.line.ToString())
                        'Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
                        Amt = CDbl(RowID.LengthSum)
                        If CDbl(Amt) = 0 Then
                            If BranchReconciliation(objFDT, RowID.branch.ToString()) = False Then
                                ErrorFlag = True
                                ': BubbleEvent = False : Exit Sub
                            End If
                        Else
                            If JournalEntry(objFDT, RowID.branch.ToString(), CDbl(Amt)) = False Then
                                ErrorFlag = True
                                objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
                            End If
                        End If
                    Next

                    If ErrorFlag = True Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Try
                            objform.Freeze(True)
                            Matrix1.FlushToDataSource()
                            For rowNum As Integer = 0 To odbdsDetails.Size - 1
                                odbdsDetails.SetValue("U_JENo", rowNum, "")
                                odbdsDetails.SetValue("U_RecoNo", rowNum, "")
                            Next
                            Matrix1.LoadFromDataSource()
                        Catch ex As Exception
                        Finally
                            objform.Freeze(False)
                        End Try
                        objform.Update()
                        objaddon.objapplication.MessageBox("Error while reconciling the transactions... " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 0, "OK")
                        objaddon.objapplication.StatusBar.SetText("Error while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        Matrix1.FlushToDataSource()
                        objaddon.objapplication.StatusBar.SetText("Internal Reconciliations Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Catch ex As Exception
                    BubbleEvent = False
                    objaddon.objapplication.MessageBox("Exception:  " + ex.Message, 0, "OK")
                    objaddon.objapplication.StatusBar.SetText("Exception:  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

            Catch ex As Exception
                BubbleEvent = False
                objaddon.objapplication.MessageBox("Form_DataAdd Exception: " + ex.Message, 0, "OK")
                objaddon.objapplication.StatusBar.SetText("Form_DataAdd Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

#End Region

    End Class
End Namespace
