Imports System.Drawing

Namespace Finance_Payment
    Public Class ClsJE
        Public Const Formtype = "392"
        Dim objform As SAPbouiCOM.Form
        Dim FormCount As Integer = 0

        Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            objform = objaddon.objapplication.Forms.Item(FormUID)
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                End Select
            Else
                Try
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    End Select
                Catch ex As Exception

                End Try
            End If
        End Sub



        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                                Dim TransId As String
                                TransId = objform.DataSources.DBDataSources.Item("OJDT").GetValue("TransId", 0)
                                If TransId = "" Then Exit Sub
                                'objform.Items.Item("2").Click()
                                'objform.Close()
                                'pModal = False
                                'objaddon.objapplication.SendKeys("{ESC}")
                                Dim objRecform As SAPbouiCOM.Form
                                objRecform = objaddon.objapplication.Forms.GetForm("FOITR", 1)
                                Load_AdjJE(objRecform, TransId)
                            End If

                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try
        End Sub

        Private Function Load_AdjJE(ByVal objRecForm As SAPbouiCOM.Form, ByVal TransId As String) As Boolean
            Try
                Dim objRs As SAPbobsCOM.Recordset
                Dim strSQL As String = ""
                Dim odbdsDetails As SAPbouiCOM.DBDataSource
                Dim Matrix1 As SAPbouiCOM.Matrix
                odbdsDetails = objRecForm.DataSources.DBDataSources.Item("@MI_ITR1")
                Matrix1 = objRecForm.Items.Item("mtxcont").Specific
                If objaddon.HANA Then
                    strSQL = "SELECT ROW_NUMBER() OVER (ORDER BY A.""CardCode"",A.""DocDate"") AS ""LineId"",* FROM  "
                    strSQL += vbCrLf + "(SELECT 'N' AS ""Selected"",T1.""TransId"",T1.""Line_ID"",T1.""DebCred"",CASE WHEN ""FCCurrency"" IS NULL THEN (SELECT ""MainCurncy"" FROM OADM) ELSE ""FCCurrency"" END AS ""DocCur"","
                    strSQL += vbCrLf + "T1.""TransType"" AS ""ObjType"",T1.""LineMemo"" as ""LineMemo"",T1.""BaseRef"" AS ""DocNum"",T1.""CreatedBy"" as ""DocEntry"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'IN' WHEN T1.""TransType""='14' THEN 'CN' WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'DT' WHEN T1.""TransType""='18' THEN 'PU' WHEN T1.""TransType""='19' THEN 'PC'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='24' THEN 'RC' WHEN T1.""TransType""='46' THEN 'PS' Else 'JE' END AS ""Origin"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'A/R Invoice' WHEN T1.""TransType""='14' THEN 'A/R Credit Memo' WHEN T1.""TransType""='203' THEN 'A/R DownPayment' WHEN T1.""TransType""='18' THEN 'A/P Invoice'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='19' THEN 'A/R Credit Memo' WHEN T1.""TransType""='24' THEN 'Incoming Payment' WHEN T1.""TransType""='46' THEN 'Outgoing Payment' ELSE 'Journal Entry' END AS ""DocType"","
                    strSQL += vbCrLf + "T1.""ShortName"" AS ""CardCode"",(SELECT ""CardName"" FROM OCRD where ""CardCode""=T1.""ShortName"") as ""CardName"", (Select ""CardType"" from OCRD where ""CardCode""=T1.""ShortName"") as ""CardType"","
                    strSQL += vbCrLf + "CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END AS ""DocTotalLC"",CASE WHEN T1.""FCCredit""<>0  THEN  -T1.""FCCredit"" Else T1.""FCDebit"" END AS ""DocTotalFC"","
                    strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN  -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END  AS ""BalanceLC"",CASE WHEN T1.""BalFcCred""<>0  THEN  -T1.""BalFcCred"" ELSE T1.""BalFcDeb"" END AS ""BalanceFC"","
                    strSQL += vbCrLf + "CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END AS ""DocTotal"","
                    strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"","
                    strSQL += vbCrLf + "To_Varchar(T0.""RefDate"",'yyyyMMdd') AS ""DocDate"",T1.""BPLId"",(SELECT ""BPLName"" FROM OBPL WHERE ""BPLId""=T1.""BPLId"") AS ""BPLName"",T0.""Ref1"",T0.""Ref2"",T0.""Ref3"""
                    strSQL += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""DprId"" is null "
                    strSQL += vbCrLf + ") A "
                    strSQL += vbCrLf + "WHERE A.""TransId""='" & TransId & "' and A.""BPLId"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y') and A.""CardType"" in ('C','S') and A.""Balance""<>0"
                    strSQL += vbCrLf + "ORDER BY A.""CardCode"",A.""DocDate"""
                Else

                End If
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(strSQL)
                odbdsDetails.Clear()
                If objRs.RecordCount > 0 Then
                    While Not objRs.EoF
                        Matrix1.AddRow()
                        Matrix1.GetLineData(Matrix1.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, Matrix1.VisualRowCount) 'objRs.Fields.Item("LineId").Value.ToString
                        odbdsDetails.SetValue("U_Select", 0, "N")
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
                        Matrix1.SetLineData(Matrix1.VisualRowCount)
                        objRs.MoveNext()
                        'Matrix1.CommonSetting.SetRowBackColor(Matrix1.VisualRowCount, Color.PeachPuff.ToArgb)
                        ''Matrix1.Columns.Item("select").Cells.Item(Matrix1.VisualRowCount).Specific.Checked = True
                        'Matrix1.Columns.Item("paytot").Cells.Item(Matrix1.VisualRowCount).Click()
                        'objaddon.objapplication.SendKeys("^{TAB}")
                        'objRecForm.Update()
                        'Matrix1.Columns.Item("select").Cells.Item(Matrix1.VisualRowCount).Click()
                    End While
                    Matrix1.AutoResizeColumns()
                    objRecForm.Update()
                End If

                Return True
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

    End Class
End Namespace