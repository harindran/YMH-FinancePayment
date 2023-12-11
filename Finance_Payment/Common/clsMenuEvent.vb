Imports SAPbouiCOM
Namespace Finance_Payment

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "MBAPSI"
                        Mul_Branch_AP_Service_Invoice_MenuEvent(pVal, BubbleEvent)
                    Case "141"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                    Case "PAYINIT"
                        PaymentInit_MenuEvent(pVal, BubbleEvent)
                    Case "PAYM"
                        Payment_Means_MenuEvent(pVal, BubbleEvent)
                    Case "FINPAY"
                        InPayments_MenuEvent(pVal, BubbleEvent)
                    Case "FOUTPAY"
                        OutPayments_MenuEvent(pVal, BubbleEvent)
                    Case "FOITR"
                        InternalReconciliation_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Dim oUDFForm As SAPbouiCOM.Form
                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1281"
                            oUDFForm.Items.Item("U_MBAPNo").Enabled = True
                        Case "1287"
                            If oUDFForm.Items.Item("U_MBAPNo").Enabled = False Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = True
                            End If
                            oUDFForm.Items.Item("U_MBAPNo").Specific.String = ""
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Mul_Branch_AP_Service_Invoice"

        Private Sub Mul_Branch_AP_Service_Invoice_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("tposdate").Enabled = True
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("tduedate").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                        Case "1282" ' Add Mode
                            objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "MIAPSI")

                        Case "1288", "1289", "1290", "1291"

                        Case "1293"
                            DeleteRow(Matrix0, "@MIPL_API1")
                        Case "1292"
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region

#Region "Payment"

        Private Sub PaymentInit_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxdata").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1293"
                            Try
                                Dim USERSource As SAPbouiCOM.UserDataSource
                                USERSource = objform.DataSources.UserDataSources.Item("UD_3")
                                objform.Freeze(True)
                                For i As Integer = 1 To Matrix0.VisualRowCount
                                    Matrix0.GetLineData(i)
                                    USERSource.Value = i
                                    Matrix0.SetLineData(i)
                                Next
                                objform.Freeze(False)
                            Catch ex As Exception
                                objform.Freeze(False)
                                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                            Finally
                            End Try

                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub InPayments_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("ttranno").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub OutPayments_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("ttranno").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub InternalReconciliation_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("tpaydate").Enabled = True
                            objform.Items.Item("btnadjent").Enabled = False
                            'objform.Items.Item("ttranno").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub Payment_Means_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcheq").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "CPYD"  'Copy Due
                            If objform.Items.Item("tcurr").Specific.selected.Value = MainCurr Then
                                If objform.ActiveItem = "tctot" Then
                                    objform.Items.Item("tctot").Specific.String = objform.Items.Item("tbaldue").Specific.String
                                ElseIf objform.ActiveItem = "tbtot" Then
                                    objform.Items.Item("tbtot").Specific.String = objform.Items.Item("tbaldue").Specific.String
                                Else
                                    Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                                    Dim RowID As Integer = Matrix0.GetCellFocus().rowIndex
                                    If ColID = 2 Then 'chamt
                                        Matrix0.Columns.Item("chamt").Cells.Item(RowID).Specific.String = CDbl(objform.Items.Item("tbaldue").Specific.String) '+ Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                                    End If
                                End If
                            Else
                                If objform.ActiveItem = "tctot" Then
                                    objform.Items.Item("tctot").Specific.String = objform.Items.Item("tbalduec").Specific.String
                                ElseIf objform.ActiveItem = "tbtot" Then
                                    objform.Items.Item("tbtot").Specific.String = objform.Items.Item("tbalduec").Specific.String
                                Else
                                    Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                                    Dim RowID As Integer = Matrix0.GetCellFocus().rowIndex
                                    If ColID = 2 Then 'chamt
                                        Matrix0.Columns.Item("chamt").Cells.Item(RowID).Specific.String = CDbl(objform.Items.Item("tbalduec").Specific.String) '+ Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                                    End If
                                End If
                            End If
                    End Select
                Else
                    Select Case pval.MenuUID
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
    End Class
End Namespace