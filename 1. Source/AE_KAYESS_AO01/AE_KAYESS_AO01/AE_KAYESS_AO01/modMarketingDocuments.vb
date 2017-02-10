Option Explicit On
Imports SAPbobsCOM
Imports SAPbouiCOM

Module modMarketingDocuments

#Region "AutoBatch Processing"

    Public Function CreateAutoBatchProcess(ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   CreateAutoBatchProcess()
        '   Purpose     :   This function will be providing to cater the processing for 
        '                   auto proceeding all items
        '                   (FIFO) Autobatch selection will be proceed the early admission date will be taking out first
        '   Parameters  :   ByVal oForm As SAPbouiCOM.Form
        '                       oForm = set the SAP UI Object Form object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Date        :   
        '   Author      :   Sri
        '   Change      :   
        '                      
        ' **********************************************************************************

        Dim oMatrixUp As SAPbouiCOM.Matrix = Nothing
        Dim oMatrixDnLeft As SAPbouiCOM.Matrix = Nothing
        Dim oDIRecordset As SAPbobsCOM.Recordset = Nothing
        Dim sFuncName As String = String.Empty
        Dim sAdmission As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemCode As String = String.Empty

        Dim sDate As String = String.Empty
        Dim dDate As Date = Nothing
        Dim dMinDate As Date = Nothing

        Dim iInnerCount As Int32 = 0
        Dim iCount As Int32 = 0
        Dim lQty As Double = -1

        Try
            sFuncName = "CreateAutoBatchProcess()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start Function", sFuncName)

            oMatrixUp = oForm.Items.Item("3").Specific
            oMatrixDnLeft = oForm.Items.Item("4").Specific

            If Not oMatrixDnLeft.Columns.Item("14").Visible Then
                sMessage = String.Format("Please be ensure to visible the [Manufacturing Date] in {0} before proceeding ... ", oForm.Items.Item("6").Specific.Caption)
                p_oSBOApplication.MessageBox(sMessage)
                GoTo Normal_Exit
            End If

            DisplayStatus(oForm, "Please wait ..... ", sErrDesc)

            oForm.Freeze(True)
            For iCount = 0 To oMatrixUp.VisualRowCount - 1
                sItemCode = oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Specific.Value
                lQty = CDbl(oMatrixUp.Columns.Item("55").Cells.Item(iCount + 1).Specific.Value)
                oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If oMatrixDnLeft.VisualRowCount > 0 And lQty <> 0.0 Then
Sort_Admission:
                    Try

                        oMatrixDnLeft.Columns.Item("14").TitleObject.Click(BoCellClickType.ct_Double)
                        oMatrixDnLeft = oForm.Items.Item("4").Specific

                        For iInnerCount = 0 To oMatrixDnLeft.VisualRowCount - 1
                            sDate = oMatrixDnLeft.Columns.Item("14").Cells.Item(iInnerCount + 1).Specific.Value
                            dDate = New Date(sDate.Substring(0, 4), sDate.Substring(4, 2), sDate.Substring(6))
                            If iInnerCount = 0 Then
                                dMinDate = dDate
                            Else
                                If dDate < dMinDate Then
                                    GoTo Sort_Admission
                                Else
                                    dMinDate = dDate
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        Throw New ArgumentException(" Sorting ... " & ex.Message)
                    End Try

                    Try

                        For iInnerCount = 1 To oMatrixDnLeft.VisualRowCount

                            'If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()) >= lQty Then
                            '    oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = Convert.ToDecimal(lQty)
                            '    oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            '    Exit For
                            'Else
                            '    lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()))
                            '    oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()
                            '    oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'End If

                            If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()) >= lQty Then
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = Convert.ToDecimal(lQty)
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit For
                            Else
                                lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString())))
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()))
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If


                        Next


                    Catch ex As Exception
                        sErrDesc = String.Format("{0} >>  Line : {1} {2}", sItemCode, iCount, ex.Message)
                        Throw New ArgumentException(sErrDesc)
                    End Try

                End If

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Next iCount
Normal_Exit:
            oForm.Freeze(False)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAutoBatchProcess = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAutoBatchProcess = RTN_ERROR
        Finally
            EndStatus(sErrDesc)
            oMatrixUp = Nothing
            oMatrixDnLeft = Nothing
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Function


    Public Function CreateAutoBatchProcess_PICKnPACK(ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   CreateAutoBatchProcess_PICKnPACK()
        '   Purpose     :   This function will be providing to cater the processing for 
        '                   auto proceeding all items
        '                   (FIFO) Autobatch selection will be proceed the early admission date will be taking out first
        '   Parameters  :   ByVal oForm As SAPbouiCOM.Form
        '                       oForm = set the SAP UI Object Form object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Date        :  
        '   Author      :   Sri
        '   Change      :   
        '                      
        ' **********************************************************************************

        Dim oMatrixUp As SAPbouiCOM.Matrix = Nothing
        Dim oMatrixDnLeft As SAPbouiCOM.Matrix = Nothing
        Dim oDIRecordset As SAPbobsCOM.Recordset = Nothing
        Dim sFuncName As String = String.Empty
        Dim sAdmission As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemCode As String = String.Empty

        Dim sDate As String = String.Empty
        Dim dDate As Date = Nothing
        Dim dMinDate As Date = Nothing

        Dim iInnerCount As Int32 = 0
        Dim iCount As Int32 = 0
        Dim lQty As Double = -1

        Try
            sFuncName = "CreateAutoBatchProcess()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start Function", sFuncName)

            oMatrixUp = oForm.Items.Item("3").Specific
            oMatrixDnLeft = oForm.Items.Item("4").Specific

            If Not oMatrixDnLeft.Columns.Item("14").Visible Then
                sMessage = String.Format("Please be ensure to visible the [Manufacturing Date] in {0} before proceeding ... ", oForm.Items.Item("6").Specific.Caption)
                p_oSBOApplication.MessageBox(sMessage)
                GoTo Normal_Exit
            End If

            DisplayStatus(oForm, "Please wait ..... ", sErrDesc)

            oForm.Freeze(True)
            For iCount = 0 To oMatrixUp.VisualRowCount - 1
                sItemCode = oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Specific.Value
                lQty = CDbl(oMatrixUp.Columns.Item("55").Cells.Item(iCount + 1).Specific.Value)
                oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If oMatrixDnLeft.VisualRowCount > 0 And lQty <> 0.0 Then
Sort_Admission:
                    Try

                        oMatrixDnLeft.Columns.Item("14").TitleObject.Click(BoCellClickType.ct_Double)
                        oMatrixDnLeft = oForm.Items.Item("4").Specific

                        For iInnerCount = 0 To oMatrixDnLeft.VisualRowCount - 1
                            sDate = oMatrixDnLeft.Columns.Item("14").Cells.Item(iInnerCount + 1).Specific.Value
                            dDate = New Date(sDate.Substring(0, 4), sDate.Substring(4, 2), sDate.Substring(6))
                            If iInnerCount = 0 Then
                                dMinDate = dDate
                            Else
                                If dDate < dMinDate Then
                                    GoTo Sort_Admission
                                Else
                                    dMinDate = dDate
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        Throw New ArgumentException(" Sorting ... " & ex.Message)
                    End Try

                    Try

                        For iInnerCount = 0 To oMatrixDnLeft.VisualRowCount - 1

                            If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()) >= lQty Then
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = Convert.ToDecimal(lQty)
                                oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(1).Specific.Value = Convert.ToDecimal(lQty)
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit For
                            Else
                                lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()))
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()
                                oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(1).Specific.Value = oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        Next
                    Catch ex As Exception
                        sErrDesc = String.Format("{0} >>  Line : {1} {2}", sItemCode, iCount, ex.Message)
                        Throw New ArgumentException(sErrDesc)
                    End Try

                End If

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Next iCount
Normal_Exit:
            oForm.Freeze(False)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAutoBatchProcess_PICKnPACK = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAutoBatchProcess_PICKnPACK = RTN_ERROR
        Finally
            EndStatus(sErrDesc)
            oMatrixUp = Nothing
            oMatrixDnLeft = Nothing
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Function


#End Region

End Module
