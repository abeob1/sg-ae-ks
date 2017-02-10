Option Explicit On

Public Class clsEventHandler

    Public WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO

    Public Sub New()
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Class_Initialize()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
            SBO_Application = p_oApps.GetApplication

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = SBO_Application.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(exc.Message, sFuncName)
        End Try
    End Sub

    Public Function SetApplication(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetApplication()
        '   Purpose    :    This function will be calling to initialize the default settings
        '                   such as Retrieving the Company Default settings, Creating Menus, and
        '                   Initialize the Event Filters
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetApplication()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
            If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
            If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetApplication = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(exc.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetApplication = RTN_ERROR
        End Try
    End Function

    Private Function SetMenus(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetMenus()
        '   Purpose    :    This function will be gathering to create the customized menu
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "SetMenus()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetMenus = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetMenus = RTN_ERROR
        End Try
    End Function

    Private Function SetFilters(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function   :    SetFilters()
        '   Purpose    :    This function will be gathering to declare the event filter 
        '                   before starting the AddOn Application
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetFilters()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
            oFilters = New SAPbouiCOM.EventFilters

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up Form Load filter", "SetFilters()")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_LOAD filter", sFuncName)
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            oFilter.Add("42")  'Batch Number sElection
            oFilter.Add("1320000126")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_ACTIVATE filter", sFuncName)
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_ITEM_PRESSED filter", sFuncName)
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            oFilter.AddEx("42") 'Batch Number sElection
            oFilter.Add("1320000126")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up Form Load filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_KEY_DOWN filter", sFuncName)
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_MENU_CLICK filter", sFuncName)
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
            SBO_Application.SetFilter(oFilters)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetFilters = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetFilters = RTN_ERROR
        End Try
    End Function

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_AppEvent()
        '   Purpose    :    This function will be handling the SAP Application Event
        '               
        '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
        '                       EventType = set the SAP UI Application Eveny Object        
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sMessage As String = String.Empty

        Try
            sFuncName = "SBO_Application_AppEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                    p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ShowErr(sErrDesc)
        Finally
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_MenuEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
        '                       pVal = set the SAP UI MenuEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SBO_Application_MenuEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            Select Case pVal.MenuUID

            End Select
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            ShowErr(exc.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
            ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        ' **********************************************************************************
        '   Function   :    SBO_Application_ItemEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByVal FormUID As String
        '                       FormUID = set the FormUID
        '                   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal = set the SAP UI ItemEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************
        Dim oForm As SAPbouiCOM.Form
        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SBO_Application_ItemEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not IsNothing(p_oDICompany) Then
                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            End If

            Select Case pVal.FormTypeEx

                Case "42"  ' Batch Number Selection
                    oForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.Before_Action = True Then
                                Try
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddButton() - AB_AUTOBT", sFuncName)
                                    If AddButton(oForm, "AB_AUTOBT", "Auto Select All", "2", 10, 100, True, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Catch ex As Exception
                                End Try
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "AB_AUTOBT" And pVal.Before_Action = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAutoBatchProcess()", sFuncName)
                                If CreateAutoBatchProcess(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If



                    End Select

                Case "1320000126"   ' Batch Number Selection (Pick & Pack)
                    oForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.Before_Action = True Then
                                Try
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddButton() - AB_AUTOBT", sFuncName)
                                    If AddButton(oForm, "AB_AUTOBT", "Auto Select All", "2", 10, 100, True, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Catch ex As Exception
                                End Try
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "AB_AUTOBT" And pVal.Before_Action = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAutoBatchProcess()", sFuncName)
                                If CreateAutoBatchProcess_PICKnPACK(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                    End Select


            End Select

Normal_Exit:
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            sErrDesc = exc.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
            ShowErr(sErrDesc)
        Finally
            oForm = Nothing
        End Try

    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_FormDataEvent()
        '   Purpose    :    This function will be handling the SAP FormData Event
        '               
        '   Parameters :    ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo
        '                       BusinessObjectInfo = set the SAP UI BusinessObjectInfo Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************
        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim sKeyValue As String = String.Empty
        Dim sMessage As String = String.Empty

        Try
            sFuncName = "SBO_Application_FormDataEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Select Case BusinessObjectInfo.FormTypeEx

              End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ShowErr(sErrDesc)
            BubbleEvent = False
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class