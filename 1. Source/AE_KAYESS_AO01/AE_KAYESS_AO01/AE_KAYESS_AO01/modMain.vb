Option Explicit On

Module modMain

    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2


    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public p_oEventHandler As clsEventHandler
    Public p_oSBOApplication As SAPbouiCOM.Application
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oUICompany As SAPbouiCOM.Company
    Public sFuncName As String
    Public sErrDesc As String


    Sub main(ByVal Args() As String)

        Try
            sFuncName = "Main()"
            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", sFuncName)

            p_oApps = New SAPbouiCOM.SboGuiApi

            Dim sconn As String = Environment.GetCommandLineArgs.GetValue(1)

            p_oApps.Connect(sconn)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
            p_oEventHandler = New clsEventHandler

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
            p_oSBOApplication = p_oEventHandler.SBO_Application

            p_oDICompany = New SAPbobsCOM.Company
            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
            Call DisplayStatus(Nothing, "Addon starting.....please wait....", sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication Function", sFuncName)
            Call p_oEventHandler.SetApplication(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
            Call EndStatus(sErrDesc)
            System.Windows.Forms.Application.Run()
        Catch exp As Exception
            Call WriteToLogFile(exp.Message, "Main()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", "Main()")
        Finally
        End Try
    End Sub


End Module
