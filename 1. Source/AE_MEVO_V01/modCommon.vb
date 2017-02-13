Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.IO





Module modCommon


    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sLicenseServer = String.Empty
            oCompDef.iServerLanguage = 3
            'oCompDef.iServerType = 7
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty



            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ServerType")) Then
                oCompDef.sServerType = ConfigurationManager.AppSettings("ServerType")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.sLicenseServer = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ProcessedDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("ProcessedDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ErrorDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("ErrorDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("JESeries")) Then
                oCompDef.sSeries = ConfigurationManager.AppSettings("JESeries")
            End If

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function


    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            ' Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            '  Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs.Tables(0)
    End Function

    Public Function ExecuteSQLQuery(ByVal sEntity As String, ByRef sErrDesc As String) As Long

        '**************************************************************
        ' Function      : ExecuteSQLQuery_DT
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDt As New DataTable
        Dim sFuncName As String = String.Empty
        Dim sQuery As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            ' Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
            sQuery = "SELECT isnull(Max(CAST( CODE as int)),0)+1 AS CODE FROM " & sEntity & ".. [@AB_STATITISTICSDATA]"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Statistics Row Count SQL " & sQuery, sFuncName)
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDt)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity." & sEntity & " Row Count " & oDt.Rows(0).Item(0).ToString.Trim, sFuncName)
            oDT_StatisticsRowCount.Rows.Add(sEntity, oDt.Rows(0).Item(0).ToString.Trim)
            '  Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            ExecuteSQLQuery = RTN_SUCCESS

        Catch ex As Exception
            ExecuteSQLQuery = RTN_ERROR
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            sErrDesc = ex.Message
        Finally
            oCon.Dispose()
        End Try

    End Function


    Public Function ExecuteInsertSQLQuery(ByVal sQuery As String, ByRef sErrDesc As String) As Long

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim sFuncName As String = String.Empty

        'Dim sConstr As String = "DRIVER={HDBODBC32};SERVERNODE={" & p_oCompDef.sServer & "};DSN=" & p_oCompDef.sDSN & ";UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";"

        Try
            sFuncName = "ExecuteInsertSQLQuery()"
            Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            'oCon.ConnectionString = "DRIVER={HDBODBC};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & " ;SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName & ""
            ' oCon.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sQuery, sFuncName)
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            Try
                da.Fill(oDs)
            Catch ex As Exception
            End Try

            Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

            Return RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return RTN_ERROR
        Finally
            oCon.Dispose()
        End Try

    End Function

    Public Function GetEntitiesDetails(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetEntitiesDetails()
        '   Purpose     :   This function will be providing information about the Entities, SAP username, SAP Password, Banking Details
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************


        Dim sSqlstr As String = String.Empty
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetEntitiesDetails()"
            ' Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)
            ' Getting the details of Entity, SAP User name, Password and Banking from the COMPANYDATA Table
            'sSqlstr = "SELECT T0.[PrcCode] [Center Code], T0.[PrcName] [Center Name], T1.[Name] [DB Name], T1.[U_AE_UPass] [Pass], T1.[U_AE_UName] [User Name] FROM OPRC T0 " & _
            '    "inner join [dbo].[@AE_COMPANYDATA]  T1 on T0.[U_AE_DBName] = T1.Name"

            sSqlstr = "SELECT T0.[PrcCode] [OUCode], T0.[PrcName] [OU Name], T0.[U_AB_ENTITY] [Entity],T0.[U_AB_REPORTCODE] [BU Code], " & _
                "T2.[U_AB_REPORTCODE] [LOS Code], T3.[U_AB_USERCODE] [User], T3.[U_AB_PASSWORD] [Pass], T0.[U_AB_OUCOMMON] [EntityFlag], T3.[U_AB_IPOWERCODE] [EntityCode] " & _
                "FROM OPRC T0  INNER JOIN ODIM T1 ON T0.[DimCode] = T1.[DimCode] left outer join OPRC T2 " & _
                "on T2.[PrcCode] = T0.[U_AB_REPORTCODE] left outer join [@AB_COMPANYDATA] T3 on T0.[U_AB_ENTITY] = T3.[Name] WHERE T1.[DimCode] = 3"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            p_oEntitesDetails = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- GL Account
            sSqlstr = "SELECT T0.[AcctCode], T0.[AcctName], T0.FrgnName [ExportCode] FROM OACT T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)
            p_oGLAccount = ExecuteSQLQuery_DT(sSqlstr)
            ' SELECT T0.[Code], T0.[Name], T0.[U_AB_STDESCRIPTION], T0.[U_AB_STNEWCODE] FROM [dbo].[@AB_IPOWERSTCODE]  T0
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- AB_IPOWERSTCODE
            sSqlstr = "SELECT T0.[Code], T0.[Name], T0.[U_AB_STDESCRIPTION], T0.[U_AB_STNEWCODE] , case when left(T0.[U_AB_STNEWCODE],2) = 'ST' then 'IMS' else 'IP' end [Cat] FROM [dbo].[@AB_IPOWERSTCODE]  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oSTOLDCODE = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- Company Data
            sSqlstr = "SELECT T0.[U_AB_IPOWERCODE] [ipEntityCode], T0.[U_AB_COMCODE]  [EntityCode], T0.[U_AB_COMPANYNAME], T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD] FROM [dbo].[@AB_COMPANYDATA]  T0"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTCompanyData = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- iPower Period
            sSqlstr = "select Code, Name,left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' ' + CAST(YEAR(getdate()) as varchar)[Month Name]," & _
"cast(month(left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' 1 2015') as varchar) + ' ' + CAST(YEAR(getdate()) as varchar) [Month Number]," & _
"cast(month(left(LEFT(name, CHARINDEX(' ', Name  )),3) + ' 1 2015') as varchar) [Month] , CAST(YEAR(getdate()) as varchar) [Year]" & _
" from [@AB_IPOWERPERIOD] "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTiPowerPeriod = ExecuteSQLQuery_DT(sSqlstr)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            '----------------------- SAP Period
            sSqlstr = "select Code, Name, cast(MONTH(F_RefDate ) as varchar) + ' ' + cast(year(F_RefDate ) as varchar), " & _
                "month(F_RefDate ) [F_Month], MONTH(T_RefDate ) [T_Month]," & _
"YEAR(F_RefDate ) [Year], F_RefDate [RefDate_F], T_RefDate [RefDate_T] from OFPR  "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL String " & sSqlstr, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            p_oDTSAPPeriod = ExecuteSQLQuery_DT(sSqlstr)



            Console.WriteLine("Completed With SUCCESS " & sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            GetEntitiesDetails = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed With ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            GetEntitiesDetails = RTN_ERROR
        End Try

    End Function

    Public Function IdentifyTXTFile_JournalEntry(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   IdentifyTXTFile_JournalEntry()
        '   Purpose     :   This function will identify the TXT file of Journal Entry
        '                    Upload the file into Dataview and provide the information to post transaction in SAP.
        '                     Transaction Success : Move the TXT file to SUCESS folder
        '                     Transaction Fail :    Move the TXT file to FAIL folder and send Error notification to concern person
        '               
        '   Parameters  :   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************


        Dim sSqlstr As String = String.Empty
        Dim bJEFileExist As Boolean
        Dim sFileType As String = String.Empty
        Dim oDTDistinct As DataTable = Nothing
        Dim oDTRowFilter As DataTable = Nothing
        Dim oDVJE_DET As DataView = Nothing
        Dim oDVJE_HDR As DataView = Nothing
        Dim oDVJE As DataView = Nothing
        Dim oDV As DataView = Nothing
        Dim oDVIMPSTS As DataView = Nothing
        Dim oDICompany() As SAPbobsCOM.Company = Nothing
        Dim sCompanyDB As String = String.Empty
        Dim oDT_Entity As DataTable = Nothing
        Dim sFuncName As String = String.Empty
        Dim oDT_File As DataTable = Nothing

        Try
            sFuncName = "IdentifyTXTFile_JournalEntry()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oDT_StatisticsRowCount = New DataTable
            oDT_StatisticsRowCount.Columns.Add("Entity", GetType(String))
            oDT_StatisticsRowCount.Columns.Add("Count", GetType(Integer))

            oDT_Entity = New DataTable()
            oDT_Entity.Columns.Add("Entity", GetType(String))

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo
            Dim sExtension As String = String.Empty
            Dim sCommonFile As String
            Dim sFileHdr As String
            Dim sFileDet As String
            Dim sFilePath As String

            files = DirInfo.GetFiles("*.*")

            ''Code Added - Shibin - 19 Aug 2016
            oDT_File = New DataTable()
            oDT_File.Columns.Add("FileName", GetType(String))
            oDT_File.Columns.Add("FilePath", GetType(String))
            oDT_File.Columns.Add("Extension", GetType(String))
            oDT_File.Columns.Add("OnlyName", GetType(String))

            ''Dim sFilename As String = File.Name
            'Dim sfileNamePart() As String = files.ToString().Split(".")
            ''Dim sFileNames = New HashSet(Of String)(sfileNamePart)
            'Dim sFileNames As String() = sfileNamePart.Distinct().ToArray()
            'For Each sCommonFile As String In sFileNames
            '    Console.WriteLine(sCommonFile)
            'Next

            For Each File As System.IO.FileInfo In files
                Dim sFilename As String = File.Name
                Dim sfileNamePart() As String = sFilename.ToString().Split(".")
                oDT_File.Rows.Add(File.Name, File.FullName, File.Extension, sfileNamePart(0).ToString())
            Next


            Dim names = From row In oDT_File.AsEnumerable()
                        Select row.Field(Of String)("OnlyName") Distinct
            Dim iLength As Integer = names.LongCount()
           
            For Value As Integer = 0 To names.LongCount()
          
                If (Value = names.LongCount) Then
                    Exit For
                Else
                    sCommonFile = names(Value).ToString()
                    sFileDet = sCommonFile + ".det"
                    sFileHdr = sCommonFile + ".hdr"
                    sFilePath = DirInfo.FullName

                    bJEFileExist = True

                    Console.WriteLine("Attempting File Name - " & sCommonFile, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting File Name - " & sCommonFile, sFuncName)
                    'sFileType = Replace(File.Name, ".txt", "").Trim
                    'upload the CSV to Dataview
                    'sExtension = File.Extension

                    If sFileDet <> "" Then
                        Console.WriteLine("GetDataViewFromTXT_DET() ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GetDataViewFromTXT() ", sFuncName)
                        'oDVJE_DET = GetDataViewFromTXT_DET(File.FullName, File.Name, sErrDesc)
                        oDVJE_DET = GetDataViewFromTXT_DET(sFilePath + "\" + sFileDet, sFileDet, sErrDesc)
                        ''  oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
                        If sErrDesc.Length > 1 Then
                            Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            'FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                            FileMoveToArchives(sFileDet, sFilePath + "\" + sFileDet, RTN_ERROR, ".det", "")
                            Write_TextFile_I("Invalid File Format , preferable format is Txt {Tab} Delimiter ", sErrDesc)
                            IdentifyTXTFile_JournalEntry = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    
                    If sFileDet <> "" Then
                        Console.WriteLine("GetDataViewFromTXT_DET() ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GetDataViewFromTXT() ", sFuncName)
                        'oDVJE_HDR = GetDataViewFromTXT_HDR(File.FullName, File.Name, sErrDesc)
                        oDVJE_HDR = GetDataViewFromTXT_HDR(sFilePath + "\" + sFileHdr, sFileHdr, sErrDesc)
                        ''  oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
                        If sErrDesc.Length > 1 Then
                            Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
                            'FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                            FileMoveToArchives(sFileHdr, sFilePath + "\" + sFileHdr, RTN_ERROR, ".hdr", "")
                            Write_TextFile_I("Invalid File Format , preferable format is Txt {Tab} Delimiter ", sErrDesc)
                            IdentifyTXTFile_JournalEntry = RTN_ERROR
                            Exit Function
                        End If
                    End If

                    If bJEFileExist = False Then
                        Console.WriteLine("No input file found  ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No input file found ", sFuncName)
                        Return RTN_SUCCESS
                    End If
                    oDV = oDVJE_HDR

                    If Not p_oCompany.InTransaction Then
                        p_oCompany.StartTransaction()
                    End If

                    For Each odr As DataRowView In oDVJE_HDR
                        oDV.RowFilter = "tran_hdr_uid='" & odr("tran_hdr_uid").ToString() & "'"
                        oDVJE_DET.RowFilter = "tran_hdr_uid='" & odr("tran_hdr_uid").ToString() & "'"
                        If Create_Sales_order(oDV, oDVJE_DET, p_oCompany, sErrDesc) <> RTN_SUCCESS Then
                            If p_oCompany.InTransaction Then
                                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            'files = DirInfo.GetFiles("*.*")
                            'For Each File As System.IO.FileInfo In files
                            '    FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
                            'Next
                            'Throw New ArgumentException(sErrDesc)
                            FileMoveToArchives(sFileDet, sFilePath + "\" + sFileDet, RTN_ERROR, ".det", "")
                            FileMoveToArchives(sFileHdr, sFilePath + "\" + sFileHdr, RTN_ERROR, ".hdr", "")
                        End If
                    Next

                    If p_oCompany.InTransaction Then
                        p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                      

                    End If

                    FileMoveToArchives(sFileDet, sFilePath + "\" + sFileDet, RTN_SUCCESS, ".det", "")
                    FileMoveToArchives(sFileHdr, sFilePath + "\" + sFileHdr, RTN_SUCCESS, ".hdr", "")
                End If
            Next

            ' '' Code ended -Shibin - 19 Aug 2016


            ''Commented on 19 Aug 2016 - Shibin
            'For Each File As System.IO.FileInfo In files
            '    bJEFileExist = True

            '    Console.WriteLine("Attempting File Name - " & File.Name, sFuncName)
            '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting File Name - " & File.Name, sFuncName)
            '    'sFileType = Replace(File.Name, ".txt", "").Trim
            '    'upload the CSV to Dataview
            '    sExtension = File.Extension

            '    Select Case sExtension

            '        Case ".det"
            '            Console.WriteLine("GetDataViewFromTXT_DET() ", sFuncName)
            '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GetDataViewFromTXT() ", sFuncName)
            '            oDVJE_DET = GetDataViewFromTXT_DET(File.FullName, File.Name, sErrDesc)
            '            ''  oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
            '            If sErrDesc.Length > 1 Then
            '                Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
            '                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
            '                FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
            '                Write_TextFile_I("Invalid File Format , preferable format is Txt {Tab} Delimiter ", sErrDesc)
            '                IdentifyTXTFile_JournalEntry = RTN_ERROR
            '                Exit Function
            '            End If

            '        Case ".hdr"
            '            Console.WriteLine("GetDataViewFromTXT_DET() ", sFuncName)
            '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GetDataViewFromTXT() ", sFuncName)
            '            oDVJE_HDR = GetDataViewFromTXT_HDR(File.FullName, File.Name, sErrDesc)
            '            ''  oDTDistinct = oDVJE.Table.DefaultView.ToTable(True, "Entity")
            '            If sErrDesc.Length > 1 Then
            '                Console.WriteLine("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
            '                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Calling FileMoveToArchive for moving CSV file to archive folder", sFuncName)
            '                FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
            '                Write_TextFile_I("Invalid File Format , preferable format is Txt {Tab} Delimiter ", sErrDesc)
            '                IdentifyTXTFile_JournalEntry = RTN_ERROR
            '                Exit Function
            '            End If
            '    End Select
            'Next

            'If bJEFileExist = False Then
            '    Console.WriteLine("No input file found  ", sFuncName)
            '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No input file found ", sFuncName)
            '    Return RTN_SUCCESS
            'End If
            'oDV = oDVJE_HDR

            'If Not p_oCompany.InTransaction Then
            '    p_oCompany.StartTransaction()
            'End If

            'For Each odr As DataRowView In oDVJE_HDR
            '    oDV.RowFilter = "tran_hdr_uid='" & odr("tran_hdr_uid").ToString() & "'"
            '    oDVJE_DET.RowFilter = "tran_hdr_uid='" & odr("tran_hdr_uid").ToString() & "'"
            '    If Create_Sales_order(oDV, oDVJE_DET, p_oCompany, sErrDesc) <> RTN_SUCCESS Then
            '        If p_oCompany.InTransaction Then
            '            p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '        End If
            '        files = DirInfo.GetFiles("*.*")
            '        For Each File As System.IO.FileInfo In files
            '            FileMoveToArchive(File, File.FullName, RTN_ERROR, "")
            '        Next
            '        Throw New ArgumentException(sErrDesc)
            '    End If
            'Next

            'If p_oCompany.InTransaction Then
            '    p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If


            'files = DirInfo.GetFiles("*.*")
            'For Each File As System.IO.FileInfo In files
            '    FileMoveToArchive(File, File.FullName, RTN_SUCCESS, "")
            'Next
            ' '' Commeneted end - 19 Aug 2016 - Shibin

            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            IdentifyTXTFile_JournalEntry = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed With ERROR", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR", sFuncName)
            IdentifyTXTFile_JournalEntry = RTN_ERROR
        End Try

    End Function

    Public Function GetDataViewFromTXT_OLD(ByVal CurrFileToUpload As String, ByVal Filename As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromCSV()
        '   Purpose     :   This function will upload the data from CSV file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim dv As DataView

        Dim sFuncName As String = String.Empty
        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        Dim dvGLAcccount As DataView = New DataView(p_oGLAccount)
        Dim sEntity As String = String.Empty
        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim oSR As StreamReader
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty

        Try
            sFuncName = "GetDataViewFromTXT"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            'Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oipower As New DataTable()
            Dim oDT_ImportStatistics As New DataTable()

            oipower.Columns.Add("GL_Code", GetType(String))
            oipower.Columns.Add("Date", GetType(DateTime)) ' Date
            oipower.Columns.Add("Col3", GetType(String))
            oipower.Columns.Add("Amount", GetType(String)) ' Amount
            oipower.Columns.Add("Ref1", GetType(String))
            oipower.Columns.Add("Col6", GetType(String))
            oipower.Columns.Add("Col7", GetType(String))
            oipower.Columns.Add("Description", GetType(String))
            oipower.Columns.Add("Col9", GetType(String))
            oipower.Columns.Add("Col10", GetType(String))
            oipower.Columns.Add("OU", GetType(String))
            oipower.Columns.Add("EntityCode", GetType(String))
            oipower.Columns.Add("Col13", GetType(String))
            oipower.Columns.Add("GST", GetType(String))
            oipower.Columns.Add("Col15", GetType(String))
            oipower.Columns.Add("Voucher", GetType(String))
            oipower.Columns.Add("Entity", GetType(String))
            oipower.Columns.Add("BUCode", GetType(String))
            oipower.Columns.Add("LOSCode", GetType(String))
            oipower.Columns.Add("Year", GetType(String))
            oipower.Columns.Add("Code", GetType(String))

            oDT_ImportStatistics.Columns.Add("GL_Code", GetType(String))
            oDT_ImportStatistics.Columns.Add("Date", GetType(DateTime)) ' Date
            oDT_ImportStatistics.Columns.Add("Col3", GetType(String))
            oDT_ImportStatistics.Columns.Add("Amount", GetType(String)) ' Amount
            oDT_ImportStatistics.Columns.Add("Ref1", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col6", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col7", GetType(String))
            oDT_ImportStatistics.Columns.Add("Description", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col9", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col10", GetType(String))
            oDT_ImportStatistics.Columns.Add("OU", GetType(String))
            oDT_ImportStatistics.Columns.Add("EntityCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col13", GetType(String))
            oDT_ImportStatistics.Columns.Add("GST", GetType(String))
            oDT_ImportStatistics.Columns.Add("Col15", GetType(String))
            oDT_ImportStatistics.Columns.Add("Voucher", GetType(String))
            oDT_ImportStatistics.Columns.Add("Entity", GetType(String))
            oDT_ImportStatistics.Columns.Add("BUCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("LOSCode", GetType(String))
            oDT_ImportStatistics.Columns.Add("Year", GetType(String))
            oDT_ImportStatistics.Columns.Add("Code", GetType(String))

            'Open the file in a stream reader.
            oSR = New StreamReader(CurrFileToUpload)

            Dim sText As String
            Dim sString(-1) As String
            Dim sDelimiter As String() = {vbTab}

            While oSR.Peek <> -1
                sText = oSR.ReadLine()
                If sText.Length > 1 Then
                    sString = sText.Split(sDelimiter, StringSplitOptions.None) ' "RemoveEmptyEntrie" I am also using the option to remove empty entries a
                    ' dtEntiry = p_oEntitesDetails.DefaultView.ToTable(True, sString(10))
                    dvEntiry.RowFilter = "OUCode='" & sString(10) & "'"

                    If dvEntiry.Count > 0 Then
                        If dvEntiry.Item(0)("EntityFlag").ToString.ToUpper = "YES" Then
                            dvEntiry.RowFilter = "EntityCode= '" & sString(11) & "'"
                            If dvEntiry.Count > 0 Then
                                sEntity = dvEntiry.Item(0)(2).ToString
                                sBUCode = dvEntiry.Item(0)(3).ToString
                                sLOS = dvEntiry.Item(0)(4).ToString
                            Else
                                sEntity = String.Empty
                                sBUCode = String.Empty
                                sLOS = String.Empty
                            End If

                        Else
                            sEntity = dvEntiry.Item(0)(2).ToString
                            sBUCode = dvEntiry.Item(0)(3).ToString
                            sLOS = dvEntiry.Item(0)(4).ToString
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("OU Code " & sString(10) & " Entity is Empty in Line " & iCount, sFuncName)
                        sEntity = ""
                        sBUCode = ""
                        sLOS = ""
                    End If
                    dvGLAcccount.RowFilter = "ExportCode='" & sString(0) & "'"
                    If dvGLAcccount.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No GL Account Found  " & sString(0) & "  Line No " & iCount, sFuncName)
                        sGLAccount = ""
                    Else
                        sGLAccount = dvGLAcccount.Item(0)(0).ToString
                    End If
                    oipower.Rows.Add(sGLAccount, DateTime.ParseExact(Right(sString(1), 8), "yyyyMMdd", Nothing), sString(2), sString(3), sString(4), sString(5), sString(6), sString(7), sString(8), _
                                     sString(9), sString(10), sString(11), sString(12), sString(13), sString(14), sString(15), sEntity, sBUCode, sLOS, Left(sString(1), 4), Right(Left(sString(1), 7), 3))

                    iCount += 1
                End If
            End While

            'Console.WriteLine("Del_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Del_schema() ", sFuncName)
            'Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oipower)
            Return dv

        Catch ex As Exception

            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally
            oSR.Close()
            oSR = Nothing
        End Try

    End Function

    Public Function ImportStatistics(ByRef oDVLineDetails As DataView, ByRef ocompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        'Function   :   ImportStatistics()
        'Purpose    :   Import Text File Data Into UDT
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   SAI
        'Date       :   22/1/2015
        'Change     :

        Dim sFuncName As String = String.Empty

        Dim oDt As DataTable
        Dim iCode As Integer
        Dim sConnString As String = String.Empty
        Dim oDTCode As DataTable = Nothing
        Dim oRset As SAPbobsCOM.Recordset = ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSql As String
        Dim oDVStatisticsrowcount As DataView = New DataView(oDT_StatisticsRowCount)
        Dim asql(100) As String
        Dim iloop As Integer = 0

        Try
            sFuncName = "ImportStatistics()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering Entity " & ocompany.CompanyDB, sFuncName)
            oDVStatisticsrowcount.RowFilter = "Entity='" & ocompany.CompanyDB & "'"

            oDt = oDVLineDetails.ToTable
            sSql = String.Empty
            iCode = oDVStatisticsrowcount.Item(0)(1)
            ReDim asql(25)
            Dim icount As Integer = 0

            For Each row As DataRow In oDt.Rows
                ' write insert statement
                If icount = 5000 Then
                    iloop += 1
                    icount = 0
                End If

                asql(iloop) += " Insert Into [@AB_STATITISTICSDATA] ( [Code],  [Name], [U_AB_PERIOD],[U_AB_OPER_UNIT], " & _
                    " [U_AB_ENTITY], [U_AB_DEBIT_CREDIT],[U_AB_AMOUNT], [U_AB_GLCODE], [U_AB_DESCRIPTION]) " & _
                    " Values ('" & iCode & "','" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(5).ToString & "','" & row.Item(6).ToString & "', " & _
                    "'" & row.Item(3).ToString & "'," & CDbl(row.Item(2).ToString) & ",'" & row.Item(0).ToString & "', '" & row.Item(4).ToString & "')"

                iCode = iCode + 1
                icount = icount + 1
            Next

            If asql.Length > 1 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Insert Data in " & ocompany.CompanyDB, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(ocompany.CompanyDB & " - Import Statistics Count " & icount, sFuncName)

                For Each element As String In asql
                    If Not String.IsNullOrEmpty(element) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(element, sFuncName)
                        oRset.DoQuery(element)
                    Else
                        Exit For
                    End If
                Next
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserted Successful", sFuncName)
            End If

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() for Inserting the data to the Table.", sFuncName)

            'ExecuteSQLQuery_DT(sSql, sConnString, sErrDesc)

            ImportStatistics = RTN_SUCCESS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            ImportStatistics = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
        End Try
    End Function

    Public Function ExecuteSQLQuery_DT(ByVal sQuery As String, ByVal sConnString As String, ByRef sErrDesc As String) As DataTable

        '**************************************************************
        ' Function      : ExecuteQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : JOHN
        ' Date          : MAY 2014 20
        ' Change        :
        '**************************************************************

        ''Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd

        Dim oCon As New SqlConnection(sConnString)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet
        Dim oDT As New DataTable
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "ExecExecuteSQLQuery_DT()"
            Console.WriteLine("Starting Function.. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query : " & sQuery, sFuncName)

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDT)
            Console.WriteLine("Completed Successfully. ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully.", sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New Exception(ex.Message)
            Return Nothing
        Finally
            If oCon.State = ConnectionState.Open Then
                oCon.Dispose()
            End If
        End Try
        Return oDT
    End Function

    Public Function GetDataViewFromTXT_DET(ByVal CurrFileToUpload As String, ByVal Filename As String, ByRef sErrDesc As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromTXT_DET()
        '   Purpose     :   This function will upload the data from CSV file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim dv As DataView

        Dim sFuncName As String = String.Empty
        sErrDesc = String.Empty
        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        ' Dim dvGLAcccount As DataView = New DataView(p_oGLAccount)
        Dim oDvNewGlCode As DataView = New DataView(p_oSTOLDCODE)
        Dim oDVCompanyData As DataView = New DataView(p_oDTCompanyData)
        Dim oDVipowerPeriod As DataView = New DataView(p_oDTiPowerPeriod)
        Dim oDVsapPeriod As DataView = New DataView(p_oDTSAPPeriod)
        Dim dperioddate As Date
        Dim sEntity As String = String.Empty
        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim oSR As StreamReader
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty
        Dim sSAPPeriodCode As String = String.Empty
        oDT_OUCODE = New DataTable()


        Try
            sFuncName = "GetDataViewFromTXT_DET"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            'Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oDT As New DataTable()
            p_oDTImportStatistics = New DataTable()

            oDT.Columns.Add("tran_det_uid", GetType(String))  ''0
            oDT.Columns.Add("tran_hdr_uid", GetType(String))
            oDT.Columns.Add("line_no", GetType(String))
            oDT.Columns.Add("type", GetType(String))
            oDT.Columns.Add("item_uid", GetType(String))
            oDT.Columns.Add("tax_rate", GetType(String))
            oDT.Columns.Add("tax_incl", GetType(String))
            oDT.Columns.Add("item_note", GetType(String))
            oDT.Columns.Add("item_note2", GetType(String))
            oDT.Columns.Add("item_note3", GetType(String))
            oDT.Columns.Add("item_note4", GetType(String))  ''10
            oDT.Columns.Add("desp", GetType(String))
            oDT.Columns.Add("desp2", GetType(String))
            oDT.Columns.Add("loc_uid", GetType(String))
            oDT.Columns.Add("serial", GetType(String))
            oDT.Columns.Add("qty", GetType(Double))   ''15
            oDT.Columns.Add("uom_uid", GetType(String))
            oDT.Columns.Add("qty_per_uom", GetType(Double))
            oDT.Columns.Add("is_foc", GetType(String))
            oDT.Columns.Add("uprice", GetType(Double))
            oDT.Columns.Add("disc", GetType(String))  ''20
            oDT.Columns.Add("disc_type", GetType(String))
            oDT.Columns.Add("disc2", GetType(String))
            oDT.Columns.Add("slf_uom_uid", GetType(String))
            oDT.Columns.Add("slf_qty_per_uom", GetType(String))
            oDT.Columns.Add("slf_qty", GetType(String))
            oDT.Columns.Add("rtn_qty", GetType(String))
            oDT.Columns.Add("net_qty", GetType(String)) ''27
            oDT.Columns.Add("rtn_item_bad", GetType(String))
            oDT.Columns.Add("sls_item_bad", GetType(String))
            oDT.Columns.Add("req_qty", GetType(String))
            oDT.Columns.Add("comm_qty", GetType(String))
            oDT.Columns.Add("shelf_from", GetType(String))
            oDT.Columns.Add("shelf_to", GetType(String)) ''33
            oDT.Columns.Add("note", GetType(String))
            oDT.Columns.Add("note2", GetType(String))
            oDT.Columns.Add("note3", GetType(String))
            oDT.Columns.Add("note4", GetType(String))

            'Open the file in a stream reader.


            Using MyReader As New Microsoft.VisualBasic.
                       FileIO.TextFieldParser(
                         CurrFileToUpload)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        If currentRow(0) = "tran_det_uid" Then Continue While
                        oDT.Rows.Add(currentRow(0), currentRow(1), currentRow(2), currentRow(3), currentRow(4), currentRow(5), currentRow(6), currentRow(7), currentRow(8), currentRow(9), currentRow(10),
                                currentRow(11), currentRow(12), currentRow(13), currentRow(14), CDbl(currentRow(15)), currentRow(16), CDbl(currentRow(17)), currentRow(18), CDbl(currentRow(19)), currentRow(20), currentRow(21),
                                currentRow(22), currentRow(24), currentRow(25), currentRow(26), currentRow(27), currentRow(28), currentRow(29), currentRow(30), currentRow(31), currentRow(32), currentRow(33),
                                currentRow(34), currentRow(35), currentRow(36), currentRow(37), currentRow(38))

                       
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
            'Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oDT)
            Return dv


        Catch ex As Exception

            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally
            
        End Try

    End Function

    Public Function GetDataViewFromTXT_HDR(ByVal CurrFileToUpload As String, ByVal Filename As String, ByRef sErrDesc As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromTXT_DET()
        '   Purpose     :   This function will upload the data from CSV file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************

        Dim dv As DataView

        Dim sFuncName As String = String.Empty
        sErrDesc = String.Empty
        Dim dvEntiry As DataView = New DataView(p_oEntitesDetails)
        ' Dim dvGLAcccount As DataView = New DataView(p_oGLAccount)
        Dim oDvNewGlCode As DataView = New DataView(p_oSTOLDCODE)
        Dim oDVCompanyData As DataView = New DataView(p_oDTCompanyData)
        Dim oDVipowerPeriod As DataView = New DataView(p_oDTiPowerPeriod)
        Dim oDVsapPeriod As DataView = New DataView(p_oDTSAPPeriod)
        Dim dperioddate As Date
        Dim sEntity As String = String.Empty
        Dim sBUCode As String = String.Empty
        Dim sLOS As String = String.Empty
        Dim oSR As StreamReader
        Dim iCount As Integer = 1
        Dim sGLAccount As String = String.Empty
        Dim sSAPPeriodCode As String = String.Empty
        oDT_OUCODE = New DataTable()


        Try
            sFuncName = "GetDataViewFromTXT_DET"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            'Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oDT As New DataTable()
            p_oDTImportStatistics = New DataTable()

            oDT.Columns.Add("tran_hdr_uid", GetType(String))  ''0
            oDT.Columns.Add("tran_uid", GetType(String))
            oDT.Columns.Add("tran_type", GetType(String))
            oDT.Columns.Add("tran_uid2", GetType(String))
            oDT.Columns.Add("tran_date", GetType(DateTime))
            oDT.Columns.Add("del_date", GetType(DateTime))
            oDT.Columns.Add("ref_uid", GetType(String))
            oDT.Columns.Add("ref_uid2", GetType(String))
            oDT.Columns.Add("cust_uid", GetType(String))
            oDT.Columns.Add("cust_name", GetType(String))
            oDT.Columns.Add("cust_name2", GetType(String))  ''10
            oDT.Columns.Add("cust_attn", GetType(String))
            oDT.Columns.Add("cust_addr", GetType(String))
            oDT.Columns.Add("cust_addr2", GetType(String))
            oDT.Columns.Add("cust_addr3", GetType(String))
            oDT.Columns.Add("cust_addr4", GetType(String))   ''15
            oDT.Columns.Add("cust_note", GetType(String))
            oDT.Columns.Add("cust_note2", GetType(String))
            oDT.Columns.Add("cust_note3", GetType(String))
            oDT.Columns.Add("cust_note4", GetType(String))
            oDT.Columns.Add("bill_to_uid", GetType(String))  ''20
            oDT.Columns.Add("slsp_uid", GetType(String))
            oDT.Columns.Add("proj_uid", GetType(String))
            oDT.Columns.Add("job_uid", GetType(String))
            oDT.Columns.Add("disc", GetType(String))
            oDT.Columns.Add("disc_type", GetType(String))
            oDT.Columns.Add("disc2", GetType(String))
            oDT.Columns.Add("disc_type2", GetType(String)) ''27
            oDT.Columns.Add("svc_charge", GetType(String))
            oDT.Columns.Add("svc_charge_type", GetType(String))
            oDT.Columns.Add("tax_rate", GetType(String))
            oDT.Columns.Add("tax_incl", GetType(String))
            oDT.Columns.Add("pay_term_uid", GetType(String))
            oDT.Columns.Add("pay_meth_uid", GetType(String)) ''33
            oDT.Columns.Add("time_in", GetType(String))
            oDT.Columns.Add("time_out", GetType(String))
            oDT.Columns.Add("whsp_uid", GetType(String))
            oDT.Columns.Add("loc_from_uid", GetType(String))
            oDT.Columns.Add("loc_to_uid", GetType(String))
            oDT.Columns.Add("confirm", GetType(String))
            oDT.Columns.Add("void_tran", GetType(String))
            oDT.Columns.Add("note", GetType(String))
            oDT.Columns.Add("note2", GetType(String))
            oDT.Columns.Add("note3", GetType(String))
            oDT.Columns.Add("note4", GetType(String))
            oDT.Columns.Add("deposit", GetType(String))
            oDT.Columns.Add("coll_cash_amt", GetType(String))
            oDT.Columns.Add("coll_cheq_amt", GetType(String))
            oDT.Columns.Add("coll_cheq_no", GetType(String))
            oDT.Columns.Add("coll_cheq_bank", GetType(String))
            oDT.Columns.Add("coll_cheq_date", GetType(String))
            oDT.Columns.Add("coll_ccard_amt", GetType(String))
            oDT.Columns.Add("coll_ccard_no", GetType(String))
            oDT.Columns.Add("coll_ccard_iss", GetType(String))

          
            Using MyReader As New Microsoft.VisualBasic.
                    FileIO.TextFieldParser(
                      CurrFileToUpload)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        If currentRow(0) = "tran_hdr_uid" Then Continue While
                        oDT.Rows.Add(currentRow(0), currentRow(1), currentRow(2), currentRow(3), DateTime.ParseExact(Left(currentRow(4), 10), "yyyy-MM-dd", Nothing),
                               DateTime.ParseExact(Left(currentRow(5), 10), "yyyy-MM-dd", Nothing), currentRow(6), currentRow(7), currentRow(8), currentRow(9), currentRow(10),
                               currentRow(11), currentRow(12), currentRow(13), currentRow(14), (currentRow(15)), currentRow(16), (currentRow(17)), currentRow(18), (currentRow(19)), currentRow(20), currentRow(21),
                               currentRow(22), currentRow(24), currentRow(25), currentRow(26), currentRow(27), currentRow(28), currentRow(29), currentRow(30), currentRow(31), currentRow(32), currentRow(33),
                               currentRow(34), currentRow(35), currentRow(36), currentRow(37), currentRow(38), currentRow(39), currentRow(40), currentRow(41), currentRow(42), currentRow(43), currentRow(44),
                               currentRow(45), currentRow(46), currentRow(47), currentRow(48), currentRow(49), currentRow(50), currentRow(51), currentRow(52), currentRow(53))


                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While
            End Using


            'Console.WriteLine("Del_schema() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
            'Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oDT)
            Return dv

        Catch ex As Exception

            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally
           
        End Try

    End Function


    Public Function Create_Sales_order(ByVal oDVHdr As DataView, ByVal oDVDet As DataView, ByRef oCompany As SAPbobsCOM.Company, ByRef sErrdesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oOrders As SAPbobsCOM.Documents
        Dim oOrdersLines As SAPbobsCOM.Document_Lines
        Dim odtLines As New Data.DataTable
        Dim intRetCode As Integer = 0
        Dim strErrMsg As String = ""
        Dim sDocEntry As String = String.Empty

        Try
            oOrders = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            sFuncName = "Create_Sales_order()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oOrders.CardCode = oDVHdr.Item(0)("bill_to_uid").ToString()
            oOrders.NumAtCard = oDVHdr.Item(0)("tran_uid").ToString()
            oOrders.DocDueDate = oDVHdr.Item(0)("del_date").ToString()
            oOrders.DocDate = oDVHdr.Item(0)("tran_date").ToString()
            oOrders.TaxDate = oDVHdr.Item(0)("tran_date").ToString()
            oOrders.SalesPersonCode = oDVHdr.Item(0)("slsp_uid").ToString()
            oOrders.UserFields.Fields.Item("U_MEVO_REF_HDR").Value = oDVHdr.Item(0)("tran_hdr_uid").ToString()
            For Each odr As DataRowView In oDVDet
                oOrdersLines = oOrders.Lines
                oOrdersLines.ItemCode = odr("item_uid").ToString()
                oOrdersLines.UnitPrice = odr("Qty").ToString()
                oOrdersLines.Quantity = odr("Uprice").ToString()
                oOrdersLines.UserFields.Fields.Item("U_MEVO_REF_DET").Value = odr("tran_det_uid").ToString()
                oOrdersLines.UserFields.Fields.Item("U_MEVO_REF_HDR").Value = odr("tran_hdr_uid").ToString()
                oOrdersLines.UserFields.Fields.Item("U_MEVO_REF_LNUM").Value = odr("line_no").ToString()
                oOrdersLines.ItemCode = odr("item_uid").ToString()
                oOrdersLines.Quantity = odr("Qty").ToString()
                oOrdersLines.Add()
            Next

            intRetCode = oOrders.Add()
            If intRetCode <> 0 Then
                p_oCompany.GetLastError(intRetCode, strErrMsg)
                sErrdesc = strErrMsg
                Throw New ArgumentException(sErrdesc)
            Else

                oCompany.GetNewObjectCode(sDocEntry)
                sErrdesc = String.Empty
            End If
            Console.WriteLine("Completed with SUCCESS " & sDocEntry, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS  " & sDocEntry, sFuncName)
            Create_Sales_order = RTN_SUCCESS
        Catch ex As Exception
            sErrdesc = ex.Message
            Console.WriteLine("Completed with ERROR  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Create_Sales_order = RTN_ERROR
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oOrders)
            oOrders = Nothing
        End Try
    End Function

    Public Function GetCompanyDetails(ByVal sEntity As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetCompanyDetails()
        '   Purpose     :   This function will get the relavent Banking informations with respective Entities 
        '   Parameters  :   ByRef sEntity AS String 
        '                       sEntity = Entity Name
        '   Author      :   JOHN
        '   Date        :   MAY 2014 20
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        sFuncName = "GetCompanyDetails()"

        Try
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Dim Findatarow() As DataRow = p_oEntitesDetails.Select("Entity = '" & sEntity.ToString.Trim & "'")

            For Each row As DataRow In Findatarow
                p_sSAPEntityName = row(2)
                p_sSAPUName = row(5)
                p_sSAPUPass = row(6)
            Next

            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)
            GetCompanyDetails = RTN_SUCCESS

        Catch ex As Exception
            Console.WriteLine("Completed With ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            GetCompanyDetails = RTN_ERROR
        End Try

    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                          ByVal sEntity As String, _
                                          ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2013 21
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet

        Try
            sFuncName = "ConnectToTargetCompany()"
            Console.WriteLine("Starting function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            Console.WriteLine("Initializing the Company Object ", sFuncName)
            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
            Console.WriteLine("Assigning the representing database name ", sFuncName)
            oCompany.Server = p_oCompDef.sServer

            If p_oCompDef.sServerType = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf p_oCompDef.sServerType = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf p_oCompDef.sServerType = "2014" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            End If


            oCompany.LicenseServer = p_oCompDef.sLicenseServer
            oCompany.CompanyDB = sEntity
            oCompany.UserName = p_oCompDef.sSAPUser
            oCompany.Password = p_oCompDef.sSAPPwd

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            Console.WriteLine("Connecting to the Company Database. ", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL Server : " & oCompany.Server & " SQL Type " & p_oCompDef.sServerType _
              & " License Server " & p_oCompDef.sLicenseServer & "  CompanyDB " & p_sSAPEntityName & " User Name " & p_sSAPUName & " pass " & p_sSAPUPass, sFuncName)

            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Sub FileMoveToArchive(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal iStatus As Integer, ByVal sErrDesc As String)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   JOHN 
        'Date       :   21 MAY 2014

        Dim sFuncName As String = String.Empty
        Dim sExtension As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            Console.WriteLine("Starting Function ", sFuncName)
            sExtension = oFile.Extension
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & sExtension

            If iStatus = RTN_SUCCESS Then
                Console.WriteLine("Moving CSV file to success folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to success folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
            Else
                Console.WriteLine("Moving CSV file to Fail folder ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in renaming/copying/moving ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Sub FileMoveToArchives(ByVal oFile As String, ByVal CurrFileToUpload As String, ByVal iStatus As Integer, ByVal sExtension As String, ByVal sErrDesc As String)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   JOHN 
        'Date       :   21 MAY 2014

        Dim sFuncName As String = String.Empty
        'Dim sExtension As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"
            Console.WriteLine("Starting Function ", sFuncName)
            Dim oFiles() As System.IO.FileInfo
            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            'sExtension = oFile.Extension
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile, 1, oFile.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & sExtension
          
            oFiles = DirInfo.GetFiles("*.*")
            For Each File As System.IO.FileInfo In oFiles
                Dim filename As String = File.Name
                If filename = oFile Then
                    If iStatus = RTN_SUCCESS Then
                        Console.WriteLine("Moving CSV file to success folder ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to success folder", sFuncName)
                        File.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
                        'oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)

                    Else
                        Console.WriteLine("Moving CSV file to Fail folder ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving CSV file to Fail folder", sFuncName)
                        File.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
                        'oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
                    End If
                End If
            Next
           
        Catch ex As Exception
            Console.WriteLine("Error in renaming/copying/moving ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function Del_schema(ByVal csvFileFolder As String) As Long

        ' ***********************************************************************************
        '   Function   :    Del_schema()
        '   Purpose    :    This function is handles - Delete the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Del_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim FileToDelete As String
            FileToDelete = csvFileFolder & "\\schema.ini"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Del_schema = RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Del_schema = RTN_ERROR
        End Try
    End Function

    Public Function Create_schema(ByVal csvFileFolder As String, ByVal FileName As String) As Long

        ' ***********************************************************************************
        '   Function   :    Create_schema()
        '   Purpose    :    This function is handles - Create the Schema file
        '   Parameters :    ByVal csvFileFolder As String
        '                       csvFileFolder = Passing file name
        '   Author     :    JOHN
        '   Date       :    26/06/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Create_schema()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            Dim csvFileName As String = FileName
            Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
            Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
            'Dim s1, s2, s3, s4, s5 As String

            srOutput.WriteLine("[" & csvFileName & "]")
            srOutput.WriteLine("ColNameHeader=False")
            srOutput.WriteLine("Format=CSVDelimited")
            srOutput.WriteLine("Col1=F1 Text")
            srOutput.WriteLine("Col2=F2 Text")
            srOutput.WriteLine("Col3=F3 Text")
            srOutput.WriteLine("Col4=F4 Text")
            srOutput.WriteLine("Col5=F5 Text")
            srOutput.WriteLine("Col6=F6 Text")
            srOutput.WriteLine("Col7=F7 Text")
            srOutput.WriteLine("Col8=F8 Text")
            srOutput.WriteLine("Col9=F9 Text")
            srOutput.WriteLine("Col10=F10 Double")
            srOutput.WriteLine("Col11=F11 Text")
            srOutput.WriteLine("Col12=F12 Double")
            srOutput.WriteLine("Col13=F13 Text")
            srOutput.WriteLine("Col14=F14 Text")
            srOutput.WriteLine("Col15=F15 Text")
            srOutput.WriteLine("MaxScanRows=0")
            srOutput.WriteLine("CharacterSet=OEM")
            'srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf)
            srOutput.Close()
            fsOutput.Close()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            Create_schema = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            Create_schema = RTN_ERROR
        End Try

    End Function

    Public Function Write_TextFile(ByVal oDTDisplay As DataTable, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty

        Try
            Dim irow As Integer
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "Validationip.txt"
            Dim sbuffer As String = String.Empty

            sFuncName = "Write_TextFile()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("Validation Error!  The following OU Codes are not Existing / No Entities Tagged for this OU Codes  ")
            sw.WriteLine("")
            sw.WriteLine("Line No.  OU Code        Message                                                       ")
            sw.WriteLine("=======================================================================================")
            sw.WriteLine(" ")

            For irow = 0 To oDTDisplay.Rows.Count - 1
                If Not String.IsNullOrEmpty(oDTDisplay.Rows(irow).Item(0).ToString) Then
                    sw.WriteLine(oDTDisplay.Rows(irow).Item(0).ToString.PadRight(10, " "c) + oDTDisplay.Rows(irow).Item(1).ToString.PadRight(17, " "c) _
                                 + oDTDisplay.Rows(irow).Item(2).ToString.PadRight(57, " "c))
                Else
                    Exit For
                End If
            Next irow

            sw.WriteLine(" ")
            sw.WriteLine("========================================================================================")
            sw.WriteLine("Please Check.")
            sw.Close()
            Process.Start(sPath & sFileName)

            Write_TextFile = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function Write_TextFile_I(ByVal sString As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty

        Try
            Dim sPath As String = System.Windows.Forms.Application.StartupPath & "\"
            Dim sFileName As String = "Validationip.txt"
            Dim sbuffer As String = String.Empty

            sFuncName = "Write_TextFile_I()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            If File.Exists(sPath & sFileName) Then
                Try
                    File.Delete(sPath & sFileName)
                Catch ex As Exception
                End Try
            End If

            Dim sw As StreamWriter = New StreamWriter(sPath & sFileName)
            ' Add some text to the file.
            sw.WriteLine("")
            sw.WriteLine("")
            sw.WriteLine("Validation Error!    " & sString)
            sw.WriteLine(" ")
            sw.WriteLine(" ")
            sw.WriteLine("========================================================================================")
            sw.WriteLine("Please Check.")
            sw.Close()
            Process.Start(sPath & sFileName)

            Write_TextFile_I = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Write_TextFile_I = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function MergeAutoNumberedToDataTable(ByVal SourceTable As DataTable, ByVal sErrDesc As String) As DataTable

        'Function   :   MergeAutoNumberedToDataTable()
        'Purpose    :   
        'Parameters :   ByVal SourceTable As DataTable
        '                   SourceTable= Source Datatable
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        '                   =
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   John
        'Date       :   19-05-2015
        'Change     :

        Dim sFuncName As String

        Try
            sFuncName = "MergeAutoNumberedToDataTable()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Dim ResultTable As DataTable = New DataTable()
            Dim AutoNumberColumn As DataColumn = New DataColumn()
            AutoNumberColumn.ColumnName = "SNo"
            AutoNumberColumn.DataType = GetType(Integer)
            AutoNumberColumn.AutoIncrement = True
            AutoNumberColumn.AutoIncrementSeed = 1
            AutoNumberColumn.AutoIncrementStep = 1
            ResultTable.Columns.Add(AutoNumberColumn)
            ResultTable.Merge(SourceTable)
            ResultTable.Columns(0).SetOrdinal(ResultTable.Columns.Count - 1)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Return ResultTable
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try


    End Function

End Module