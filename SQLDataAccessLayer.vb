Imports System
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Xml
Imports PrePress.BusinessLogicLayer


Namespace DataAccessLayer



    '*********************************************************************
    '
    ' SQLDataAccessLayer Class
    '
    ' The SQLDataAccessLayer contains the data access layer for Microsoft
    ' SQL Server. This class implements the abstract methods in the
    ' DataAccessLayerBaseClass class.
    '
    '*********************************************************************

    Public Class SQLDataAccessLayer
        Inherits DataAccessLayerBaseClass


        '*********************************************************************
        '
        ' Constants
        '
        ' Each of the constants below represent the name of a SQL Stored
        ' Procedure. If you need to change the name of any stored procedure
        ' used by the CTCR, modify one of the constants.
        '
        '*********************************************************************

        Private Const SP_COMPANY_CREATE As String = "spI_COMPANY"
        Private Const SP_COMPANY_UPDATE As String = "spU_COMPANY"
        Private Const SP_COMPANY_DELETE As String = "spD_COMPANY"
        Private Const SP_COMPANY_SELECT As String = "spS_COMPANY"

        Private Const SP_PERSONPREF_CREATE As String = "spI_PERSONPREF"
        Private Const SP_PERSONPREF_UPDATE As String = "spU_PERSONPREF"
        Private Const SP_PERSONPREF_DELETE As String = "spD_PERSONPREF"
        Private Const SP_PERSONPREF_SELECT As String = "spS_PERSONPREF"



        Private Const SP_EVENTLOG_CREATE As String = "spI_EVENTLOG"
        Private Const SP_EVENTLOG_UPDATE As String = "spU_EVENTLOG"
        Private Const SP_EVENTLOG_DELETE As String = "spD_EVENTLOG"
        Private Const SP_EVENTLOG_SELECT As String = "spS_EVENTLOG"

        Private Const SP_PERSON_CREATE As String = "spI_PERSON"
        Private Const SP_PERSON_UPDATE As String = "spU_PERSON"
        Private Const SP_PERSON_DELETE As String = "spD_PERSON"
        Private Const SP_PERSON_SELECT As String = "spS_PERSON"

        Private Const SP_ATTACHMENT_CREATE As String = "spI_ATTACHMENT"
        Private Const SP_ATTACHMENT_UPDATE As String = "spU_ATTACHMENT"
        Private Const SP_ATTACHMENT_DELETE As String = "spD_ATTACHMENT"
        Private Const SP_ATTACHMENT_SELECT As String = "spS_ATTACHMENT"

        Private Const SP_TASKORDER_CREATE As String = "spI_TASKORDER"
        Private Const SP_TASKORDER_UPDATE As String = "spU_TASKORDER"
        Private Const SP_TASKORDER_DELETE As String = "spD_TASKORDER"
        Private Const SP_TASKORDER_SELECT As String = "spS_TASKORDER"
        Private Const SP_TASKORDER_GETLIST As String = "spS_GetTaskOrderList"
        Private Const SP_TASKORDER_CHECK_DUPLICATE As String = "spS_CheckForDuplicateTaskOrder"


        Private Const SP_STATUS_CREATE As String = "spI_STATUS"
        Private Const SP_STATUS_UPDATE As String = "spU_STATUS"
        Private Const SP_STATUS_DELETE As String = "spD_STATUS"
        Private Const SP_STATUS_SELECT As String = "spS_STATUS"

        Private Const SP_SECURITYLEVEL_CREATE As String = "spI_SECURITYLEVEL"
        Private Const SP_SECURITYLEVEL_UPDATE As String = "spU_SECURITYLEVEL"
        Private Const SP_SECURITYLEVEL_DELETE As String = "spD_SECURITYLEVEL"
        Private Const SP_SECURITYLEVEL_SELECT As String = "spS_SECURITYLEVEL"


        Private Const SP_ITEMIZE_CREATE As String = "spI_ITEMIZE"
        Private Const SP_ITEMIZE_UPDATE As String = "spU_ITEMIZE"
        Private Const SP_ITEMIZE_DELETE As String = "spD_ITEMIZE"
        Private Const SP_ITEMIZE_SELECT As String = "spS_ITEMIZE"

        Private Const spS_TASKORDER_ITEMIZE As String = "spS_TASKORDER_ITEMIZE"
        Private Const spS_BIDSEARCHMODAL As String = "spS_BID_MODAL_TABLE"
        Private Const spS_PERSON_QUALIFY As String = "spS_PERSON_QUALIFY_V2"
        Private Const spS_TASKORDERPastDue As String = "spS_TASKORDER_PastDue"
        Private Const spS_TASKORDERFutureDuebyDate As String = "spS_TASKORDER_FutureDuebyDate"
        Private Const sPS_EDITORPERFORMANCEMETRICS As String = "sPS_EDITORPERFORMANCEMETRICS"
        Private Const SP_TASKORDER_Person As String = "spS_PersonTASKORDER"

        Private Const SP_COPYEDITLEVEL_CREATE As String = "spI_COPYEDITSKILLS"
        Private Const SP_COPYEDITLEVEL_UPDATE As String = "spU_COPYEDITSKILLS"
        Private Const SP_COPYEDITLEVEL_DELETE As String = "spD_COPYEDITSKILLS"
        Private Const SP_COPYEDITLEVEL_SELECT As String = "spS_COPYEDITSKILLS"

        Private Const SP_LK_USERDISCIPLINE_SELECT As String = "spS_lk_UserDiscipline"

        Private Const SP_UserDisciplineSet_CREATE As String = "spI_USER_Discipline"
        Private Const SP_UserDisciplineSet_UPDATE As String = "spU_USER_Discipline"
        Private Const SP_UserDisciplineSet_DELETE As String = "spD_USER_Discipline"
        Private Const SP_UserDisciplineSet_SELECT As String = "spS_USER_Discipline"

        Private Const SP_STYLETYPELOOKUP_SELECT As String = "spS_lk_UserStyleSkill"

        Private Const SP_USERSTYLE_CREATE As String = "spI_USER_STYLE"
        Private Const SP_USERSTYLE_UPDATE As String = "spU_USER_STYLE"
        Private Const SP_USERSTYLE_DELETE As String = "spD_USER_STYLE"
        Private Const SP_USERSTYLE_SELECT As String = "spS_USER_STYLE"



        Private Const SP_USERCOPYEDITSKILLSET_CREATE As String = "spI_USER_COPYEDITSKILL"
        Private Const SP_USERCOPYEDITSKILLSET_UPDATE As String = "spU_USER_COPYEDITSKILL"
        Private Const SP_USERCOPYEDITSKILLSET_DELETE As String = "spD_USER_COPYEDITSKILL"
        Private Const SP_USERCOPYEDITSKILLSET_SELECT As String = "spS_USER_COPYEDITSKILL"

        Private Const SP_COPYEDITLOOKUP_SELECT As String = "spS_lk_CopyEditSkill"

        Private Const SP_USERLANGUAGE_CREATE As String = "spI_USER_LANGUAGE"
        Private Const SP_USERLANGUAGE_DELETE As String = "spD_USER_LANGUAGE"
        Private Const SP_USERLANGUAGE_SELECT As String = "spS_USER_LANGUAGEs"

        Private Const SP_LANGUAGELOOKUP_SELECT As String = "spS_lk_LANGUAGE"


        Private Const SP_BID_CREATE As String = "spI_BID"
        Private Const SP_BID_UPDATE As String = "spU_BID"
        Private Const SP_BID_DELETE As String = "spD_BID"
        Private Const SP_BID_SELECT As String = "spS_BID"

        Private Const SP_BIDSTATUS_CREATE As String = "spI_BID_STATUS"
        Private Const SP_BIDSTATUS_UPDATE As String = "spU_BID_STATUS"
        Private Const SP_BIDSTATUS_DELETE As String = "spD_BID_STATUS"
        Private Const SP_BIDSTATUS_SELECT As String = "spS_BID_STATUS"


        Private Const SP_BIDEVENT_CREATE As String = "spI_BID_EVENT"
        Private Const SP_BIDEVENT_UPDATE As String = "spU_BID_EVENT"
        Private Const SP_BIDEVENT_DELETE As String = "spD_BID_EVENT"
        Private Const SP_BIDEVENT_SELECT As String = "spS_BID_EVENT"

        Private Const SP_BLACKOUTWINDOW_CREATE As String = "spI_USER_BLACKOUTWINDOW"
        Private Const SP_BLACKOUTWINDOW_UPDATE As String = "spU_USER_BLACKOUTWINDOW"
        Private Const SP_BLACKOUTWINDOW_DELETE As String = "spD_USER_BLACKOUTWINDOW"
        Private Const SP_BLACKOUTWINDOW_SELECT As String = "spS_USER_BLACKOUTWINDOW"
        Private Const SP_BLACKOUTWINDOWTO_CREATE As String = "spI_USER_BLACKOUTWINDOWTaskOrder"

        Private Const SP_JOB_CREATE As String = "spI_JOB"
        Private Const SP_JOB_UPDATE As String = "spU_JOB"
        Private Const SP_JOB_DELETE As String = "spD_JOB"
        Private Const SP_JOB_SELECT As String = "spS_JOB"

        Private Const SP_LIBRARYATTACHMENT_CREATE As String = "spI_LIBRARY_ATTACHMENT"
        Private Const SP_LIBRARYATTACHMENT_UPDATE As String = "spU_LIBRARY_ATTACHMENT"
        Private Const SP_LIBRARYATTACHMENT_DELETE As String = "spD_LIBRARY_ATTACHMENT"
        Private Const SP_LIBRARYATTACHMENT_SELECT As String = "spS_LIBRARY_ATTACHMENT"
        '*** INSTANCE PROPERTIES ***

        Private Const SP_COPYEDITLEVELLIST_SELECT As String = "spS_COPYEDITLEVEL"
        Private Const SP_COPYEDITLEVELLOOKUP_DELETE As String = "spD_COPYEDITLEVEL"
        Private Const SP_COPYEDITLEVELLOOKUP_CREATE As String = "spI_COPYEDITLEVEL"
        Private Const SP_COPYEDITLEVELLOOKUP_UPDATE As String = "spU_COPYEDITLEVEL"
        Private Const SP_COPYEDITLEVELLOOKUP_MOVE As String = "sp_COPYEDITLEVEL_MOVE"

        Private Const SP_CURRENCY_SELECT As String = "spS_Currencies"

        Private Const SP_ACCOUNT_SELECT As String = "spS_Accounts"
        Private Const SP_EditorPerformanceRpt_SELECT As String = "spS_GetEditorPerformanceReportData"
        Private Const SP_GetEditorialMetricsRpt_SELECT As String = "rptsp_GetEditorialMetrics" '
        Private Const SP_GetCertificationRpt_SELECT As String = "sps_GetCertificationReport"
        Private Const SP_GetRptCopyEditLevels As String = "rptsp_GetCopyEditLevels"
        Private Const SP_GetAssignedFreelancersReport_SELECT As String = "rptsp_GetAssignedFreelancersReport"
        Private Const SP_GetPersonsBySkills_SELECT As String = "spS_GetPersonsBySkills"

        Private Const SP_TaskOrderPaymentLog_CREATE As String = "spI_TaskOrderPaymentLog"
        Private Const SP_TaskOrderPaymentLog_UPDATE As String = "spU_TaskOrderPaymentLog"
        Private Const SP_TaskOrderPaymentLog_SELECT As String = "spS_TaskOrderPaymentLog"
        Private Const spS_PERSON_GT As String = "spS_PERSON_GT"
        Private Const SP_PaymentHistory As String = "rptsp_PaymentHistory"
        Private Const SP_GetTaskorderComments_SELECT As String = "spS_GetTaskorderComments"
        Private Const SP_GetTaskorderSubBatchDetails_SELECT As String = "spS_GetTaskorderSubBatchDetails"

        Private Const SP_Renegotiation_Select As String = "spS_GetRenegotiationDetails"
        Private Const SP_Renegotiation_Insert As String = "spI_Renegotiation"
        Private Const SP_Renegotiation_Update As String = "spU_Renegotiation"

        Private Const SP_TimesheetLog_Select As String = "spS_TimesheetLog"
        Private Const SP_TimesheetLogWeek_Select As String = "spS_TimesheetWeekView"
        Private Const SP_TimesheetLog_Insert As String = "spI_TimesheetLog"
        Private Const SP_TimesheetLog_Update As String = "spU_TimesheetLog"
        Private Const SP_TimesheetLog_Delete As String = "spD_TimesheetLog"
        Private Const SP_Timesheet_Submit As String = "spTS_SubmitTimesheet"
        Private Const SP_Timesheet_Certify As String = "spTS_CertifyTimesheet"
        Private Const SP_TimesheetReport_Select As String = "spS_TimesheetReport"
        Private Const SP_TimesheetManagerReport_Select As String = "spS_TimesheetReportForManager"

        Private Const SP_USERCurrency_CREATE As String = "spI_USER_Currency"
        Private Const SP_USERCurrency_DELETE As String = "spD_USER_Currency"
        Private Const SP_USERCurrency_SELECT As String = "spS_USER_Currencies"

        Private Const SP_DIRECTPAYMENTS_DELETE As String = "spD_DIRECTPAYMENTS"
        Private Const SP_DIRECTPAYMENTS_CREATE As String = "spI_DIRECTPAYMENTS"
        Private Const SP_DIRECTPAYMENTS_UPDATE As String = "spU_DIRECTPAYMENTS"
        Private Const SP_DIRECTPAYMENTS_PAID As String = "spU_MARK_DIRECTPAYMENTS_PAID"

        '*** INSTANCE METHODS ***

#Region "*********** BLACKOUTWINDOW Methods ***********"

        '***************************************************
        '**
        '** BLACKOUTWINDOW Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewBLACKOUTWINDOW(ByVal objClass As BLACKOUTWINDOW) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewBLACKOUTWINDOW")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            'AddParamToSQLCmd(sqlCmd, "@UserBlackoutId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UserBlackoutId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@StartDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.StartDate)
            AddParamToSQLCmd(sqlCmd, "@EndDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.EndDate)
            AddParamToSQLCmd(sqlCmd, "@BlackoutDescription", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.BlackoutDescription)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOW_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateBLACKOUTWINDOW(ByVal objClass As BLACKOUTWINDOW) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateBLACKOUTWINDOW")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@UserBlackoutId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UserBlackoutId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@StartDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.StartDate)
            AddParamToSQLCmd(sqlCmd, "@EndDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.EndDate)
            AddParamToSQLCmd(sqlCmd, "@BlackoutDescription", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.BlackoutDescription)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOW_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteBLACKOUTWINDOWById(ByVal UserBlackoutId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@UserBlackoutId", SqlDbType.Int, 0, ParameterDirection.Input, UserBlackoutId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOW_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetBLACKOUTWINDOWByID(ByVal UserBlackoutId As Integer) As BLACKOUTWINDOW
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@UserBlackoutId", SqlDbType.Int, 0, ParameterDirection.Input, UserBlackoutId)
            AddParamToSQLCmd(sqlCmd, "@Today", SqlDbType.DateTime, 4, ParameterDirection.Input, DateTime.Now().Date)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOW_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBLACKOUTWINDOWCollectionFromReader)
            Dim iCollection As BLACKOUTWINDOWCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BLACKOUTWINDOWCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function getUserBlackoutWindowByPersonId(ByVal PersonId As Integer, Optional ByVal ShowAll As Boolean = False) As BLACKOUTWINDOWCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            If (ShowAll = False) Then
                AddParamToSQLCmd(sqlCmd, "@Today", SqlDbType.DateTime, 4, ParameterDirection.Input, DateTime.Now().Date)
            End If
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOW_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBLACKOUTWINDOWCollectionFromReader)
            Dim iCollection As BLACKOUTWINDOWCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BLACKOUTWINDOWCollection)

            Return iCollection
        End Function


        Public Overrides Function CreateNewBLACKOUTWINDOWTaskOrder(ByVal PersonId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal BDescription As String, ByVal TaskOrderId As Integer, ByVal Context As String) As Integer

            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            'AddParamToSQLCmd(sqlCmd, "@UserBlackoutId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UserBlackoutId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@StartDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EndDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)
            AddParamToSQLCmd(sqlCmd, "@BlackoutDescription", SqlDbType.VarChar, 300, ParameterDirection.Input, BDescription)
            AddParamToSQLCmd(sqlCmd, "@TaskOrderId", SqlDbType.Int, 4, ParameterDirection.Input, TaskOrderId)
            AddParamToSQLCmd(sqlCmd, "@Context", SqlDbType.VarChar, 200, ParameterDirection.Input, Context)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BLACKOUTWINDOWTO_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


#End Region '***********BLACKOUTWINDOW Methods ***********

#Region "*********** COMPANY Methods ***********"

        '***************************************************
        '**
        '** COMPANY Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewCOMPANY(ByVal objClass As COMPANY) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewCOMPANY")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@companyname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.companyname)
            AddParamToSQLCmd(sqlCmd, "@isinternalTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.isinternalTF)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.isenabledTF)
            AddParamToSQLCmd(sqlCmd, "@address1", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address1)
            AddParamToSQLCmd(sqlCmd, "@address2", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address2)
            AddParamToSQLCmd(sqlCmd, "@city", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.city)
            AddParamToSQLCmd(sqlCmd, "@statecode", SqlDbType.VarChar, 2, ParameterDirection.Input, objClass.statecode)
            AddParamToSQLCmd(sqlCmd, "@zipcode", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.zipcode)
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateCOMPANY(ByVal objClass As COMPANY) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateCOMPANY")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.company_id)
            AddParamToSQLCmd(sqlCmd, "@companyname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.companyname)
            AddParamToSQLCmd(sqlCmd, "@isinternalTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.isinternalTF)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.isenabledTF)
            AddParamToSQLCmd(sqlCmd, "@address1", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address1)
            AddParamToSQLCmd(sqlCmd, "@address2", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address2)
            AddParamToSQLCmd(sqlCmd, "@city", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.city)
            AddParamToSQLCmd(sqlCmd, "@statecode", SqlDbType.VarChar, 2, ParameterDirection.Input, objClass.statecode)
            AddParamToSQLCmd(sqlCmd, "@zipcode", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.zipcode)
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteCOMPANYById(ByVal company_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 0, ParameterDirection.Input, company_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetCOMPANYByID(ByVal company_id As Integer) As COMPANY
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 0, ParameterDirection.Input, company_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOMPANYCollectionFromReader)
            Dim iCollection As COMPANYCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COMPANYCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetALLCOMPANYSBySearch(ByVal IsInternal As Integer, ByVal IsEnabled As Integer, ByVal CountryId As Integer) As COMPANYCollection
            Dim sqlCmd As New SqlCommand

            '-1 is wildcard


            If IsInternal > -1 Then
                AddParamToSQLCmd(sqlCmd, "@isinternalTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, IsInternal)
            End If

            If IsEnabled > -1 Then
                AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, IsEnabled)
            End If
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOMPANYCollectionFromReader)
            Dim iCollection As COMPANYCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COMPANYCollection)

            Return iCollection

        End Function

        Public Overrides Function GetALLCOMPANYSByCompanyName(ByVal companyname As String, ByVal CountryId As Integer) As COMPANYCollection
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@companyname", SqlDbType.VarChar, 50, ParameterDirection.Input, companyname)
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOMPANYCollectionFromReader)
            Dim iCollection As COMPANYCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COMPANYCollection)

            Return iCollection

        End Function

        Public Overrides Function GetALLCOMPANYS(ByVal CountryId As Integer) As COMPANYCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COMPANY_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOMPANYCollectionFromReader)
            Dim iCollection As COMPANYCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COMPANYCollection)

            Return iCollection

        End Function




#End Region '***********COMPANY Methods ***********

#Region "*********** PERSONPREF Methods ***********"

        '***************************************************
        '**
        '** PERSONPREF Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewPERSONPREF(ByVal objClass As PERSONPREF) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewPERSONPREF")
            End If

            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)

            AddParamToSQLCmd(sqlCmd, "@personID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.personID)
            AddParamToSQLCmd(sqlCmd, "@upkey", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.upkey)
            AddParamToSQLCmd(sqlCmd, "@upvalue", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.upvalue)
            AddParamToSQLCmd(sqlCmd, "@updescription", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.updescription)
            AddParamToSQLCmd(sqlCmd, "@upseq", SqlDbType.Int, 4, ParameterDirection.Input, objClass.upseq)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSONPREF_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function



        Public Overrides Function UpdatePERSONPREF(ByVal objClass As PERSONPREF) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdatePERSONPREF")
            End If

            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@personID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.personID)
            AddParamToSQLCmd(sqlCmd, "@upkey", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.upkey)
            AddParamToSQLCmd(sqlCmd, "@upvalue", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.upvalue)
            AddParamToSQLCmd(sqlCmd, "@updescription", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.updescription)
            AddParamToSQLCmd(sqlCmd, "@upseq", SqlDbType.Int, 4, ParameterDirection.Input, objClass.upseq)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSONPREF_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function



        Public Overrides Function DeletePERSONPREFById(ByVal personID As Integer, ByVal upkey As String) As Boolean
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)

            AddParamToSQLCmd(sqlCmd, "@personID", SqlDbType.Int, 4, ParameterDirection.Input, personID)
            AddParamToSQLCmd(sqlCmd, "@upkey", SqlDbType.VarChar, 100, ParameterDirection.Input, upkey)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSONPREF_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function



        Public Overrides Function GetPERSONPREFByID(ByVal personID As Integer, ByVal upkey As String) As PERSONPREF

            If personID <= 0 Then
                Throw New Exception("GetPERSONPREFByID:personID should be a positive number.")
            End If

            If Len(upkey) = 0 Then
                Throw New Exception("GetPERSONPREFByID:upkey must not be empty.")
            End If

            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@personID", SqlDbType.Int, 4, ParameterDirection.Input, personID)
            AddParamToSQLCmd(sqlCmd, "@upkey", SqlDbType.VarChar, 100, ParameterDirection.Input, upkey)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSONPREF_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONPREFCollectionFromReader)
            Dim iCollection As PERSONPREFCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONPREFCollection)

            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If

        End Function

        Public Overrides Function GetALLPERSONPREFByID(ByVal personID As Integer, ByVal upkey As String) As PERSONPREFCollection

            If personID <= 0 Then
                Throw New Exception("GetPERSONPREFByID:personID should be a positive number.")
            End If

            If Len(upkey) = 0 Then
                Throw New Exception("GetPERSONPREFByID:upkey must not be empty.")
            End If

            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@personID", SqlDbType.Int, 4, ParameterDirection.Input, personID)
            AddParamToSQLCmd(sqlCmd, "@upkey", SqlDbType.VarChar, 100, ParameterDirection.Input, upkey)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSONPREF_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONPREFCollectionFromReader)
            Dim iCollection As PERSONPREFCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONPREFCollection)

            Return iCollection

        End Function



#End Region '*********** PERSONPREF Methods ***********

#Region "*********** SECURITYLEVEL Methods ***********"

        '***************************************************
        '**
        '** SECURITYLEVEL Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewSECURITYLEVEL(ByVal objClass As SECURITYLEVEL) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewSECURITYLEVEL")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@securityname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.securityname)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_SECURITYLEVEL_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateSECURITYLEVEL(ByVal objClass As SECURITYLEVEL) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateSECURITYLEVEL")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.securitylevel_code)
            AddParamToSQLCmd(sqlCmd, "@securityname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.securityname)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_SECURITYLEVEL_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteSECURITYLEVELById(ByVal securitylevel_code As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 0, ParameterDirection.Input, securitylevel_code)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_SECURITYLEVEL_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetSECURITYLEVELByID(ByVal securitylevel_code As Integer) As SECURITYLEVEL
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 0, ParameterDirection.Input, securitylevel_code)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_SECURITYLEVEL_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSECURITYLEVELCollectionFromReader)
            Dim iCollection As SECURITYLEVELCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), SECURITYLEVELCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetALLSECURITYLEVELS() As SECURITYLEVELCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_SECURITYLEVEL_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSECURITYLEVELCollectionFromReader)
            Dim iCollection As SECURITYLEVELCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), SECURITYLEVELCollection)
            Return iCollection

        End Function


#End Region '***********SECURITYLEVEL Methods ***********

#Region "*********** EVENTLOG Methods ***********"

        '***************************************************
        '**
        '** EVENTLOG Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewGLOBALEVENTLOG(ByVal objClass As GLOBALEVENTLOG) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewEVENTLOG")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@object_code", SqlDbType.Char, 4, ParameterDirection.Input, objClass.object_code)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.object_id)
            AddParamToSQLCmd(sqlCmd, "@person_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.person_id)
            AddParamToSQLCmd(sqlCmd, "@eventtype", SqlDbType.VarChar, 3000, ParameterDirection.Input, objClass.eventtype)
            AddParamToSQLCmd(sqlCmd, "@eventDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.eventDT)
            AddParamToSQLCmd(sqlCmd, "@reporting_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.reporting_code)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EVENTLOG_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function UpdateGLOBALEVENTLOG(ByVal objClass As GLOBALEVENTLOG) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateEVENTLOG")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@object_code", SqlDbType.Char, 4, ParameterDirection.Input, objClass.object_code)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.object_id)
            AddParamToSQLCmd(sqlCmd, "@eventDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.eventDT)
            AddParamToSQLCmd(sqlCmd, "@reporting_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.reporting_code)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EVENTLOG_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteGLOBALEVENTLOGById(ByVal object_code As String, ByVal object_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@object_code", SqlDbType.Char, 4, ParameterDirection.Input, object_code)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, object_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EVENTLOG_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CBool(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetGLOBALEVENTLOG(ByVal object_code As String, ByVal object_id As Integer) As GLOBALEVENTLOGCollection
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@object_code", SqlDbType.Char, 4, ParameterDirection.Input, object_code)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, object_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EVENTLOG_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateEVENTLOGCollectionFromReader)
            Dim iCollection As GLOBALEVENTLOGCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), GLOBALEVENTLOGCollection)

            Return iCollection

        End Function

        Public Overrides Function GetALLGLOBALEVENTLOG() As GLOBALEVENTLOGCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EVENTLOG_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateEVENTLOGCollectionFromReader)
            Dim iCollection As GLOBALEVENTLOGCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), GLOBALEVENTLOGCollection)

            Return iCollection
        End Function

#End Region '***********EVENTLOG Methods ***********

#Region "*********** PERSON Methods ***********"

        '***************************************************
        '**
        '** PERSON Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewPERSON(ByVal objClass As PERSON) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewPERSON")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.company_id)
            AddParamToSQLCmd(sqlCmd, "@firstname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.firstname)
            AddParamToSQLCmd(sqlCmd, "@lastname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.lastname)
            AddParamToSQLCmd(sqlCmd, "@middleinitial", SqlDbType.Char, 1, ParameterDirection.Input, objClass.middleinitial)
            AddParamToSQLCmd(sqlCmd, "@suffixname", SqlDbType.VarChar, 5, ParameterDirection.Input, objClass.suffixname)
            AddParamToSQLCmd(sqlCmd, "@address1", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address1)
            AddParamToSQLCmd(sqlCmd, "@address2", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address2)
            AddParamToSQLCmd(sqlCmd, "@city", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.city)
            AddParamToSQLCmd(sqlCmd, "@statecode", SqlDbType.Char, 2, ParameterDirection.Input, objClass.statecode)
            AddParamToSQLCmd(sqlCmd, "@zipcode", SqlDbType.VarChar, 10, ParameterDirection.Input, objClass.zipcode)
            AddParamToSQLCmd(sqlCmd, "@workphone", SqlDbType.VarChar, 25, ParameterDirection.Input, objClass.workphone)
            AddParamToSQLCmd(sqlCmd, "@mobilephone", SqlDbType.VarChar, 25, ParameterDirection.Input, objClass.mobilephone)
            AddParamToSQLCmd(sqlCmd, "@email", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.email)
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.securitylevel_code)
            AddParamToSQLCmd(sqlCmd, "@UserNotes", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.UserNotes)
            AddParamToSQLCmd(sqlCmd, "@password", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.password)
            AddParamToSQLCmd(sqlCmd, "@resetpasswordTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.resetpasswordTF)
            AddParamToSQLCmd(sqlCmd, "@failcount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.failcount)
            AddParamToSQLCmd(sqlCmd, "@attid", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.attid)
            AddParamToSQLCmd(sqlCmd, "@expireDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.expireDT)
            AddParamToSQLCmd(sqlCmd, "@enabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.enabledTF)
            AddParamToSQLCmd(sqlCmd, "@deletedTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.deletedTF)
            AddParamToSQLCmd(sqlCmd, "@ECNotes", SqlDbType.NVarChar, 1000, ParameterDirection.Input, objClass.ECNotes)
            'AddParamToSQLCmd(sqlCmd, "@Currency", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.Currency)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.CountryID)
            AddParamToSQLCmd(sqlCmd, "@AccountId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.AccountId)
            AddParamToSQLCmd(sqlCmd, "@GTType", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.GTType)
            AddParamToSQLCmd(sqlCmd, "@CreatedOn", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.CreatedOn)
            AddParamToSQLCmd(sqlCmd, "@CreatedBY", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.CreatedBy)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdatePERSON(ByVal objClass As PERSON) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdatePERSON")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@person_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.person_id)
            AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.company_id)
            AddParamToSQLCmd(sqlCmd, "@firstname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.firstname)
            AddParamToSQLCmd(sqlCmd, "@lastname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.lastname)
            AddParamToSQLCmd(sqlCmd, "@middleinitial", SqlDbType.Char, 1, ParameterDirection.Input, objClass.middleinitial)
            AddParamToSQLCmd(sqlCmd, "@suffixname", SqlDbType.VarChar, 5, ParameterDirection.Input, objClass.suffixname)
            AddParamToSQLCmd(sqlCmd, "@address1", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address1)
            AddParamToSQLCmd(sqlCmd, "@address2", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.address2)
            AddParamToSQLCmd(sqlCmd, "@city", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.city)
            AddParamToSQLCmd(sqlCmd, "@statecode", SqlDbType.Char, 2, ParameterDirection.Input, objClass.statecode)
            AddParamToSQLCmd(sqlCmd, "@zipcode", SqlDbType.VarChar, 10, ParameterDirection.Input, objClass.zipcode)
            AddParamToSQLCmd(sqlCmd, "@workphone", SqlDbType.VarChar, 25, ParameterDirection.Input, objClass.workphone)
            AddParamToSQLCmd(sqlCmd, "@mobilephone", SqlDbType.VarChar, 25, ParameterDirection.Input, objClass.mobilephone)
            AddParamToSQLCmd(sqlCmd, "@email", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.email)
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.securitylevel_code)
            AddParamToSQLCmd(sqlCmd, "@UserNotes", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.UserNotes)
            AddParamToSQLCmd(sqlCmd, "@password", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.password)
            AddParamToSQLCmd(sqlCmd, "@resetpasswordTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.resetpasswordTF)
            AddParamToSQLCmd(sqlCmd, "@failcount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.failcount)
            AddParamToSQLCmd(sqlCmd, "@attid", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.attid)
            AddParamToSQLCmd(sqlCmd, "@expireDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.expireDT)
            AddParamToSQLCmd(sqlCmd, "@enabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.enabledTF)
            AddParamToSQLCmd(sqlCmd, "@deletedTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, objClass.deletedTF)
            AddParamToSQLCmd(sqlCmd, "@ECNotes", SqlDbType.NVarChar, 1000, ParameterDirection.Input, objClass.ECNotes)
            'AddParamToSQLCmd(sqlCmd, "@Currency", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.Currency)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.CountryID)
            AddParamToSQLCmd(sqlCmd, "@AccountId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.AccountId)
            AddParamToSQLCmd(sqlCmd, "@GTType", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.GTType)
            AddParamToSQLCmd(sqlCmd, "@ModifiedOn", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.ModifiedOn)
            AddParamToSQLCmd(sqlCmd, "@ModifiedBy", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.ModifiedBy)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeletePERSONById(ByVal person_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@person_id", SqlDbType.Int, 0, ParameterDirection.Input, person_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetPERSONByID(ByVal person_id As Integer, Optional ByVal Enabled As Boolean = False) As PERSON
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@person_id", SqlDbType.Int, 0, ParameterDirection.Input, person_id)

            If Enabled Then
                AddParamToSQLCmd(sqlCmd, "@enabledTF", SqlDbType.Int, 0, ParameterDirection.Input, 0)
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetPERSONByEmail(ByVal email As String) As PERSON
            Dim sqlCmd As New SqlCommand

            If Len(Trim(email)) = 0 Then
                Return Nothing
            End If

            AddParamToSQLCmd(sqlCmd, "@email", SqlDbType.VarChar, 150, ParameterDirection.Input, email)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetPERSONByAttId(ByVal attid As String, ByVal curPersonId As Integer) As PERSON
            Dim sqlCmd As New SqlCommand

            If Len(Trim(attid)) = 0 Then
                Return Nothing
            End If

            AddParamToSQLCmd(sqlCmd, "@AttId", SqlDbType.VarChar, 50, ParameterDirection.Input, attid)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)
            If iCollection.Count > 0 Then
                For Each itm As PERSON In iCollection
                    If itm.person_id <> curPersonId Then
                        Return itm
                    End If
                Next
                Return Nothing
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetALLPERSONS() As PERSONCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)

            Return iCollection

        End Function


        Public Overrides Function GetALLPERSONSBySearch(ByVal opensearch As String, ByVal SecurityLevelCode As Integer, ByVal MinFailCount As Integer, ByVal company_id As Integer, ByVal enabledTF As Integer, ByVal CountryId As Integer, ByVal AccountId As Integer) As PERSONCollection
            Dim sqlCmd As New SqlCommand

            If company_id > 0 Then
                AddParamToSQLCmd(sqlCmd, "@company_id", SqlDbType.Int, 4, ParameterDirection.Input, company_id)
            End If

            If Len(opensearch) > 0 Then
                AddParamToSQLCmd(sqlCmd, "@opensearch", SqlDbType.VarChar, 50, ParameterDirection.Input, opensearch)
            End If

            If SecurityLevelCode > 0 Then
                AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.Int, 4, ParameterDirection.Input, SecurityLevelCode)
            End If

            If MinFailCount > 0 Then
                AddParamToSQLCmd(sqlCmd, "@failcount", SqlDbType.Int, 4, ParameterDirection.Input, MinFailCount)
            End If

            If enabledTF > -1 Then
                AddParamToSQLCmd(sqlCmd, "@enabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, enabledTF)
            End If

            'Only return non-deleted users
            AddParamToSQLCmd(sqlCmd, "@deletedTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, 0)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)

            If AccountId > -1 Then
                AddParamToSQLCmd(sqlCmd, "@AccountId", SqlDbType.Int, 4, ParameterDirection.Input, AccountId)
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)

            Return iCollection

        End Function

        Public Overrides Function GetEditorPerformanceReportData() As System.Data.DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_EditorPerformanceRpt_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds)
            Return ds
        End Function

        Public Overrides Function GetVendorReportData() As PERSONCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PERSON_SELECT)
            AddParamToSQLCmd(sqlCmd, "@enabledTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, 1)
            AddParamToSQLCmd(sqlCmd, "@deletedTF", SqlDbType.TinyInt, 1, ParameterDirection.Input, 0)
            AddParamToSQLCmd(sqlCmd, "@securitylevel_code", SqlDbType.TinyInt, 1, ParameterDirection.Input, 20)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePERSONCollectionFromReader)
            Dim iCollection As PERSONCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), PERSONCollection)

            Return iCollection

        End Function

        Public Overrides Function GetPersonsBySkillIdList(ByVal skills As String) As System.Data.DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "Skills", SqlDbType.NVarChar, 250, ParameterDirection.Input, skills)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetPersonsBySkills_SELECT)

            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds)
            Return ds
        End Function

#End Region '***********PERSON Methods ***********

#Region "*********** Attachment Methods ***********"

        '*********************************************************************
        ' Attachment Methods
        '
        ' The following methods are used for working with Attachments.
        '
        '*********************************************************************

        Public Overrides Function InsertAttachment(ByVal oAttachment As ATTACHMENT) As Integer
            ' Validate Parameters
            If oAttachment Is Nothing Then
                Throw New ArgumentNullException("oAttachment is Nothing")
            End If

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@unit_id", SqlDbType.Int, 0, ParameterDirection.Input, oAttachment.unit_id)
            AddParamToSQLCmd(sqlCmd, "@FileName", SqlDbType.VarChar, 255, ParameterDirection.Input, oAttachment.filename)
            AddParamToSQLCmd(sqlCmd, "@FileSize", SqlDbType.BigInt, 0, ParameterDirection.Input, oAttachment.filesize)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NText, 0, ParameterDirection.Input, oAttachment.contenttype)
            sqlCmd.Parameters.Add("@Attachment", SqlDbType.Image, oAttachment.Attachment.Length)
            sqlCmd.Parameters("@Attachment").Value = oAttachment.Attachment
            AddParamToSQLCmd(sqlCmd, "@objecttype", SqlDbType.VarChar, 20, ParameterDirection.Input, oAttachment.objecttype)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, oAttachment.object_id)
            AddParamToSQLCmd(sqlCmd, "@attachmentnote", SqlDbType.VarChar, 500, ParameterDirection.Input, oAttachment.attachmentnote)
            AddParamToSQLCmd(sqlCmd, "@round", SqlDbType.Int, 4, ParameterDirection.Input, oAttachment.Round)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_CREATE)
            ExecuteScalarCmd(sqlCmd)

            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)

        End Function 'InsertAttachment


        Public Overrides Function UpdateAttachment(ByVal oAttachment As ATTACHMENT) As Boolean
            ' Validate Parameters
            If oAttachment Is Nothing Then
                Throw New ArgumentNullException("oAttachment is Nothing")
            End If
            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@unit_id", SqlDbType.Int, 0, ParameterDirection.Input, oAttachment.unit_id)
            AddParamToSQLCmd(sqlCmd, "@AttachmentID", SqlDbType.Int, 0, ParameterDirection.Input, oAttachment.attachment_id)
            AddParamToSQLCmd(sqlCmd, "@FileName", SqlDbType.NText, 0, ParameterDirection.Input, oAttachment.filename)
            AddParamToSQLCmd(sqlCmd, "@FileSize", SqlDbType.BigInt, 0, ParameterDirection.Input, oAttachment.filesize)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NText, 0, ParameterDirection.Input, oAttachment.contenttype)
            sqlCmd.Parameters.Add("@Attachment", SqlDbType.Image, oAttachment.Attachment.Length)
            sqlCmd.Parameters("@Attachment").Value = oAttachment.Attachment
            AddParamToSQLCmd(sqlCmd, "@objecttype", SqlDbType.VarChar, 20, ParameterDirection.Input, oAttachment.objecttype)
            AddParamToSQLCmd(sqlCmd, "@object_id", SqlDbType.Int, 4, ParameterDirection.Input, oAttachment.object_id)
            AddParamToSQLCmd(sqlCmd, "@attachmentnote", SqlDbType.VarChar, 500, ParameterDirection.Input, oAttachment.attachmentnote)
            AddParamToSQLCmd(sqlCmd, "@round", SqlDbType.Int, 4, ParameterDirection.Input, oAttachment.Round)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)

        End Function 'UpdateAttachment



        Public Overrides Function DeleteAttachment(ByVal AttachmentID As Integer) As Boolean
            ' Validate Parameters
            If AttachmentID <= 0 Then
                Throw New ArgumentOutOfRangeException("AttachmentID")
            End If

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@Attachment_ID", SqlDbType.Int, 0, ParameterDirection.Input, AttachmentID)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Dim returnValue As Integer = CInt(sqlCmd.Parameters("@ReturnValue").Value)
            Return IIf(returnValue = 0, True, False)

        End Function 'DeleteAttachment



        Public Overrides Function GetAttachmentByID(ByVal AttachmentID As Integer) As ATTACHMENT

            Dim oAttachment As ATTACHMENT = Nothing

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@Attachment_ID", SqlDbType.Int, 0, ParameterDirection.Input, AttachmentID)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_SELECT)

            ' Execute Reader
            If ConnectionString = String.Empty Then
                Throw New ArgumentOutOfRangeException("ConnectionString")
            End If

            Dim cn As New SqlConnection(Me.ConnectionString)
            Try
                sqlCmd.Connection = cn
                cn.Open()
                Dim dtr As SqlDataReader = sqlCmd.ExecuteReader()
                If dtr.Read() Then
                    oAttachment = New ATTACHMENT(CInt(dtr("Attachment_ID")), CInt(dtr("unit_id")), CStr(dtr("FileName")), CInt(dtr("FileSize")), CStr(dtr("ContentType")), CType(dtr("Attachment"), Byte()), CStr(dtr("ObjectType")), CInt(dtr("Object_ID")), CStr(dtr("attachmentnote")), CInt(dtr("Round")))
                End If
            Catch ex As Exception
                'Do Something with this error
            Finally
                cn.Close()
            End Try


            Return oAttachment

        End Function 'GetAttachmentByID



        Public Overrides Function GetAttachmentByObject(ByVal ObjectType As String, ByVal ObjectID As Integer) As AttachmentCollection

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ObjectType", SqlDbType.VarChar, 20, ParameterDirection.Input, ObjectType)
            AddParamToSQLCmd(sqlCmd, "@Object_ID", SqlDbType.Int, 0, ParameterDirection.Input, ObjectID)


            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateAttachmentCollectionFromReader)
            Dim iCollection As AttachmentCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), AttachmentCollection)
            Return iCollection

        End Function

        Public Overrides Function GetAttachmentByObjectAndRound(ByVal ObjectType As String, ByVal ObjectID As Integer, ByVal Round As Integer) As AttachmentCollection

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ObjectType", SqlDbType.VarChar, 20, ParameterDirection.Input, ObjectType)
            AddParamToSQLCmd(sqlCmd, "@Object_ID", SqlDbType.Int, 0, ParameterDirection.Input, ObjectID)
            AddParamToSQLCmd(sqlCmd, "@round", SqlDbType.Int, 0, ParameterDirection.Input, Round)


            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ATTACHMENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateAttachmentCollectionFromReader)
            Dim iCollection As AttachmentCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), AttachmentCollection)
            Return iCollection

        End Function

#End Region '*********** Attachment Methods ***********

#Region "*********** TASKORDER Methods ***********"

        '***************************************************
        '**
        '** TASKORDER Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewTASKORDER(ByVal objClass As TASKORDER) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewTASKORDER")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.assignedperson_id)
            AddParamToSQLCmd(sqlCmd, "@apexmgrperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.apexmgrperson_id)
            AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.status_code)
            AddParamToSQLCmd(sqlCmd, "@jobnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.jobnumber)
            AddParamToSQLCmd(sqlCmd, "@batchnumber", SqlDbType.VarChar, 10, ParameterDirection.Input, objClass.batchnumber)
            AddParamToSQLCmd(sqlCmd, "@bookid", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bookid)
            AddParamToSQLCmd(sqlCmd, "@booktitle", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.booktitle)
            AddParamToSQLCmd(sqlCmd, "@bookshortname", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.bookshortname)
            AddParamToSQLCmd(sqlCmd, "@roundnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.roundnumber)
            AddParamToSQLCmd(sqlCmd, "@requestedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.requestedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@proposedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.proposedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@agreedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.agreedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@actualcompletedDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.actualcompletedDT)
            AddParamToSQLCmd(sqlCmd, "@instructions", SqlDbType.VarChar, 3000, ParameterDirection.Input, objClass.instructions)
            AddParamToSQLCmd(sqlCmd, "@designtemplate", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.designtemplate)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevel)
            AddParamToSQLCmd(sqlCmd, "@hardcopyeditsTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.hardcopyeditsTF)
            AddParamToSQLCmd(sqlCmd, "@refmanuattachedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.refmanuattachedTF)
            AddParamToSQLCmd(sqlCmd, "@styledpagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.priceperpage)
            AddParamToSQLCmd(sqlCmd, "@perpageoverrideTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.perpageoverrideTF)
            AddParamToSQLCmd(sqlCmd, "@totalprice", SqlDbType.Money, 8, ParameterDirection.Input, objClass.totalprice)
            AddParamToSQLCmd(sqlCmd, "@authorname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.authorname)
            AddParamToSQLCmd(sqlCmd, "@coordperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.coordperson_id)
            AddParamToSQLCmd(sqlCmd, "@ispaidTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ispaidTF)
            AddParamToSQLCmd(sqlCmd, "@subbatchnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.subbatchnumber)
            AddParamToSQLCmd(sqlCmd, "@remarks", SqlDbType.VarChar, 3000, ParameterDirection.Input, objClass.remarks)
            AddParamToSQLCmd(sqlCmd, "@isreviewrequired", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isReviewRequired)
            AddParamToSQLCmd(sqlCmd, "@currencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyID)
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)
            AddParamToSQLCmd(sqlCmd, "@isInterimPaidTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isInterimPaidTF)
            AddParamToSQLCmd(sqlCmd, "@InterimPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.InterimPaymentDT)
            AddParamToSQLCmd(sqlCmd, "@IsInterimCertified", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsInterimCertified)
            AddParamToSQLCmd(sqlCmd, "@IsCompletedByEC", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsCompletedByEC)
            AddParamToSQLCmd(sqlCmd, "@PASScore", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PASScore)
            AddParamToSQLCmd(sqlCmd, "@Comments", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.CertificationComments)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateTASKORDER(ByVal objClass As TASKORDER) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateTASKORDER")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.taskorder_id)
            AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.assignedperson_id)
            AddParamToSQLCmd(sqlCmd, "@apexmgrperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.apexmgrperson_id)
            AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.status_code)
            AddParamToSQLCmd(sqlCmd, "@jobnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.jobnumber)
            AddParamToSQLCmd(sqlCmd, "@batchnumber", SqlDbType.VarChar, 10, ParameterDirection.Input, objClass.batchnumber)
            AddParamToSQLCmd(sqlCmd, "@bookid", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bookid)
            AddParamToSQLCmd(sqlCmd, "@booktitle", SqlDbType.VarChar, 100, ParameterDirection.Input, objClass.booktitle)
            AddParamToSQLCmd(sqlCmd, "@bookshortname", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.bookshortname)
            AddParamToSQLCmd(sqlCmd, "@roundnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.roundnumber)
            AddParamToSQLCmd(sqlCmd, "@requestedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.requestedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@proposedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.proposedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@agreedreturnDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.agreedreturnDT)
            AddParamToSQLCmd(sqlCmd, "@actualcompletedDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.actualcompletedDT)
            AddParamToSQLCmd(sqlCmd, "@instructions", SqlDbType.VarChar, 3000, ParameterDirection.Input, objClass.instructions)
            AddParamToSQLCmd(sqlCmd, "@designtemplate", SqlDbType.VarChar, 20, ParameterDirection.Input, objClass.designtemplate)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevel)
            AddParamToSQLCmd(sqlCmd, "@hardcopyeditsTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.hardcopyeditsTF)
            AddParamToSQLCmd(sqlCmd, "@refmanuattachedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.refmanuattachedTF)
            AddParamToSQLCmd(sqlCmd, "@styledpagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.priceperpage)
            AddParamToSQLCmd(sqlCmd, "@perpageoverrideTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.perpageoverrideTF)
            AddParamToSQLCmd(sqlCmd, "@totalprice", SqlDbType.Money, 8, ParameterDirection.Input, objClass.totalprice)
            AddParamToSQLCmd(sqlCmd, "@authorname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.authorname)
            AddParamToSQLCmd(sqlCmd, "@coordperson_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.coordperson_id)
            AddParamToSQLCmd(sqlCmd, "@ispaidTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ispaidTF)
            AddParamToSQLCmd(sqlCmd, "@subbatchnumber", SqlDbType.Int, 4, ParameterDirection.Input, objClass.subbatchnumber)
            AddParamToSQLCmd(sqlCmd, "@remarks", SqlDbType.VarChar, 3000, ParameterDirection.Input, objClass.remarks)
            AddParamToSQLCmd(sqlCmd, "@isreviewrequired", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isReviewRequired)
            AddParamToSQLCmd(sqlCmd, "@currencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyID)
            AddParamToSQLCmd(sqlCmd, "@countryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)
            AddParamToSQLCmd(sqlCmd, "@paymentDT", SqlDbType.DateTime, 4, ParameterDirection.Input, objClass.paymentDT)
            AddParamToSQLCmd(sqlCmd, "@isInterimPaidTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isInterimPaidTF)
            AddParamToSQLCmd(sqlCmd, "@InterimPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.InterimPaymentDT)
            AddParamToSQLCmd(sqlCmd, "@IsInterimCertified", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsInterimCertified)
            AddParamToSQLCmd(sqlCmd, "@IsCompletedByEC", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsCompletedByEC)
            AddParamToSQLCmd(sqlCmd, "@PASScore", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PASScore)
            AddParamToSQLCmd(sqlCmd, "@Comments", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.CertificationComments)
            AddParamToSQLCmd(sqlCmd, "@AgreedFinalReturnDateDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.AgreedFinalReturnDateDT)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteTASKORDERById(ByVal taskorder_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 0, ParameterDirection.Input, taskorder_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetTASKORDERByID(ByVal taskorder_id As Integer) As TASKORDER
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 0, ParameterDirection.Input, taskorder_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateTASKORDERCollectionFromReader)
            Dim iCollection As TASKORDERCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TASKORDERCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function


        Public Overrides Function GetALLTASKORDERBySearch(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal coordpersonid As Integer, ByVal Context As String, ByVal CountryId As Integer) As TASKORDERCollection
            Dim sqlCmd As New SqlCommand

            If assignedperson_id > 0 Then
                AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 4, ParameterDirection.Input, assignedperson_id)
            End If

            If _status_code > 0 Then
                AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, _status_code)
            End If

            If coordpersonid > 0 Then

                AddParamToSQLCmd(sqlCmd, "@coordperson_id", SqlDbType.Int, 4, ParameterDirection.Input, coordpersonid)
            End If
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            AddParamToSQLCmd(sqlCmd, "@Context", SqlDbType.VarChar, 50, ParameterDirection.Input, Context)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateTASKORDERCollectionFromReader)
            Dim iCollection As TASKORDERCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TASKORDERCollection)

            Return iCollection

        End Function

        Public Overrides Function GetTASKORDERList(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal coordpersonid As Integer, ByVal Context As String, ByVal CountryId As Integer, ByVal ViewerSecurityCode As Integer, ByVal JobBatchNumber As String) As TASKORDERCollection
            Dim sqlCmd As New SqlCommand

            If assignedperson_id > 0 Then
                AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 4, ParameterDirection.Input, assignedperson_id)
            End If

            If _status_code > 0 Then
                AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, _status_code)
            End If

            If coordpersonid > 0 Then

                AddParamToSQLCmd(sqlCmd, "@coordperson_id", SqlDbType.Int, 4, ParameterDirection.Input, coordpersonid)
            End If
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            AddParamToSQLCmd(sqlCmd, "@Context", SqlDbType.VarChar, 50, ParameterDirection.Input, Context)
            AddParamToSQLCmd(sqlCmd, "@ViewerSecurityCode", SqlDbType.Int, 4, ParameterDirection.Input, ViewerSecurityCode)
            AddParamToSQLCmd(sqlCmd, "@JobBatchNumber", SqlDbType.VarChar, 100, ParameterDirection.Input, JobBatchNumber)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_GETLIST)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateTASKORDERCollectionFromReader)
            Dim iCollection As TASKORDERCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TASKORDERCollection)

            Return iCollection

        End Function

        Public Overrides Function CheckForDuplicateTaskOrder(ByVal taskorder_id As Integer, ByVal jobNumber As String, ByVal batchNumber As String, ByVal subBatchNumber As String) As Boolean
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@TaskOrder_id", SqlDbType.Int, 4, ParameterDirection.Input, taskorder_id)
            AddParamToSQLCmd(sqlCmd, "@JobNumber", SqlDbType.VarChar, 50, ParameterDirection.Input, jobNumber)
            AddParamToSQLCmd(sqlCmd, "@BatchNumber", SqlDbType.VarChar, 150, ParameterDirection.Input, batchNumber)
            AddParamToSQLCmd(sqlCmd, "@SubBatchNumber", SqlDbType.VarChar, 100, ParameterDirection.Input, subBatchNumber)
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_CHECK_DUPLICATE)
            ExecuteScalarCmd(sqlCmd)
            Return sqlCmd.Parameters("@ReturnValue").Value

        End Function

        'Public Overrides Function GetALLTASKORDERBySearch(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal startrangeDT As Date, ByVal endrangeDT As Date, ByVal IsPaid As Integer) As TASKORDERCollection
        '    Dim sqlCmd As New SqlCommand

        '    If assignedperson_id > 0 Then
        '        AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 4, ParameterDirection.Input, assignedperson_id)
        '    End If

        '    If _status_code > 0 Then
        '        AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, _status_code)
        '    End If


        '    AddParamToSQLCmd(sqlCmd, "@StartRangeDT", SqlDbType.DateTime, 10, ParameterDirection.Input, startrangeDT)

        '    AddParamToSQLCmd(sqlCmd, "@EndRangeDT", SqlDbType.DateTime, 10, ParameterDirection.Input, endrangeDT)

        '    If IsPaid >= 0 Then
        '        '-1 = both
        '        '0 = not paid
        '        '1 = paid
        '        AddParamToSQLCmd(sqlCmd, "@ispaidTF", SqlDbType.Int, 4, ParameterDirection.Input, IsPaid)
        '    End If


        '    SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_SELECT)
        '    Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateTASKORDERCollectionFromReader)
        '    Dim iCollection As TASKORDERCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TASKORDERCollection)

        '    Return iCollection

        'End Function



#End Region '***********TASKORDER Methods ***********

#Region "*********** STATUS Methods ***********"

        '***************************************************
        '**
        '** STATUS Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewSTATUS(ByVal objClass As STATUS) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewSTATUS")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@statusname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.statusname)
            AddParamToSQLCmd(sqlCmd, "@statusaction", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.statusaction)
            AddParamToSQLCmd(sqlCmd, "@iconfile", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.iconfile)
            AddParamToSQLCmd(sqlCmd, "@isclosedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isclosedTF)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isenabledTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STATUS_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateSTATUS(ByVal objClass As STATUS) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateSTATUS")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.status_code)
            AddParamToSQLCmd(sqlCmd, "@statusname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.statusname)
            AddParamToSQLCmd(sqlCmd, "@statusaction", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.statusaction)
            AddParamToSQLCmd(sqlCmd, "@iconfile", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.iconfile)
            AddParamToSQLCmd(sqlCmd, "@isclosedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isclosedTF)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isenabledTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STATUS_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteSTATUSById(ByVal status_code As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 0, ParameterDirection.Input, status_code)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STATUS_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetSTATUSByID(ByVal status_code As Integer) As STATUS
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 0, ParameterDirection.Input, status_code)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STATUS_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSTATUSCollectionFromReader)
            Dim iCollection As STATUSCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), STATUSCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetALLSTATUSES() As STATUSCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STATUS_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSTATUSCollectionFromReader)
            Dim iCollection As STATUSCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), STATUSCollection)

            Return iCollection

        End Function


#End Region '***********STATUS Methods ***********


#Region "*********** ITEMIZE Methods ***********"

        '***************************************************
        '**
        '** ITEMIZE Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewITEMIZE(ByVal objClass As ITEMIZE) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewITEMIZE")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.taskorder_id)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevel)
            AddParamToSQLCmd(sqlCmd, "@styledpagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.priceperpage)
            AddParamToSQLCmd(sqlCmd, "@subtotal", SqlDbType.Money, 8, ParameterDirection.Input, objClass.subtotal)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@jobtype", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.jobtype)
            AddParamToSQLCmd(sqlCmd, "@manual_pagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.manual_styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@manual_priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.manual_priceperpage)
            AddParamToSQLCmd(sqlCmd, "@manual_selectedpercent", SqlDbType.Decimal, 8, ParameterDirection.Input, objClass.manual_selectedpercent)
            AddParamToSQLCmd(sqlCmd, "@jobpagetype", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.jobpagetype)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ITEMIZE_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateITEMIZE(ByVal objClass As ITEMIZE) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateITEMIZE")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@itemize_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.itemize_id)
            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.taskorder_id)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevel)
            AddParamToSQLCmd(sqlCmd, "@styledpagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.priceperpage)
            AddParamToSQLCmd(sqlCmd, "@subtotal", SqlDbType.Money, 8, ParameterDirection.Input, objClass.subtotal)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@jobtype", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.jobtype)
            AddParamToSQLCmd(sqlCmd, "@manual_pagecount", SqlDbType.Float, 8, ParameterDirection.Input, objClass.manual_styledpagecount)
            AddParamToSQLCmd(sqlCmd, "@manual_priceperpage", SqlDbType.Money, 8, ParameterDirection.Input, objClass.manual_priceperpage)
            AddParamToSQLCmd(sqlCmd, "@manual_selectedpercent", SqlDbType.Decimal, 8, ParameterDirection.Input, objClass.manual_selectedpercent)
            AddParamToSQLCmd(sqlCmd, "@jobpagetype", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.jobpagetype)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ITEMIZE_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteITEMIZEById(ByVal itemize_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@itemize_id", SqlDbType.Int, 0, ParameterDirection.Input, itemize_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ITEMIZE_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetITEMIZEByID(ByVal itemize_id As Integer) As ITEMIZE
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@itemize_id", SqlDbType.Int, 0, ParameterDirection.Input, itemize_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ITEMIZE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateITEMIZECollectionFromReader)
            Dim iCollection As ITEMIZECollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), ITEMIZECollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetITEMIZEByTaskOrderID(ByVal taskorder_id As Integer) As ITEMIZECollection
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 4, ParameterDirection.Input, taskorder_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ITEMIZE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateITEMIZECollectionFromReader)
            Dim iCollection As ITEMIZECollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), ITEMIZECollection)

            Return iCollection

        End Function


#End Region '***********ITEMIZE Methods ***********

#Region "*** ITEMIZE LOOKUP TABLE (NEW) (10/11/2012) ***"




#End Region
        '*** REPORTS ***
#Region "*** Custom DataSets Report ***"

        Public Overrides Function gen_accounting_itemized_report(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal ispaidTF As Integer, ByVal jobBatchNumber As String, ByVal CountryId As Integer, ByVal CurrentPersonId As Integer, Optional ByVal CurrencyId As Integer = -1) As DataSet

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)


            'If assignedperson_id > 0 Then
            '    AddParamToSQLCmd(sqlCmd, "@assignedperson_id", SqlDbType.Int, 10, ParameterDirection.Input, assignedperson_id)
            'End If

            'If status_code > 0 Then
            '    AddParamToSQLCmd(sqlCmd, "@status_code", SqlDbType.Int, 10, ParameterDirection.Input, status_code)
            'End If

            If Not Util.IsNullDate(StartRangeDT) Then

                AddParamToSQLCmd(sqlCmd, "@StartRangeDT", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(StartRangeDT, "M/d/yyyy"))
                AddParamToSQLCmd(sqlCmd, "@EndRangeDT", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(EndRangeDT, "M/d/yyyy"))
            End If


            If ispaidTF = -1 Then
                'do not add parameter
            Else
                AddParamToSQLCmd(sqlCmd, "@ispaidTF", SqlDbType.Int, 10, ParameterDirection.Input, ispaidTF)
            End If

            If jobBatchNumber <> "" AndAlso jobBatchNumber <> "All" Then
                AddParamToSQLCmd(sqlCmd, "@jobBatchNumber", SqlDbType.VarChar, 250, ParameterDirection.Input, jobBatchNumber)
            End If

            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, CurrencyId)
            AddParamToSQLCmd(sqlCmd, "@CurrentPersonId", SqlDbType.Int, 4, ParameterDirection.Input, CurrentPersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_TASKORDER_ITEMIZE)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_TASKORDER_ITEMIZE")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function GetPaymentReport(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal jobBatchNumber As String) As DataSet

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            If Not Util.IsNullDate(StartRangeDT) Then

                AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(StartRangeDT, "M/d/yyyy"))
                AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(EndRangeDT, "M/d/yyyy"))
            End If

            If jobBatchNumber <> "" AndAlso jobBatchNumber <> "All" Then
                AddParamToSQLCmd(sqlCmd, "@jobBatchNumber", SqlDbType.VarChar, 250, ParameterDirection.Input, jobBatchNumber)
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, "spS_PaymentHistory")
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "Payments")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function gen_bidsearch_modal_result(ByVal BidId As Integer) As DataSet

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            If BidId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 10, ParameterDirection.Input, BidId)
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_BIDSEARCHMODAL)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_BID_MODAL_TABLE")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function gen_bidsearchmodal_person_qualify(ByVal RequiredUserSkillId As Integer, ByVal RequiredCopyEditSkillId As Integer, ByVal RequiredStyleSkillId As Integer, ByVal BidId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal CountryId As Integer, ByVal AccountId As Integer) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            If RequiredUserSkillId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@RequiredUserSkillId", SqlDbType.Int, 10, ParameterDirection.Input, RequiredUserSkillId)
            End If

            If RequiredCopyEditSkillId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@RequiredCopyEditSkillId", SqlDbType.Int, 10, ParameterDirection.Input, RequiredCopyEditSkillId)
            End If

            If RequiredCopyEditSkillId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@RequiredStyleSkillId", SqlDbType.Int, 10, ParameterDirection.Input, RequiredStyleSkillId)
            End If

            If BidId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 10, ParameterDirection.Input, BidId)
            End If

            If Not Util.IsNullDate(StartDate) Then

                AddParamToSQLCmd(sqlCmd, "@WorkStartDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(StartDate, "M/d/yyyy"))
                AddParamToSQLCmd(sqlCmd, "@WorkEndDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(EndDate, "M/d/yyyy"))
            End If

            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)

            If AccountId > -1 Then
                AddParamToSQLCmd(sqlCmd, "@AccountId", SqlDbType.Int, 4, ParameterDirection.Input, AccountId)
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_PERSON_QUALIFY)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_PERSON_QUALIFY_V2")
            'MyCommand.Fill(ds)

            Return ds
        End Function



        Public Overrides Function gen_WorkPastDue(ByVal CountryId As Integer) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)


            ''AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 10, ParameterDirection.Input, taskorder_id)


            'If status_code > 0 Then
            '    AddParamToSQLCmd(sqlCmd, "status_code", SqlDbType.Int, 10, ParameterDirection.Input, status_code)
            'End If



            AddParamToSQLCmd(sqlCmd, "CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_TASKORDERPastDue)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_TASKORDERPastDue")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function gen_WorkClosing(ByVal CountryId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)


            ''AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 10, ParameterDirection.Input, taskorder_id)


            'If status_code > 0 Then
            '    AddParamToSQLCmd(sqlCmd, "status_code", SqlDbType.Int, 10, ParameterDirection.Input, status_code)
            'End If

            AddParamToSQLCmd(sqlCmd, "CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            AddParamToSQLCmd(sqlCmd, "startdate", SqlDbType.DateTime, 4, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "enddate", SqlDbType.DateTime, 4, ParameterDirection.Input, EndDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_TASKORDERFutureDuebyDate)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_TASKORDERFutureDuebyDate")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function GetEditorMetricsReport(ByVal ESType As String, ByVal rptType As String, startDate As Nullable(Of DateTime), endDate As Nullable(Of DateTime)) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "ESType", SqlDbType.NVarChar, 50, ParameterDirection.Input, ESType)
            AddParamToSQLCmd(sqlCmd, "Type", SqlDbType.NVarChar, 50, ParameterDirection.Input, rptType)
            AddParamToSQLCmd(sqlCmd, "SDate", SqlDbType.DateTime, 4, ParameterDirection.Input, startDate)
            AddParamToSQLCmd(sqlCmd, "EDate", SqlDbType.DateTime, 4, ParameterDirection.Input, endDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetEditorialMetricsRpt_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_GetEditorialMetricsRpt_SELECT")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function GetCertificationReport(startDate As Nullable(Of DateTime), endDate As Nullable(Of DateTime)) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "SDate", SqlDbType.DateTime, 4, ParameterDirection.Input, startDate)
            AddParamToSQLCmd(sqlCmd, "EDate", SqlDbType.DateTime, 4, ParameterDirection.Input, endDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetCertificationRpt_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "sps_GetCertificationReport")
            'MyCommand.Fill(ds)

            Return ds

        End Function


        Public Overrides Function GetRptCopyEditLevels() As DataTable
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetRptCopyEditLevels)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_GetRptCopyEditLevels")
            Return ds.Tables(0)
        End Function


        Public Overrides Function GetDraftTOByPersonCurrency(ByVal Person_id As Integer, ByVal Currency As String) As Integer
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, Person_id)
            AddParamToSQLCmd(sqlCmd, "@Currency", SqlDbType.VarChar, 8, ParameterDirection.Input, Currency)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TASKORDER_Person)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function gen_EditorPerformance() As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)


            ''AddParamToSQLCmd(sqlCmd, "@taskorder_id", SqlDbType.Int, 10, ParameterDirection.Input, taskorder_id)


            'If status_code > 0 Then
            '    AddParamToSQLCmd(sqlCmd, "status_code", SqlDbType.Int, 10, ParameterDirection.Input, status_code)
            'End If


            SetCommandType(sqlCmd, CommandType.StoredProcedure, sPS_EDITORPERFORMANCEMETRICS)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "sPS_EDITORPERFORMANCEMETRICS")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function gen_searchmodal_person_GT(ByVal ReturnDate As DateTime) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "@RequestReturnDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(ReturnDate, "M/d/yyyy"))
              
            SetCommandType(sqlCmd, CommandType.StoredProcedure, spS_PERSON_GT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "spS_PERSON_GT")
            'MyCommand.Fill(ds)

            Return ds
        End Function

        Public Overrides Function payment_history_report(ByVal ReportType As String, ByVal StartRangeDT As Date, ByVal EndRangeDT As Date) As DataSet

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "@ReportType", SqlDbType.NVarChar, 100, ParameterDirection.Input, ReportType)
            If Not Util.IsNullDate(StartRangeDT) Then
                AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(StartRangeDT, "M/d/yyyy"))
                AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(EndRangeDT, "M/d/yyyy"))
            End If

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_PaymentHistory)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_PaymentHistory")
            'MyCommand.Fill(ds)

            Return ds

        End Function


        Public Overrides Function GetAssignedFreelancersReport(ByVal jobBatchNumber As String, startDate As Nullable(Of DateTime), endDate As Nullable(Of DateTime)) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "JobBatchNumber", SqlDbType.NVarChar, 50, ParameterDirection.Input, jobBatchNumber)
            AddParamToSQLCmd(sqlCmd, "SDate", SqlDbType.DateTime, 4, ParameterDirection.Input, startDate)
            AddParamToSQLCmd(sqlCmd, "EDate", SqlDbType.DateTime, 4, ParameterDirection.Input, endDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetAssignedFreelancersReport_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds)
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function GetTaskorderComments(ByVal TaskId As Integer, ByVal PersonId As Integer, ByVal ReportType As Integer) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "TaskId", SqlDbType.Int, 50, ParameterDirection.Input, TaskId)
            AddParamToSQLCmd(sqlCmd, "PersonId", SqlDbType.Int, 50, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "IsTaskOrderDetails", SqlDbType.Int, 50, ParameterDirection.Input, ReportType)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetTaskorderComments_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "GetTaskorderComments")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        'GetTaskorderSubBatchDetails
        Public Overrides Function GetTaskorderSubBatchDetails(ByVal JobNumber As String, ByVal BatchNumber As String) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "JobNumber", SqlDbType.NVarChar, 50, ParameterDirection.Input, JobNumber)
            AddParamToSQLCmd(sqlCmd, "BatchNumber", SqlDbType.NVarChar, 100, ParameterDirection.Input, BatchNumber)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_GetTaskorderSubBatchDetails_SELECT)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "GetTaskorderSubBatchDetails")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function GetRenegotiationDetails(ByVal TaskOrderId As Integer, ByVal StatusCode As Integer) As DataTable
            Dim sqlCmd As New SqlCommand
            Dim dt As New DataTable
            Dim cn As New SqlConnection(Me.ConnectionString)

            AddParamToSQLCmd(sqlCmd, "TaskorderId", SqlDbType.Int, 4, ParameterDirection.Input, TaskOrderId)
            AddParamToSQLCmd(sqlCmd, "StatusCode", SqlDbType.Int, 4, ParameterDirection.Input, StatusCode)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_Renegotiation_Select)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(dt)
            'MyCommand.Fill(ds)

            Return dt

        End Function


        Public Overrides Function InsertRenegotiation(ByVal objClass As Renegotiation) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewRenegotiation")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@TaskOrderId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.TaskOrderId)
            AddParamToSQLCmd(sqlCmd, "@CurrentReturnDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.CurrentReturnDate)
            AddParamToSQLCmd(sqlCmd, "@NewDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.NewDate)
            AddParamToSQLCmd(sqlCmd, "@IsInterimOrFinal", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.IsInterimOrFinal)
            AddParamToSQLCmd(sqlCmd, "@CreatedOn", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.CreatedOn)
            AddParamToSQLCmd(sqlCmd, "@CreatedBy", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CreatedBy)
            AddParamToSQLCmd(sqlCmd, "@StatusCode", SqlDbType.Int, 4, ParameterDirection.Input, objClass.StatusCode)
            AddParamToSQLCmd(sqlCmd, "@Comments", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.Comments)
            AddParamToSQLCmd(sqlCmd, "@Round", SqlDbType.Int, 4, ParameterDirection.Input, objClass.Round)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_Renegotiation_Insert)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        
#End Region


#Region "*********** COPYEDITLEVEL Methods ***********"

        '***************************************************
        '**
        '** COPYEDITLEVEL Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewCOPYEDITLEVEL(ByVal objClass As COPYEDITLEVEL) As Integer
            If objClass Is Nothing Then

                Throw New ArgumentNullException("NewCOPYEDITLEVEL")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevelcode_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevelcode_id)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel_name", SqlDbType.Text, 300, ParameterDirection.Input, objClass.copyeditlevel_name)
            AddParamToSQLCmd(sqlCmd, "@CopyEditLevelPrice", SqlDbType.Decimal, 300, ParameterDirection.Input, objClass.CopyEditLevelPrice)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@isEditableTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isEditableTF)





            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVEL_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateCOPYEDITLEVEL(ByVal objClass As COPYEDITLEVEL) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateCOPYEDITLEVEL")
            End If


            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevelcode_id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.copyeditlevelcode_id)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel_name", SqlDbType.Text, 300, ParameterDirection.Input, objClass.copyeditlevel_name)
            AddParamToSQLCmd(sqlCmd, "@CopyEditLevelPrice", SqlDbType.Decimal, 300, ParameterDirection.Input, objClass.CopyEditLevelPrice)
            AddParamToSQLCmd(sqlCmd, "@displayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.displayorder)
            AddParamToSQLCmd(sqlCmd, "@isEditableTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isEditableTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVEL_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetCopyEditLevelbyId(ByVal copyeditcode_id As Integer) As COPYEDITLEVEL
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@copyeditcode_id", SqlDbType.Int, 0, ParameterDirection.Input, copyeditcode_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVEL_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITLEVELCollectionFromReader)
            Dim iCollection As COPYEDITLEVELCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITLEVELCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

#End Region '***********COPYEDITLEVEL Methods ***********


#Region "************COPYEDITTYPELOOKUPVALUES***************"

        Public Overrides Function getCopyEditbyLookupId(ByVal CopyEditSkillId As Integer) As COPYEDITTYPESLOOKUP

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)



            AddParamToSQLCmd(sqlCmd, "@CopyEditSkillId", SqlDbType.Int, 10, ParameterDirection.Input, CopyEditSkillId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLOOKUP_SELECT)

            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITTYPELOOKUPCollectionFromReader)
            Dim iCollection As COPYEDITTYPESLOOKUPCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITTYPESLOOKUPCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetAllDisciplineTypes() As DisciplineTypeLookupCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LK_USERDISCIPLINE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateDisciplineTypeLookupFromReader)
            Dim iCollection As DisciplineTypeLookupCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), DisciplineTypeLookupCollection)
            Return iCollection

        End Function

#End Region '***********DisciplineTypeLookup Methods ***********

#Region "*********** UserDisciplineSet Methods ***********"

        '***************************************************
        '**
        '** UserDisciplineSet Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewUserDisciplineSet(ByVal objClass As UserDisciplineSet) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewUserDisciplineSet")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@UserSkillId", SqlDbType.Int, 0, ParameterDirection.Input, objClass.UserSkillId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, objClass.PersonID)
            AddParamToSQLCmd(sqlCmd, "@SkillDescription", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.SkillDescription)


            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_UserDisciplineSet_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateUserDisciplineSet(ByVal objClass As UserDisciplineSet) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateUserDisciplineSet")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@UserSkillId", SqlDbType.Int, 0, ParameterDirection.Input, objClass.UserSkillId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, objClass.PersonID)
            AddParamToSQLCmd(sqlCmd, "@SkillDescription", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.SkillDescription)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_UserDisciplineSet_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteUserDisciplineSetById(ByVal personId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@personId", SqlDbType.Int, 0, ParameterDirection.Input, personId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_UserDisciplineSet_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetUSERDISCIPLINESETByPERSONId(ByVal PersonId As Integer) As UserDisciplineSetCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_UserDisciplineSet_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUserDisciplineSetCollectionFromReader)
            Dim iCollection As UserDisciplineSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserDisciplineSetCollection)

            Return iCollection
        End Function

        Public Overrides Function GetAllUserDisciplineSets() As UserDisciplineSetCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_UserDisciplineSet_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUserDisciplineSetCollectionFromReader)
            Dim iCollection As UserDisciplineSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserDisciplineSetCollection)

            Return iCollection

        End Function


#End Region


#Region "*********** USERSTYLE Methods ***********"

        '***************************************************
        '**
        '** USERSTYLE Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewUSERSTYLESet(ByVal objClass As UserStyleSet) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewUSERSTYLE")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@userStyleId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UserStyleId)
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@StyleSkillName", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.StyleSkillName)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERSTYLE_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateUSERSTYLESet(ByVal objClass As UserStyleSet) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateUSERSTYLE")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@userStyleId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UserStyleId)
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@StyleSkillName", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.StyleSkillName)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERSTYLE_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetUserStyleSetByPersonID(ByVal PersonID As Integer) As UserStyleSetCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 0, ParameterDirection.Input, PersonID)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERSTYLE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERSTYLECollectionFromReader)
            Dim iCollection As UserStyleSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserStyleSetCollection)

            Return iCollection
        End Function


        Public Overrides Function GetAllUserStyleSets() As UserStyleSetCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERSTYLE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUserDisciplineSetCollectionFromReader)
            Dim iCollection As UserStyleSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserStyleSetCollection)

            Return iCollection

        End Function

        Public Overrides Function DeleteUserStyleSetById(ByVal PersonID As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 0, ParameterDirection.Input, PersonID)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERSTYLE_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

#End Region '***********USERTYLESETS Methods ***********


#Region "************STYLESETLOOKUPVALUES***************"

        Public Overrides Function getStyleTypeByLookupId(ByVal userStyleId As Integer) As STYLETYPEPLOOKUP

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)



            AddParamToSQLCmd(sqlCmd, "@StyleSkillId", SqlDbType.Int, 10, ParameterDirection.Input, userStyleId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STYLETYPELOOKUP_SELECT)

            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSTYLETYPELOOKUPCollectionFromReader)
            Dim iCollection As STYLETYPELOOKUPCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), STYLETYPELOOKUPCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetAllStyleTypes() As STYLETYPELOOKUPCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_STYLETYPELOOKUP_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateSTYLETYPELOOKUPCollectionFromReader)
            Dim iCollection As STYLETYPELOOKUPCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), STYLETYPELOOKUPCollection)
            Return iCollection

        End Function

#End Region



#Region "************DisciplineTypeLookupVALUES***************"

        Public Overrides Function getSkillByLookupId(ByVal UserSkillId As Integer) As DisciplineTypeLookup

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)



            AddParamToSQLCmd(sqlCmd, "@StyleSkillId", SqlDbType.Int, 10, ParameterDirection.Input, UserSkillId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LK_USERDISCIPLINE_SELECT)

            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateDisciplineTypeLookupFromReader)
            Dim iCollection As DisciplineTypeLookupCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), DisciplineTypeLookupCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetAllCopyEditTypes() As COPYEDITTYPESLOOKUPCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLOOKUP_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITTYPELOOKUPCollectionFromReader)
            Dim iCollection As COPYEDITTYPESLOOKUPCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITTYPESLOOKUPCollection)
            Return iCollection

        End Function

#End Region '***********DisciplineTypeLookup Methods ***********

#Region "*********** USERCOPYEDITSKILLSET Methods ***********"

        '***************************************************
        '**
        '** USERCOPYEDITSKILLSET Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewUSERCOPYEDITSKILLSET(ByVal objClass As UserCopyEditLevelSet) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewUSERCOPYEDITSKILLSET")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@userCopyEditSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.userCopyEditSkillId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@CopyEditSkillName", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.CopyEditSkillName)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateUSERCOPYEDITSKILLSET(ByVal objClass As UserCopyEditLevelSet) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateUSERCOPYEDITSKILLSET")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@userCopyEditSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.userCopyEditSkillId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@CopyEditSkillName", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.CopyEditSkillName)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteUSERCOPYEDITSKILLSETById(ByVal PersonId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetUSERCOPYEDITSKILLSETByID(ByVal userCopyEditSkillId As Integer) As UserCopyEditLevelSet
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@userCopyEditSkillId", SqlDbType.Int, 0, ParameterDirection.Input, userCopyEditSkillId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERCOPYEDITSKILLSETCollectionFromReader)
            Dim iCollection As UserCopyEditLevelSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserCopyEditLevelSetCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function getUserCopyEditSkillSetbyPersonId(ByVal PersonId As Integer) As UserCopyEditLevelSetCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERCOPYEDITSKILLSETCollectionFromReader)
            Dim iCollection As UserCopyEditLevelSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserCopyEditLevelSetCollection)

            Return iCollection
        End Function

        Public Overrides Function GetAllUserCopyEditSets() As UserCopyEditLevelSetCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCOPYEDITSKILLSET_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERCOPYEDITSKILLSETCollectionFromReader)
            Dim iCollection As UserCopyEditLevelSetCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserCopyEditLevelSetCollection)

            Return iCollection

        End Function
#End Region '***********USERCOPYEDITSKILLSET Methods ***********



#Region "************LANGUAGELOOKUPVALUES***************"

       
        Public Overrides Function GetAllLanguages() As LanguageLookUpCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LANGUAGELOOKUP_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateLanguageLOOKUPCollectionFromReader)
            Dim iCollection As LanguageLookUpCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), LanguageLookUpCollection)
            Return iCollection

        End Function

#End Region

#Region "*********** USERLanguage Methods ***********"

        '***************************************************
        '**
        '** USERLanguage Methods
        '**
        '**
        '***************************************************


        Public Overrides Function AddUserLanguage(ByVal objClass As UserLanguageSet) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UserLanguageSet")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@LanguageId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.LanguageId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERLANGUAGE_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function DeleteUSERLanguageByID(ByVal PersonId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERLANGUAGE_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function getUserLanguagesbyPersonId(ByVal PersonId As Integer) As UserLanguageCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERLANGUAGE_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERLanguageCollectionFromReader)
            Dim iCollection As UserLanguageCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserLanguageCollection)

            Return iCollection
        End Function
        
#End Region '***********USERCOPYEDITSKILLSET Methods ***********


#Region "*********** BID Methods ***********"

        '***************************************************
        '**
        '** BID Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewBID(ByVal objClass As BID) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewBID")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidStatusID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidStatusID)
            AddParamToSQLCmd(sqlCmd, "@PersonIDAwardedBid", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonIDAwardedBid)
            AddParamToSQLCmd(sqlCmd, "@ApexManagerID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ApexManagerID)
            AddParamToSQLCmd(sqlCmd, "@RequiredUserSkillID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredUserSkillID)
            AddParamToSQLCmd(sqlCmd, "@RequiredCopyEditSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredCopyEditSkillId)
            AddParamToSQLCmd(sqlCmd, "@RequiredStyleSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredStyleSkillId)
            AddParamToSQLCmd(sqlCmd, "@RespondByDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.RespondByDate)
            AddParamToSQLCmd(sqlCmd, "@WorkStartDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.WorkStartDate)
            AddParamToSQLCmd(sqlCmd, "@WorkEndDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.WorkEndDate)
            AddParamToSQLCmd(sqlCmd, "@BidTitle", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.BidTitle)
            AddParamToSQLCmd(sqlCmd, "@ConvertedTaskOrderID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ConvertedTaskOrderID)

            AddParamToSQLCmd(sqlCmd, "@ProjectExtent", SqlDbType.NVarChar, 500, ParameterDirection.Input, objClass.BidProjectExtent)
            AddParamToSQLCmd(sqlCmd, "@Notes", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.BidNotes)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)
            'AddParamToSQLCmd(sqlCmd, "@PageCount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PageCount)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateBID(ByVal objClass As BID) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateBID")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidId)
            AddParamToSQLCmd(sqlCmd, "@BidStatusID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidStatusID)
            AddParamToSQLCmd(sqlCmd, "@PersonIDAwardedBid", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonIDAwardedBid)
            AddParamToSQLCmd(sqlCmd, "@ApexManagerID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ApexManagerID)
            AddParamToSQLCmd(sqlCmd, "@RequiredUserSkillID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredUserSkillID)
            AddParamToSQLCmd(sqlCmd, "@RequiredCopyEditSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredCopyEditSkillId)
            AddParamToSQLCmd(sqlCmd, "@RequiredStyleSkillId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.RequiredStyleSkillId)
            AddParamToSQLCmd(sqlCmd, "@RespondByDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.RespondByDate)
            AddParamToSQLCmd(sqlCmd, "@WorkStartDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.WorkStartDate)
            AddParamToSQLCmd(sqlCmd, "@WorkEndDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.WorkEndDate)
            AddParamToSQLCmd(sqlCmd, "@BidTitle", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.BidTitle)
            AddParamToSQLCmd(sqlCmd, "@ConvertedTaskOrderID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.ConvertedTaskOrderID)

            AddParamToSQLCmd(sqlCmd, "@ProjectExtent", SqlDbType.NVarChar, 500, ParameterDirection.Input, objClass.BidProjectExtent)
            AddParamToSQLCmd(sqlCmd, "@Notes", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.BidNotes)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.countryId)
            'AddParamToSQLCmd(sqlCmd, "@PageCount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PageCount)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteBIDById(ByVal BidId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 0, ParameterDirection.Input, BidId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetBIDByID(ByVal BidId As Integer) As BID
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 0, ParameterDirection.Input, BidId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDCollectionFromReader)
            Dim iCollection As BIDCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetBIDByPersonIDAwardedBid(ByVal PersonIDAwardedBid As Integer) As BIDCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonIDAwardedBid", SqlDbType.Int, 0, ParameterDirection.Input, PersonIDAwardedBid)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDCollectionFromReader)
            Dim iCollection As BIDCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDCollection)

            Return iCollection
        End Function

        Public Overrides Function GetALLBIDSBySearch(ByVal BidID As Integer, ByVal BidStatusID As Integer, ByVal CountryId As Integer) As BIDCollection
            Dim sqlCmd As New SqlCommand

            If BidID > 0 Then
                AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 4, ParameterDirection.Input, BidID)
            End If

            If BidStatusID > 0 Then
                AddParamToSQLCmd(sqlCmd, "@BidStatusID", SqlDbType.Int, 4, ParameterDirection.Input, BidStatusID)
            End If
            If CountryId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            End If


            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDCollectionFromReader)
            Dim iCollection As BIDCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDCollection)

            Return iCollection

        End Function
        Public Overrides Function GetAllBids(ByVal Context As String, ByVal CountryId As Integer) As BIDCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@Context", SqlDbType.VarChar, 20, ParameterDirection.Input, Context)
            AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDCollectionFromReader)
            Dim iCollection As BIDCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDCollection)

            Return iCollection
        End Function

        Public Overrides Function GetALLBIDSByStatus(ByVal BidStatusID As Integer, ByVal Context As String, ByVal CountryId As Integer) As BIDCollection
            Dim sqlCmd As New SqlCommand

            If BidStatusID > 0 Then
                AddParamToSQLCmd(sqlCmd, "@BidStatusID", SqlDbType.Int, 4, ParameterDirection.Input, BidStatusID)
            End If
            If CountryId > 0 Then
                AddParamToSQLCmd(sqlCmd, "@CountryId", SqlDbType.Int, 4, ParameterDirection.Input, CountryId)
            End If
            AddParamToSQLCmd(sqlCmd, "@Context", SqlDbType.VarChar, 20, ParameterDirection.Input, Context)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BID_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDCollectionFromReader)
            Dim iCollection As BIDCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDCollection)

            Return iCollection

        End Function
#End Region '***********BID Methods ***********


#Region "*********** BID_STATUS Methods ***********"

        '***************************************************
        '**
        '** BIDSTATUS Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewBIDSTATUS(ByVal objClass As BIDSTATUS) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewBIDSTATUS")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@bidstatusname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bidstatusname)
            AddParamToSQLCmd(sqlCmd, "@bidstatusaction", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bidstatusaction)
            AddParamToSQLCmd(sqlCmd, "@bidiconfile", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.bidiconfile)
            AddParamToSQLCmd(sqlCmd, "@isClosedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isClosedTF)
            AddParamToSQLCmd(sqlCmd, "@bidstatusdisplayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.bidstatusdisplayorder)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isenabledTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDSTATUS_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateBIDSTATUS(ByVal objClass As BIDSTATUS) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateBIDSTATUS")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@bid_status_code", SqlDbType.Int, 4, ParameterDirection.Input, objClass.bid_status_code)
            AddParamToSQLCmd(sqlCmd, "@bidstatusname", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bidstatusname)
            AddParamToSQLCmd(sqlCmd, "@bidstatusaction", SqlDbType.VarChar, 50, ParameterDirection.Input, objClass.bidstatusaction)
            AddParamToSQLCmd(sqlCmd, "@bidiconfile", SqlDbType.VarChar, 300, ParameterDirection.Input, objClass.bidiconfile)
            AddParamToSQLCmd(sqlCmd, "@isClosedTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isClosedTF)
            AddParamToSQLCmd(sqlCmd, "@bidstatusdisplayorder", SqlDbType.Int, 4, ParameterDirection.Input, objClass.bidstatusdisplayorder)
            AddParamToSQLCmd(sqlCmd, "@isenabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isenabledTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDSTATUS_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteBIDSTATUSById(ByVal bid_status_code As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@bid_status_code", SqlDbType.Int, 0, ParameterDirection.Input, bid_status_code)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDSTATUS_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetBIDSTATUSByID(ByVal bid_status_code As Integer) As BIDSTATUS
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@bid_status_code", SqlDbType.Int, 0, ParameterDirection.Input, bid_status_code)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDSTATUS_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDSTATUSCollectionFromReader)
            Dim iCollection As BIDSTATUSCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDSTATUSCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function



        Public Overrides Function GetBIDEVENTByBIDId(ByVal BidId As Integer) As BIDEVENTCollection
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 4, ParameterDirection.Input, BidId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDEVENTCollectionFromReader)
            Dim iCollection As BIDEVENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDEVENTCollection)

            Return iCollection

        End Function

        Public Overrides Function GetALLBIDSTATUSES() As BIDSTATUSCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDSTATUS_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDSTATUSCollectionFromReader)
            Dim iCollection As BIDSTATUSCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDSTATUSCollection)

            Return iCollection

        End Function
#End Region '***********BIDSTATUS Methods ***********


#Region "*********** BIDEVENT Methods ***********"

        '***************************************************
        '**
        '** BIDEVENT Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewBIDEVENT(ByVal objClass As BIDEVENT) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewBIDEVENT")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidId)
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonID)
            AddParamToSQLCmd(sqlCmd, "@PersonAcceptTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonAcceptTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateBIDEVENT(ByVal objClass As BIDEVENT) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateBIDEVENT")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidEventId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidEventId)
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.BidId)
            AddParamToSQLCmd(sqlCmd, "@PersonID", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonID)
            AddParamToSQLCmd(sqlCmd, "@PersonAcceptTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonAcceptTF)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function DeleteBIDEVENTById(ByVal BidId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 0, ParameterDirection.Input, BidId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetBIDEVENTByID(ByVal BidEventId As Integer) As BIDEVENT
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 0, ParameterDirection.Input, BidEventId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDEVENTCollectionFromReader)
            Dim iCollection As BIDEVENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDEVENTCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetBidEventByPersonId(ByVal PersonId As Integer) As BIDEVENTCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDEVENTCollectionFromReader)
            Dim iCollection As BIDEVENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDEVENTCollection)
            Return iCollection
        End Function


        Public Overrides Function GetAllBidEvents() As BIDEVENTCollection
            Dim sqlCmd As New SqlCommand
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDEVENTCollectionFromReader)
            Dim iCollection As BIDEVENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDEVENTCollection)

            Return iCollection
        End Function


        Public Overrides Function GetBIDEVENTByAcceptance(ByVal BidId As Integer, ByVal PersonAcceptance As Integer) As BIDEVENTCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@BidId", SqlDbType.Int, 0, ParameterDirection.Input, BidId)
            AddParamToSQLCmd(sqlCmd, "@PersonAcceptTF", SqlDbType.Int, 0, ParameterDirection.Input, PersonAcceptance)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_BIDEVENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateBIDEVENTCollectionFromReader)
            Dim iCollection As BIDEVENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), BIDEVENTCollection)
            If iCollection.Count > 0 Then
                Return iCollection
            Else
                Return Nothing
            End If
        End Function



#End Region '***********BIDEVENT Methods ***********


#Region "*********** JOB Methods ***********"

        '***************************************************
        '**
        '** JOB Methods
        '**
        '**
        '***************************************************


        Public Overrides Function CreateNewJOB(ByVal objClass As JOB) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewJOB")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.int, 4, ParameterDirection.Input, objClass.job_id)
            AddParamToSQLCmd(sqlCmd, "@job_number", SqlDbType.Int, 4, ParameterDirection.Input, objClass.job_number)
            AddParamToSQLCmd(sqlCmd, "@job_description", SqlDbType.varchar, 300, ParameterDirection.Input, objClass.job_description)
            AddParamToSQLCmd(sqlCmd, "@isEnabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isEnabledTF)
            AddParamToSQLCmd(sqlCmd, "@isInterimProductRequiredTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isInterimProductRequiredTF)
            AddParamToSQLCmd(sqlCmd, "@interimCompensationPercent", SqlDbType.Int, 4, ParameterDirection.Input, objClass.interimCompensationPercent)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateJOB(ByVal objClass As JOB) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateJOB")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.int, 4, ParameterDirection.Input, objClass.job_id)
            AddParamToSQLCmd(sqlCmd, "@job_number", SqlDbType.Int, 4, ParameterDirection.Input, objClass.job_number)
            AddParamToSQLCmd(sqlCmd, "@job_description", SqlDbType.varchar, 300, ParameterDirection.Input, objClass.job_description)
            AddParamToSQLCmd(sqlCmd, "@isEnabledTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isEnabledTF)
            AddParamToSQLCmd(sqlCmd, "@isInterimProductRequiredTF", SqlDbType.Int, 4, ParameterDirection.Input, objClass.isInterimProductRequiredTF)
            AddParamToSQLCmd(sqlCmd, "@interimCompensationPercent", SqlDbType.Int, 4, ParameterDirection.Input, objClass.interimCompensationPercent)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function GetJOBByID(ByVal job_id As Integer) As JOB
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.Int, 0, ParameterDirection.Input, job_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateJOBCollectionFromReader)
            Dim iCollection As JOBCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), JOBCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetJOBByJobNumber(ByVal job_number As Integer) As JOB
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@job_number", SqlDbType.Int, 0, ParameterDirection.Input, job_number)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateJOBCollectionFromReader)
            Dim iCollection As JOBCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), JOBCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function


        Public Overrides Function DeleteJOBByID(ByVal job_id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.Int, 0, ParameterDirection.Input, job_id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function GetAllJobs() As JOBCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_JOB_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateJOBCollectionFromReader)
            Dim iCollection As JOBCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), JOBCollection)

            Return iCollection
        End Function

#End Region '***********JOB Methods ***********


#Region "*********** LIBRARY Attachment Methods ***********"

        '*********************************************************************
        ' Attachment Methods
        '
        ' The following methods are used for working with Attachments.
        '
        '*********************************************************************

        Public Overrides Function InsertLibraryAttachment(ByVal oLAttachment As LIBRARYATTACHMENT) As Integer
            ' Validate Parameters
            If oLAttachment Is Nothing Then
                Throw New ArgumentNullException("oAttachment is Nothing")
            End If

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@unit_id", SqlDbType.Int, 0, ParameterDirection.Input, oLAttachment.unit_id)
            AddParamToSQLCmd(sqlCmd, "@FileName", SqlDbType.VarChar, 255, ParameterDirection.Input, oLAttachment.filename)
            AddParamToSQLCmd(sqlCmd, "@FileSize", SqlDbType.BigInt, 0, ParameterDirection.Input, oLAttachment.filesize)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NText, 0, ParameterDirection.Input, oLAttachment.contenttype)
            sqlCmd.Parameters.Add("@Attachment", SqlDbType.Image, oLAttachment.LibraryAttachment.Length)
            sqlCmd.Parameters("@Attachment").Value = oLAttachment.LibraryAttachment
            AddParamToSQLCmd(sqlCmd, "@objecttype", SqlDbType.VarChar, 20, ParameterDirection.Input, oLAttachment.objecttype)
            AddParamToSQLCmd(sqlCmd, "@attachmentnotes", SqlDbType.VarChar, 500, ParameterDirection.Input, oLAttachment.LibraryAttachmentnotes)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LIBRARYATTACHMENT_CREATE)
            ExecuteScalarCmd(sqlCmd)

            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)

        End Function 'InsertAttachment


        Public Overrides Function UpdateLibraryAttachment(ByVal oLAttachment As LIBRARYATTACHMENT) As Boolean
            ' Validate Parameters
            If oLAttachment Is Nothing Then
                Throw New ArgumentNullException("oAttachment is Nothing")
            End If
            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand



            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@unit_id", SqlDbType.Int, 0, ParameterDirection.Input, oLAttachment.unit_id)
            AddParamToSQLCmd(sqlCmd, "@FileName", SqlDbType.VarChar, 255, ParameterDirection.Input, oLAttachment.filename)
            AddParamToSQLCmd(sqlCmd, "@FileSize", SqlDbType.BigInt, 0, ParameterDirection.Input, oLAttachment.filesize)
            AddParamToSQLCmd(sqlCmd, "@ContentType", SqlDbType.NText, 0, ParameterDirection.Input, oLAttachment.contenttype)
            sqlCmd.Parameters.Add("@Attachment", SqlDbType.Image, oLAttachment.LibraryAttachment.Length)
            sqlCmd.Parameters("@Attachment").Value = oLAttachment.LibraryAttachment
            AddParamToSQLCmd(sqlCmd, "@objecttype", SqlDbType.VarChar, 20, ParameterDirection.Input, oLAttachment.objecttype)
            AddParamToSQLCmd(sqlCmd, "@attachmentnotes", SqlDbType.VarChar, 500, ParameterDirection.Input, oLAttachment.LibraryAttachmentnotes)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LIBRARYATTACHMENT_UPDATE)
            ExecuteScalarCmd(sqlCmd)

            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)

        End Function 'UpdateAttachment



        Public Overrides Function DeleteLibraryAttachment(ByVal LibraryAttachment_ID As Integer) As Boolean
            ' Validate Parameters
            If LibraryAttachment_ID <= 0 Then
                Throw New ArgumentOutOfRangeException("LibAttachmentID")
            End If

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@LibraryAttachment_id", SqlDbType.Int, 0, ParameterDirection.Input, LibraryAttachment_ID)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LIBRARYATTACHMENT_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Dim returnValue As Integer = CInt(sqlCmd.Parameters("@ReturnValue").Value)
            Return IIf(returnValue = 0, True, False)

        End Function 'DeleteAttachment



        Public Overrides Function GetLibraryAttachmentByID(ByVal LibraryAttachment_ID As Integer) As LIBRARYATTACHMENT

            Dim oLAttachment As LIBRARYATTACHMENT = Nothing

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@LibraryAttachment_ID", SqlDbType.Int, 0, ParameterDirection.Input, LibraryAttachment_ID)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LIBRARYATTACHMENT_SELECT)

            ' Execute Reader
            If ConnectionString = String.Empty Then
                Throw New ArgumentOutOfRangeException("ConnectionString")
            End If

            Dim cn As New SqlConnection(Me.ConnectionString)
            Try
                sqlCmd.Connection = cn
                cn.Open()
                Dim dtr As SqlDataReader = sqlCmd.ExecuteReader()
                If dtr.Read() Then
                    oLAttachment = New LIBRARYATTACHMENT(CInt(dtr("LibraryAttachment_ID")), CInt(dtr("unit_id")), CStr(dtr("FileName")), CInt(dtr("FileSize")), CStr(dtr("ContentType")), CType(dtr("Attachment"), Byte()), CStr(dtr("ObjectType")), CStr(dtr("attachmentnotes")))
                End If
            Catch ex As Exception
                'Do Something with this error
            Finally
                cn.Close()
            End Try


            Return oLAttachment

        End Function 'GetAttachmentByID



        Public Overrides Function GetLibraryAttachmentbyObject(ByVal ObjectType As String, ByVal unit_id As Integer) As LIBRARYATTACHMENTCollection

            ' Execute SQL Command
            Dim sqlCmd As New SqlCommand

            AddParamToSQLCmd(sqlCmd, "@ObjectType", SqlDbType.VarChar, 20, ParameterDirection.Input, ObjectType)
            AddParamToSQLCmd(sqlCmd, "@unit_id", SqlDbType.Int, 0, ParameterDirection.Input, unit_id)


            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_LIBRARYATTACHMENT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateLIBRARYATTACHMENTCollectionFromReader)
            Dim iCollection As LIBRARYATTACHMENTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), LIBRARYATTACHMENTCollection)
            Return iCollection

        End Function

#End Region '*********** Attachment Methods ***********


        Public Overrides Function GetCopyEditLevelbyLookupId(ByVal copyeditlevel_id As Integer, ByVal currency_id As Integer) As COPYEDITLEVELLOOKUP_V3
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel_id", SqlDbType.Int, 0, ParameterDirection.Input, copyeditlevel_id)
            AddParamToSQLCmd(sqlCmd, "@currencyId", SqlDbType.Int, 0, ParameterDirection.Input, currency_id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLIST_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITLEVELLookupCollectionFromReader)
            Dim iCollection As COPYEDITLEVELLOOKUP_V3Collection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITLEVELLOOKUP_V3Collection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetCopyEditLevelByName(ByVal copyeditlevelname As String, ByVal jobId As Integer) As COPYEDITLEVELLOOKUP_V3
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@copyeditlevelname", SqlDbType.NVarChar, 1000, ParameterDirection.Input, copyeditlevelname)
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.Int, 4, ParameterDirection.Input, jobId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLIST_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITLEVELLookupCollectionFromReader)
            Dim iCollection As COPYEDITLEVELLOOKUP_V3Collection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITLEVELLOOKUP_V3Collection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function


        Public Overrides Function GetAllCopyEditLevels(ByVal job_id As Integer, Optional ByVal CurrencyId As Integer = 0) As COPYEDITLEVELLOOKUP_V3Collection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@job_id", SqlDbType.Int, 0, ParameterDirection.Input, job_id)
            If CurrencyId <> 0 Then
                AddParamToSQLCmd(sqlCmd, "@currencyId", SqlDbType.Int, 0, ParameterDirection.Input, CurrencyId)
            End If
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLIST_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCOPYEDITLEVELLookupCollectionFromReader)
            Dim iCollection As COPYEDITLEVELLOOKUP_V3Collection = CType(ExecuteReaderCmd(sqlCmd, sqlData), COPYEDITLEVELLOOKUP_V3Collection)
            If iCollection.Count > 0 Then
                Return iCollection
            Else
                Return Nothing
            End If
        End Function


        Public Overrides Function DeleteCopyEditLevelLookupByID(ByVal copyeditlevel_id As Integer, ByVal CurrencyId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@copyeditlevel_id", SqlDbType.Int, 4, ParameterDirection.Input, copyeditlevel_id)
            AddParamToSQLCmd(sqlCmd, "@currencyId", SqlDbType.Int, 4, ParameterDirection.Input, CurrencyId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLOOKUP_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function CreateNewCopyEditLevelLookUp(ByVal objClass As COPYEDITLEVELLOOKUP_V3) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewCopyEditLevel")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@Name", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.CopyEditLevelName)
            AddParamToSQLCmd(sqlCmd, "@Description", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.CopyEditLevelDescription)
            AddParamToSQLCmd(sqlCmd, "@Job_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.JobId)
            AddParamToSQLCmd(sqlCmd, "@UnitType", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.UnitType)
            AddParamToSQLCmd(sqlCmd, "@UnitPrice", SqlDbType.Money, 8, ParameterDirection.Input, objClass.UnitPrice)
            AddParamToSQLCmd(sqlCmd, "@IsEnabled", SqlDbType.Char, 1, ParameterDirection.Input, objClass.IsEnabled)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyId)
            AddParamToSQLCmd(sqlCmd, "@JobType", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.AccountingJobType)
            AddParamToSQLCmd(sqlCmd, "@DefaultUnitCount", SqlDbType.Int, 4, ParameterDirection.Input, IIf(objClass.DefaultUnitCount.HasValue, objClass.DefaultUnitCount.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@IsUnitCountEditable", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.IsUnitCountEditable)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLOOKUP_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateCopyEditLevelLookup(ByVal objClass As COPYEDITLEVELLOOKUP_V3) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateCopyEditLevel")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@CopyEditLevel_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CopyEditLevelId)
            AddParamToSQLCmd(sqlCmd, "@Name", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.CopyEditLevelName)
            AddParamToSQLCmd(sqlCmd, "@Description", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.CopyEditLevelDescription)
            AddParamToSQLCmd(sqlCmd, "@Job_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.JobId)
            AddParamToSQLCmd(sqlCmd, "@UnitType", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.UnitType)
            AddParamToSQLCmd(sqlCmd, "@UnitPrice", SqlDbType.Money, 8, ParameterDirection.Input, objClass.UnitPrice)
            AddParamToSQLCmd(sqlCmd, "@Display_order", SqlDbType.Int, 4, ParameterDirection.Input, objClass.DisplayOrder)
            AddParamToSQLCmd(sqlCmd, "@IsEnabled", SqlDbType.Char, 1, ParameterDirection.Input, objClass.IsEnabled)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyId)
            AddParamToSQLCmd(sqlCmd, "@JobType", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.AccountingJobType)
            AddParamToSQLCmd(sqlCmd, "@DefaultUnitCount", SqlDbType.Int, 4, ParameterDirection.Input, IIf(objClass.DefaultUnitCount.HasValue, objClass.DefaultUnitCount.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@IsUnitCountEditable", SqlDbType.NVarChar, 50, ParameterDirection.Input, objClass.IsUnitCountEditable)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLOOKUP_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function MoveCopyEditLevelLookupItem(ByVal CurId As Integer, DestId As Integer, Position As String) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@CurId", SqlDbType.Int, 4, ParameterDirection.Input, CurId)
            AddParamToSQLCmd(sqlCmd, "@DestId", SqlDbType.Int, 4, ParameterDirection.Input, DestId)
            AddParamToSQLCmd(sqlCmd, "@Position", SqlDbType.VarChar, 20, ParameterDirection.Input, Position)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_COPYEDITLEVELLOOKUP_MOVE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function
        Public Overrides Function GetCurrencyDetails(Optional ByVal Currency As String = "") As CurrencyCollection
            Dim sqlCmd As New SqlCommand
            If (Currency <> "") Then
                AddParamToSQLCmd(sqlCmd, "@Currency", SqlDbType.NVarChar, 50, ParameterDirection.Input, Currency)
            End If
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_CURRENCY_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateCurrencyCollectionFromReader)
            Dim iCollection As CurrencyCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), CurrencyCollection)
            If iCollection.Count > 0 Then
                Return iCollection
            Else
                Return Nothing
            End If
        End Function


        Public Overrides Function GetAccountByID(ByVal AccountId As Integer) As ACCOUNTS
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@AccountId", SqlDbType.Int, 4, ParameterDirection.Input, AccountId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ACCOUNT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateACCOUNTCollectionFromReader)
            Dim iCollection As ACCOUNTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), ACCOUNTCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetALLAccounts() As ACCOUNTCollection
            Dim sqlCmd As New SqlCommand

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_ACCOUNT_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateACCOUNTCollectionFromReader)
            Dim iCollection As ACCOUNTCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), ACCOUNTCollection)
            Return iCollection

        End Function

        'Taskorder Payment Log
        Public Overrides Function CreateNewPaymentLog(ByVal objClass As TaskOrderPaymentLog) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewPaymentLog")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@TaskOrder_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.TaskOrder_Id)
            AddParamToSQLCmd(sqlCmd, "@InterimAmount", SqlDbType.Money, 8, ParameterDirection.Input, objClass.InterimAmount)
            AddParamToSQLCmd(sqlCmd, "@InterimPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.InterimPaymentDT)
            AddParamToSQLCmd(sqlCmd, "@FinalPaymentAmount", SqlDbType.Money, 8, ParameterDirection.Input, objClass.FinalPaymentAmount)
            AddParamToSQLCmd(sqlCmd, "@FinalPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.FinalPaymentDT)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TaskOrderPaymentLog_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdatePaymentLog(ByVal objClass As TaskOrderPaymentLog) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdatePaymentLog")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PaymentLog_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PaymentLog_Id)
            AddParamToSQLCmd(sqlCmd, "@TaskOrder_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.TaskOrder_Id)
            AddParamToSQLCmd(sqlCmd, "@InterimAmount", SqlDbType.Money, 8, ParameterDirection.Input, objClass.InterimAmount)
            AddParamToSQLCmd(sqlCmd, "@InterimPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.InterimPaymentDT)
            AddParamToSQLCmd(sqlCmd, "@FinalPaymentAmount", SqlDbType.Money, 8, ParameterDirection.Input, objClass.FinalPaymentAmount)
            AddParamToSQLCmd(sqlCmd, "@FinalPaymentDT", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.FinalPaymentDT)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TaskOrderPaymentLog_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function GetPaymentLogById(ByVal PaymentLog_Id As Integer) As TaskOrderPaymentLog
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PaymentLog_Id", SqlDbType.Int, 4, ParameterDirection.Input, PaymentLog_Id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TaskOrderPaymentLog_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePaymentLogCollectionFromReader)
            Dim iCollection As TaskorderPaymentLogCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TaskorderPaymentLogCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetPaymentLogByTaskOrder(ByVal TaskOrder_Id As Integer) As TaskOrderPaymentLog
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@TaskOrder_Id", SqlDbType.Int, 4, ParameterDirection.Input, TaskOrder_Id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TaskOrderPaymentLog_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GeneratePaymentLogCollectionFromReader)
            Dim iCollection As TaskorderPaymentLogCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), TaskorderPaymentLogCollection)
            If iCollection.Count > 0 Then
                Return iCollection(0)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function GetNewSubBatchNumber(ByVal JobNumber As String, ByVal BatchNumber As String) As Integer
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@NewNumber", SqlDbType.Int, 4, ParameterDirection.Output, Nothing)
            AddParamToSQLCmd(sqlCmd, "@JobNumber", SqlDbType.NVarChar, 100, ParameterDirection.Input, JobNumber)
            AddParamToSQLCmd(sqlCmd, "@BatchNumber", SqlDbType.NVarChar, 100, ParameterDirection.Input, BatchNumber)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, "sps_GetNextNumber")
            ExecuteScalarCmd(sqlCmd)
            Dim value As Integer
            If (sqlCmd.Parameters("@NewNumber").Value Is Nothing) Then
                value = 0
            Else
                value = CInt(sqlCmd.Parameters("@NewNumber").Value)
            End If
            Return value
        End Function


#Region "Timesheet"
  
        Public Overrides Function CreateTimesheetLog(ByVal objClass As TimesheetLog) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("NewTimesheetLog")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@Person_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.Person_Id)
            AddParamToSQLCmd(sqlCmd, "@Job_Number", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.Job_Number)
            AddParamToSQLCmd(sqlCmd, "@EntryDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.Entry_Date)
            AddParamToSQLCmd(sqlCmd, "@EditorialService", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.Editorial_Service)
            AddParamToSQLCmd(sqlCmd, "@Notes", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.Notes)
            AddParamToSQLCmd(sqlCmd, "@Hours", SqlDbType.Decimal, 10, ParameterDirection.Input, objClass.HoursInt)
            AddParamToSQLCmd(sqlCmd, "@CreatedOn", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.CreatedOn)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetLog_Insert)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function UpdateTimesheetLog(ByVal objClass As TimesheetLog) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateTimesheetLog")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@LogId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.Log_Id)
            AddParamToSQLCmd(sqlCmd, "@Person_Id", SqlDbType.Int, 4, ParameterDirection.Input, objClass.Person_Id)
            AddParamToSQLCmd(sqlCmd, "@Job_Number", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.Job_Number)
            AddParamToSQLCmd(sqlCmd, "@EntryDate", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.Entry_Date)
            AddParamToSQLCmd(sqlCmd, "@EditorialService", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.Editorial_Service)
            AddParamToSQLCmd(sqlCmd, "@Notes", SqlDbType.NVarChar, 4000, ParameterDirection.Input, objClass.Notes)
            AddParamToSQLCmd(sqlCmd, "@Hours", SqlDbType.Decimal, 10, ParameterDirection.Input, objClass.HoursInt)
            AddParamToSQLCmd(sqlCmd, "@CreatedOn", SqlDbType.DateTime, 8, ParameterDirection.Input, objClass.CreatedOn)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetLog_Update)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function DeleteTimesheetLogById(ByVal LogId As Integer, ByVal Person_Id As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@LogId", SqlDbType.Int, 4, ParameterDirection.Input, LogId)
            AddParamToSQLCmd(sqlCmd, "@Person_Id", SqlDbType.Int, 4, ParameterDirection.Input, Person_Id)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetLog_Delete)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function getTimesheetLogByPersonId(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime, Optional GetPendingApproval As Boolean = False, Optional LogId As Integer = 0) As DataTable
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)
            AddParamToSQLCmd(sqlCmd, "@GetPendingApproval", SqlDbType.Int, 0, ParameterDirection.Input, GetPendingApproval)
            AddParamToSQLCmd(sqlCmd, "@LogId", SqlDbType.Int, 4, ParameterDirection.Input, LogId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetLog_Select)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_TimesheetLog_Select")
            cn.Close()
            cn.Dispose()
            Return ds.Tables(0)
        End Function

        Public Overrides Function getTimesheetLogByWeek(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime) As DataTable
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetLogWeek_Select)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_TimesheetLogWeek_Select")
            cn.Close()
            cn.Dispose()
            Return ds.Tables(0)
        End Function

        Public Overrides Function SubmitForApproval(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_Timesheet_Submit)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function CertifyTimesheet(ByVal PersonId As Integer, Approval As Integer, StartDate As DateTime, EndDate As DateTime, CertifiedBy As String) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@Approved", SqlDbType.Int, 4, ParameterDirection.Input, Approval)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)
            AddParamToSQLCmd(sqlCmd, "@CertifiedBy", SqlDbType.NVarChar, 250, ParameterDirection.Input, CertifiedBy)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_Timesheet_Certify)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        Public Overrides Function GetTimesheetReport(ByVal PersonId As Integer, RoleId As Integer, StartDate As DateTime, EndDate As DateTime) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, PersonId)
            AddParamToSQLCmd(sqlCmd, "@RoleId", SqlDbType.Int, 4, ParameterDirection.Input, RoleId)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetReport_Select)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_TimesheetReport_Select")
            cn.Close()
            cn.Dispose()
            Return ds
        End Function

        Public Overrides Function GetTimesheetManagerReport(ByVal CurrentRoleId As Integer, StartDate As DateTime, EndDate As DateTime) As DataSet
            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)
            AddParamToSQLCmd(sqlCmd, "@CurRoleId", SqlDbType.Int, 4, ParameterDirection.Input, CurrentRoleId)
            AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 8, ParameterDirection.Input, StartDate)
            AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 8, ParameterDirection.Input, EndDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_TimesheetManagerReport_Select)
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "SP_TimesheetManagerReport_Select")
            cn.Close()
            cn.Dispose()
            Return ds
        End Function
      
#End Region

        Public Overrides Function GetDirectPaymentsByEC(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal Person_Id As Integer) As DataSet

            Dim sqlCmd As New SqlCommand
            Dim ds As New DataSet
            Dim cn As New SqlConnection(Me.ConnectionString)

            If Not Util.IsNullDate(StartRangeDT) Then

                AddParamToSQLCmd(sqlCmd, "@SDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(StartRangeDT, "M/d/yyyy"))
                AddParamToSQLCmd(sqlCmd, "@EDate", SqlDbType.DateTime, 10, ParameterDirection.Input, Format(EndRangeDT, "M/d/yyyy"))
            End If
            AddParamToSQLCmd(sqlCmd, "@EC_Person_Id", SqlDbType.Int, 4, ParameterDirection.Input, Person_Id)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, "sps_GetECDirectPaymentList")
            sqlCmd.Connection = cn

            Dim MyCommand As New SqlDataAdapter(sqlCmd)
            MyCommand.Fill(ds, "DirectPayments")
            'MyCommand.Fill(ds)

            Return ds

        End Function

        Public Overrides Function InsertDirectPayment(ByVal objClass As DirectPaymentModel) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("InsertDirectPayment")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@JobId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.JobId)
            AddParamToSQLCmd(sqlCmd, "@BatchNumber", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.BatchNumber)
            AddParamToSQLCmd(sqlCmd, "@EditorialService", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.EditorialService)
            AddParamToSQLCmd(sqlCmd, "@JobType", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.JobType)
            AddParamToSQLCmd(sqlCmd, "@UnitType", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.UnitType)
            AddParamToSQLCmd(sqlCmd, "@IsLumpsum", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsLumpsum)
            AddParamToSQLCmd(sqlCmd, "@UnitCount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UnitCount)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyId)
            AddParamToSQLCmd(sqlCmd, "@PricePerPage", SqlDbType.Money, 8, ParameterDirection.Input, IIf(objClass.PricePerPage.HasValue, objClass.PricePerPage.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@TotalPrice", SqlDbType.Money, 8, ParameterDirection.Input, IIf(objClass.TotalPrice.HasValue, objClass.TotalPrice.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@LoggedOn", SqlDbType.DateTime, 4, ParameterDirection.Input, objClass.LoggedOn)
            AddParamToSQLCmd(sqlCmd, "@LoggedBy", SqlDbType.Int, 4, ParameterDirection.Input, objClass.LoggedBy)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_DIRECTPAYMENTS_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function UpdateDirectPayment(ByVal objClass As DirectPaymentModel) As Boolean
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UpdateDirectPayment")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PaymentLogId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PaymentLogId)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@JobId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.JobId)
            AddParamToSQLCmd(sqlCmd, "@BatchNumber", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.BatchNumber)
            AddParamToSQLCmd(sqlCmd, "@EditorialService", SqlDbType.NVarChar, 250, ParameterDirection.Input, objClass.EditorialService)
            AddParamToSQLCmd(sqlCmd, "@JobType", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.JobType)
            AddParamToSQLCmd(sqlCmd, "@UnitType", SqlDbType.NVarChar, 100, ParameterDirection.Input, objClass.UnitType)
            AddParamToSQLCmd(sqlCmd, "@IsLumpsum", SqlDbType.Int, 4, ParameterDirection.Input, objClass.IsLumpsum)
            AddParamToSQLCmd(sqlCmd, "@UnitCount", SqlDbType.Int, 4, ParameterDirection.Input, objClass.UnitCount)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyId)
            AddParamToSQLCmd(sqlCmd, "@PricePerPage", SqlDbType.Money, 8, ParameterDirection.Input, IIf(objClass.PricePerPage.HasValue, objClass.PricePerPage.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@TotalPrice", SqlDbType.Money, 8, ParameterDirection.Input, IIf(objClass.TotalPrice.HasValue, objClass.TotalPrice.Value, -1))
            AddParamToSQLCmd(sqlCmd, "@LoggedOn", SqlDbType.DateTime, 4, ParameterDirection.Input, objClass.LoggedOn)
            AddParamToSQLCmd(sqlCmd, "@LoggedBy", SqlDbType.Int, 4, ParameterDirection.Input, objClass.LoggedBy)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_DIRECTPAYMENTS_UPDATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function DeleteDirectPaymentByID(ByVal PaymentLogId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PaymentLogId", SqlDbType.Int, 4, ParameterDirection.Input, PaymentLogId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_DIRECTPAYMENTS_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function MarkDirectPaymentAsPaid(ByVal PaymentLogId As Integer, ByVal PaymentDate As DateTime) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PaymentLogId", SqlDbType.Int, 4, ParameterDirection.Input, PaymentLogId)
            AddParamToSQLCmd(sqlCmd, "@PaymentDate", SqlDbType.DateTime, 4, ParameterDirection.Input, PaymentDate)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_DIRECTPAYMENTS_PAID)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function


        '*********************************************************************
        '
        ' SQL Helper Methods
        '
        ' The following utility methods are used to interact with SQL Server.
        '
        '*********************************************************************

        Private Sub AddParamToSQLCmd(ByVal sqlCmd As SqlCommand, ByVal paramId As String, ByVal sqlType As SqlDbType, ByVal paramSize As Integer, ByVal paramDirection As ParameterDirection, ByVal paramvalue As Object)
            ' Validate Parameter Properties
            If sqlCmd Is Nothing Then
                Throw New ArgumentNullException("sqlCmd")
            End If
            If paramId = String.Empty Then
                Throw New ArgumentOutOfRangeException("paramId")
            End If
            ' Add Parameter
            Dim newSqlParam As New SqlParameter
            newSqlParam.ParameterName = paramId
            newSqlParam.SqlDbType = sqlType
            newSqlParam.Direction = paramDirection

            If paramSize > 0 Then
                newSqlParam.Size = paramSize
            End If
            If Not (paramvalue Is Nothing) Then
                newSqlParam.Value = paramvalue
            End If
            sqlCmd.Parameters.Add(newSqlParam)
        End Sub 'AddParamToSQLCmd



        Private Function ExecuteScalarCmd(ByVal sqlCmd As SqlCommand) As [Object]
            ' Validate Command Properties
            If ConnectionString = String.Empty Then
                Throw New ArgumentOutOfRangeException("ConnectionString")
            End If
            If sqlCmd Is Nothing Then
                Throw New ArgumentNullException("sqlCmd")
            End If
            Dim result As [Object] = Nothing

            Dim cn As New SqlConnection(Me.ConnectionString)
            Try
                sqlCmd.Connection = cn
                cn.Open()
                result = sqlCmd.ExecuteScalar()
            Finally
                cn.Dispose()
            End Try

            Return result
        End Function 'ExecuteScalarCmd



        Private Function ExecuteReaderCmd(ByVal sqlCmd As SqlCommand, ByVal gcfr As GenerateCollectionFromReader) As CollectionBase
            If ConnectionString = String.Empty Then
                Throw New ArgumentOutOfRangeException("ConnectionString")
            End If
            If sqlCmd Is Nothing Then
                Throw New ArgumentNullException("sqlCmd")
            End If
            Dim cn As New SqlConnection(Me.ConnectionString)
            Try
                sqlCmd.Connection = cn
                cn.Open()
                Dim temp As CollectionBase = gcfr(sqlCmd.ExecuteReader())
                Return temp
            Finally
                cn.Dispose()
            End Try
        End Function 'ExecuteReaderCmd



        Private Sub SetCommandType(ByVal sqlCmd As SqlCommand, ByVal cmdType As CommandType, ByVal cmdText As String)
            sqlCmd.CommandType = cmdType
            sqlCmd.CommandText = cmdText
        End Sub 'SetCommandType


#Region "*********** USER Currency Methods ***********"

        '***************************************************
        '**
        '** USER Currency Methods
        '**
        '**
        '***************************************************


        Public Overrides Function AddUserCurrency(ByVal objClass As UserCurrencySet) As Integer
            If objClass Is Nothing Then
                Throw New ArgumentNullException("UserCurrencySet")
            End If
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.PersonId)
            AddParamToSQLCmd(sqlCmd, "@CurrencyId", SqlDbType.Int, 4, ParameterDirection.Input, objClass.CurrencyId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCurrency_CREATE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function DeleteUSERCurrencyByID(ByVal PersonId As Integer) As Boolean
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@ReturnValue", SqlDbType.Int, 0, ParameterDirection.ReturnValue, Nothing)
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)

            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCurrency_DELETE)
            ExecuteScalarCmd(sqlCmd)
            Return CInt(sqlCmd.Parameters("@ReturnValue").Value)
        End Function

        Public Overrides Function getUserCurrenciesByPersonId(ByVal PersonId As Integer) As UserCurrencyCollection
            Dim sqlCmd As New SqlCommand
            AddParamToSQLCmd(sqlCmd, "@PersonId", SqlDbType.Int, 0, ParameterDirection.Input, PersonId)
            SetCommandType(sqlCmd, CommandType.StoredProcedure, SP_USERCurrency_SELECT)
            Dim sqlData As New GenerateCollectionFromReader(AddressOf GenerateUSERCurrencyCollectionFromReader)
            Dim iCollection As UserCurrencyCollection = CType(ExecuteReaderCmd(sqlCmd, sqlData), UserCurrencyCollection)

            Return iCollection
        End Function

#End Region


    End Class 'SQLDataAccessLayer
End Namespace 'DataAccessLayer



