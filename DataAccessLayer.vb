Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Configuration
Imports PrePress.BusinessLogicLayer



Namespace DataAccessLayer


    '*********************************************************************
    '
    ' DataAccessLayerBaseClass Class
    '
    ' The DataAccessLayerBaseClass lists all the abstract methods
    ' that each data access layer provider (SQL Server, Access, etc.)
    ' must implement.
    '
    '*********************************************************************

    Public MustInherit Class DataAccessLayerBaseClass

        '*** PRIVATE FIELDS ***

        Private _connectionString As String

        '*** PROPERTIES ***

        Public Property ConnectionString() As String
            Get
                Dim str As String = System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
                If str Is Nothing OrElse str.Length <= 0 Then
                    Throw New ApplicationException("ConnectionString configuration is missing from your web.config. It should contain  <appSettings><add key=""ConnectionString"" value=""database=IssueTrackerStarterKit;server=localhost;Trusted_Connection=true"" /></appSettings> ")
                Else
                    Return str
                End If
            End Get
            Set(ByVal Value As String)
                _connectionString = Value
            End Set
        End Property


        '*** ABSTRACT PROPERTIES ***




        '*** ABSTRACT METHODS ***


        ' PERSONPREF
        Public MustOverride Function CreateNewPERSONPREF(ByVal oPERSONPREF As PERSONPREF) As Integer
        Public MustOverride Function UpdatePERSONPREF(ByVal oPERSONPREF As PERSONPREF) As Boolean
        Public MustOverride Function DeletePERSONPREFByID(ByVal person_id As Integer, ByVal upkey As String) As Boolean
        Public MustOverride Function GetPERSONPREFByID(ByVal person_id As Integer, ByVal upkey As String) As PERSONPREF
        Public MustOverride Function GetALLPERSONPREFByID(ByVal personID As Integer, ByVal upkey As String) As PERSONPREFCollection

        ' SECURITYLEVEL
        Public MustOverride Function CreateNewSECURITYLEVEL(ByVal oSECURITYLEVEL As SECURITYLEVEL) As Integer
        Public MustOverride Function UpdateSECURITYLEVEL(ByVal oSECURITYLEVEL As SECURITYLEVEL) As Boolean
        Public MustOverride Function DeleteSECURITYLEVELByID(ByVal securitylevel_code As Integer) As Boolean
        Public MustOverride Function GetSECURITYLEVELByID(ByVal securitylevel_code As Integer) As SECURITYLEVEL
        Public MustOverride Function GetALLSECURITYLEVELS() As SECURITYLEVELCollection

        ' EVENTLOG
        Public MustOverride Function CreateNewGLOBALEVENTLOG(ByVal oEVENTLOG As GLOBALEVENTLOG) As Integer
        Public MustOverride Function UpdateGLOBALEVENTLOG(ByVal oEVENTLOG As GLOBALEVENTLOG) As Integer
        Public MustOverride Function DeleteGLOBALEVENTLOGById(ByVal object_code As String, ByVal object_id As Integer) As Boolean
        Public MustOverride Function GetGLOBALEVENTLOG(ByVal object_code As String, ByVal object_id As Integer) As GLOBALEVENTLOGCollection
        Public MustOverride Function GetALLGLOBALEVENTLOG() As GLOBALEVENTLOGCollection
        ' COMPANY
        Public MustOverride Function CreateNewCOMPANY(ByVal oCOMPANY As COMPANY) As Integer
        Public MustOverride Function UpdateCOMPANY(ByVal oCOMPANY As COMPANY) As Boolean
        Public MustOverride Function DeleteCOMPANYByID(ByVal company_id As Integer) As Boolean
        Public MustOverride Function GetCOMPANYByID(ByVal company_id As Integer) As COMPANY
        Public MustOverride Function GetALLCOMPANYSBySearch(ByVal IsInternal As Integer, ByVal IsEnabled As Integer, ByVal CountryId As Integer) As COMPANYCollection
        Public MustOverride Function GetALLCOMPANYSByCompanyName(ByVal companyname As String, ByVal CountryId As Integer) As COMPANYCollection
        Public MustOverride Function GetALLCOMPANYS(ByVal CountryId As Integer) As COMPANYCollection

        ' PERSON
        Public MustOverride Function CreateNewPERSON(ByVal oPERSON As PERSON) As Integer
        Public MustOverride Function UpdatePERSON(ByVal oPERSON As PERSON) As Boolean
        Public MustOverride Function DeletePERSONByID(ByVal person_id As Integer) As Boolean
        Public MustOverride Function GetPERSONByID(ByVal person_id As Integer, Optional ByVal Enabled As Boolean = False) As PERSON
        Public MustOverride Function GetPERSONByEmail(ByVal email As String) As PERSON
        Public MustOverride Function GetPERSONByAttId(ByVal attid As String, ByVal curPersonId As Integer) As PERSON
        Public MustOverride Function GetALLPERSONS() As PERSONCollection
        Public MustOverride Function GetALLPERSONSBySearch(ByVal opensearch As String, ByVal SecurityLevelCode As Integer, ByVal MinFailCount As Integer, ByVal company_id As Integer, ByVal enabledTF As Integer, ByVal CountryId As Integer, ByVal AccountId As Integer) As PERSONCollection
        

        ' ATTACHMENTS
        Public MustOverride Function InsertAttachment(ByVal oAttachment As ATTACHMENT) As Integer
        Public MustOverride Function UpdateAttachment(ByVal oAttachment As ATTACHMENT) As Boolean
        Public MustOverride Function DeleteAttachment(ByVal AttachmentID As Integer) As Boolean
        Public MustOverride Function GetAttachmentByID(ByVal AttachmentID As Integer) As ATTACHMENT
        Public MustOverride Function GetAttachmentByObject(ByVal ObjectType As String, ByVal ObjectID As Integer) As AttachmentCollection
        Public MustOverride Function GetAttachmentByObjectAndRound(ByVal ObjectType As String, ByVal ObjectID As Integer, ByVal Round As Integer) As AttachmentCollection


        ' TASKORDER
        Public MustOverride Function CreateNewTASKORDER(ByVal oTASKORDER As TASKORDER) As Integer
        Public MustOverride Function UpdateTASKORDER(ByVal oTASKORDER As TASKORDER) As Boolean
        Public MustOverride Function DeleteTASKORDERByID(ByVal taskorder_id As Integer) As Boolean
        Public MustOverride Function GetTASKORDERByID(ByVal taskorder_id As Integer) As TASKORDER
        Public MustOverride Function GetALLTASKORDERBySearch(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal coordpersonid As Integer, ByVal Context As String, ByVal CountryId As Integer) As TASKORDERCollection
        Public MustOverride Function GetDraftTOByPersonCurrency(ByVal Personid As Integer, ByVal Currency As String) As Integer
        Public MustOverride Function GetTASKORDERList(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal coordpersonid As Integer, ByVal Context As String, ByVal CountryId As Integer, ByVal ViewerSecurityCode As Integer, ByVal JobBatchNumber As String) As TASKORDERCollection
        'Public MustOverride Function GetALLTASKORDERBySearch(ByVal assignedperson_id As Integer, ByVal _status_code As Integer, ByVal startrangeDT As Date, ByVal endrangeDT As Date, ByVal IsPaid As Integer) As TASKORDERCollection
        Public MustOverride Function GetTaskorderComments(ByVal TaskId As Integer, ByVal PersonId As Integer, ByVal ReportType As Integer) As DataSet
        Public MustOverride Function GetTaskorderSubBatchDetails(ByVal JobNumber As String, ByVal BatchNumber As String) As DataSet
        Public MustOverride Function CheckForDuplicateTaskOrder(ByVal taskorder_id As Integer, ByVal jobNumber As String, ByVal batchNumber As String, ByVal subBatchNumber As String) As Boolean

        ' STATUS
        Public MustOverride Function CreateNewSTATUS(ByVal oSTATUS As STATUS) As Integer
        Public MustOverride Function UpdateSTATUS(ByVal oSTATUS As STATUS) As Boolean
        Public MustOverride Function DeleteSTATUSByID(ByVal status_code As Integer) As Boolean
        Public MustOverride Function GetSTATUSByID(ByVal status_code As Integer) As STATUS
        Public MustOverride Function GetALLSTATUSES() As STATUSCollection


        ' ITEMIZE
        Public MustOverride Function CreateNewITEMIZE(ByVal oITEMIZE As ITEMIZE) As Integer
        Public MustOverride Function UpdateITEMIZE(ByVal oITEMIZE As ITEMIZE) As Boolean
        Public MustOverride Function DeleteITEMIZEByID(ByVal itemize_id As Integer) As Boolean
        Public MustOverride Function GetITEMIZEByID(ByVal itemize_id As Integer) As ITEMIZE
        Public MustOverride Function GetITEMIZEByTaskOrderID(ByVal taskorder_id As Integer) As ITEMIZECollection

        'REPORTS / DataSets
        Public MustOverride Function gen_accounting_itemized_report(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal ispaidTF As Integer, ByVal jobBatchNumber As String, ByVal CountryId As Integer, ByVal CurrentPersonId As Integer, Optional ByVal CurrencyId As Integer = -1) As DataSet
        Public MustOverride Function gen_bidsearch_modal_result(ByVal BidId As Integer) As DataSet
        Public MustOverride Function gen_bidsearchmodal_person_qualify(ByVal RequiredUserSkillId As Integer, ByVal RequiredCopyEditSkillId As Integer, ByVal RequiredStyleSkillId As Integer, ByVal BidId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal CountryId As Integer, ByVal AccountId As Integer) As DataSet
        Public MustOverride Function gen_WorkPastDue(ByVal CountryId As Integer) As DataSet
        Public MustOverride Function gen_WorkClosing(ByVal CountryId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
        Public MustOverride Function gen_EditorPerformance() As DataSet
        Public MustOverride Function GetEditorPerformanceReportData() As System.Data.DataSet
        Public MustOverride Function GetVendorReportData() As PERSONCollection
        Public MustOverride Function GetEditorMetricsReport(ByVal ESType As String, ByVal rptType As String, startDate As Nullable(Of DateTime), endDate As Nullable(Of DateTime)) As DataSet
        Public MustOverride Function GetRptCopyEditLevels() As DataTable
        Public MustOverride Function payment_history_report(ByVal ReportType As String, ByVal StartRangeDT As Date, ByVal EndRangeDT As Date) As DataSet
        Public MustOverride Function GetAssignedFreelancersReport(ByVal jobBatchNumber As String, startDate As Nullable(Of DateTime), endDate As Nullable(Of DateTime)) As DataSet
        Public MustOverride Function GetPersonsBySkillIdList(ByVal skills As String) As System.Data.DataSet
        Public MustOverride Function GetCertificationReport(ByVal StartDate As Nullable(Of DateTime), ByVal EndDate As Nullable(Of DateTime)) As System.Data.DataSet
        Public MustOverride Function GetPaymentReport(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal jobBatchNumber As String) As DataSet

        'COPYEDITLEVEL
        Public MustOverride Function CreateNewCOPYEDITLEVEL(ByVal oCOPYEDITLEVEL As COPYEDITLEVEL) As Integer
        Public MustOverride Function UpdateCOPYEDITLEVEL(ByVal oCOPYEDITLEVEL As COPYEDITLEVEL) As Boolean
        Public MustOverride Function GetCOPYEDITLEVELBYId(ByVal copyeditcode_id As Integer) As COPYEDITLEVEL

        'Skill Lookup - 1/24/2013
        Public MustOverride Function getSkillByLookupId(ByVal userSkillId As Integer) As DisciplineTypeLookup
        Public MustOverride Function GetAllDisciplineTypes() As DisciplineTypeLookupCollection


        'USER DISCIPLINE SET- 1/24/2013
        Public MustOverride Function CreateNewUserDisciplineSet(ByVal oSkillSet As UserDisciplineSet) As Integer
        Public MustOverride Function UpdateUserDisciplineSet(ByVal oSkillSet As UserDisciplineSet) As Boolean
        Public MustOverride Function DeleteUSERDISCIPLINESETById(ByVal personId As Integer) As Boolean
        Public MustOverride Function GetUSERDISCIPLINESETByPERSONId(ByVal PersonId As Integer) As UserDisciplineSetCollection
        Public MustOverride Function GetAllUserDisciplineSets() As UserDisciplineSetCollection

        'STYLE LOOKUP - 1/24/2013
        Public MustOverride Function getStyleTypeByLookupId(ByVal StyleSkillId As Integer) As STYLETYPEPLOOKUP
        Public MustOverride Function GetAllStyleTypes() As STYLETYPELOOKUPCollection

        'USERSTYLE 1/24/2013
        Public MustOverride Function CreateNewUSERSTYLESET(ByVal oUSERSTYLE As UserStyleSet) As Integer
        Public MustOverride Function UpdateUSERSTYLESET(ByVal oUSERSTYLE As UserStyleSet) As Boolean
        Public MustOverride Function GetUserStyleSetbyPersonID(ByVal PersonID As Integer) As UserStyleSetCollection
        Public MustOverride Function GetAllUSERSTYLESets() As UserStyleSetCollection
        Public MustOverride Function DeleteUserStyleSetById(ByVal PersonID As Integer) As Boolean


        'COPY EDIT LEVEL TYPES - LOOKUP
        Public MustOverride Function getCopyEditbyLookupId(ByVal CopyEditSkillId As Integer) As COPYEDITTYPESLOOKUP
        Public MustOverride Function GetAllCopyEditTypes() As COPYEDITTYPESLOOKUPCollection

        ' USERCOPYEDITSKILLSET - 1/24/2013
        Public MustOverride Function CreateNewUSERCOPYEDITSKILLSET(ByVal oUSERCOPYEDITSKILLSET As UserCopyEditLevelSet) As Integer
        Public MustOverride Function UpdateUSERCOPYEDITSKILLSET(ByVal oUSERCOPYEDITSKILLSET As UserCopyEditLevelSet) As Boolean
        Public MustOverride Function DeleteUSERCOPYEDITSKILLSETByID(ByVal userCopyEditSkillId As Integer) As Boolean
        Public MustOverride Function GetUSERCOPYEDITSKILLSETByID(ByVal userCopyEditSkillId As Integer) As UserCopyEditLevelSet
        Public MustOverride Function getUserCopyEditSkillSetbyPersonId(ByVal PersonId As Integer) As UserCopyEditLevelSetCollection
        Public MustOverride Function GetAllUserCopyEditSets() As UserCopyEditLevelSetCollection

        'Languages - LOOKUP
        Public MustOverride Function GetAllLanguages() As LanguageLookUpCollection

        ' USERLANGUAGES - 7/21/2016
        Public MustOverride Function AddUserLanguage(ByVal oUSERLanguage As UserLanguageSet) As Integer
        Public MustOverride Function DeleteUSERLanguageByID(ByVal userLanguageId As Integer) As Boolean
        Public MustOverride Function getUserLanguagesbyPersonId(ByVal PersonId As Integer) As UserLanguageCollection


        ' BID - 1/31/2013
        Public MustOverride Function CreateNewBID(ByVal oBID As BID) As Integer
        Public MustOverride Function UpdateBID(ByVal oBID As BID) As Boolean
        Public MustOverride Function DeleteBIDByID(ByVal BidId As Integer) As Boolean
        Public MustOverride Function GetBIDByID(ByVal BidId As Integer) As BID
        Public MustOverride Function GetBIDByPersonIDAwardedBid(ByVal PersonIDAwardedBid As Integer) As BIDCollection
        Public MustOverride Function GetALLBIDSBySearch(ByVal BidID As Integer, ByVal BidStatusID As Integer, ByVal CountryId As Integer) As BIDCollection
        Public MustOverride Function GetAllBids(ByVal Context As String, ByVal CountryId As Integer) As BIDCollection
        Public MustOverride Function GetALLBIDSByStatus(ByVal BidStatusID As Integer, ByVal Context As String, ByVal CountryId As Integer) As BIDCollection
        ' BIDSTATUS
        Public MustOverride Function CreateNewBIDSTATUS(ByVal oBIDSTATUS As BIDSTATUS) As Integer
        Public MustOverride Function UpdateBIDSTATUS(ByVal oBIDSTATUS As BIDSTATUS) As Boolean
        Public MustOverride Function DeleteBIDSTATUSByID(ByVal bid_status_code As Integer) As Boolean
        Public MustOverride Function GetBIDSTATUSByID(ByVal bid_status_code As Integer) As BIDSTATUS
        Public MustOverride Function GetALLBIDSTATUSES() As BIDSTATUSCollection


        ' BIDEVENT
        Public MustOverride Function CreateNewBIDEVENT(ByVal oBIDEVENT As BIDEVENT) As Integer
        Public MustOverride Function UpdateBIDEVENT(ByVal oBIDEVENT As BIDEVENT) As Boolean
        Public MustOverride Function DeleteBIDEVENTByID(ByVal BidEventId As Integer) As Boolean
        Public MustOverride Function GetBIDEVENTByID(ByVal BidEventId As Integer) As BIDEVENT
        Public MustOverride Function GetBIDEVENTByBIDId(ByVal BidId As Integer) As BIDEVENTCollection
        Public MustOverride Function GetBIDEVENTByPersonId(ByVal PersonId As Integer) As BIDEVENTCollection
        Public MustOverride Function GetAllBidEvents() As BIDEVENTCollection
        Public MustOverride Function GetBIDEVENTByAcceptance(ByVal PersonId As Integer, ByVal PersonAcceptance As Integer) As BIDEVENTCollection

        ' BLACKOUTWINDOW
        Public MustOverride Function CreateNewBLACKOUTWINDOW(ByVal oBLACKOUTWINDOW As BLACKOUTWINDOW) As Integer
        Public MustOverride Function CreateNewBLACKOUTWINDOWTaskOrder(ByVal PersonId As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal BDescription As String, ByVal TaskOrderId As Integer, ByVal Context As String) As Integer
        Public MustOverride Function UpdateBLACKOUTWINDOW(ByVal oBLACKOUTWINDOW As BLACKOUTWINDOW) As Boolean
        Public MustOverride Function DeleteBLACKOUTWINDOWByID(ByVal UserBlackoutId As Integer) As Boolean
        Public MustOverride Function GetBLACKOUTWINDOWByID(ByVal UserBlackoutId As Integer) As BLACKOUTWINDOW
        Public MustOverride Function getUserBlackoutWindowByPersonId(ByVal PersonId As Integer, Optional ByVal ShowAll As Boolean = False) As BLACKOUTWINDOWCollection

        ' JOB
        Public MustOverride Function CreateNewJOB(ByVal oJOB As JOB) As Integer
        Public MustOverride Function UpdateJOB(ByVal oJOB As JOB) As Boolean
        Public MustOverride Function GetJOBByID(ByVal job_id As Integer) As JOB
        Public MustOverride Function GetJOBByJobNumber(ByVal job_number As Integer) As JOB
        Public MustOverride Function DeleteJOBByID(ByVal job_id As Integer) As Boolean
        Public MustOverride Function GetAllJobs() As JOBCollection

        'LIBRARYATTACHMENTS
        Public MustOverride Function InsertLibraryAttachment(ByVal oLAttachment As LIBRARYATTACHMENT) As Integer
        Public MustOverride Function UpdateLibraryAttachment(ByVal oLAttachment As LIBRARYATTACHMENT) As Boolean
        Public MustOverride Function DeleteLibraryAttachment(ByVal LibraryAttachment_id As Integer) As Boolean
        Public MustOverride Function GetLibraryAttachmentByID(ByVal LibraryAttachment_id As Integer) As LIBRARYATTACHMENT
        Public MustOverride Function GetLibraryAttachmentByObject(ByVal ObjectType As String, ByVal unitId As Integer) As LIBRARYATTACHMENTCollection

        'COPY EDIT LEVEL V# - LOOKUP
        Public MustOverride Function getCopyEditLevelbyLookupId(ByVal CopyEditLevelId As Integer, ByVal currency_Id As Integer) As COPYEDITLEVELLOOKUP_V3
        Public MustOverride Function GetAllCopyEditLevels(ByVal JobId As Integer, Optional ByVal CurrencyId As Integer = 0) As COPYEDITLEVELLOOKUP_V3Collection
        Public MustOverride Function CreateNewCopyEditLevelLookup(ByVal oCEL As COPYEDITLEVELLOOKUP_V3) As Integer
        Public MustOverride Function UpdateCopyEditLevelLookup(ByVal oCEL As COPYEDITLEVELLOOKUP_V3) As Boolean
        Public MustOverride Function DeleteCopyEditLevelLookupByID(ByVal CopyEditLevelId As Integer, ByVal CurrencyId As Integer) As Boolean
        Public MustOverride Function MoveCopyEditLevelLookupItem(ByVal CurId As Integer, DestId As Integer, Position As String) As Boolean
        Public MustOverride Function GetCopyEditLevelByName(ByVal CopyEditLevelName As String, ByVal JobId As Integer) As COPYEDITLEVELLOOKUP_V3

        Public MustOverride Function GetCurrencyDetails(Optional ByVal Currency As String = "") As CurrencyCollection

        ' ACCOUNTS    
        Public MustOverride Function GetAccountByID(ByVal AccountId As Integer) As ACCOUNTS
        Public MustOverride Function GetALLAccounts() As ACCOUNTCollection

        ' PAYMENT LOGS    
        Public MustOverride Function CreateNewPaymentLog(ByVal oPLog As TaskOrderPaymentLog) As Integer
        Public MustOverride Function UpdatePaymentLog(ByVal oPLog As TaskOrderPaymentLog) As Boolean
        Public MustOverride Function GetPaymentLogById(ByVal PaymentLog_Id As Integer) As TaskOrderPaymentLog
        Public MustOverride Function GetPaymentLogByTaskOrder(ByVal TaskOrder_Id As Integer) As TaskOrderPaymentLog

        Public MustOverride Function gen_searchmodal_person_GT(ByVal ReturnDate As DateTime) As DataSet

        Public MustOverride Function GetNewSubBatchNumber(ByVal JobNumber As String, ByVal BatchNumber As String) As Integer
        Public MustOverride Function GetRenegotiationDetails(ByVal TaskorderId As Integer, ByVal StatusCode As Integer) As DataTable
        Public MustOverride Function InsertRenegotiation(ByVal objClass As Renegotiation) As Integer

        Public MustOverride Function CreateTimesheetLog(ByVal tsLog As TimesheetLog) As Integer
        Public MustOverride Function UpdateTimesheetLog(ByVal tsLog As TimesheetLog) As Integer
        Public MustOverride Function DeleteTimesheetLogById(ByVal LogId As Integer, ByVal Person_Id As Integer) As Boolean
        Public MustOverride Function GetTimesheetLogByPersonId(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime, Optional GetPendingApproval As Boolean = False, Optional LogId As Integer = 0) As DataTable
        Public MustOverride Function GetTimesheetLogByWeek(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime) As DataTable
        Public MustOverride Function SubmitForApproval(ByVal PersonId As Integer, StartDate As DateTime, EndDate As DateTime) As Boolean
        Public MustOverride Function GetTimesheetReport(ByVal PersonId As Integer, RoleId As Integer, StartDate As DateTime, EndDate As DateTime) As DataSet
        Public MustOverride Function GetTimesheetManagerReport(ByVal CurrentRoleId As Integer, StartDate As DateTime, EndDate As DateTime) As DataSet
        Public MustOverride Function CertifyTimesheet(ByVal PersonId As Integer, Approval As Integer, StartDate As DateTime, EndDate As DateTime, CertifiedBy As String) As Boolean

        ' USER Currencies
        Public MustOverride Function AddUserCurrency(ByVal oUSERCurrencye As UserCurrencySet) As Integer
        Public MustOverride Function DeleteUSERCurrencyByID(ByVal userCurrencyId As Integer) As Boolean
        Public MustOverride Function getUserCurrenciesByPersonId(ByVal PersonId As Integer) As UserCurrencyCollection

        Public MustOverride Function GetDirectPaymentsByEC(ByVal StartRangeDT As Date, ByVal EndRangeDT As Date, ByVal Person_Id As Integer) As DataSet
        Public MustOverride Function InsertDirectPayment(ByVal objClass As DirectPaymentModel) As Boolean
        Public MustOverride Function UpdateDirectPayment(ByVal objClass As DirectPaymentModel) As Boolean
        Public MustOverride Function DeleteDirectPaymentByID(ByVal PaymentLogId As Integer) As Boolean
        Public MustOverride Function MarkDirectPaymentAsPaid(ByVal PaymentLogId As Integer, ByVal PaymentDate As DateTime) As Boolean
        '*********************************************************************
        '
        ' GenerateCollectionFromReader Delegate
        '
        ' The GenerateCollectionFromReader delegate represents any method
        ' which returns a collection from a SQL Data Reader.
        '
        '*********************************************************************
        Delegate Function GenerateCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase



        '*********************************************************************
        '
        ' Collection Helper Methods
        '
        ' The following methods are used to generate collections of objects
        ' from a SqlDataReader.
        '
        '*********************************************************************









        ' COMPANY
        Protected Function GenerateCOMPANYCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New COMPANYCollection
            While returnData.Read()
                Dim objClass As New COMPANY(CInt(returnData("company_id")), CStr(returnData("companyname")), CInt(returnData("isinternalTF")), CInt(returnData("isenabledTF")), CStr(returnData("address1")), CStr(returnData("address2")), CStr(returnData("city")), CStr(returnData("statecode")), CStr(returnData("zipcode")), CInt(returnData("CountryId")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        ' PERSONPREF
        Protected Function GeneratePERSONPREFCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New PERSONPREFCollection
            While returnData.Read()
                Dim objClass As New PERSONPREF(CInt(returnData("personID")), CStr(returnData("upkey")), CStr(returnData("upvalue")), CStr(returnData("updescription")), CInt(returnData("upseq")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' SECURITYLEVEL
        Protected Function GenerateSECURITYLEVELCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New SECURITYLEVELCollection
            While returnData.Read()
                Dim objClass As New SECURITYLEVEL(CInt(returnData("securitylevel_code")), CStr(returnData("securityname")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        ' EVENTLOG
        Protected Function GenerateEVENTLOGCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New GLOBALEVENTLOGCollection
            While returnData.Read()
                Dim objClass As New GLOBALEVENTLOG(CStr(returnData("object_code")), CInt(returnData("object_id")), CInt(returnData("person_id")), CStr(returnData("persondisplayname")), CStr(returnData("eventtype")), CDate(returnData("eventDT")), CInt(returnData("reporting_code")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        Protected Function GenerateEditorRptCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New PERSONCollection
            While returnData.Read()
                'Dim objClass As New PERSON(CInt(returnData("person_id")), CStr(returnData("firstname")), CStr(returnData("lastname")), CStr(returnData("Companyname")), CStr(returnData("Address")), CStr(returnData("mobilephone")), CStr(returnData("Email")), CStr(returnData("attid")), CStr(returnData("ECNotes")), CStr(returnData("UserNotes")), CStr(returnData("Country")), CStr(returnData("Discipline")), CStr(returnData("WorkType")), CStr(returnData("Style")), CInt(returnData("TaskOrderCount")), CInt(returnData("MultipleRoundCount")), CDate(returnData("expireDT")), CStr(returnData("LateCount")))
                Dim ecNotes As String = ""
                If returnData("ECNotes") IsNot DBNull.Value Then
                    ecNotes = CStr(returnData("ECNotes"))
                End If
                Dim AccountId As Integer = 0
                If returnData("AccountId") IsNot DBNull.Value Then
                    AccountId = CStr(returnData("AccountId"))
                End If
                Dim discipline As String = ""
                If returnData("Discipline") IsNot DBNull.Value Then
                    discipline = CStr(returnData("Discipline"))
                End If
                Dim style As String = ""
                If returnData("Style") IsNot DBNull.Value Then
                    style = CStr(returnData("Style"))
                End If
                Dim workType As String = ""
                If returnData("WorkType") IsNot DBNull.Value Then
                    workType = CStr(returnData("WorkType"))
                End If

                Dim AveragePASScore As Nullable(Of Double) = Nothing
                If returnData("AveragePASScore") IsNot DBNull.Value Then
                    AveragePASScore = CDbl(returnData("AveragePASScore"))
                End If

                Dim AveragePASScoreLast As Nullable(Of Double) = Nothing
                If returnData("AveragePASScoreLast") IsNot DBNull.Value Then
                    AveragePASScoreLast = CDbl(returnData("AveragePASScoreLast"))
                End If

                Dim objClass As New PERSON(CInt(returnData("person_id")), 0, CStr(returnData("firstname")), CStr(returnData("lastname")), "", "", CStr(returnData("Address")), "", "", "", "", "", CStr(returnData("mobilephone")), CStr(returnData("email")), 20, CStr(returnData("UserNotes")), "", 0, 0, CStr(returnData("attid")), CDate(returnData("expireDT")), 1, 0, ecNotes, "", "", -1, "", CInt(returnData("CountryId")), CStr(returnData("Country")), "", AccountId, "", "", Nothing, Nothing, "", "", CStr(returnData("Companyname")), CInt(returnData("TaskOrderCount")), CInt(returnData("PercentageOnTime")), CInt(returnData("FirstRoundAcceptanceRate")), discipline, style, workType, AveragePASScore, AveragePASScoreLast, CInt(returnData("PercentageOnTimeLast")), CInt(returnData("FirstRoundAcceptanceRateLast")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' PERSON
        Protected Function GeneratePERSONCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New PERSONCollection
            While returnData.Read()
                Dim ecNotes As String = ""
                If returnData("ECNotes") IsNot DBNull.Value Then
                    ecNotes = CStr(returnData("ECNotes"))
                End If
                Dim currency As String = ""
                If returnData("currency") IsNot DBNull.Value Then
                    currency = CStr(returnData("currency"))
                End If

                Dim currencyDisplay As String = ""
                If returnData("CurrencyDisplay") IsNot DBNull.Value Then
                    currencyDisplay = CStr(returnData("CurrencyDisplay"))
                End If

                Dim currencyName As String = ""
                If returnData("CurrencyName") IsNot DBNull.Value Then
                    currencyName = CStr(returnData("CurrencyName"))
                End If

                Dim CountryName As String = ""
                If returnData("Country") IsNot DBNull.Value Then
                    CountryName = CStr(returnData("Country"))
                End If

                Dim RoleName As String = ""
                If returnData("RoleName") IsNot DBNull.Value Then
                    RoleName = CStr(returnData("RoleName"))
                End If

                Dim AccountId As Integer = 0
                If returnData("AccountId") IsNot DBNull.Value Then
                    AccountId = CStr(returnData("AccountId"))
                End If

                Dim AccountName As String = ""
                If returnData("AccountName") IsNot DBNull.Value Then
                    AccountName = CStr(returnData("AccountName"))
                End If

                Dim CompanyName As String = ""
                If returnData("CompanyName") IsNot DBNull.Value Then
                    CompanyName = CStr(returnData("CompanyName"))
                End If

                Dim FullAddress As String = ""
                If returnData("FullAddress") IsNot DBNull.Value Then
                    FullAddress = CStr(returnData("FullAddress"))
                End If

                Dim GTType As String = ""
                If returnData("GTType") IsNot DBNull.Value Then
                    GTType = CStr(returnData("GTType"))
                End If

                Dim CreatedOn As Nullable(Of DateTime) = Nothing
                If returnData("CreatedOn") IsNot DBNull.Value Then
                    CreatedOn = CStr(returnData("CreatedOn"))
                End If

                Dim ModifiedOn As Nullable(Of DateTime) = Nothing
                If returnData("ModifiedOn") IsNot DBNull.Value Then
                    ModifiedOn = CStr(returnData("ModifiedOn"))
                End If

                Dim CreatedBy As String = ""
                If returnData("CreatedBy") IsNot DBNull.Value Then
                    CreatedBy = CStr(returnData("CreatedBy"))
                End If

                Dim ModifiedBy As String = ""
                If returnData("ModifiedBy") IsNot DBNull.Value Then
                    ModifiedBy = CStr(returnData("ModifiedBy"))
                End If

                Dim objClass As New PERSON(CInt(returnData("person_id")), CInt(returnData("company_id")), CStr(returnData("firstname")), CStr(returnData("lastname")), CStr(returnData("middleinitial")), CStr(returnData("suffixname")), CStr(returnData("address1")), CStr(returnData("address2")), CStr(returnData("city")), CStr(returnData("statecode")), CStr(returnData("zipcode")), CStr(returnData("workphone")), CStr(returnData("mobilephone")), CStr(returnData("email")), CInt(returnData("securitylevel_code")), CStr(returnData("UserNotes")), CStr(returnData("password")), CInt(returnData("resetpasswordTF")), CInt(returnData("failcount")), CStr(returnData("attid")), CDate(returnData("expireDT")), CInt(returnData("enabledTF")), CInt(returnData("deletedTF")), ecNotes, currency, currencyDisplay, CInt(returnData("CurrencyID")), currencyName, CInt(returnData("CountryId")), CountryName, RoleName, AccountId, AccountName, GTType:=GTType, CreatedOn:=CreatedOn, ModifiedOn:=ModifiedOn, CreatedBy:=CreatedBy, ModifiedBy:=ModifiedBy, CompanyName:=CompanyName, FullAddress:=FullAddress)
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' Attachments
        Protected Function GenerateAttachmentCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New AttachmentCollection
            While returnData.Read()
                Dim newClass As New ATTACHMENT(CInt(returnData("Attachment_ID")), CInt(returnData("unit_id")), CStr(returnData("FileName")), CInt(returnData("FileSize")), CStr(returnData("ContentType")), CStr(returnData("ObjectType")), CInt(returnData("Object_ID")), CStr(returnData("attachmentnote")), CInt(returnData("Round")))
                mlsCollection.Add(newClass)
            End While
            Return mlsCollection
        End Function

        ' TASKORDER
        Protected Function GenerateTASKORDERCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New TASKORDERCollection
            While returnData.Read()
                Dim currencyDisplay As String = ""
                If returnData("CurrencyDisplay") IsNot DBNull.Value Then
                    currencyDisplay = CStr(returnData("CurrencyDisplay"))
                End If

                Dim currencyId As Integer = -1
                If returnData("CurrencyID") IsNot DBNull.Value Then
                    currencyId = CInt(returnData("CurrencyID"))
                End If

                Dim paymentDT As Nullable(Of DateTime) = Nothing
                If returnData("paymentDT") IsNot DBNull.Value Then
                    paymentDT = CDate(returnData("paymentDT"))
                End If

                Dim InterimPaymentDT As Nullable(Of DateTime) = Nothing
                If returnData("InterimPaymentDT") IsNot DBNull.Value Then
                    InterimPaymentDT = CDate(returnData("InterimPaymentDT"))
                End If

                Dim pasScore As Integer = 0, pasScoreName As String = ""
                If returnData("PAS_Score") IsNot DBNull.Value Then
                    pasScore = CInt(returnData("PAS_Score"))
                End If

                If returnData("PASScoreName") IsNot DBNull.Value Then
                    pasScoreName = CStr(returnData("PASScoreName"))
                End If

                Dim AgreedFinalReturnDateDT As Nullable(Of DateTime) = Nothing
                If returnData("AgreedFinalReturnDateDT") IsNot DBNull.Value Then
                    AgreedFinalReturnDateDT = CDate(returnData("AgreedFinalReturnDateDT"))
                End If

                Dim objClass As New TASKORDER(CInt(returnData("taskorder_id")), CInt(returnData("assignedperson_id")), CStr(returnData("assignedpersonname")), CStr(returnData("apexmgr")), CInt(returnData("apexmgrperson_id")), CInt(returnData("status_code")), CStr(returnData("statusname")), CStr(returnData("iconfile")), CInt(returnData("jobnumber")), CStr(returnData("batchnumber")), CStr(returnData("bookid")), CStr(returnData("booktitle")), CStr(returnData("bookshortname")), CInt(returnData("roundnumber")), CDate(returnData("requestedreturnDT")), CDate(returnData("proposedreturnDT")), CDate(returnData("agreedreturnDT")), CDate(returnData("actualcompletedDT")), CStr(returnData("instructions")), CStr(returnData("designtemplate")), CInt(returnData("copyeditlevel")), CInt(returnData("hardcopyeditsTF")), CInt(returnData("refmanuattachedTF")), CDec(returnData("styledpagecount")), CDec(returnData("priceperpage")), CInt(returnData("perpageoverrideTF")), CDec(returnData("totalprice")), CStr(returnData("authorname")), CInt(returnData("coordperson_id")), CInt(returnData("ispaidTF")), CInt(returnData("subbatchnumber")), CStr(returnData("remarks")), (IIf(returnData("IsReviewRequired") = True, 1, 0)), currencyDisplay, currencyId, CInt(returnData("CountryId")), CInt(returnData("isInterimPaidTF")), CDbl(returnData("interimCompensationPercent")), CInt(returnData("IsInterimCertified")), CInt(returnData("IsCompletedByEC")), pasScore, pasScoreName, CStr(returnData("TaskDetails")), CStr(returnData("ColorCode")), InterimPaymentDT, paymentDT, AgreedFinalReturnDateDT)
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' STATUS
        Protected Function GenerateSTATUSCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New STATUSCollection
            While returnData.Read()
                Dim objClass As New STATUS(CInt(returnData("status_code")), CStr(returnData("statusname")), CStr(returnData("statusaction")), CStr(returnData("iconfile")), CInt(returnData("isclosedTF")), CInt(returnData("displayorder")), CInt(returnData("isenabledTF")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' ITEMIZE
        Protected Function GenerateITEMIZECollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New ITEMIZECollection



            While returnData.Read()
                Dim jobpagetype As String = ""
                If (Not returnData("jobpagetype") Is DBNull.Value) Then
                    jobpagetype = CStr(returnData("jobpagetype"))
                End If

                Dim objClass As New ITEMIZE(CInt(returnData("itemize_id")), CInt(returnData("taskorder_id")), CInt(returnData("copyeditlevel")), CDec(returnData("styledpagecount")), CDec(returnData("priceperpage")), CDec(returnData("subtotal")), CInt(returnData("displayorder")), CStr(returnData("jobtype")), CStr(returnData("CopyEditLevelName")), CDec(returnData("manual_styledpagecount")), CDec(returnData("manual_priceperpage")), CDec(returnData("manual_selectedpercent")), jobpagetype)
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        'COPYEDITLEVEL

        Protected Function GenerateCOPYEDITLEVELCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New COPYEDITLEVELCollection
            While returnData.Read()
                Dim objClass As New COPYEDITLEVEL(CInt(returnData("copyeditcode_id")), CStr(returnData("copyeditlevel_name")), CDbl(returnData("CopyEditLevelPrice")), CInt(returnData("displayorder")), CDec(returnData("isEditableTF")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        'SKILLLOOKUP
        Protected Function GenerateDisciplineTypeLookupFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New DisciplineTypeLookupCollection
            While returnData.Read()
                Dim objClass As New DisciplineTypeLookup(CInt(returnData("UserSkillId")), CStr(returnData("userSkillName")), CStr(returnData("userSkillDescription")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        'UserDisciplineSet
        Protected Function GenerateUserDisciplineSetCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New UserDisciplineSetCollection
            While returnData.Read()
                Dim objClass As New UserDisciplineSet(CInt(returnData("userSkillId")), CInt(returnData("PersonId")), CStr(returnData("SkillDescription")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        'STYLELOOKUP
        Protected Function GenerateSTYLETYPELOOKUPCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New STYLETYPELOOKUPCollection
            While returnData.Read()
                Dim objClass As New STYLETYPEPLOOKUP(CInt(returnData("StyleSkillId")), CStr(returnData("StyleName")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' USERSTYLE

        Protected Function GenerateUSERSTYLECollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New UserStyleSetCollection
            While returnData.Read()
                Dim objClass As New UserStyleSet(CInt(returnData("UserStyleId")), CInt(returnData("PersonID")), CStr(returnData("StyleSkillName")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function







        ' USERCOPYEDITSKILLLOOKUP

        Protected Function GenerateCOPYEDITTYPELOOKUPCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New COPYEDITTYPESLOOKUPCollection
            While returnData.Read()
                Dim objClass As New COPYEDITTYPESLOOKUP(CInt(returnData("CopyEditSkillId")), CStr(returnData("CopyEditSkillName")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' USERCOPYEDITSKILLSET

        Protected Function GenerateUSERCOPYEDITSKILLSETCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New UserCopyEditLevelSetCollection
            While returnData.Read()
                Dim objClass As New UserCopyEditLevelSet(CInt(returnData("userCopyEditSkillId")), CInt(returnData("PersonId")), CStr(returnData("CopyEditSkillName")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' LANGUAGELOOKUP

        Protected Function GenerateLanguageLOOKUPCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New LanguageLookUpCollection
            While returnData.Read()
                Dim objClass As New LanguagesLookUp(CInt(returnData("LanguageId")), CStr(returnData("Language")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' USERLANGUAGESET

        Protected Function GenerateUSERLanguageCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New UserLanguageCollection
            While returnData.Read()
                Dim objClass As New UserLanguageSet(CInt(returnData("LanguageId")), CInt(returnData("PersonId")), CStr(returnData("Language")), CInt(returnData("IsSelected")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' BID

        Protected Function GenerateBIDCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New BIDCollection
            While returnData.Read()
                Dim ProjectExtent As String = ""
                If Not returnData("ProjectExtent") Is DBNull.Value Then
                    ProjectExtent = returnData("ProjectExtent")
                End If

                Dim Notes As String = ""
                If Not returnData("Notes") Is DBNull.Value Then
                    Notes = returnData("Notes")
                End If

                'Dim PageCount As Integer = 0
                'If returnData("PageCount") IsNot DBNull.Value Then
                '    PageCount = CInt(returnData("PageCount"))
                'End If

                Dim objClass As New BID(CInt(returnData("BidId")), CInt(returnData("BidStatusID")), CInt(returnData("PersonIDAwardedBid")), CInt(returnData("ApexManagerID")), CInt(returnData("RequiredUserSkillID")), CInt(returnData("RequiredCopyEditSkillId")), CInt(returnData("RequiredStyleSkillId")), CDate(returnData("RespondByDate")), CDate(returnData("WorkStartDate")), CDate(returnData("WorkEndDate")), CStr(returnData("BidTitle")), CInt(returnData("ConvertedTaskOrderID")), CStr(returnData("statusname")), CStr(returnData("iconfile")), CStr(returnData("assignedpersonname")), ProjectExtent, Notes, CInt(returnData("CountryId")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        'BIDSTATUS

        Protected Function GenerateBIDSTATUSCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New BIDSTATUSCollection
            While returnData.Read()
                Dim objClass As New BIDSTATUS(CInt(returnData("bid_status_code")), CStr(returnData("bidstatusname")), CStr(returnData("bidstatusaction")), CStr(returnData("bidiconfile")), CInt(returnData("isClosedTF")), CInt(returnData("bidstatusdisplayorder")), CInt(returnData("isenabledTF")))

                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        'BIDEVENT

        Protected Function GenerateBIDEVENTCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New BIDEVENTCollection
            While returnData.Read()
                Dim objClass As New BIDEVENT(CInt(returnData("BidEventId")), CInt(returnData("BidId")), CInt(returnData("PersonID")), CInt(returnData("PersonAcceptTF")))

                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        'BLACKOUTWINDOW

        Protected Function GenerateBLACKOUTWINDOWCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New BLACKOUTWINDOWCollection
            While returnData.Read()
                Dim objClass As New BLACKOUTWINDOW(CInt(returnData("UserBlackoutId")), CInt(returnData("PersonId")), CDate(returnData("StartDate")), CDate(returnData("EndDate")), CStr(returnData("BlackoutDescription")))

                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' JOB

        Protected Function GenerateJOBCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New JOBCollection
            While returnData.Read()
                Dim objClass As New JOB(CInt(returnData("job_id")), CInt(returnData("job_number")), CStr(returnData("job_description")), CInt(returnData("isEnabledTF")), CInt(returnData("isInterimProductRequired")), CInt(returnData("interimCompensationPercent")))

                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        ' LIBRARYATTACHMENT
        Protected Function GenerateLIBRARYATTACHMENTCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New LIBRARYATTACHMENTCollection
            While returnData.Read()
                Dim newClass As New LIBRARYATTACHMENT(CInt(returnData("LibraryAttachment_id")), CInt(returnData("unit_id")), CStr(returnData("FileName")), CInt(returnData("FileSize")), CStr(returnData("ContentType")), CStr(returnData("ObjectType")), CStr(returnData("attachmentnotes")))
                mlsCollection.Add(newClass)
            End While
            Return mlsCollection
        End Function


        Protected Function GenerateCOPYEDITLEVELLookupCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New COPYEDITLEVELLOOKUP_V3Collection
            While returnData.Read()
                Dim strDesc As String = ""
                If returnData("Description") IsNot DBNull.Value Then
                    strDesc = CStr(returnData("Description"))
                End If

                Dim JobId As Integer = -1
                If returnData("Job_Id") IsNot DBNull.Value Then
                    JobId = CInt(returnData("Job_Id"))
                End If

                Dim JobNumber As String = ""
                If returnData("JobNumber") IsNot DBNull.Value Then
                    JobNumber = CInt(returnData("JobNumber"))
                End If

                Dim defaultUnitCount As Nullable(Of Integer) = Nothing
                If returnData("DefaultUnitCount") IsNot DBNull.Value Then
                    defaultUnitCount = CInt(returnData("DefaultUnitCount"))
                End If
                Dim objClass As New COPYEDITLEVELLOOKUP_V3(CInt(returnData("CopyEditLevel_id")), CStr(returnData("Name")), strDesc, JobId, CStr(returnData("UnitType")), CDec(returnData("UnitPrice")), CInt(returnData("DisplayOrder")), CStr(returnData("IsEnabled")), JobNumber, CInt(returnData("CurrencyId")), CStr(returnData("DisplayText")), CStr(returnData("JobType")), defaultUnitCount, CBool(returnData("IsUnitCountEditable")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function

        Protected Function GenerateCurrencyCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New CurrencyCollection
            While returnData.Read()
                Dim objClass As New CURRENCY(CStr(returnData("Name")), CStr(returnData("Currency")), CStr(returnData("DisplayText")), CStr(returnData("CurrencyId")), CStr(returnData("CountryId")), CStr(returnData("Country")))

                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function


        ' ACCOUNT
        Protected Function GenerateACCOUNTCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New ACCOUNTCollection
            While returnData.Read()
                Dim objClass As New ACCOUNTS(CInt(returnData("AccountId")), CStr(returnData("AccountName")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        ' TaskOrderPaymentLog
        Protected Function GeneratePaymentLogCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New TaskorderPaymentLogCollection
            While returnData.Read()
                Dim iAmt As Double = 0, fAmt As Double = 0
                If (returnData("InterimAmount") IsNot DBNull.Value) Then
                    iAmt = CDbl(returnData("InterimAmount"))
                End If
                If (returnData("FinalPaymentAmount") IsNot DBNull.Value) Then
                    fAmt = CDbl(returnData("FinalPaymentAmount"))
                End If
                Dim iPaymentDT As Nullable(Of DateTime) = Nothing
                If returnData("InterimPaymentDT") IsNot DBNull.Value Then
                    iPaymentDT = CDate(returnData("InterimPaymentDT"))
                End If
                Dim iFinalPaymentDT As Nullable(Of DateTime) = Nothing
                If returnData("FinalPaymentDT") IsNot DBNull.Value Then
                    iFinalPaymentDT = CDate(returnData("FinalPaymentDT"))
                End If
                Dim newClass As New TaskOrderPaymentLog(CInt(returnData("PaymentLog_Id")), CInt(returnData("TaskOrder_Id")), iAmt, fAmt, iPaymentDT, iFinalPaymentDT)
                mlsCollection.Add(newClass)
            End While
            Return mlsCollection
        End Function

        ' USERCurrencySET

        Protected Function GenerateUSERCurrencyCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
            Dim mlsCollection As New UserCurrencyCollection
            While returnData.Read()
                Dim objClass As New UserCurrencySet(CInt(returnData("CurrencyId")), CInt(returnData("PersonId")), CStr(returnData("Currency")), CInt(returnData("IsSelected")))
                mlsCollection.Add(objClass)
            End While
            Return mlsCollection
        End Function



        ' TaskOrderPaymentLog
        'Protected Function GenerateTimesheetLogCollectionFromReader(ByVal returnData As IDataReader) As CollectionBase
        '    Dim mlsCollection As New TimesheetLogCollection
        '    While returnData.Read()
        '        Dim entryDate As Nullable(Of DateTime) = Nothing
        '        If returnData("Entry_Date") IsNot DBNull.Value Then
        '            entryDate = CDate(returnData("Entry_Date"))
        '        End If
        '        Dim createdOn As Nullable(Of DateTime) = Nothing
        '        If returnData("CreatedOn") IsNot DBNull.Value Then
        '            createdOn = CDate(returnData("CreatedOn"))
        '        End If
        '        Dim modifiedOn As Nullable(Of DateTime) = Nothing
        '        If returnData("ModifiedOn") IsNot DBNull.Value Then
        '            modifiedOn = CDate(returnData("ModifiedOn"))
        '        End If
        '        Dim newClass As New TimesheetLog(CInt(returnData("Log_Id")), CInt(returnData("Person_Id")), CInt(returnData("Year_Number")), CInt(returnData("Week_Number")), CStr(returnData("Job_Number")), CStr(returnData("Editorial_Service")), CStr(returnData("Notes")), CStr(returnData("CertifiedBy")), CDec(returnData("Hours")), CBool(returnData("IsSubmittedTF")), CBool(returnData("IsApprovedTF")), entryDate, createdOn, modifiedOn)
        '        mlsCollection.Add(newClass)
        '    End While
        '    Return mlsCollection
        'End Function

        'DataAccessLayerBaseClass
    End Class

    '*********************************************************************
    '
    ' DataAccessLyerBaseClassHelper Class
    '
    ' Loads different data access layers depending on the configuration
    ' setting in the Web.Config file.
    '
    '*********************************************************************


    Public Class DataAccessLayerBaseClassHelper

        Public Shared Function GetDataAccessLayer() As DataAccessLayerBaseClass
            Dim trp As Type = Type.GetType(Globals.DataAccessType, True)
            ' Throw an error if wrong base type
            If Not trp.BaseType Is Type.GetType("PrePress.DataAccessLayer.DataAccessLayerBaseClass") Then
                Throw New Exception("Data Access Layer does not inherit DataAccessLayerBaseClass!")
            End If
            Dim dc As DataAccessLayerBaseClass = CType(Activator.CreateInstance(trp), DataAccessLayerBaseClass)
            Return dc
        End Function 'GetDataAccessLayer

    End Class 'DataAccessLayerBaseClassHelper


End Namespace 'DataAccessLayer




