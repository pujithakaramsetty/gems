Imports PrePress.BusinessLogicLayer
Imports Telerik.Web.UI
Imports PrePress.DataAccessLayer

Public Class DirectPayments
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '*************************
        '*** Get User Security ***
        '*************************
        Dim oCurUser As PERSON
        oCurUser = PERSON.GetPERSONByEmail(Page.User.Identity.Name)
        If oCurUser Is Nothing Then
            Response.Redirect(ResolveUrl("~/Common/logoff.aspx"))
        End If


        If oCurUser.security_code <= 20 Then
            radDirectPayments.Enabled = False
            radDirectPayments.Visible = False
            radDirectPayments.Visible = False
            PageMessage.DisplayMessage(PrePress.PageMessage.MessageType.Critical, "You do not have the required permissions to view this page. ")
        End If
        If Page.IsPostBack = False Then
            StartDT.SelectedDate = DateTime.Now.AddMonths(-1)
            EndDT.SelectedDate = DateTime.Now
        End If
        Dim profileData As String = Request.QueryString("profile")
        RadWindowManager1.Windows.Clear()

    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender

        ''*************************
        ''*** Get User Security ***
        ''*************************
        'Dim oCurUser As PERSON
        'oCurUser = PERSON.GetPERSONByEmail(Page.User.Identity.Name)
        'If oCurUser Is Nothing Then
        '    Response.Redirect(ResolveUrl("~/Common/logoff.aspx"))
        'End If

        ''Check security for reports
        'If oCurUser.security_code <= 20 Then
        '    'NOT ACCOUNTING USER
        '    PageMessage.DisplayMessage(PrePress.PageMessage.MessageType.Information, "Security: You do not have the permissions required to view this page.")
        '    radDirectPayments.Enabled = False
        '    radDirectPayments.Visible = False
        '    radDirectPayments.Visible = False
        '    Exit Sub
        'End If
        Dim oCurUser As PERSON
        oCurUser = PERSON.GetPERSONByEmail(Page.User.Identity.Name)

        fntFillSearchResults(oCurUser.person_id)

    End Sub


    Private Sub fntFillSearchResults(ByVal Person_Id As Integer)


        Dim DBLayer As DataAccessLayerBaseClass = DataAccessLayerBaseClassHelper.GetDataAccessLayer
        Dim ds As New System.Data.DataSet
        'CUSTOM SEARCH RANGE
        Dim SDate As Date
        Dim EDate As Date



        If StartDT.SelectedDate.HasValue Then
            SDate = StartDT.SelectedDate
            EDate = EndDT.SelectedDate
        End If

        If Util.IsNullDate(SDate) Then
            ds = DBLayer.GetDirectPaymentsByEC(Util.NullDate, DateTime.Now, Person_Id)

        Else
            ds = DBLayer.GetDirectPaymentsByEC(SDate, EDate, Person_Id)
        End If

        radDirectPayments.DataSource = ds
        radDirectPayments.DataBind()
    End Sub

    Private Sub btnExportExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExport.Click

        'radDirectPayments.MasterTableView.GetColumn("assignedpersonname").Visible = True

        radDirectPayments.ExportSettings.OpenInNewWindow = True
        radDirectPayments.ExportSettings.FileName = "Direct Payments"
        radDirectPayments.ExportSettings.IgnorePaging = True
        radDirectPayments.MasterTableView.ExportToCSV()
    End Sub

    Protected Sub radDirectPayments_InsertCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles radDirectPayments.InsertCommand
        'ON INSERT

        Dim item As GridDataItem = TryCast(e.Item, GridDataItem)
        Dim ht As New Hashtable()

        Dim editedItem As GridEditableItem = CType(e.Item, GridEditableItem)
        Dim ucObj As UserControl = CType(e.Item.FindControl(GridEditFormItem.EditFormUserControlID), UserControl)

        'ht.Add("PaymentLogId", CType(ucObj.FindControl("hdnPaymentLogId"), HiddenField).Value)
        ht.Add("PersonId", CType(ucObj.FindControl("cmbFreelancer"), RadComboBox).SelectedValue)
        ht.Add("JobId", CType(ucObj.FindControl("cmbJob"), RadComboBox).SelectedValue)
        ht.Add("BatchNumber", CType(ucObj.FindControl("txtBatchNumber"), RadTextBox).Text)
        ht.Add("EditorialService", CType(ucObj.FindControl("txtEditorialService"), RadTextBox).Text)
        ht.Add("UnitType", CType(ucObj.FindControl("txtUnitType"), RadTextBox).Text)
        ht.Add("JobType", CType(ucObj.FindControl("txtJobType"), RadTextBox).Text)
        ht.Add("IsLumpsum", CType(ucObj.FindControl("cmbIsLumpsum"), RadComboBox).SelectedValue)
        ht.Add("CurrencyId", CType(ucObj.FindControl("cmbCurrency"), RadComboBox).SelectedValue)
        ht.Add("PageCount", CType(ucObj.FindControl("rntPageCount"), RadNumericTextBox).Text)

        If (CType(ucObj.FindControl("rntPricePerPage"), RadNumericTextBox).Text = "") Then
            ht.Add("PricePerPage", -1)
        Else
            ht.Add("PricePerPage", CType(ucObj.FindControl("rntPricePerPage"), RadNumericTextBox).Text)
        End If

        If (CType(ucObj.FindControl("rntTotalPrice"), RadNumericTextBox).Text = "") Then
            ht.Add("TotalPrice", -1)
        Else
            ht.Add("TotalPrice", CType(ucObj.FindControl("rntTotalPrice"), RadNumericTextBox).Text)
        End If

        'Clear Message Area
        PageMessage.ClearMessages()

        'If txtTaskOrderID.Text = 0 Then
        '    cpyform.fntSaveInsert()
        'End If


        'Save Insert
        If fntSave(ht, e, "Insert") = True Then
            'Successful
            PageMessage.DisplayMessage(PrePress.PageMessage.MessageType.Successful, "Successful: This Item was saved.")
        Else
            'Failed to Save
            'e.Canceled = True 'Stay in Edit Mode
        End If
    End Sub

    Protected Sub radDirectPayments_UpdateCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles radDirectPayments.UpdateCommand
        'ON INSERT

        Dim item As GridDataItem = TryCast(e.Item, GridDataItem)
        Dim ht As New Hashtable()

        Dim editedItem As GridEditableItem = CType(e.Item, GridEditableItem)
        Dim ucObj As UserControl = CType(e.Item.FindControl(GridEditFormItem.EditFormUserControlID), UserControl)

        ht.Add("PaymentLogId", CType(ucObj.FindControl("hdnPaymentLogId"), HiddenField).Value)
        ht.Add("PersonId", CType(ucObj.FindControl("cmbFreelancer"), RadComboBox).SelectedValue)
        ht.Add("JobId", CType(ucObj.FindControl("cmbJob"), RadComboBox).SelectedValue)
        ht.Add("BatchNumber", CType(ucObj.FindControl("txtBatchNumber"), RadTextBox).Text)
        ht.Add("EditorialService", CType(ucObj.FindControl("txtEditorialService"), RadTextBox).Text)
        ht.Add("UnitType", CType(ucObj.FindControl("txtUnitType"), RadTextBox).Text)
        ht.Add("JobType", CType(ucObj.FindControl("txtJobType"), RadTextBox).Text)
        ht.Add("IsLumpsum", CType(ucObj.FindControl("cmbIsLumpsum"), RadComboBox).SelectedValue)
        ht.Add("CurrencyId", CType(ucObj.FindControl("cmbCurrency"), RadComboBox).SelectedValue)
        ht.Add("PageCount", CType(ucObj.FindControl("rntPageCount"), RadNumericTextBox).Text)

        If (CType(ucObj.FindControl("rntPricePerPage"), RadNumericTextBox).Text = "") Then
            ht.Add("PricePerPage", -1)
        Else
            ht.Add("PricePerPage", CType(ucObj.FindControl("rntPricePerPage"), RadNumericTextBox).Text)
        End If

        If (CType(ucObj.FindControl("rntTotalPrice"), RadNumericTextBox).Text = "") Then
            ht.Add("TotalPrice", -1)
        Else
            ht.Add("TotalPrice", CType(ucObj.FindControl("rntTotalPrice"), RadNumericTextBox).Text)
        End If

        'Clear Message Area
        PageMessage.ClearMessages()

        'If txtTaskOrderID.Text = 0 Then
        '    cpyform.fntSaveInsert()
        'End If


        'Save Insert
        If fntSave(ht, e, "Update") = True Then
            'Successful
            PageMessage.DisplayMessage(PrePress.PageMessage.MessageType.Successful, "Successful: This Item was saved.")
        Else
            'Failed to Save
            'e.Canceled = True 'Stay in Edit Mode
        End If
    End Sub

    Protected Sub radDirectPayments_DeleteCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles radDirectPayments.DeleteCommand

        Dim item As GridDataItem = TryCast(e.Item, GridDataItem)
        Dim ht As New Hashtable()

        ht.Add("PaymentLogId", item.GetDataKeyValue("PaymentLogId"))
        'Extract Values
        item.ExtractValues(ht)
        'Clear Message Area
        PageMessage.ClearMessages()
        'Save Insert
        If fntDeleteRecord(ht, e) = True Then
            'Successful
            PageMessage.DisplayMessage(PrePress.PageMessage.MessageType.Successful, "Successful: This Item was deleted.")
        Else
            'Failed to Save
            'e.Canceled = True 'Stay in Edit Mode
        End If
    End Sub

    Public Function fntSave(ByVal ht As Hashtable, ByVal e As Telerik.Web.UI.GridCommandEventArgs, ByVal strType As String) As Boolean

        '*************************
        '*** Get User Security ***
        '*************************
        Dim oCurUser As PERSON
        oCurUser = PERSON.GetPERSONByEmail(Page.User.Identity.Name)
        If oCurUser Is Nothing Then
            Response.Redirect(ResolveUrl("~/Common/logoff.aspx"))
        End If

        Dim dpm As New DirectPaymentModel()
        If (strType = "Update") Then
            dpm.PaymentLogId = ht("PaymentLogId")
        End If

        Dim strFreelancer As String = ht("PersonId")
        Dim PersonId As Integer = Convert.ToInt32(ht("PersonId").ToString().Split(New Char() {"_"c})(0).ToString())

        dpm.PersonId = PersonId
        dpm.JobId = ht("JobId")
        dpm.BatchNumber = ConvertNullToEmpty(ht("BatchNumber"))
        dpm.EditorialService = ht("EditorialService")
        dpm.JobType = ht("JobType")
        dpm.UnitType = ht("UnitType")
        dpm.IsLumpsum = ht("IsLumpsum")
        dpm.UnitCount = ht("PageCount")
        dpm.CurrencyId = ht("CurrencyId")
        dpm.PricePerPage = Convert.ToDouble(ht("PricePerPage"))
        dpm.TotalPrice = Convert.ToDouble(ht("TotalPrice"))
        dpm.LoggedBy = oCurUser.person_id
        dpm.LoggedOn = DateTime.Now

        Dim DBLayer As DataAccessLayerBaseClass = DataAccessLayerBaseClassHelper.GetDataAccessLayer
        If (strType = "Insert") Then
            DBLayer.InsertDirectPayment(dpm)
        ElseIf (strType = "Update") Then
            DBLayer.UpdateDirectPayment(dpm)
        End If
        Return True
    End Function

    Private Function fntDeleteRecord(ByVal ht As Hashtable, ByVal e As Telerik.Web.UI.GridCommandEventArgs) As Boolean
        ' ON DELETE
        Dim PaymentLogId As Integer

        If IsNumeric(ConvertNullToEmpty(ht("PaymentLogId"))) Then
            PaymentLogId = CInt(ConvertNullToEmpty(ht("PaymentLogId")))
        Else
            'Item id is invalid
            Exit Function
        End If
        Dim DBLayer As DataAccessLayerBaseClass = DataAccessLayerBaseClassHelper.GetDataAccessLayer
        DBLayer.DeleteDirectPaymentByID(PaymentLogId)
        Return True
    End Function

    Private Function ConvertNullToEmpty(ByVal obj As Object) As String
        If obj Is Nothing Then
            Return [String].Empty
        Else
            Return obj.ToString()
        End If
    End Function

End Class