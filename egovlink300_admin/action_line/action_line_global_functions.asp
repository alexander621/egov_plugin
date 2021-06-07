<%
 'lcl_dsn = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
 lcl_dsn = Application("DSN")

'------------------------------------------------------------------------------
function checkRecordsPerPageFilter(p_default,p_recordsPer)
  lcl_default = 25

 'Set up the default number of records to pull.
 'This is in preparation of user search defaults
  if p_default <> "" then
     if isnumeric(p_default) then
        lcl_default = p_default
     end if
  end if

  lcl_return = lcl_default

 'Validate the value the user has entered.
  if p_recordsPer <> "" then
     if isnumeric(p_recordsPer) then
        if clng(p_recordsPer) > 0 then
           lcl_return = p_recordsPer
        end if
     end if
  end if

  checkRecordsPerPageFilter = lcl_return

end function

'   if recordsPer = "" or IsNull(recordsPer) then
'      recordsPer = 25
'   else
'      if isnumeric(recordsPer) then
'         if clng(recordsPer) = 0 then
'            recordsPer = 25
'         end if
'      else
'         recordsPer = 25
'      end if
'   end if

'------------------------------------------------------------------------------
 sub displayReportTypesList(p_reporttype,p_orghasfeature_actionline_listfull,p_orghasfeature_responsetimereporting)
   lcl_selected_detail          = ""
   lcl_selected_summary         = ""
   lcl_selected_listfull        = ""
   lcl_selected_responsesummary = ""
   lcl_selected_responsedetail  = ""
   lcl_selected_statussummary   = ""

  'Determine which option is selected
   if UCASE(p_reporttype) = "DETAIL" then
      lcl_selected_detail = " selected=""selected"""

   elseif UCASE(p_reporttype) = "SUMMARY" then
      lcl_selected_summary = " selected=""selected"""

   elseif UCASE(p_reporttype) = "LISTFULL" then
      lcl_selected_listfull = " selected=""selected"""

   elseif UCASE(p_reporttype) = "RESPONSESUMMARY" then
      lcl_selected_responsesummary = " selected=""selected"""

   elseif UCASE(p_reporttype) = "RESPONSEDETAIL" then
      lcl_selected_responsedetail = " selected=""selected"""

   elseif UCASE(p_reporttype) = "STATUSSUMMARY" then
      lcl_selected_statussummary = " selected=""selected"""

   end if

  'Display Report Type Option List
   response.write "  <option value=""List"">List</option>" & vbcrlf

   if p_orghasfeature_actionline_listfull then
      response.write "  <option value=""ListFull""" & lcl_selected_listfull & ">List (Full)</option>" & vbcrlf
   end if

   response.write "  <option value=""Summary""" & lcl_selected_summary & ">Summary</option>" & vbcrlf
   response.write "  <option value=""Detail"""  & lcl_selected_detail  & ">Detail</option>" & vbcrlf

   if p_orghasfeature_responsetimereporting then
      response.write "  <option value=""ResponseSummary""" & lcl_selected_responsesummary & ">Response Summary</option>" & vbcrlf
      response.write "  <option value=""responsedetail"""  & lcl_selected_responsedetail  & ">Response Detail</option>" & vbcrlf
   end if

   response.write "  <option value=""statussummary""" & lcl_selected_statussummary & ">Status (Summary)</option>" & vbcrlf

 end sub

'------------------------------------------------------------------------------
 sub displayOrderByList(p_orderby, p_orghasfeature_issuelocation, _
                        p_orghasfeature_actionline_maintain_duedate, _
                        p_userhaspermission_actionline_maintain_duedate)

   lcl_selected_assigned_name = ""
   lcl_selected_action_formid = ""
   lcl_selected_submit_date   = ""
   lcl_selected_due_date      = ""
   lcl_selected_deptid        = ""
   lcl_selected_streetname    = ""
   lcl_selected_status        = ""

   if UCASE(p_orderby) = "ASSIGNED_NAME" then
      lcl_selected_assigned_name = " selected=""selected"""
   elseif UCASE(p_orderby) = "ACTION_FORMID" then
      lcl_selected_action_formid = " selected=""selected"""
   elseif UCASE(p_orderby) = "SUBMIT_DATE" then
      lcl_selected_submit_date = " selected=""selected"""
   elseif UCASE(p_orderby) = "DUE_DATE" then
      lcl_selected_due_date = " selected=""selected"""
   elseif UCASE(p_orderby) = "DEPTID" then
      lcl_selected_deptid = " selected=""selected"""
   elseif UCASE(p_orderby) = "STREETNAME" then
      lcl_selected_streetname = " selected=""selected"""
   elseif UCASE(p_orderby) = "SUBMITTEDBY" then
      lcl_selected_submittedby = " selected=""selected"""
   elseif UCASE(p_orderby) = "STATUS" then
      lcl_selected_status = " selected=""selected"""
   end if

   response.write "  <option value=""assigned_Name""" & lcl_selected_assigned_name & ">Assigned To</option>" & vbcrlf
   response.write "  <option value=""action_Formid""" & lcl_selected_action_formid & ">Form</option>" & vbcrlf
   response.write "  <option value=""submit_date"""   & lcl_selected_submit_date   & ">Date Descending</option>" & vbcrlf

   if p_orghasfeature_actionline_maintain_duedate AND p_userhaspermission_actionline_maintain_duedate then
      response.write "  <option value=""due_date"""   & lcl_selected_due_date      & ">Due Date</option>" & vbcrlf
   end if

   response.write "  <option value=""deptId"""        & lcl_selected_deptid        & ">Department</option>" & vbcrlf

   if p_orghasfeature_issuelocation then
      response.write "  <option value=""streetname""" & lcl_selected_streetname & ">Issue/Problem Location Street Name</option>" & vbcrlf
   end if

   response.write "  <option value=""submittedby""" & lcl_selected_submittedby & ">Submitted By</option>" & vbcrlf
   response.write "  <option value=""status"""      & lcl_selected_status      & ">Status</option>" & vbcrlf

 end sub

'------------------------------------------------------------------------------
 function buildQueryForExportToCSV(iSQL, p_orghasfeature_action_line_substatus, _
                                         p_orghasfeature_issuelocation, _
                                         p_orghasfeature_actionline_maintain_duedate, _
                                         p_userhaspermission_actionline_maintain_duedate)
   lcl_return = ""

   if iSQL <> "" then
      lcl_return = "SELECT "
      lcl_return = lcl_return & "CAST(action_autoid AS varchar) + RIGHT('0'+ CAST(DATEPART(hh, submit_date) AS varchar),2) + RIGHT('0' + CAST(DATEPART(mi, submit_date) AS varchar),2) as [E-Gov Link Tracking #], "
      lcl_return = lcl_return & "action_formtitle as [Action Form Name], "
      lcl_return = lcl_return & "submit_date as [Date Submitted], "
      lcl_return = lcl_return & "status as [Status], "

      if p_orghasfeature_action_line_substatus then
          lcl_return = lcl_return & "sub_status_desc as [Sub-Status], "
      end if

      lcl_return = lcl_return & "userfname + ' ' + userlname as [Submitted By], "
      lcl_return = lcl_return & "submittedbytype as [Submitted By Type], "

      if p_orghasfeature_actionline_maintain_duedate AND p_userhaspermission_actionline_maintain_duedate then
         lcl_return = lcl_return & "due_date as [Due Date], "
      end if

      lcl_return = lcl_return & "assigned_Name as [Assigned To], "
      lcl_return = lcl_return & "FirstActionDate as [FirstActionDate], "
      lcl_return = lcl_return & "DateDiff(d,submit_date,firstactiondate) as [Days to First Action], "
      lcl_return = lcl_return & "complete_date as [Resolved Date], "
      lcl_return = lcl_return & "DateDiff(d,submit_date,complete_date) as [Days to Resolve], "
      lcl_return = lcl_return & "groupname as [Department], "

      if p_orghasfeature_issuelocation then
       		lcl_return = lcl_return & "streetnumber as [Issue Street Number], "
       		lcl_return = lcl_return & "streetname as [Issue Street Name], "
       		lcl_return = lcl_return & "city as [Issue City], "
   	  	  lcl_return = lcl_return & "state as [Issue State], "
      		 lcl_return = lcl_return & "zip as [Issue Zip], "
         lcl_return = lcl_return & "validstreet as [Non-Listed Address], "
         lcl_return = lcl_return & "comments as [Additional Info], "
      end if

      lcl_return = lcl_return & "userfname as [First Name], "
      lcl_return = lcl_return & "userlname as [Last Name], "
      lcl_return = lcl_return & "userbusinessname as [Business Name], "
      lcl_return = lcl_return & "useremail as [Email], "
      'lcl_return = lcl_return & "userworkphone as [Daytime Phone], "
      lcl_return = lcl_return & "userhomephone as [Daytime Phone], "
      lcl_return = lcl_return & "userfax as [Fax], "
      lcl_return = lcl_return & "useraddress as [Address], "
      lcl_return = lcl_return & "useraddress2 as [Address 2], "
      lcl_return = lcl_return & "usercity as [City], "
      lcl_return = lcl_return & "userstate as [State], "
      lcl_return = lcl_return & "userzip as [Zip], "
      lcl_return = lcl_return & "contactmethodid as [Preferred Contact Method], "
      lcl_return = lcl_return & "comment as [Form Values], "
      lcl_return = lcl_return & "'' as [Internal Values] "
      lcl_return = lcl_return & RIGHT(iSQL,Len(iSQL)-instr(iSQL,"FROM")+1)
   end if

   buildQueryForExportToCSV = lcl_return

 end function

'------------------------------------------------------------------------------
 function buildQueryForDownloadActivityLog(p_WhereClause,p_order_by,p_orderby)
  'p_order_by: columns identified to be used in ORDER BY
  'p_orderby:  field identified by user on how list should be ordered.
   lcl_return = ""

   lcl_return = "SELECT "
   lcl_return = lcl_return & "CAST(egov_action_request_view.action_autoid AS varchar) + RIGHT('0'+ CAST(DATEPART(hh, egov_action_request_view.submit_date) AS varchar),2) + RIGHT('0' + CAST(DATEPART(mi, egov_action_request_view.submit_date) AS varchar),2) as [E-Gov Link Tracking #], "
   lcl_return = lcl_return & "egov_action_request_view.submit_date as [submit_date], "
   lcl_return = lcl_return & "es.status_name as [status_name], "
   lcl_return = lcl_return & "Users.FirstName + ' ' + Users.LastName + ' - ' + UPPER(egr.action_status) + '[sub_status]' + ' - ' + CAST(egr.action_editdate AS varchar) AS [status_line], "
   lcl_return = lcl_return & "Users.FirstName AS [users_first_name], "
   lcl_return = lcl_return & "Users.LastName AS [users_last_name], "
   lcl_return = lcl_return & "egr.action_status AS [action_status], "
   lcl_return = lcl_return & "egr.action_editdate AS [action_editdate], "
   lcl_return = lcl_return & "egr.action_externalcomment as [action_externalcomment], "
   lcl_return = lcl_return & "egr.action_citizen as [action_citizen], "
   lcl_return = lcl_return & "egov_users.userfname as [userfname], "
   lcl_return = lcl_return & "egov_users.userlname as [userlname], "
   lcl_return = lcl_return & "egov_users.userfname + ' ' + egov_users.userlname as [user_name], "
   lcl_return = lcl_return & "egr.action_internalcomment as [action_internalcomment], "
   lcl_return = lcl_return & "egov_action_request_view.action_autoid as [action_autoid], "
   lcl_return = lcl_return & "egov_action_request_view.assigned_userid as [assignedemployeeid], "
   lcl_return = lcl_return & "(users.FirstName + ' ' + users.LastName) as [EmployeeSubmitName], "
   lcl_return = lcl_return & "egov_action_request_view.employeesubmitid as [employeesubmitid], "
   lcl_return = lcl_return & "egov_action_request_view.action_form_resolved_status as [action_form_resolved_status] "
   lcl_return = lcl_return & " FROM egov_action_responses egr "
   lcl_return = lcl_return &      " LEFT OUTER JOIN users ON egr.action_userid = users.UserID "
   lcl_return = lcl_return &      " LEFT OUTER JOIN egov_actionline_requests_statuses AS es ON egr.action_sub_status_id = es.action_status_id "
   lcl_return = lcl_return &      " INNER JOIN egov_action_request_view ON egr.action_autoid = egov_action_request_view.action_autoid "
   lcl_return = lcl_return &      " LEFT OUTER JOIN egov_users ON egov_users.userid = egov_action_request_view.userid "
   lcl_return = lcl_return &                 " LEFT OUTER JOIN Groups ON DeptID = GroupID "

  'Format the WHERE clause for this query
   lcl_where_clause = replace(p_WhereClause,"action_autoid","egov_action_request_view.action_autoid")

   lcl_return = lcl_return & lcl_where_clause

  'Build the ORDER BY
   lcl_return = lcl_return & " ORDER BY " & p_order_by

   if p_orderBy = "submit_date" then
      lcl_return = lcl_return & " desc"
   end if

   lcl_return = lcl_return & ", egr.action_editdate DESC"

   buildQueryForDownloadActivityLog = lcl_return

 end function

'------------------------------------------------------------------------------
 function buildQueryForCustomReports(p_reportname, p_WhereClause, p_order_by, p_orderby)
  'p_order_by: columns identified to be used in ORDER BY
  'p_orderby:  field identified by user on how list should be ordered.
   lcl_return = ""

   if UCASE(p_reportname) = "CODESECTIONS" then
      lcl_return = "SELECT "
      lcl_return = lcl_return & "CAST(egov_action_request_view.action_autoid AS varchar) + RIGHT('0'+ CAST(DATEPART(hh, egov_action_request_view.submit_date) AS varchar),2) + RIGHT('0' + CAST(DATEPART(mi, egov_action_request_view.submit_date) AS varchar),2) as TrackingNumber, "
      lcl_return = lcl_return & "submit_date, complete_date, due_date, sub_status_desc, "
      lcl_return = lcl_return & "dbo.fn_buildAddress(streetnumber, streetprefix, streetaddress, streetsuffix, streetdirection) as issuelocation_address, "
      lcl_return = lcl_return & "egov_action_request_view.action_autoid, status, sc.submitted_action_code_id, "
      lcl_return = lcl_return & "isnull(dbo.getCodeSectionName(sc.submitted_action_code_id),'') as code_name, assignedname "
      lcl_return = lcl_return & " FROM egov_action_request_view "
      lcl_return = lcl_return &   " LEFT JOIN egov_submitted_request_code_sections AS sc ON sc.submitted_request_id = egov_action_request_view.action_autoid "

     'Format the WHERE clause for this query
      lcl_where_clause = replace(p_WhereClause,"action_autoid","egov_action_request_view.action_autoid")

      lcl_return = lcl_return & lcl_where_clause

     'Build the ORDER BY
      'lcl_return = lcl_return & " ORDER BY submit_date desc"
      lcl_return = lcl_return & " ORDER BY " & p_order_by

      if UCASE(p_orderBy) = "SUBMIT_DATE" then
         lcl_return = lcl_return & " desc"
      end if

   end if

   buildQueryForCustomReports = lcl_return

 end function

'------------------------------------------------------------------------------
 function valueExistsOrRedQuestionMarks(p_value)
   lcl_return = ""

   if trim(p_value) <> "" then
      lcl_return = p_value
   else
      lcl_return = "<font style=""color:#ff0000;font-weight:bold"">???</font>"
   end if

   valueExistsOrRedQuestionMarks = lcl_return

 end function

'------------------------------------------------------------------------------
 function buildCityStateZip(p_usercity, p_userstate, p_userzip)
   lcl_return = ""

   lcl_user_city  = p_usercity
   lcl_user_state = p_userstate
   lcl_user_zip   = p_userzip

   if lcl_user_city <> "" OR lcl_user_status <> "" OR lcl_user_zip <> "" then

      lcl_user_csz   = lcl_user_city

     'Add the userstate
      if lcl_user_csz = "" then
         lcl_user_csz = lcl_user_state
      else
         if lcl_user_state <> "" then
            lcl_user_csz = lcl_user_csz & " / " & lcl_user_state
         else
            lcl_user_csz = lcl_user_csz & " / -- "
         end if
      end if

     'Add the userzip
      if lcl_user_csz = "" then
         lcl_user_csz = lcl_user_zip
      else
         if lcl_user_zip <> "" then
            lcl_user_csz = lcl_user_csz & " / " & lcl_user_zip
         end if
      end if
   end if

   buildCityStateZip = lcl_return

end function

'------------------------------------------------------------------------------
function AddCommentTaskComment(sInternalMsg, sExternalMsg, sStatus, iRequestID, iUserID, _
                               iOrgID, sSubStatusID, iCitizenID, iCitizenEmail)

  dim lcl_status, lcl_substatusid

  lcl_status      = ""
  lcl_substatusid = 0
  lcl_submit_date = ConvertDateTimetoTimeZone()

  if sStatus <> "" then
     lcl_status = ucase(sStatus)
  end if

  if sSubStatusID = "" OR IsNULL(sSubStatusID) then
     lcl_substatusid = 0
  else
     lcl_substatusid = clng(sSubStatusID)
  end if

 'Set up citizen info. All citizen fields are based off of the Citizen Email.
 'If it is NULL then make all of the other "citizen" fields NULL as well.
 'If it is NOT NULL then the "citizen" fields must be populated.
  lcl_citizen_sentby_id      = "NULL"
  lcl_citizen_sentby_name    = "NULL"
  lcl_citizen_sentto_id      = "NULL"
  lcl_citizen_sentto_email   = "NULL"
  lcl_citizen_emailsent_date = "NULL"

  if iCitizenEmail <> "" then
     lcl_citizen_sentby_id      = iUserID
     lcl_citizen_sentby_name    = getAdminName(iUserID)
     lcl_citizen_sentto_id      = iCitizenID
     lcl_citizen_sentto_email   = iCitizenEmail
     lcl_citizen_emailsent_date = "'" & lcl_submit_date & "'"

     if trim(lcl_citizen_sentby_name) <> "" then
        lcl_citizen_sentby_name = "'" & dbsafe(lcl_citizen_sentby_name) & "'"
     else
        lcl_citizen_sentby_name = "NULL"
     end if

     if trim(lcl_citizen_sentto_email) <> "" then
        lcl_citizen_sentto_email = "'" & dbsafe(lcl_citizen_sentto_email) & "'"
    else
        lcl_citizen_sentto_email = "NULL"
     end if
  end if

		sSQL = "INSERT INTO egov_action_responses ("
		sSQL = sSQL & "action_status, "
		sSQL = sSQL & "action_internalcomment, "
		sSQL = sSQL & "action_externalcomment, "
		sSQL = sSQL & "action_editdate, "
		sSQL = sSQL & "action_userid, "
		sSQL = sSQL & "action_orgid, "
		sSQL = sSQL & "action_autoid, "
		sSQL = sSQL & "action_sub_status_id, "
    sSQL = sSQL & "citizen_sentby_id, "
    sSQL = sSQL & "citizen_sentby_name, "
    sSQL = sSQL & "citizen_sentto_id, "
    sSQL = sSQL & "citizen_sentto_email, "
    sSQL = sSQL & "citizen_emailsent_date "
		sSQL = sSQL & ") VALUES ( "
		sSQL = sSQL & "'" & dbsafe(lcl_status)       & "', "
		sSQL = sSQL & "'" & DBsafe(sInternalMsg)     & "', " 
		sSQL = sSQL & "'" & DBsafe(sExternalMsg)     & "', " 
    sSQL = sSQL & "'" & lcl_submit_date          & "', "
		sSQL = sSQL & "'" & iUserID                  & "', " 
		sSQL = sSQL & "'" & iOrgID                   & "', " 
		sSQL = sSQL & "'" & iRequestID               & "', "
		sSQL = sSQL &       lcl_substatusid          & ", "
		sSQL = sSQL &       lcl_citizen_sentby_id    & ", "
		sSQL = sSQL &       lcl_citizen_sentby_name  & ", "
		sSQL = sSQL &       lcl_citizen_sentto_id    & ", "
		sSQL = sSQL &       lcl_citizen_sentto_email & ", "
		sSQL = sSQL &       lcl_citizen_emailsent_date
    sSQL = sSQL & ")"

    'session("oTaskSQL") = sSql

		set oTaskComment = Server.CreateObject("ADODB.Recordset")
		'oTaskComment.Open sSQL, Application("DSN"), 0, 1
		oTaskComment.Open sSQL, lcl_dsn, 0, 1

  'oTaskComment.close
		set oTaskComment = nothing

	if mobileID <> "0" then UpdateRegistryStatusNote sExternalMsg, mobileID

end function

'------------------------------------------------------------------------------
sub getActivityLogInfo(ByVal iRequestID, ByRef lcl_action_editdate, ByRef lcl_previous_status)

  dim sSQL, lcl_requestid

  lcl_action_editdate = ""
  lcl_previous_status = ""
  lcl_requestid       = 0

  if iRequestID <> "" then
     lcl_requestid = clng(iRequestID)
  end if

  sSQL = "SELECT TOP 1 "
  sSQL = sSQL & " action_editdate, "
  sSQL = sSQL & " upper(action_status) as action_status "
  sSQL = sSQL & " FROM egov_action_responses "
  sSQL = sSQL & " WHERE action_autoid = " & lcl_requestid
  sSQL = sSQL & " ORDER BY action_editdate desc "

		set oGetActivityLogInfo = Server.CreateObject("ADODB.Recordset")
		oGetActivityLogInfo.Open sSQL, lcl_dsn, 0, 1

  if not oGetActivityLogInfo.eof then
     lcl_action_editdate = oGetActivityLogInfo("action_editdate")
     lcl_previous_status = oGetActivityLogInfo("action_status")
  end if

  oGetActivityLogInfo.close
  set oGetActivityLogInfo = nothing

end sub

'------------------------------------------------------------------------------
sub getPreviousActionLineInfo(ByVal iRequestID, ByRef lcl_previous_status, ByRef lcl_previous_complete_date)

  dim sSQL, lcl_requestid

  lcl_requestid              = 0
  lcl_previous_status        = ""
  lcl_previous_complete_date = ""

  if iRequestID <> "" then
     lcl_requestid = clng(iRequestID)
  end if

  if lcl_requestid > 0 then
     sSQL = "SELECT status, "
     sSQL = sSQL & " complete_date "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE action_autoid = " & lcl_requestid

   		set oGetPreviousActionLineInfo = Server.CreateObject("ADODB.Recordset")
   		oGetPreviousActionLineInfo.Open sSQL, lcl_dsn, 0, 1

     if not oGetPreviousActionLineInfo.eof then
        if oGetPreviousActionLineInfo("status") <> "" then
           lcl_previous_status = ucase(oGetPreviousActionLineInfo("status"))
        end if

        lcl_previous_complete_date = oGetPreviousActionLineInfo("complete_date")
     end if

     oGetPreviousActionLineInfo.close
     set oGetPreviousActionLineInfo = nothing

  end if

end sub

'------------------------------------------------------------------------------
 sub displayStatusCheckbox(iCheckBoxName,iIsChecked)

    if iCheckBoxName <> "" then
      'Set up the NAME and ID values for the checkbox field.
       lcl_checkboxname = UCASE(iCheckBoxName)
       lcl_checkboxname = replace(lcl_checkboxname," ","")

       response.write "<input type=""checkbox"" name=""status" & lcl_checkboxname & """ id=""status" & lcl_checkboxname & """ value=""yes""" & iIsChecked & " />" & iCheckBoxName & vbcrlf
    end if

    'response.write "<input type=""checkbox"" name=""statusSUBMITTED"" id=""selStatusSUBMITTED"" value=""yes"""   & check1 & " />Submitted"   & vbcrlf
    'response.write "<input type=""checkbox"" name=""statusINPROGRESS"" id=""selStatusINPROGRESS"" value=""yes""" & check2 & " />In Progress" & vbcrlf
    'response.write "<input type=""checkbox"" name=""statusWAITING"" id=""selStatusWAITING"" value=""yes"""       & check3 & " />Waiting"     & vbcrlf
    'response.write "<input type=""checkbox"" name=""statusRESOLVED"" id=""selStatusRESOLVED"" value=""yes"""     & check4 & " />Resolved"    & vbcrlf
    'response.write "<input type=""checkbox"" name=""statusDISMISSED"" id=""selStatusDISMISSED"" value=""yes"""   & check5 & " />Dismissed"   & vbcrlf
 end sub

'------------------------------------------------------------------------------
 sub getPermissionLevels(ByVal iLevel, ByRef blnCanViewAllActionItems, ByRef blnCanViewOwnActionItems, ByRef blnCanViewDeptActionItems)

 'Get their view permission level
  iPermissionLevelId = GetUserPermissionLevel(session("userid"),"requests")

  if clng(iPermissionLevelId) > 0 then
    	sPermissionLevel = GetPermissionLevelName(iPermissionLevelId)
  else
    	response.redirect iLevel & "permissiondenied.asp"
  end if

 'Set to use new permission levels
  blnCanViewAllActionItems  = False  
  blnCanViewOwnActionItems  = False 
  blnCanViewDeptActionItems = False

 'Override the flags for what they get
  Select Case sPermissionLevel
	   Case "View All - Edit All"
    		blnCanViewAllActionItems  = True
    		blnCanViewOwnActionItems  = True
    		blnCanViewDeptActionItems = True
  	 Case "View All - Edit Dept"
   	 	blnCanViewAllActionItems  = True
   		 blnCanViewOwnActionItems  = True
    		blnCanViewDeptActionItems = True
   	Case "View All - Edit Own"
    		blnCanViewAllActionItems  = True
    		blnCanViewOwnActionItems  = True
   	 	blnCanViewDeptActionItems = True
     '----------------------------------
   	Case "View Dept - Edit Dept"
    		blnCanViewOwnActionItems  = True
   	 	blnCanViewDeptActionItems = True
   	Case "View Dept - Edit Own"
    		blnCanViewOwnActionItems  = True
    		blnCanViewDeptActionItems = True
  	 Case "View Own - Edit Own"
   	 	blnCanViewOwnActionItems  = True
   	Case "View All"
    		blnCanViewAllActionItems  = True
    		blnCanViewOwnActionItems  = True
    		blnCanViewDeptActionItems = True
  End Select

end sub

'------------------------------------------------------------------------------
 sub displaySubStatusOptions(iUserHasPermission_ActionLineSubStatus, iSubStatusHidden, iShowHideSubStatus)
  if iUserHasPermission_ActionLineSubStatus then
	    'Cycle through each main status and determine if there are any active sub-statuses.
	    'Retrieve all of the sub-statuses for the organization for each parent_status
      sSQL1 = "SELECT s1.action_status_id, "
      sSQL1 = sSQL1 & " s1.status_name, "
      sSQL1 = sSQL1 & " s1.orgid, "
      sSQL1 = sSQL1 & " s1.parent_status, "
      sSQL1 = sSQL1 & " s1.display_order, "
      sSQL1 = sSQL1 & " s1.active_flag, "
 					sSQL1 = sSQL1 & " s2.action_status_id AS parent_status_id, "
      sSQL1 = sSQL1 & " s2.display_order AS parent_display_order "
 					sSQL1 = sSQL1 & " FROM egov_actionline_requests_statuses s1, "
      sSQL1 = sSQL1 &      " egov_actionline_requests_statuses s2"
 					sSQL1 = sSQL1 & " WHERE s1.parent_status = s2.status_name "
 					sSQL1 = sSQL1 & " AND s1.active_flag = 'Y' "
 					sSQL1 = sSQL1 & " AND s2.active_flag = 'Y' "
 					sSQL1 = sSQL1 & " AND s1.orgid = " & session("orgid")
 					sSQL1 = sSQL1 & " ORDER BY 8, s1.parent_status, s1.display_order, s1.status_name "

	 				set oExists = Server.CreateObject("ADODB.Recordset")
      'oExists.Open sSQL1, Application("DSN"), 0, 1
      oExists.Open sSQL1, lcl_dsn, 0, 1

      i = 0
			 		if not oExists.eof then

         response.write "					<table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""sub_status_row"">" & vbcrlf
         response.write "       <tr>" & vbcrlf
         response.write "           <td>" & vbcrlf
         response.write "               <p>" & vbcrlf
         response.write "                 &nbsp;&nbsp;&nbsp;<strong>Sub-Status:&nbsp;</strong>" & vbcrlf
         response.write "                 <input type=""hidden"" name=""substatus_hidden"" id=""substatus_hidden"" value=""" & iSubStatusHidden & """ />" & vbcrlf
         response.write "                 <input type=""hidden"" name=""show_hide_substatus"" id=""show_hide_substatus"" value=""" & iShowHideSubStatus & """ />" & vbcrlf
         response.write "                 <span id=""display_substatus""></span>" & vbcrlf
         response.write "               </p>" & vbcrlf
         'response.write "               <input type=""button"" name=""showHideSubStatusButton"" id=""showHideSubStatusButton"" value=""Show/Hide Sub-Status List"" class=""button"" onclick=""showhide_substatus_criteria();"" />" & vbcrlf
         response.write "               <input type=""button"" name=""showHideSubStatusButton"" id=""showHideSubStatusButton"" value=""Show/Hide Sub-Status List"" class=""button ui-button ui-widget ui-corner-all"" />" & vbcrlf
         response.write "           </td>" & vbcrlf
         response.write "       </tr>" & vbcrlf
         response.write "       <tr valign=""top"">" & vbcrlf
         response.write "					      <td>" & vbcrlf
         response.write "      							  <span id=""selectSubStatus"" align=""center"">" & vbcrlf
         response.write "               <table border=""0"" cellspacing=""1"" cellpadding=""0"" bgcolor=""#000000"">" & vbcrlf
         response.write "       								  <tr>" & vbcrlf
         response.write "                     <td>" & vbcrlf
         response.write "                 								<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#c0c0c0"">" & vbcrlf
         response.write "                 								  <tr valign=""top"">" & vbcrlf

         lcl_parent_status = ""
					    lcl_line_count    = 0

					   'Loop through all of the Sub-Statuses
					    while not oExists.eof
           lcl_line_count = lcl_line_count + 1
						     i = i + 1

           if instr(iSubStatusHidden,"(" & oExists("action_status_id") & ")") > 0 then
              lcl_click = " checked=""checked"" "
           else
              lcl_click = ""
  						   end if

           if lcl_line_count = 1 then
						        lcl_parent_status = oExists("parent_status")

              response.write "                 								      <td>" & vbcrlf
              response.write "                 								          <table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#ffffff"">" & vbcrlf
              response.write "                 								            <tr bgcolor=""#efefef"">" & vbcrlf
              response.write "                 								                <td colspan=""2""><strong>" & oExists("parent_status") & "</strong></td>" & vbcrlf
              response.write "                 								            </tr>" & vbcrlf
              response.write "                 								            <tr>" & vbcrlf
              response.write "                 								                <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
              response.write "                 								                <td>" & oExists("status_name") & "</td>" & vbcrlf
              response.write "                 								            </tr>" & vbcrlf

           else
   					      if UCASE(lcl_parent_status) <> UCASE(oExists("parent_status")) then
     							     if lcl_line_count > 1 then
					         			   lcl_line_count = 1
                    lcl_parent_status = oExists("parent_status")

                    response.write "                 								</table>" & vbcrlf
                    response.write "                 				</td>" & vbcrlf
                    response.write "                 				<td>" & vbcrlf
                    response.write "                 								<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#ffffff"">" & vbcrlf
                    response.write "                 								  <tr bgcolor=""#efefef"">" & vbcrlf
                    response.write "                 								      <td colspan=""2""><strong>" & oExists("parent_status") & "</strong></td>" & vbcrlf
                    response.write "                 								  </tr>" & vbcrlf
                    response.write "                 								  <tr>" & vbcrlf
                    response.write "                 								      <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
                    response.write "                 								      <td>" & oExists("status_name") & "</td>" & vbcrlf
                    response.write "                 								  </tr>" & vbcrlf
				             end if
              else
                 lcl_parent_status = lcl_parent_status

                 response.write "                 								  <tr>" & vbcrlf
                 response.write "      											              <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
                 response.write "     								                  <td>" & oExists("status_name") & "</td>" & vbcrlf
                 response.write "      											          </tr>" & vbcrlf

              end if
           end if
		 				    oExists.movenext
			 		   wend

         response.write "                    </table>" & vbcrlf
         response.write "						          </td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "          </table>" & vbcrlf
         response.write "		    </td>" & vbcrlf
         response.write "  </tr>" & vbcrlf
         response.write "</table>" & vbcrlf
         response.write "</span>" & vbcrlf

      else

         response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""sub_status_row"">" & vbcrlf
         response.write "  <tr>" & vbcrlf
         response.write "      <td>&nbsp;&nbsp;&nbsp;<strong>Sub-Status:&nbsp;</strong>No Sub-Statuses"
         response.write "          <span id=""selectSubStatus""></span>" & vbcrlf
         response.write "          <input type=""hidden"" name=""substatus_hidden"" id=""substatus_hidden"" value=""" & iSubStatusHidden & """ />" & vbcrlf
         response.write "          <input type=""hidden"" name=""show_hide_substatus"" id=""show_hide_substatus"" value=""" & iShowHideSubStatus & """ />" & vbcrlf
         response.write "      </td>" & vbcrlf
         response.write "  </tr>" & vbcrlf
      end if

      oExists.close
      set oExists = nothing

      response.write "  </td>" & vbcrlf
      response.write "</tr>" & vbcrlf
      response.write "</table>" & vbcrlf
  end if

 end sub

'------------------------------------------------------------------------------
 sub displaySubStatusOptions_new(iUserHasPermission_ActionLineSubStatus, iSubStatusHidden, iShowHideSubStatus)
  if iUserHasPermission_ActionLineSubStatus then
	    'Cycle through each main status and determine if there are any active sub-statuses.
	    'Retrieve all of the sub-statuses for the organization for each parent_status
      sSQL1 = "SELECT s1.action_status_id, "
      sSQL1 = sSQL1 & " s1.status_name, "
      sSQL1 = sSQL1 & " s1.orgid, "
      sSQL1 = sSQL1 & " s1.parent_status, "
      sSQL1 = sSQL1 & " s1.display_order, "
      sSQL1 = sSQL1 & " s1.active_flag, "
 					sSQL1 = sSQL1 & " s2.action_status_id AS parent_status_id, "
      sSQL1 = sSQL1 & " s2.display_order AS parent_display_order "
 					sSQL1 = sSQL1 & " FROM egov_actionline_requests_statuses s1, "
      sSQL1 = sSQL1 &      " egov_actionline_requests_statuses s2"
 					sSQL1 = sSQL1 & " WHERE s1.parent_status = s2.status_name "
 					sSQL1 = sSQL1 & " AND s1.active_flag = 'Y' "
 					sSQL1 = sSQL1 & " AND s2.active_flag = 'Y' "
 					sSQL1 = sSQL1 & " AND s1.orgid = " & session("orgid")
 					sSQL1 = sSQL1 & " ORDER BY 8, s1.parent_status, s1.display_order, s1.status_name "

	 				set oExists = Server.CreateObject("ADODB.Recordset")
      'oExists.Open sSQL1, Application("DSN"), 0, 1
      oExists.Open sSQL1, lcl_dsn, 0, 1

      i = 0

         response.write "					<table border=""0"" cellspacing=""0"" width=""100%"" cellpadding=""2"" id=""sub_status_row"">" & vbcrlf
         response.write "       <tr>" & vbcrlf
         response.write "           <td>" & vbcrlf
         response.write "                 <input type=""hidden"" name=""substatus_hidden"" id=""substatus_hidden"" value=""" & iSubStatusHidden & """ />" & vbcrlf
         response.write "                 <input type=""hidden"" name=""show_hide_substatus"" id=""show_hide_substatus"" value=""" & iShowHideSubStatus & """ />" & vbcrlf
	 response.write "		<div id=""accordion2"">"
	 response.write "		<h3><strong>Sub-Status</strong>&nbsp;&nbsp;&nbsp;<span id=""display_substatus""></span></h3>"
	 response.write "		<div>"
	if not oExists.eof then
         response.write "               <table border=""0"" cellspacing=""1"" cellpadding=""0"">" & vbcrlf
         response.write "       								  <tr>" & vbcrlf
         response.write "                     <td>" & vbcrlf
         response.write "                 								<table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
         response.write "                 								  <tr valign=""top"">" & vbcrlf

         lcl_parent_status = ""
					    lcl_line_count    = 0

					   'Loop through all of the Sub-Statuses
					    while not oExists.eof
           lcl_line_count = lcl_line_count + 1
						     i = i + 1

           if instr(iSubStatusHidden,"(" & oExists("action_status_id") & ")") > 0 then
              lcl_click = " checked=""checked"" "
           else
              lcl_click = ""
  						   end if

           if lcl_line_count = 1 then
						        lcl_parent_status = oExists("parent_status")

              response.write "                 								      <td>" & vbcrlf
              response.write "                 								          <table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#ffffff"">" & vbcrlf
              response.write "                 								            <tr>" & vbcrlf
              response.write "                 								                <td colspan=""2""><strong>" & oExists("parent_status") & "</strong></td>" & vbcrlf
              response.write "                 								            </tr>" & vbcrlf
              response.write "                 								            <tr>" & vbcrlf
              response.write "                 								                <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
              response.write "                 								                <td>" & oExists("status_name") & "</td>" & vbcrlf
              response.write "                 								            </tr>" & vbcrlf

           else
   					      if UCASE(lcl_parent_status) <> UCASE(oExists("parent_status")) then
     							     if lcl_line_count > 1 then
					         			   lcl_line_count = 1
                    lcl_parent_status = oExists("parent_status")

                    response.write "                 								</table>" & vbcrlf
                    response.write "                 				</td>" & vbcrlf
                    response.write "                 				<td>" & vbcrlf
                    response.write "                 								<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#ffffff"">" & vbcrlf
                    response.write "                 								  <tr>" & vbcrlf
                    response.write "                 								      <td colspan=""2""><strong>" & oExists("parent_status") & "</strong></td>" & vbcrlf
                    response.write "                 								  </tr>" & vbcrlf
                    response.write "                 								  <tr>" & vbcrlf
                    response.write "                 								      <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
                    response.write "                 								      <td>" & oExists("status_name") & "</td>" & vbcrlf
                    response.write "                 								  </tr>" & vbcrlf
				             end if
              else
                 lcl_parent_status = lcl_parent_status

                 response.write "                 								  <tr>" & vbcrlf
                 response.write "      											              <td><input type=""checkbox"" name=""substatus"" id=""SS_" & oExists("action_status_id") & """ value=""" & oExists("action_status_id") & """ onclick=""javascript:change_substatus_filter();""" & lcl_click & " /></td>" & vbcrlf
                 response.write "     								                  <td>" & oExists("status_name") & "</td>" & vbcrlf
                 response.write "      											          </tr>" & vbcrlf

              end if
           end if
		 				    oExists.movenext
			 		   wend

         response.write "                    </table>" & vbcrlf
         response.write "						          </td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "          </table>" & vbcrlf
      else

         response.write "<table border=""0"" cellspacing=""0"" width=""100%"" cellpadding=""2"" id=""sub_status_row"">" & vbcrlf
         response.write "  <tr>" & vbcrlf
         response.write "      <td>No Sub-Statuses"
         response.write "      </td>" & vbcrlf
         response.write "  </tr>" & vbcrlf
      end if
	 response.write "		</div>"
	 response.write "		</div>"
         response.write "           </td>" & vbcrlf
         response.write "       </tr>" & vbcrlf
         response.write "</table>" & vbcrlf
         response.write "</span>" & vbcrlf


      oExists.close
      set oExists = nothing

      response.write "  </td>" & vbcrlf
      response.write "</tr>" & vbcrlf
      response.write "</table>" & vbcrlf
  end if

 end sub

'------------------------------------------------------------------------------
Function GetGroups(iUserID)

	sSQL = "SELECT Groups.GroupID, "
 sSQL = sSQL & " Users.OrgID, "
 sSQL = sSQL & " Groups.GroupName, "
 sSQL = sSQL & " Groups.GroupDescription "
 sSQL = sSQL & " FROM Users "
 sSQL = sSQL &      " INNER JOIN UsersGroups ON Users.UserID = UsersGroups.UserID "
 sSQL = sSQL &      " INNER JOIN Groups ON UsersGroups.GroupID = Groups.GroupID "
 sSQL = sSQL & " WHERE (Groups.GroupType = 2) "
 sSQL = sSQL & " AND (Users.OrgID = "  & Session("OrgID")  & ") "
 sSQL = sSQL & " AND (Users.UserID = " & iUserID & ") "
 sSQL = sSQL & " ORDER BY Groups.GroupName "

	set oDepts = Server.CreateObject("ADODB.Recordset")
	'oDepts.Open sSQL, Application("DSN") , 0, 1
	oDepts.Open sSQL, lcl_dsn , 0, 1

	If NOT oDepts.EOF Then

  		do while not oDepts.EOF
		    	sReturnValue = sReturnValue & "'" & oDepts("GroupID") & "',"
    			oDepts.MoveNext
  		loop

  		sReturnValue = LEFT(sReturnValue,LEN(sReturnValue)-1) 

 else
    sReturnValue = 0
	End If 
			
 oDepts.close
	Set oDepts = Nothing

	GetGroups = sReturnValue

End Function

'------------------------------------------------------------------------------
Function fnListDepts(iSelectedDeptID)
 'SET SELECTED DEPARTMENT ID
  If iSelectedDeptID = "" Then
    	iSelectedDeptID = 0
  End If

 'Get all of the departments for the org
  sSQLd = "SELECT distinct deptid "
  sSQLd = sSQLd & " FROM egov_action_request_view "
  sSQLd = sSQLd & " WHERE orgid = " & session("orgid")
  sSQLd = sSQLd & " ORDER BY 1 "

  set rsd = Server.CreateObject("ADODB.Recordset")
  'rsd.Open sSQLd, Application("DSN") , 0, 1
  rsd.Open sSQLd, lcl_dsn , 0, 1

  if not rsd.eof then
     while not rsd.eof
        if lcl_depts <> "" then
           lcl_depts = lcl_depts & "," & rsd("deptid")
        else
           lcl_depts = rsd("deptid")
        end if
        rsd.movenext
     wend

    'GET LIST OF DEPARTMENTS
    	sSQL = "SELECT groupid, orgid, groupname, groupdescription, isInactive "
     sSQL = sSQL & " FROM groups "
     sSQL = sSQL & " WHERE grouptype = 2 "
     sSQL = sSQL & " AND orgid = " & Session("OrgID")
     sSQL = sSQL & " AND groupid IN (" & lcl_depts & ") "

     If blnCanViewAllActionItems Then
   	   'GET ALL DEPARTMENTS
     Else
  	    'GET DEPARTMENTS ASSIGNED TO CURRENTLY LOGGED ON ADMIN USER
        sSQL = sSQL & " AND groupid IN (" & GetGroups(session("userid")) & ") "
     End If

     sSQL = sSQL & " ORDER BY isInactive, groupname "

     Set oDepts = Server.CreateObject("ADODB.Recordset")
     'oDepts.Open sSQL, Application("DSN") , 0, 1
     oDepts.Open sSQL, lcl_dsn , 0, 1

    'LOOP THRU GROUPS DISPLAYING AS NECESSARY
     if not oDepts.eof then
       	while not oDepts.EOF
	     	   'SET SELECTED DEPARTMENT
         		If IsNumeric(iSelectedDeptID) Then
			           If clng(iSelectedDeptID) = oDepts("groupid") Then
         			   	'SET SELECTED FLAG
          		   		sSelected = " selected" 
           			Else 
              		'CLEAR SELECTED FLAG
             				sSelected = ""
           			End If
         		End If

           if oDepts("isInactive") then
              lcl_isInactive = " [inactive]"
           else
              lcl_isInactive = ""
           end if

        		'DISPLAY GROUP AS OPTION
         		response.write "  <option " & sSelected  & " value=""" & oDepts("groupid") & """>" & oDepts("groupname") & lcl_isInactive & "</option>" & vbcrlf

         		oDepts.MoveNext
        wend
     end if

    'CLEAN UP OBJECTS
     oDepts.close
    	Set oDepts = Nothing

  end if

  set rsd = nothing

End Function

'------------------------------------------------------------------------------
Function fnListForms(iSelectFormID)
	sLastCategory = "NONE_START"

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM dbo.egov_formlist "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " ORDER BY form_category_sequence, action_form_name "

	set oForms = Server.CreateObject("ADODB.Recordset")
	'oForms.Open sSQL, Application("DSN") , 0, 1
	oForms.Open sSQL, lcl_dsn , 0, 1

	if not oForms.eof then
		
		  while not oForms.eof

				   sCurrentCategory = oForms("form_category_name")

   				if sLastCategory = "NONE_START" then
     					if iSelectFormID = "C" & oForms("form_category_id") & "" then
       						selectA = "selected"
     					else
					       	selectA = ""
      				end if

     					response.write "  <option value=""C" & oForms("form_category_id") & """ " & selectA & ">----Category: " & sCurrentCategory & "</option>" & vbcrlf
       end if

   				if iSelectFormID = "C" & oForms("form_category_id") & "" then
			     		selectA = "selected"
   				else
			     		selectA = ""
   				end if

   				if (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") then
      				response.write "  <option value=""C" & oForms("form_category_id") & """ " & selectA & ">----Category: " & sCurrentCategory &  "</option>" & vbcrlf
   				end if

   				if cStr(selectFormId)=cStr(oForms("action_form_id")) then 
			     		selectA = "selected"
   				else
			     		selectA = ""
   				end if

   				response.write "  <option value=""" & oForms("action_form_id") & """ " & selectA & ">" & oForms("action_form_name") &  "</option>" & vbcrlf

    			oForms.movenext
    			sLastCategory = sCurrentCategory
    wend
	end if

 oForms.close
	set oForms = nothing
	end function

'------------------------------------------------------------------------------
Sub DrawAssignedEmployeeSelection(iSelectAssignedTo)

	sSQLassignedto = "SELECT FirstName + ' ' + LastName as assigned_Name, "
 sSQLassignedto = sSQLassignedto & " UserID, "
 sSQLassignedto = sSQLassignedto & " isdeleted "
 sSQLassignedto = sSQLassignedto & " FROM USERS "
 sSQLassignedto = sSQLassignedto & " WHERE OrgID = " & Session("OrgID")
 sSQLassignedto = sSQLassignedto & " ORDER BY isdeleted, LastName, FirstName"

	set oAssigned = Server.CreateObject("ADODB.Recordset")
	'oAssigned.Open sSQLassignedto, Application("DSN"), 0, 1
	oAssigned.Open sSQLassignedto, lcl_dsn, 0, 1

'IF THERE ARE ASSIGNED USERS THEN LIST
	If NOT oAssigned.EOF Then
	
	'BEGIN SELECTION BOX
		response.write "<select name=""selectAssignedto"" id=""selectAssignedto"">" & vbcrlf
  response.write "  <option value=""all"">Anyone</option>" & vbcrlf
		
	'LOOP THRU ASSIGNED USERS
		while not oAssigned.eof

 			'SET SELECT BOX TO DISPLAY CURRENTLY SELECTED NAME
   		selectAssign   = ""
     sOptionStyle   = ""
     sAssignedValue = oAssigned("assigned_Name")

     if iSelectAssignedTo <> "all" then
     			if clng(iSelectAssignedTo) = clng(oAssigned("userid")) then
		       		selectAssign = "selected"
     			end if
     end if

    'Is the person deleted?
     if oAssigned("isdeleted") then
        sAssignedValue = "[" & sAssignedValue & " - Inactive]"
        sOptionStyle   = " style=""color:#ff0000;"""

     end if

 			'DISPLAY ASSIGNED EMPLOYEE AS OPTION
  			response.write "  <option value=""" & oAssigned("userid") & """ " & selectAssign & sOptionStyle & ">" & sAssignedValue & "</option>" & vbcrlf
		  	oAssigned.MoveNext
		wend

		response.write "</select>&nbsp;&nbsp;&nbsp;" & vbcrlf
		
		oAssigned.Close
	
	End If

	Set oAssigned = Nothing

End Sub

'------------------------------------------------------------------------------
function formatTitle(iValue)
  lcl_return = "<font style=""color:#ff0000; font-weight:bold;"">???</font>"

  if trim(iValue) <> "" then
     lcl_return = iValue
		end if

  formatTitle = lcl_return

end function

'------------------------------------------------------------------------------
function checkForNullSetToALL(iValue)

  lcl_return = iValue

  if iValue = "" OR IsNull(iValue) OR iValue = "0" then
     lcl_return = "all"
  end if

  checkForNullSetToAll = lcl_return

end function

'------------------------------------------------------------------------------
sub displayCustomSearchOptions(iCustomReportID)

  lcl_button_width       = "300px"
  lcl_userSearchDefaults = "System Options"

  if iCustomReportID <> "" then
     sSQL = "SELECT isUserDefault "
     sSQL = sSQL & " FROM egov_customreports "
     sSQL = sSQL & " WHERE customreportid = " & iCustomReportID

     set oUserDefault = Server.CreateObject("ADODB.Recordset")
     'oUserDefault.Open sSQL, Application("DSN"), 1, 3
     oUserDefault.Open sSQL, lcl_dsn, 1, 3

     if not oUserDefault.eof then
        if oUserDefault("isUserDefault") then
           lcl_userSearchDefaults = "My Saved Options"
        end if
     end if

     oUserDefault.close
     set oUserDefault = nothing
  end if

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Default Search Options&nbsp;" & vbcrlf
                      displayHelpIcon("actionline_savesearchoptions")
  response.write "    &nbsp;</legend>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-top:5px"">" & vbcrlf
  response.write "    <tr valign=""top"">" & vbcrlf
  response.write "        <td align=""center"" nowrap=""nowrap"">" & vbcrlf
  response.write "            Currently using<br />" & vbcrlf
  response.write "            <span id=""customSearchDisplay"" class=""redText"">" & lcl_userSearchDefaults & "</span>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "        <td align=""right""><input type=""button"" value=""Run Default Search"" class=""button ui-button ui-widget ui-corner-all"" onclick=""location.href='action_line_list.asp?init=Y';"" /></td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  response.write "    <tr><td colspan=""2""><input type=""button"" name=""useMyDefaults"" id=""useMyDefaults"" class=""button ui-button ui-widget ui-corner-all"" style=""width:" & lcl_button_width & """ value=""Set Current Search Options as Default"" onclick=""updateCustomReport('" & iCustomReportID & "','USER','Y');"" /></td></tr>" & vbcrlf
  response.write "    <tr><td colspan=""2"" style=""padding-top:5px""><input type=""button"" name=""useSystemDefaults"" id=""useSystemDefaults"" class=""button ui-button ui-widget ui-corner-all"" style=""width:" & lcl_button_width & """ value=""Set System Search Options as Default"" onclick=""updateCustomReport('" & iCustomReportID & "','SYSTEM','N');"" /></td></tr>" & vbcrlf
  'response.write "    <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  'response.write "    <tr><td colspan=""2""><input type=""button"" name=""saveSearchOptionsButton"" id=""saveSearchOptionsButton"" class=""button"" style=""width:" & lcl_button_width & """ value=""Save Current Search Options"" onclick=""saveSearchOptions('" & iCustomReportID & "');"" /></td></tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
  'response.write "<input type=""checkbox"" name=""setSearchOptionsAsDefault"" id=""setSearchOptionsAsDefault"" value=""on""" & lcl_userSearchDefaults & " onclick=""enableDisableSaveSearchButton();updateCustomReport('" & lcl_customreportid_actionline_user & "');"" />Use my saved search options as defaults<br />" & vbcrlf
  response.write "<input type=""hidden"" name=""userReportName"" id=""userReportName"" value=""ActionLine - User Saved Search Options"" size=""10"" maxlength=""200"" />" & vbcrlf
  'response.write "                      <input type=""button"" name=""saveSearchOptionsButton"" id=""saveSearchOptionsButton"" class=""button"" value=""Save Search Options"" onclick=""updateCustomReport('" & lcl_customreportid_actionline_user & "');saveSearchOptions('" & lcl_customreportid_actionline_user & "');"" /><br />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayCustomSearchOptions_new(iCustomReportID)

  lcl_button_width       = "300px"
  lcl_userSearchDefaults = "System Options"

  if iCustomReportID <> "" then
     sSQL = "SELECT isUserDefault "
     sSQL = sSQL & " FROM egov_customreports "
     sSQL = sSQL & " WHERE customreportid = " & iCustomReportID

     set oUserDefault = Server.CreateObject("ADODB.Recordset")
     'oUserDefault.Open sSQL, Application("DSN"), 1, 3
     oUserDefault.Open sSQL, lcl_dsn, 1, 3

     if not oUserDefault.eof then
        if oUserDefault("isUserDefault") then
           lcl_userSearchDefaults = "My Saved Options"
        end if
     end if

     oUserDefault.close
     set oUserDefault = nothing
  end if

  response.write "<div id=""defaultsearchaccord"">" & vbcrlf
  response.write "  <h3>Default Search Options&nbsp;" & vbcrlf
                      'displayHelpIcon("actionline_savesearchoptions")
  response.write "    &nbsp;</h3>" & vbcrlf
  response.write "<div>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-top:5px"">" & vbcrlf
  response.write "    <tr valign=""top"">" & vbcrlf
  response.write "        <td align=""center"" nowrap=""nowrap"">" & vbcrlf
  response.write "            Currently using<br />" & vbcrlf
  response.write "            <span id=""customSearchDisplay"" class=""redText"">" & lcl_userSearchDefaults & "</span>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "        <td align=""right""><input type=""button"" value=""Run Default Search"" class=""button ui-button ui-widget ui-corner-all"" onclick=""location.href='action_line_list.asp?init=Y';"" /></td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  response.write "    <tr><td colspan=""2""><input type=""button"" name=""useMyDefaults"" id=""useMyDefaults"" class=""button ui-button ui-widget ui-corner-all"" style=""width:" & lcl_button_width & """ value=""Set Current Search Options as Default"" onclick=""updateCustomReport('" & iCustomReportID & "','USER','Y');"" /></td></tr>" & vbcrlf
  response.write "    <tr><td colspan=""2"" style=""padding-top:5px""><input type=""button"" name=""useSystemDefaults"" id=""useSystemDefaults"" class=""button ui-button ui-widget ui-corner-all"" style=""width:" & lcl_button_width & """ value=""Set System Search Options as Default"" onclick=""updateCustomReport('" & iCustomReportID & "','SYSTEM','N');"" /></td></tr>" & vbcrlf
  'response.write "    <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  'response.write "    <tr><td colspan=""2""><input type=""button"" name=""saveSearchOptionsButton"" id=""saveSearchOptionsButton"" class=""button"" style=""width:" & lcl_button_width & """ value=""Save Current Search Options"" onclick=""saveSearchOptions('" & iCustomReportID & "');"" /></td></tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</div>" & vbcrlf
  'response.write "<input type=""checkbox"" name=""setSearchOptionsAsDefault"" id=""setSearchOptionsAsDefault"" value=""on""" & lcl_userSearchDefaults & " onclick=""enableDisableSaveSearchButton();updateCustomReport('" & lcl_customreportid_actionline_user & "');"" />Use my saved search options as defaults<br />" & vbcrlf
  response.write "<input type=""hidden"" name=""userReportName"" id=""userReportName"" value=""ActionLine - User Saved Search Options"" size=""10"" maxlength=""200"" />" & vbcrlf
  'response.write "                      <input type=""button"" name=""saveSearchOptionsButton"" id=""saveSearchOptionsButton"" class=""button"" value=""Save Search Options"" onclick=""updateCustomReport('" & lcl_customreportid_actionline_user & "');saveSearchOptions('" & lcl_customreportid_actionline_user & "');"" /><br />" & vbcrlf

end sub

'------------------------------------------------------------------------------
'sub subDisplayAttachments(iScreenType, _
'                          p_requestid, _
'                          p_status, _
'                          iOrgHasFeature_SecureAttachments, _
'                          iUserHasPermission_SecureAttachments, _
'                          iOrgHasFeature_DisplayAttachmentsToPublic, _
'                          iUserHasPermission_DisplayAttachmentsToPublic, _
'                          iIsMobile, _
'                          iFormName, _
'                          iFormPostDirectoryLevel)

sub subDisplayAttachments(iScreenType, _
                          p_requestid, _
                          p_status, _
                          iOrgHasFeature_SecureAttachments, _
                          iUserHasPermission_SecureAttachments, _
                          iOrgHasFeature_DisplayAttachmentsToPublic, _
                          iUserHasPermission_DisplayAttachmentsToPublic, _
                          iFormName, _
                          iFormPostDirectoryLevel)

'N = New ActionLine Request
'E = Existing ActionLine Request
 lcl_screenType = "E"

 if iScreenType <> "" then
    lcl_screenType = ucase(iScreenType)
 end if

 'if iScreenType = "" then
 '   lcl_screenType = "E"
 'else
 '   lcl_screenType = ucase(iScreenType)
 'end if

'Is this the "Mobile Options" section?
 'sIsMobile = "N"

 'if iIsMobile <> "" then
 '   if not containsApostrophe(iIsMobile) then
 '      sIsMobile = ucase(iIsMobile)
 '   end if
 'end if

'Determine the directory level to post the form to
 lcl_formpost_directorylevel = ""

 if iFormPostDirectoryLevel <> "" then
    if not containsApostrophe(iFormPostDirectoryLevel) then
       lcl_formpost_directorylevel = iFormPostDirectoryLevel
    end if
 end if

'If this section shows up in multiple places on the same screen (i.e. action_respond.asp) then we need to differentiate the form name.
'*** "frmAddAttachment" must ALWAYS be the initial form name.  We just concatenate on to it. ***
 lcl_formName = "frmAddAttachment"

 if iFormName <> "" then
    lcl_formName = lcl_formName & iFormName
 end if

	response.write "<form name=""" & lcl_formName & """ id=""" & lcl_formName & """ action=""" & lcl_formpost_directorylevel & "attachment_save.asp"" method=""POST"" enctype=""multipart/form-data"">" & vbcrlf
	response.write "  <input type=""hidden"" name=""status"" id=""status"" value="""         & p_status       & """ />" & vbcrlf
	response.write "  <input type=""hidden"" name=""irequestid"" id=""irequestid"" value=""" & p_requestid    & """ />" & vbcrlf
	response.write "  <input type=""hidden"" name=""screentype"" id=""screentype"" value=""" & lcl_screenType & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""attachmentFormName"" id=""attachmentFormName"" value=""" & lcl_formName & """ />" & vbcrlf
 'response.write "  <input type=""hidden"" name=""isMobilePic"" id=""isMobilePic"" value=""" & sIsMobile & """ />" & vbcrlf

 if not iOrgHasFeature_SecureAttachments AND not iUserHasPermission_SecureAttachments then
    response.write "<input type=""hidden"" name=""attachmentIsSecure"" id=""attachmentIsSecure"" value=""off"" checked=""checked"" />" & vbcrlf
 end if

 if lcl_screenType = "N" then
    response.write "  <div id=""uploadNewAttachmentLabel"" class=""groupsmall"">Upload an attachment:</div>" & vbcrlf
 end if

 response.write "<div id=""file_upload"" class=""divSection"">" & vbcrlf

'BEGIN: Attachment Form -------------------------------------------------------
 lcl_tablebg = ""

 if lcl_screenType = "N" then
    lcl_tablebg = " bgcolor=""#E0E0E0"
 end if

 response.write "<div id=""uploadAttachmentForm"">" & vbcrlf
 response.write "<table border=""0"" style=""width:675px;""" & lcl_tablebg & ">" & vbcrlf
 response.write "  <tr valign=""top"">" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <p>" & vbcrlf
 response.write "            <ol>" & vbcrlf
 response.write "                <li>Press <strong>Browse</strong> to find the file to upload.</li>" & vbcrlf
 response.write "                <li>Enter a description for the file (Max 1024 characters).</li>" & vbcrlf
 response.write "                <li>Press <strong>Save</strong>.</li>" & vbcrlf
 response.write "            </ol>  Note: It may take a few minutes to upload depending on the file size and your internet connection." & vbcrlf
 response.write "          </p>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "      <td align=""right"" nowrap=""nowrap"">&nbsp;<span id=""screenMsgAttachment"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"" align=""right"">&nbsp;</td>" & vbcrlf
 response.write "      <td rowspan=""5"">&nbsp;</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td align=""right""><strong>Name: </strong></td>" & vbcrlf
 response.write "      <td><input type=""file"" name=""filAttachment"" id=""filAttachment"" style=""width:650px;"" onchange=""validateAttachment();"" /></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td align=""right"" valign=""top""><strong>Description: </strong></td>" & vbcrlf
 response.write "      <td><textarea style=""width:575px;height:50px;"" name=""attachmentdesc""></textarea></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf

 if iOrgHasFeature_SecureAttachments AND iUserHasPermission_SecureAttachments then
   	response.write "  <tr>" & vbcrlf
    response.write "      <td>&nbsp;</td>" & vbcrlf
   	response.write "      <td><input type=""checkbox"" name=""attachmentIsSecure"" id=""attachmentIsSecure"" value=""on"" />&nbsp;Confidential</td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"" align=""right""><input type=""submit"" name=""saveAttachmentButton"" id=""saveAttachmentButton"" value=""Save"" class=""button ui-button ui-widget ui-corner-all"" onclick=""validateAttachment();"" /></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "</div>" & vbcrlf
'END: Attachment Form ---------------------------------------------------------
	
'BEGIN: Attachment List -------------------------------------------------------
 if lcl_screenType = "E" then
   	response.write "<div style=""background-color:#ffffff;padding-bottom: 10px;margin-bottom:5px;"">" & vbcrlf

    'subListAttachments p_requestid, _
    '                   iOrgHasFeature_SecureAttachments, _
    '                   iUserHasPermission_SecureAttachments, _
    '                   "Y", _
    '                   iOrgHasFeature_DisplayAttachmentsToPublic, _
    '                   iUserHasPermission_DisplayAttachmentsToPublic, _
    '                   sIsMobile

    subListAttachments p_requestid, _
                       iOrgHasFeature_SecureAttachments, _
                       iUserHasPermission_SecureAttachments, _
                       "Y", _
                       iOrgHasFeature_DisplayAttachmentsToPublic, _
                       iUserHasPermission_DisplayAttachmentsToPublic

   	response.write "</div>" & vbcrlf
 end if
'END: Attachment List ---------------------------------------------------------

	response.write "</div>" & vbcrlf
	response.write "</form>" & vbcrlf

end sub

'------------------------------------------------------------------------------
'sub subListAttachments(iRequestID, _
'                       iOrgHasFeature_SecureAttachments, _
'                       iUserHasPermission_SecureAttachments, _
'                       iCanMaintain, _
'                       iOrgHasFeature_DisplayAttachmentsToPublic, _
'                       iUserHasPermission_DisplayAttachmentsToPublic, _
'                       iIsMobile)

sub subListAttachments(iRequestID, _
                       iOrgHasFeature_SecureAttachments, _
                       iUserHasPermission_SecureAttachments, _
                       iCanMaintain, _
                       iOrgHasFeature_DisplayAttachmentsToPublic, _
                       iUserHasPermission_DisplayAttachmentsToPublic)

 'Check to see if we show the delete button and the "Secure" attachment checkbox
  lcl_canMaintain = "Y"

  if iCanMaintain <> "" then
     lcl_canMaintain = UCASE(iCanMaintain)
  end if

 'Is this the "Mobile Options" section?
  'dim sIsMobile, sIsMobileDBValue

  'sIsMobile        = "N"
  'sIsMobileDBValue = "0"

  'if iIsMobile <> "" then
  '   if not containsApostrophe(iIsMobile) then
  '      sIsMobile = ucase(iIsMobile)
  '   end if
  'end if

  'if sIsMobile = "Y" then
  '   sIsMobileDBValue = "1"
  'end if

  sBGColor = "#ffffff"

	'Retrieve all of the attachments for this requests
	 sSQL = "SELECT attachmentid, "
  sSQL = sSQL & " submitted_request_id, "
  sSQL = sSQL & " attachment_name, "
  sSQL = sSQL & " attachment_desc, "
  sSQL = sSQL & " adminuserid, "
  sSQL = sSQL & " date_added, "
  sSQL = sSQL & " firstname, "
  sSQL = sSQL & " lastname, "
  sSQL = sSQL & " isSecure, "
  sSQL = sSQL & " displayToPublic "
  sSQL = sSQL & " FROM egov_submitted_request_attachments "
  sSQL = sSQL &      " INNER JOIN users on adminuserid = userid "
  sSQL = sSQL & " WHERE submitted_request_id='" & iRequestID & "'"

 'Determine if the pic is for the Attachments or Mobile Options section
  'sSQL = sSQL & " AND isMobilePic = " & sIsMobileDBValue

 'If the org and user do not have the proper roles then only pull the non-secured attachments.
  if not iOrgHasFeature_SecureAttachments OR not iUserHasPermission_SecureAttachments then
     sSQL = sSQL & " AND isSecure = 0 "
  end if

  sSQL = sSQL & " ORDER BY date_added DESC"

 	set oAttachmentList = Server.CreateObject("ADODB.Recordset")
	 'oAttachmentList.Open sSQL,Application("DSN"),1,3
	 oAttachmentList.Open sSQL,lcl_dsn,1,3

 	response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">" & vbcrlf
 	response.write "<table class=""listAttachmentsHeaders"">" & vbcrlf
 	response.write "  <tr><th>Date Added - Added By - Name - Action</th></tr>" & vbcrlf
 	response.write "</table>" & vbcrlf
 	response.write "</div>" & vbcrlf
	
 	if not oAttachmentList.eof then
   		do while not oAttachmentList.eof
        sBGColor = changeBGColor(sBGColor,"#e0e0e0","#ffffff")

        if oAttachmentList("isSecure") then
           lcl_checked_attachment = " checked=""checked"""
        else
           lcl_checked_attachment = ""
        end if

        if oAttachmentList("displayToPublic") then
           lcl_checked_displayToPublic = " checked=""checked"""
        else
           lcl_checked_displayToPublic = ""
        end if

       'Build Attachment URL
        lcl_file_ext = lcase(right(oAttachmentList("attachment_name"), len(oAttachmentList("attachment_name")) - instrrev(oAttachmentList("attachment_name"),".")))

        lcl_attachment_url = oAttachmentList("attachmentid")
        lcl_attachment_url = lcl_attachment_url & "." & lcl_file_ext

     			response.write "<div style=""background-color:" & sBGColor & ";"">" & vbcrlf
        response.write "<table class=""listAttachments"">" & vbcrlf
        response.write "  <tr>" & vbcrlf
 		 				response.write "      <td>" & oAttachmentList("date_added") & " - </td>" & vbcrlf
 			 			response.write "      <td>" & oAttachmentList("firstname") & " " &  oAttachmentList("lastname") & " - </td>" & vbcrlf
 				 		response.write "      <td>" & oAttachmentList("attachment_name") & " - </td>" & vbcrlf
 					 	response.write "      <td>" & vbcrlf
        'response.write "          <input type=""button"" name=""viewAttachment" & oAttachmentList("attachmentid") & """ id=""viewAttachment" & oAttachmentList("attachmentid") & """ style=""button"" value=""View"" class=""button"" onclick=""location.href='attachment_view.asp?attachmentid=" & oAttachmentList("attachmentid") & "';"" />" & vbcrlf
        response.write "          <input type=""button"" name=""viewAttachment" & oAttachmentList("attachmentid") & """ id=""viewAttachment" & oAttachmentList("attachmentid") & """ style=""button"" value=""View"" class=""button"" onclick=""viewAttachment('" & lcl_attachment_url & "');"" />" & vbcrlf

        if lcl_canMaintain = "Y" then
           response.write "          <input type=""button"" name=""deleteAttachment" & oAttachmentList("attachmentid") & """ id=""deleteAttachment" & oAttachmentList("attachmentid") & """ style=""button"" value=""Delete"" class=""button"" onclick=""confirm_delete('" & oAttachmentList("attachmentid") & "','" & iTrackID & "');"" />" & vbcrlf
        end if

  						response.write "      </td>" & vbcrlf

       'If the org and user do not have the proper roles then only pull the non-secured attachments.
       'Also check to see if the user has the permission to maintain requests.
        if iOrgHasFeature_SecureAttachments AND iUserHasPermission_SecureAttachments AND lcl_canMaintain = "Y" then
    				 		response.write "      <td><input type=""checkbox"" name=""secureAttachment" & oAttachmentList("attachmentid") & """ id=""secureAttachment" & oAttachmentList("attachmentid") & """ value=""on"" onclick=""modifyAttachmentSecurity('" & oAttachmentList("attachmentid") & "');""" & lcl_checked_attachment & " /> Confidential</td>" & vbcrlf
        end if

        if iOrgHasFeature_DisplayAttachmentsToPublic AND iUserHasPermission_DisplayAttachmentsToPublic AND lcl_canMaintain = "Y" then
    				 		response.write "      <td><input type=""checkbox"" name=""displayToPublic" & oAttachmentList("attachmentid") & """ id=""displayToPublic" & oAttachmentList("attachmentid") & """ value=""on"" onclick=""modifyAttachmentsPublicDisplay('" & oAttachmentList("attachmentid") & "');""" & lcl_checked_displayToPublic & " /> Display To Public</td>" & vbcrlf
        end if

 	 					response.write "  </tr>" & vbcrlf
 		 				response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf

 				 		response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">" & vbcrlf

        if oAttachmentList("attachment_desc") <> "" then
           response.write "<table>" & vbcrlf
           response.write "  <tr>" & vbcrlf
           response.write "      <td colspan=""3""><i>" & oAttachmentList("attachment_desc") & "</i></td>" & vbcrlf
           response.write "  </tr>" & vbcrlf
  	   					response.write "</table>" & vbcrlf
        end if

        response.write "</div>" & vbcrlf

  						oAttachmentList.MoveNext
	  		loop
  else
   		response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">" & vbcrlf
     response.write "<table>" & vbcrlf
     response.write "  <tr><td colspan=""3""><i>No Attachments added.</i></td></tr>" & vbcrlf
	 	  response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf
     lcl_collapse = lcl_collapse & "$('#aaccord').accordion('option','active', false );"
 	end if

  oAttachmentList.close
  set oAttachmentList = nothing

end sub

'------------------------------------------------------------------------------
function fnPlainText( ByVal sValue )
	'sValue = UCASE(sValue)
	sValue = replace(sValue,"<B>","")
	sValue = replace(sValue,"</B>","")
	sValue = replace(sValue,"<P>","")
	sValue = replace(sValue,"</P>",vbcrlf)
	sValue = replace(sValue,"<BR>",vbcrlf)
	sValue = replace(sValue,"</BR>",vbcrlf)
	sValue = replace(sValue,"<STRONG>",vbcrlf)
	sValue = replace(sValue,"</STRONG>",vbcrlf)

	sValue = replace(sValue,"<b>","")
	sValue = replace(sValue,"</b>","")
	sValue = replace(sValue,"<p>","")
	sValue = replace(sValue,"</p>",vbcrlf)
	sValue = replace(sValue,"<br>",vbcrlf)
	sValue = replace(sValue,"</br>",vbcrlf)
	sValue = replace(sValue,"<strong>",vbcrlf)
	sValue = replace(sValue,"</strong>",vbcrlf)

	fnPlainText = sValue

end function

'------------------------------------------------------------------------------
function formatActivityLogComment(iComment)
  lcl_return = ""

  if iComment <> "" then
     lcl_return = iComment
     lcl_return = replace(lcl_return,"default_novalue","")
     lcl_return = replace(lcl_return,chr(10),"<br />")
     lcl_return = replace(lcl_return,chr(13),"")
     lcl_return = replace(lcl_return,"</p><br /><p>","</p><p>")
     lcl_return = replace(lcl_return,"<br /></p><p><br />","</p><p>")
     lcl_return = replace(lcl_return,"<p><br /><b>","<p><b>")
     lcl_return = replace(lcl_return,"</b><br><br />","</b><br />")
  end if

  formatActivityLogComment = lcl_return

end function

'------------------------------------------------------------------------------
function getActionLineFormID(p_orgid, p_action_autoid)

  lcl_return = 0

  if p_action_autoid <> "" then
    	sSQL = "SELECT category_id "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE orgid=" & p_orgid
     sSQL = sSQL & " AND action_autoid = " & p_action_autoid

     set oALFormID = Server.CreateObject("ADODB.Recordset")
     oALFormID.Open sSQL, Application("DSN"), 0, 1

     if not oALFormID.eof then
        lcl_return = oALFormID("category_id")
     end if

     oALFormID.close
     set oALFormID = nothing
  end if

  getActionLineFormID = lcl_return

end function

'------------------------------------------------------------------------------
function checkForAlertNotifications(p_orgid, p_form_id, p_email_action)

   lcl_return = False
   lcl_exists = "N"

   if p_form_id <> "" then
      sSQL = "SELECT distinct 'Y' as lcl_exists "
      sSQL = sSQL & " FROM egov_action_notifications "
      sSQL = sSQL & " WHERE orgid = " & p_orgid
      sSQL = sSQL & " AND action_form_id = " & p_form_id
      sSQL = sSQL & " AND email_action = '" & p_email_action & "' "

      set oNotifyExist = Server.CreateObject("ADODB.Recordset")
      oNotifyExist.Open sSQL, Application("DSN"), 0, 1

      if not oNotifyExist.eof then
         lcl_exists = oNotifyExist("lcl_exists")
      end if

      oNotifyExist.close
      set oNotifyExist = nothing

   end if

   if lcl_exists = "Y" then
      lcl_return = True
   end if

   checkForAlertNotifications = lcl_return


end function

'------------------------------------------------------------------------------
sub setupAlertNotificationsEmail(p_orgid, _
                                 p_action_autoid, _
                                 p_tracking_number, _
                                 p_form_id, _
                                 p_email_action)

  lcl_alert_notifications_exist = checkForAlertNotifications(p_orgid, _
                                                             p_form_id, _
                                                             p_email_action)

  if lcl_alert_notifications_exist then

     lcl_email_action = "updated"

     if p_email_action = "request_closed" then
        lcl_email_action = "closed"
     end if

     lcl_email_body = ""
     lcl_email_body = lcl_email_body & "Action Line Request (" & cstr(p_tracking_number) & ") has been " & lcl_email_action & ".<br /><br />"
     lcl_email_body = lcl_email_body & "<strong>Click the following link to view this Action Line Request:</strong><br />"
     lcl_email_body = lcl_email_body & "<a href=""" & getEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & p_action_autoid & "&e=Y"">" & getEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & p_action_autoid & "&e=Y</a>"

     sSQL = "SELECT distinct sendto "
     sSQL = sSQL & " FROM egov_action_notifications "
     sSQL = sSQL & " WHERE  orgid = " & p_orgid
     sSQL = sSQL & " AND action_form_id = " & p_form_id
     sSQL = sSQL & " AND email_action = '" & p_email_action & "' "

     set oSendAlertEmails = Server.CreateObject("ADODB.Recordset")
     oSendAlertEmails.Open sSQL, Application("DSN"), 0, 1

     if not oSendAlertEmails.eof then
        do while not oSendAlertEmails.eof

          'Get the user name and email
           lcl_notifyusername  = getAdminName(oSendAlertEmails("sendto"))
           lcl_notifyuseremail = getUserEmail(oSendAlertEmails("sendto"))

          'Check for a delegate
           getDelegateInfo lcl_notifyuserid, _
                           lcl_delegateid, _
                           lcl_delegate_username, _
                           lcl_delegate_useremail

          'Send the email
           setupSendEmail "NOTIFY", _
                           p_action_autoid, _
                           lcl_email_body, _
                           lcl_notifyuseremail, _
                           "Y", _
                           lcl_delegate_username, _
                           lcl_delegate_useremail

           oSendAlertEmails.movenext
        loop
     end if

     oSendAlertEmails.close
     set oSendAlertEmails = nothing

  end if

end sub

'------------------------------------------------------------------------------
'sub getActionLineAdminEmail(ByVal p_userid, ByVal iInternalEmail, ByRef lcl_userid, ByRef lcl_firstname, ByRef lcl_lastname, ByRef lcl_email)
sub getActionLineAdminEmail()

  lcl_userid    = ""
  lcl_firstname = ""
  lcl_lastname  = ""
  lcl_email     = ""

'  if p_userid <> "" then
'     sSQL = "SELECT email, userid, lastname, firstname "
'     sSQL = sSQL & " FROM users "
'		   sSQL = sSQL & " WHERE userid = " & p_userid

'		   set oAddress = Server.CreateObject("ADODB.Recordset")
'		   oAddress.Open sSQL, Application("DSN"), 0, 1

'		   if not oAddress.eof then
'        lcl_userid    = oAddress("userid")
'        lcl_firstname = oAddress("firstname")
'        lcl_lastname  = oAddress("lastname")
'        lcl_email     = oAddress("email")
'     end if

'     oAddress.close
'		   set oAddress = nothing
'  end if

end sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  dim lcl_return, sSuccess

  lcl_return = ""
  sSuccess   = ""

  if iSuccess <> "" then
     sSuccess = ucase(iSuccess)

     if sSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif sSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif sSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif sSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif sSuccess = "NE" then
        lcl_return = "Does not exist..."
     elseif sSuccess = "ERROR" then
        lcl_return = "ERROR"
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function isCategoryUsedOnRequest(iOrgID, iFormCategoryID)
  dim sOrgID, sFormCategoryID, lcl_return

  sOrgID          = 0
  sFormCategoryID = 0
  lcl_return      = false

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iFormCategoryID <> "" then
     sFormCategoryID = clng(iFormCategoryID)
  end if

  if sOrgID > 0 AND sFormCategoryID > 0 then
     sSQLc = "SELECT count(action_autoid) as total_requests "
     sSQLc = sSQLc & " FROM egov_actionline_requests "
     sSQLc = sSQLc & " WHERE orgid = " & sOrgID
     sSQLc = sSQLc & " AND category_id IN (select distinct(ftc.action_form_id) "
     sSQLc = sSQLc &                     " from egov_forms_to_categories ftc "
     sSQLc = sSQLc &                     " where ftc.orgid = " & sOrgID
     sSQLc = sSQLc &                     " and ftc.form_category_id = " & sFormCategoryID
     sSQLc = sSQLc &                     ") "

     set oCheckForCategoryOnRequests = Server.CreateObject("ADODB.Recordset")
     oCheckForCategoryOnRequests.Open sSQLc, Application("DSN"), 0, 1

     if not oCheckForCategoryOnRequests.eof then
        if oCheckForCategoryOnRequests("total_requests") > 0 then
           lcl_return = true
        end if
     end if

     oCheckForCategoryOnRequests.close
     set oCheckForCategoryOnRequests = nothing

  end if

  isCategoryUsedOnRequest = lcl_return

end function

'------------------------------------------------------------------------------
'function getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, p_todate, p_ownership, p_status)

'  lcl_return = 0

 'p_ownership represents the columns in the grid (who is the data limited to).  MINE, DEPT, ALL
 'p_status represents the rows in the grid (what status is the data limited to).  NEW, OPEN, CLOSED
'  if p_ownership <> "" AND p_status <> "" then
    'Determine which dates to search on

'     if UCASE(p_selectDateType) = "ACTIVE" then
'        varRequestCntClause = " AND ("
'        varRequestCntClause = varRequestCntClause & " (submit_date >= '" & p_fromdate & "' AND submit_date < '" & p_todate & "') OR "
'        varRequestCntClause = varRequestCntClause & " ( IsNull(complete_date,'" & Now & "') >= '" & p_fromdate & "' AND IsNull(complete_date,'" & Now & "') < '" & p_todate & "' ) OR "
'        varRequestCntClause = varRequestCntClause & " (submit_date < '" & p_fromdate & "' AND IsNull(complete_date,'" & Now & "') > '" & p_todate & "')) "
'     else 'selectDateType = SUBMIT
'        varRequestCntClause = " AND (submit_date BETWEEN '" & p_fromdate & "' AND '" & p_todate & "') "
'     end if

    'Build the SQL query.
'     sSQL = "SELECT count(action_autoid) AS total_requests "
'     sSQL = sSQL & " FROM egov_action_request_view "
'     sSQL = sSQL & " WHERE orgid = " & p_orgid
'     sSQL = sSQL & varRequestCntClause

    'Build the SQL statement for the p_ownership
'     if UCASE(p_ownership) = "MINE" then
'        sSQL = sSQL & " AND assignedemployeeid = " & p_userid
'     elseif UCASE(p_ownership) = "DEPT" then
'        sSQL = sSQL & " AND deptid IN (select distinct ug.groupid "
'        sSQL = sSQL &                " from usersgroups ug, groups g "
'        sSQL = sSQL &                " where ug.groupid = g.groupid "
'        sSQL = sSQL &                " and ug.userid = " & p_userid & ") "
'     else
       'For the StatusSummary report the "p_ownership" value is the GROUP BY value + the ID.
        'if instr(UCASE(p_ownership),"SUBMIT_DATE")       > 0 then
'        if instr(UCASE(p_ownership),"SUBMITDATESHORT")       > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"SUBMITDATESHORT","")
'        elseif instr(UCASE(p_ownership),"STREETNAME")    > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"STREETNAME","")
'        elseif instr(UCASE(p_ownership),"ACTION_FORMID") > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"ACTION_FORMID","")
'        elseif instr(UCASE(p_ownership),"DEPTID")        > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"DEPTID","")
'        elseif instr(UCASE(p_ownership),"ASSIGNED_NAME") > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"ASSIGNED_NAME","")
'        elseif instr(UCASE(p_ownership),"SUBMITTEDBY")   > 0 then
'           lcl_columnvalue = replace(UCASE(p_ownership),"SUBMITTEDBY","")
'        end if

'        if UCASE(p_ownership) = "SUBMITTEDBY" then
'           lcl_columnname = "assigned_userid"
'        else
'           lcl_columnname = replace(p_ownership,lcl_columnvalue,"")
'        end if

'        if UCASE(lcl_columnname) <> "ALL" AND UCASE(lcl_columnvalue) <> "ALL" then
'           sSQL = sSQL & " AND UPPER(" & lcl_columnname & ") = '" & UCASE(lcl_columnvalue) & "' "
'        end if

'     end if

    'Build the SQL statement for the p_status
'     sSQL = sSQL & " AND UPPER(status) = '" & UCASE(p_status) & "' "
     'if UCASE(p_status) = "NEW" then
     '   sSQL = sSQL & " AND UPPER(status) = 'SUBMITTED' "
     'elseif UCASE(p_status) = "OPEN" then
     '   sSQL = sSQL & " AND UPPER(status) IN ('SUBMITTED','INPROGRESS','WAITING') "
     'elseif UCASE(p_status) = "CLOSED" then
     '   sSQL = sSQL & " AND UPPER(status) IN ('RESOLVED','DISMISSED') "
     'end if
'dtb_debug(sSQL)
'   		set oWidget = Server.CreateObject("ADODB.Recordset")
   		'oWidget.Open sSQL, Application("DSN"), 0, 1
'   		oWidget.Open sSQL, lcl_dsn, 0, 1

'     if not oWidget.eof then
'        lcl_return = formatnumber(oWidget("total_requests"),0)
'     end if

'     oWidget.close
'     set oWidget = nothing
'  end if

'  getRequestCount = lcl_return

'end function

'------------------------------------------------------------------------------
'sub displayActionLineWidget(p_orgid, p_userid, p_selectDateType, p_fromdate, p_todate)

'    lcl_displayToDate = p_todate
'    lcl_queryToDate   = dateAdd("d",1,p_toDate)

   'Get request counts
'    lcl_mineSubmitted  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "SUBMITTED")
'    lcl_mineInProgress = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "INPROGRESS")
'    lcl_mineWaiting    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "WAITING")
'    lcl_mineResolved   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "RESOLVED")
'    lcl_mineDismissed  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "DISMISSED")

'    lcl_deptSubmitted  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "SUBMITTED")
'    lcl_deptInProgress = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "INPROGRESS")
'    lcl_deptWaiting    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "WAITING")
'    lcl_deptResolved   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "RESOLVED")
'    lcl_deptDismissed  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "DISMISSED")

'    lcl_allSubmitted   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "SUBMITTED")
'    lcl_allInProgress  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "INPROGRESS")
'    lcl_allWaiting     = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "WAITING")
'    lcl_allResolved    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "RESOLVED")
'    lcl_allDismissed   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "DISMISSED")


    'lcl_mineSubmitted  = 0
    'lcl_mineInProgress = 0
    'lcl_mineWaiting    = 0
    'lcl_mineResolved   = 0
    'lcl_mineDismissed  = 0

    'lcl_deptSubmitted  = 0
    'lcl_deptInProgress = 0
    'lcl_deptWaiting    = 0
    'lcl_deptResolved   = 0
    'lcl_deptDismissed  = 0

    'lcl_allSubmitted   = 0
    'lcl_allInProgress  = 0
    'lcl_allWaiting     = 0
    'lcl_allResolved    = 0
    'lcl_allDismissed   = 0

   'Get the sub-totals and totals
'    lcl_subtotal_mine_open = formatnumber(CLng(replace(lcl_mineSubmitted,"&nbsp;",0)) + CLng(replace(lcl_mineInProgress,"&nbsp;",0)) + CLng(replace(lcl_mineWaiting,"&nbsp;",0)),0)
'    lcl_subtotal_dept_open = formatnumber(CLng(replace(lcl_deptSubmitted,"&nbsp;",0)) + CLng(replace(lcl_deptInProgress,"&nbsp;",0)) + CLng(replace(lcl_deptWaiting,"&nbsp;",0)),0)
'    lcl_subtotal_all_open  = formatnumber(CLng(replace(lcl_allSubmitted,"&nbsp;",0))  + CLng(replace(lcl_allInProgress,"&nbsp;",0))  + CLng(replace(lcl_allWaiting,"&nbsp;",0)),0)

'    lcl_subtotal_mine_closed = formatnumber(CLng(replace(lcl_mineResolved,"&nbsp;",0)) + CLng(replace(lcl_mineDismissed,"&nbsp;",0)),0)
'    lcl_subtotal_dept_closed = formatnumber(CLng(replace(lcl_deptResolved,"&nbsp;",0)) + CLng(replace(lcl_deptDismissed,"&nbsp;",0)),0)
'    lcl_subtotal_all_closed  = formatnumber(CLng(replace(lcl_allResolved,"&nbsp;",0))  + CLng(replace(lcl_allDismissed,"&nbsp;",0)),0)

'    lcl_total_mine = CLng(lcl_subtotal_mine_open) + CLng(lcl_subtotal_mine_closed)
'    lcl_total_dept = CLng(lcl_subtotal_dept_open) + CLng(lcl_subtotal_dept_closed)
'    lcl_total_all  = CLng(lcl_subtotal_all_open)  + CLng(lcl_subtotal_all_closed)

'    response.write "<fieldset>" & vbcrlf
'    response.write "  <legend>Request Summary&nbsp;</legend>" & vbcrlf
'    response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px; margin-left:5px;"">" & vbcrlf
'    response.write "  <tr>" & vbcrlf
'    response.write "      <td align=""center"">" & vbcrlf
'    response.write "          <div style=""font-size:11px;""><strong>From: </strong><span style=""color:#800000;"">" & p_fromdate & "</span>&nbsp;<strong>To: </strong><span style=""color:#800000;"">" & p_todate & "</span></div><br />" & vbcrlf
'    response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" style=""margin-top:5px; margin-left:5px; border:1pt solid #000000;"">" & vbcrlf
'    response.write "            <tr align=""center"">" & vbcrlf
'    response.write "                <td colspan=""4"" style=""font-weight:bold; border-bottom:1pt solid #000000;"" bgcolor=""#93BEE1"">Action Line Requests</td>" & vbcrlf
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"" bgcolor=""#336699"">" & vbcrlf
'                                    displayActionLineWidgetCell "Status", "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell "Mine",   "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell "Dept",   "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell "All",    "", "ffffff", "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"">" & vbcrlf
'                                    displayActionLineWidgetCell "Submitted",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_mineSubmitted, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_deptSubmitted, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_allSubmitted,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"">" & vbcrlf
'                                    displayActionLineWidgetCell "In Progress",      "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_mineInProgress, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_deptInProgress, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_allInProgress,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"">" & vbcrlf
'                                    displayActionLineWidgetCell "Waiting",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_mineWaiting, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_deptWaiting, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_allWaiting,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"" bgcolor=""#93BEE1"">" & vbcrlf
'                                    displayActionLineWidgetCell "Total Open",           "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_mine_open, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_dept_open, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_all_open,  "", "ffffff", "Y", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf

'    response.write "            <tr align=""center"">" & vbcrlf
'                                    displayActionLineWidgetCell "Resolved",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_mineResolved, "", "",      "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_deptResolved, "", "",      "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_allResolved,  "", "",      "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"">" & vbcrlf
'                                    displayActionLineWidgetCell "Dismissed",        "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_mineDismissed,  "", "",     "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_deptDismissed,  "", "",     "", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_allDismissed,   "", "",     "", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"" bgcolor=""#93BEE1"">" & vbcrlf
'                                    displayActionLineWidgetCell "Total Closed",           "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_mine_closed, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_dept_closed, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_subtotal_all_closed,  "", "ffffff", "Y", "", "Y", "1pt", "000000", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "            <tr align=""center"" bgcolor=""#336699"">" & vbcrlf
'                                    displayActionLineWidgetCell "Grand Total",  "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_total_mine, "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_total_dept, "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000"
'                                    displayActionLineWidgetCell lcl_total_all,  "", "ffffff", "Y", "", "N", "", "", "N", "",    ""
'    response.write "            </tr>" & vbcrlf
'    response.write "          </table><br />" & vbcrlf
'    response.write "      </td>" & vbcrlf
'    response.write "  </tr>" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td style=""font-size:11px;"">" & vbcrlf
    'response.write "          <center style=""color:#800000; font-weight:bold;"">STATUSES</center>" & vbcrlf
    'response.write "          <strong>NEW</strong> - Submitted status (NOT included in Totals)<br />" & vbcrlf
    'response.write "          <strong>OPEN</strong> - Submitted, In Progress, and Waiting statuses<br />" & vbcrlf
    'response.write "          <strong>CLOSED</strong> - Resolved and Dismissed statuses" & vbcrlf
    'response.write "      </td>" & vbcrlf
    'response.write "  </tr>" & vbcrlf
'    response.write "</table>" & vbcrlf
'    response.write "</fieldset>" & vbcrlf
'end sub

'------------------------------------------------------------------------------
'sub displayActionLineWidgetCell(p_text, p_colspan, p_textColor, p_textBold, p_BGColor, _
'                                p_borderBottom, p_borderBottomSize, p_borderBottomColor, _
'                                p_borderRight, p_borderRightSize, p_borderRightColor)
'  lcl_rowStyle          = ""

'  lcl_colspan           = ""
'  lcl_textColor         = "color:#000000;"
'  lcl_textBold          = ""
'  lcl_BGColor           = ""

'  lcl_borderBottom      = ""
'  lcl_borderBottomSize  = "1pt"
'  lcl_borderBottomColor = "000000"

'  lcl_borderRight      = ""
'  lcl_borderRightSize  = "1pt"
'  lcl_borderRightColor = "000000"

 'Colspan
'  if p_colspan <> "" then
'     lcl_colspan = " colspan=""" & p_colspan & """"
'  end if

 'Text Color
'  if replace(p_textColor,"#","") <> "" then
'     lcl_textColor = "color:#" & replace(p_textColor,"#","") & ";"
'  end if

 'Text Bold
'  if p_textBold <> "" then
'     lcl_textBold = "font-weight:bold;"
'  end if

 'Background Color
'  if replace(p_BGColor,"#","") <> "" then
'     lcl_BGColor = "background-color:#" & replace(p_BGColor,"#","") & ";"
'  end if

 'Border Bottom
'  if p_borderBottom = "Y" then
'     if p_borderBottomSize <> "" then
'        lcl_borderBottomSize = p_borderBottomSize
'     end if

'     if replace(p_borderBottomColor,"#","") <> "" then
'        lcl_borderBottomColor = replace(p_borderBottomColor,"#","")
'     end if

'     lcl_borderBottom = "border-bottom:" & lcl_borderBottomSize & " solid #" & lcl_borderBottomColor & ";"
'  end if

 'Border Right
'  if p_borderRight = "Y" then
'     if p_borderRightSize <> "" then
'        lcl_borderRightSize = p_borderRightSize
'     end if

'     if replace(p_borderRightColor,"#","") <> "" then
'        lcl_borderRightColor = replace(p_borderRightColor,"#","")
'     end if

'     lcl_borderRight = "border-right:" & lcl_borderRightSize & " solid #" & lcl_borderRightColor & ";"
'  end if

'  response.write "<td style=""" & lcl_textColor & lcl_textBold & lcl_BGColor & lcl_borderBottom & lcl_borderRight & """" & lcl_colspan & ">" & p_text & "</td>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  if p_value <> "" then
     lcl_value = p_value
  else
     lcl_value = "EMPTY"
  end if

  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & REPLACE(lcl_value,"'","''") & "')"
  set rsi = Server.CreateObject("ADODB.Recordset")
  'rsi.Open sSQLi, Application("DSN"), 1, 3
  rsi.Open sSQLi, lcl_dsn, 1, 3

end sub


sub UpdateRegistryStatus(strStatus,intMobileID)
	sSQL = "UPDATE egovlinkRegistry.dbo.ServiceRequests SET status = '" & strStatus & "',updated_datetime='" & now() & "' WHERE service_request_id = '" & intMobileID & "'"
	Set oCmdMobile = Server.CreateObject("ADODB.Connection")
	oCmdMobile.Open Application("DSN")
	oCmdMobile.Execute(sSQL)
	oCmdMobile.Close
	Set oCmdMobile = Nothing
end sub
sub UpdateRegistryStatusNote(strNote, intMobileID)
	strCombinedNote = ""
	sSQL = "SELECT status_notes FROM egovlinkRegistry.dbo.ServiceRequests WHERE service_request_id = '" & intMobileID & "'"
	Set oNote = Server.CreateObject("ADODB.RecordSet")
	oNote.Open sSQL, Application("DSN"), 3, 1
	if not oNote.EOF then strCombinedNote = oNote("status_notes")
	oNote.Close
	Set oNote = Nothing

	if strCombinedNote <> "" then
		strCombinedNote = strNote & "<br />------------------------------<br />" & strCombinedNote
	else
		strCombinedNote = strNote
	end if


	sSQL = "UPDATE egovlinkRegistry.dbo.ServiceRequests SET status_notes = '" & dbsafe(strCombinedNote) & "',updated_datetime='" & now() & "' WHERE service_request_id = '" & intMobileID & "'"
	Set oCmdMobile = Server.CreateObject("ADODB.Connection")
	oCmdMobile.Open Application("DSN")
	oCmdMobile.Execute(sSQL)
	oCmdMobile.Close
	Set oCmdMobile = Nothing
end sub
sub SendPushNotification(strChannels, strAlert, strAction, intID, intJurisdictionID)
       	postData = "{ ""channels"": [ " + strChannels + " ], ""data"": { ""alert"": """ & strAlert & """, ""title"": ""E-Gov Alert!"", ""action"":""" & strAction & """, ""id"":""" & intID & """, ""jurisdiction_id"":""" & intJurisdictionID & """  } }"
	'response.write postData
	'response.end

	Set xmlHttp = Server.CreateObject("Microsoft.XMLHTTP") 
	xmlHttp.Open "POST", "https://api.parse.com/1/push", False
	xmlHttp.setRequestHeader "X-Parse-Application-Id", "ZduHB7Y8MfDGwzGGN6O2nWehZGbZRhwrgxDv0oVj"
	xmlHttp.setRequestHeader "X-Parse-REST-API-Key", "RCXZvGgfGBhwpiU870Ow0wuyue02YqlzuFJPMkX0"
	xmlHttp.setRequestHeader "Content-Type", "application/json"
	xmlHttp.Send postData
	set xmlHttp = Nothing 
end sub
sub SendSNS(strChannels, strAlert, strAction, intID, intJurisdictionID)
       	postData = "{ ""channels"": [ " + strChannels + " ], ""data"": { ""alert"": """ & strAlert & """, ""title"": ""E-Gov Alert!"", ""action"":""" & strAction & """, ""id"":""" & intID & """, ""jurisdiction_id"":""" & intJurisdictionID & """  } }"
	'response.write postData
	'response.end

	    Set xmlHttp = Server.CreateObject("Microsoft.XMLHTTP") 
	    xmlHttp.Open "POST", "http://registry2.eclinkhost.com/messages/sendtoawssns", False
	    xmlHttp.setRequestHeader "Content-Type", "application/json"
	    xmlHttp.Send postData
	    set xmlHttp = Nothing 
end sub
%>
