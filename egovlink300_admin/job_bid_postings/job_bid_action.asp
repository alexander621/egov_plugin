<!-- #include file="../includes/common.asp" //-->
<%
'Retrieve all of the parameters
 lcl_sc_name                    = request("sc_name")
 lcl_sc_title                   = request("sc_title")
 lcl_sc_description             = request("sc_description")
 lcl_sc_publicly_viewable       = request("sc_publicly_viewable")
 lcl_sc_list_type               = request("sc_list_type")
 lcl_sc_orderby                 = request("sc_orderby")

 lcl_posting_id                 = request("posting_id")
 lcl_jobbid_id                  = request("jobbid_id")
	lcl_posting_type               = request("posting_type")
 lcl_title                      = request("title")
	lcl_start_date                 = request("start_date")
 lcl_start_hour                 = request("start_hour")
 lcl_start_minute               = request("start_minute")
 lcl_start_ampm                 = request("start_ampm")
	lcl_end_date                   = request("end_date")
 lcl_end_hour                   = request("end_hour")
 lcl_end_minute                 = request("end_minute")
 lcl_end_ampm                   = request("end_ampm")
 lcl_status_id                  = request("status_id")
 lcl_status_name                = request("status_name")
 lcl_additional_status_info     = request("additional_status_info")
 lcl_description                = request("description")
 lcl_qualifications             = request("qualifications")
 lcl_special_requirements       = request("special_requirements")
 lcl_misc_info                  = request("misc_info")
 lcl_active_flag                = request("active_flag")
 lcl_download_available         = request("download_available")

'Check to see if this record is to be deleted.
 if request("cmd") = "D" then
   'Remove any/all jobs/bids related to this posting_id
    sSQLd2 = "DELETE FROM egov_distributionlists_jobbids WHERE posting_id = " & lcl_posting_id

   	set rsd2 = Server.CreateObject("ADODB.Recordset")
   	rsd2.Open sSQLd2, Application("DSN") , 3, 1

   'Delete the job/bid
    sSQLd = "DELETE FROM egov_jobs_bids WHERE posting_id = " & lcl_posting_id

   	set rsd = Server.CreateObject("ADODB.Recordset")
   	rsd.Open sSQLd, Application("DSN") , 3, 1

    response.redirect "job_bid_list.asp?success=SD&sc_name=" & lcl_sc_name & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby

 end if

'Retrieve the columns related to the listtype
 if lcl_sc_list_type = "JOB" then
    lcl_job_salary              = request("job_salary")

    if request("public_apply_for_position_actionline") <> "" then
       lcl_public_apply_for_position_actionline = request("public_apply_for_position_actionline")
    else
       lcl_public_apply_for_position_actionline = "NULL"
    end if
 elseif lcl_sc_list_type = "BID" then
    lcl_bid_publication_info    = request("bid_publication_info")
    lcl_bid_submittal_info      = request("bid_submittal_info")
    lcl_bid_opening_info        = request("bid_opening_info")
    lcl_bid_recipient           = request("bid_recipient")
    lcl_bid_addendum_date       = request("bid_addendum_date")
    lcl_bid_pre_bid_meeting     = request("bid_pre_bid_meeting")
    lcl_bid_contact_person      = request("bid_contact_person")
    lcl_bid_fee                 = request("bid_fee")
    lcl_bid_plan_spec_available = request("bid_plan_spec_available")
    lcl_bid_business_hours      = request("bid_business_hours")
    lcl_bid_fax_number          = request("bid_fax_number")
    lcl_bid_plan_holders        = request("bid_plan_holders")
 end if

'Set up specific fields
 if lcl_status_id = "" or isnull(lcl_status_id) then
    lcl_status_id = 0
 end if

'Build the Start Date/Time fields
 if lcl_start_date <> "" then
    lcl_start_date = lcl_start_date & " " & lcl_start_hour & ":" & lcl_start_minute & ":00 " & lcl_start_ampm
 else
    lcl_start_date = NULL
 end if

 if lcl_end_date <> "" then
    lcl_end_date = lcl_end_date & " " & lcl_end_hour & ":" & lcl_end_minute & ":00 " & lcl_end_ampm
 else
    lcl_end_date = NULL
 end if

'Determine which message is to be displayed when returning to the main screen
 if lcl_sc_list_type = "JOB" then
    lcl_error_msg = "SNJ"
 else  'lcl_sc_list_type = "BID"
    lcl_error_msg = "SNB"
 end if

if request("posting_id") = 0 then
  'Set up the parameters that are to be inserted.
   sSQL = "INSERT INTO egov_jobs_bids ("
   sSQL = sSQL & "orgid, "
   sSQL = sSQL & "jobbid_id, "
   sSQL = sSQL & "posting_type, "
   sSQL = sSQL & "title, "

   if lcl_start_date <> "" then
      sSQL = sSQL & "start_date, "
   end if

   if lcl_end_date <> "" then
      sSQL = sSQL & "end_date, "
   end if

   sSQL = sSQL & "status_id, "
   sSQL = sSQL & "additional_status_info, "
   sSQL = sSQL & "description, "
   sSQL = sSQL & "qualifications, "
   sSQL = sSQL & "special_requirements, "
   sSQL = sSQL & "misc_info, "
   sSQL = sSQL & "active_flag, "

   if lcl_sc_list_type = "JOB" then
      sSQL = sSQL & "job_salary, "
      sSQL = sSQL & "public_apply_for_position_actionline, "
   else  'lcl_sc_list_type = "BID"
      sSQL = sSQL & "bid_publication_info, "
      sSQL = sSQL & "bid_submittal_info, "
      sSQL = sSQL & "bid_opening_info, "
      sSQL = sSQL & "bid_recipient, "

      if lcl_bid_addendum_date <> "" then
         sSQL = sSQL & "bid_addendum_date, "
      end if

      sSQL = sSQL & "bid_pre_bid_meeting, "
      sSQL = sSQL & "bid_contact_person, "
      sSQL = sSQL & "bid_fee, "
      sSQL = sSQL & "bid_plan_spec_available, "
      sSQL = sSQL & "bid_business_hours, "
      sSQL = sSQL & "bid_fax_number, "
      sSQL = sSQL & "bid_plan_holders, "
   end if

   sSQL = sSQL & "download_available "

  'set up the values to be inserted
   sSQL = sSQL & ") VALUES ("
   sSQL = sSQL       & session("orgid")                        & ", "
   sSQL = sSQL & "'" & dbsafe(lcl_jobbid_id)                   & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_posting_type)                & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_title)                       & "', "

   if lcl_start_date <> "" then
      sSQL = sSQL & "'" & dbsafe(lcl_start_date)               & "', "
   end if

   if lcl_end_date <> "" then
      sSQL = sSQL & "'" & dbsafe(lcl_end_date)                 & "', "
   end if

   sSQL = sSQL       & lcl_status_id                           & ", "
   sSQL = sSQL & "'" & dbsafe(lcl_additional_status_info)      & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_description)                 & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_qualifications)              & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_special_requirements)        & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_misc_info)                   & "', "
   sSQL = sSQL & "'" & dbsafe(lcl_active_flag)                 & "', "

   if lcl_sc_list_type = "JOB" then
      sSQL = sSQL & "'" & dbsafe(lcl_job_salary)               & "', "
      sSQL = sSQL       &  lcl_public_apply_for_position_actionline & ", "
   else  'lcl_sc_list_type = "BID"
      sSQL = sSQL & "'" & dbsafe(lcl_bid_publication_info)     & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_submittal_info)       & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_opening_info)         & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_recipient)            & "', "

      if lcl_bid_addendum_date <> "" then
         sSQL = sSQL & "'" & dbsafe(lcl_bid_addendum_date)     & "', "
      end if

      sSQL = sSQL & "'" & dbsafe(lcl_bid_pre_bid_meeting)      & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_contact_person)       & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_fee)                  & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_plan_spec_available)  & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_business_hours)       & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_fax_number)           & "', "
      sSQL = sSQL & "'" & dbsafe(lcl_bid_plan_holders)         & "', "
   end if

   sSQL = sSQL & "'" & dbsafe(lcl_download_available)          & "' "

   sSQL = sSQL & ")"

  'Set up the redirect parameters
   if request("autosend_email") = "Y" then
      session("return_url") = "sc_name=" & lcl_sc_name & "&sc_title=" & lcl_sc_title & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby
   else
      lcl_redirect_url = "job_bid_list.asp?success=" & lcl_error_msg & "&sc_name=" & lcl_sc_name & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby
   end if

 else
 		'Update existing record
		  sSQL = "UPDATE egov_jobs_bids SET "
    sSQL = sSQL & " jobbid_id = '"                  & dbsafe(lcl_jobbid_id)                & "', "
    sSQL = sSQL & " title = '"                      & dbsafe(lcl_title)                    & "', "
    sSQL = sSQL & " start_date = '"                 & dbsafe(lcl_start_date)               & "', "
    sSQL = sSQL & " end_date = '"                   & dbsafe(lcl_end_date)                 & "', "
    sSQL = sSQL & " status_id = "                   & lcl_status_id                        & ", "
    sSQL = sSQL & " additional_status_info = '"     & dbsafe(lcl_additional_status_info)   & "', "
    sSQL = sSQL & " description = '"                & dbsafe(lcl_description)              & "', "
    sSQL = sSQL & " qualifications = '"             & dbsafe(lcl_qualifications)           & "', "
    sSQL = sSQL & " special_requirements = '"       & dbsafe(lcl_special_requirements)     & "', "
    sSQL = sSQL & " misc_info = '"                  & dbsafe(lcl_misc_info)                & "', "
    sSQL = sSQL & " active_flag = '"                & dbsafe(lcl_active_flag)              & "', "

    if lcl_sc_list_type = "JOB" then
       sSQL = sSQL & " job_salary = '"              & dbsafe(lcl_job_salary)               & "', "
       sSQL = sSQL & " public_apply_for_position_actionline = " & lcl_public_apply_for_position_actionline & ", "
    else  'lcl_sc_list_type = "BID"
       sSQL = sSQL & " bid_publication_info = '"    & dbsafe(lcl_bid_publication_info)     & "', "
       sSQL = sSQL & " bid_submittal_info = '"      & dbsafe(lcl_bid_submittal_info)       & "', "
       sSQL = sSQL & " bid_opening_info = '"        & dbsafe(lcl_bid_opening_info)         & "', "
       sSQL = sSQL & " bid_recipient = '"           & dbsafe(lcl_bid_recipient)            & "', "
       sSQL = sSQL & " bid_addendum_date = '"       & dbsafe(lcl_bid_addendum_date)        & "', "
       sSQL = sSQL & " bid_pre_bid_meeting = '"     & dbsafe(lcl_bid_pre_bid_meeting)      & "', "
       sSQL = sSQL & " bid_contact_person = '"      & dbsafe(lcl_bid_contact_person)       & "', "
       sSQL = sSQL & " bid_fee = '"                 & dbsafe(lcl_bid_fee)                  & "', "
       sSQL = sSQL & " bid_plan_spec_available = '" & dbsafe(lcl_bid_plan_spec_available)  & "', "
       sSQL = sSQL & " bid_business_hours = '"      & dbsafe(lcl_bid_business_hours)       & "', "
       sSQL = sSQL & " bid_fax_number = '"          & dbsafe(lcl_bid_fax_number)           & "', "
       sSQL = sSQL & " bid_plan_holders = '"        & dbsafe(lcl_bid_plan_holders)         & "', "
    end if

    sSQL = sSQL & " download_available = '"         & dbsafe(lcl_download_available)       & "' "
		  sSQL = sSQL & " WHERE posting_id = " & lcl_posting_id & ""

   if request("autosend_email") = "Y" then
      session("return_url") = "sc_name=" & lcl_sc_name & "&sc_title=" & lcl_sc_title & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby
   else
      lcl_redirect_url = "job_bid_maint.asp?success=SU&posting_id=" & lcl_posting_id & "&sc_name=" & lcl_sc_name & "&sc_description=" & lcl_sc_description & "&sc_publicy_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby
   end if

	End If

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSQL, Application("DSN") , 3, 1

'Retrieve the posting_id that was just inserted
 sSQLid = "SELECT IDENT_CURRENT('egov_jobs_bids') as NewID"
 rs.Open sSQLid, Application("DSN") , 3, 1
 lcl_identity = rs.Fields("NewID").value

	Set oCmd = Nothing

'Update the job/bid posting (categories) associated to this job/bid.
 if lcl_identity <> "" AND (lcl_posting_id = 0 OR isnull(lcl_posting_id)) then
    lcl_posting_id = lcl_identity
 end if

'Remove all existing assignments for the job/bid (posting_id)
 if clng(lcl_posting_id) > 0 then
    sSQLd = "DELETE FROM egov_distributionlists_jobbids WHERE posting_id = " & lcl_posting_id

   	set rsd = Server.CreateObject("ADODB.Recordset")
   	rsd.Open sSQLd, Application("DSN") , 3, 1

   'Loop through all of the form fields and pull out the distribution lists (categories)
    lcl_email_dlids = ""
    for each oField in request.form
      	 if left(oField,9) = "category_" then
           lcl_dlid = REPLACE(oField,"category_","")

           sSQLi = "INSERT INTO egov_distributionlists_jobbids (distributionlistid,posting_id) VALUES ("
           sSQLi = sSQLi & lcl_dlid & ", "
           sSQLi = sSQLi & lcl_posting_id & ")"

          	set rsi = Server.CreateObject("ADODB.Recordset")
          	rsi.Open sSQLi, Application("DSN") , 3, 1

          'Check to see if the user has selected to generate a notification email to send to users
           if request("autosend_email") = "Y" then
             'Track the redirect url
'              session("return_url") = lcl_redirect_url

             'Track the distribution list ids that emails need to be sent to
              if lcl_email_dlids <> "" then
                 lcl_email_dlids = lcl_email_dlids & "," & lcl_dlid
              else
                 lcl_email_dlids = lcl_dlid
              end if
           end if


        end if
    next

    session("email_dlids") = lcl_email_dlids

    if request("autosend_email") = "Y" then
       session("jobbid_id")    = lcl_jobbid_id
       session("jobbid_title") = lcl_title
       response.redirect "../classes/dl_sendmail.asp?listtype=" & lcl_sc_list_type & "&screen_mode=AUTOSEND"
    end if

 end if

 if request("autosend_email") <> "Y" then
    response.redirect lcl_redirect_url
 end if

'---------------------------------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_return = REPLACE(p_value,"'","''")
  else
     lcl_return = ""
  end if

  dbsafe = lcl_return

end function


'----------------------------------------------------------------
function getPostingsName(p_value)
  sSQLp = "SELECT title "
  sSQLp = sSQLp & " FROM egov_jobs_bids "
  sSQLp = sSQLp & " WHERE orgid = "    & session("orgid")
  sSQLp = sSQLp & " AND posting_id = " & p_value

  set rsp = Server.CreateObject("ADODB.Recordset")
  rsp.Open sSQLp, Application("DSN"), 0, 1

  if not rsp.eof then
     lcl_return = rsp("title")
  else
     lcl_return = ""
  end if

  getPostingsName = lcl_return

  set rsp = nothing

end function

'----------------------------------------------------------------
function getCategoryName(p_value)
  sSQLc = "SELECT distributionlistname "
  sSQLc = sSQLc & " FROM egov_class_distributionlist "
  sSQLc = sSQLc & " WHERE distributionlistid = " & p_value
  sSQLc = sSQLc & " AND orgid = " & session("orgid")

  set rsc = Server.CreateObject("ADODB.Recordset")
  rsc.Open sSQLc, Application("DSN"), 0, 1

  if not rsc.eof then
     lcl_return = rsc("distributionlistname")
  else
     lcl_return = ""
  end if

  getCategoryName = lcl_return

  set rsc = nothing

end function
%>
