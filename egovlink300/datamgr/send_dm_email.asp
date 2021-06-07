<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!--#include file="../include_top_functions.asp"-->
<%
 lcl_orgid      = 0
 lcl_userid     = 0
 lcl_dmid       = 0
 lcl_action     = ""
 lcl_isAjax     = "N"

 if request("orgid") <> "" then
    if isnumeric(request("orgid")) then
       lcl_orgid = clng(request("orgid"))
    end if
 end if

 if request("userid") <> "" then
    if isnumeric(request("userid")) then
       lcl_userid = clng(request("userid"))
    end if
 end if

 if request("dmid") <> "" then
    if isnumeric(request("dmid")) then
       lcl_dmid = clng(request("dmid"))
    end if
 end if

 if request("action") <> "" then
    if not containsApostrophe(request("action")) then
       lcl_action = ucase(request("action"))
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 end if

 if lcl_userid > 0 AND lcl_dmid > 0 then

   'BEGIN: Get DM Type and User Info ------------------------------------------
    lcl_dm_typeid            = 0
    lcl_description          = ""
    lcl_accountInfoSectionID = 0
    lcl_feature_maintain     = ""
    lcl_feature_owners       = ""
    lcl_approvedeniedbydate  = ""
    lcl_categoryid           = 0
    lcl_assignedto           = 0
    lcl_assignedto_name      = ""
    lcl_assignedto_email     = ""

    sSQL = "SELECT "
    sSQL = sSQL & " dmt.dm_typeid, "
    sSQL = sSQL & " dmt.description, "
    sSQL = sSQL & " dmt.assignedto, "
    sSQL = sSQL & " dmt.accountInfoSectionID, "
    sSQL = sSQL & " dmt.feature_maintain, "
    sSQL = sSQL & " dmt.feature_owners, "
    sSQL = sSQL & " dmd.approvedeniedbydate, "
    sSQL = sSQL & " dmd.categoryid, "
    sSQL = sSQL & " u.firstname as assignedto_firstname, "
    sSQL = sSQL & " u.lastname as assignedto_lastname, "
    sSQL = sSQL & " u.email as assignedto_email "
    sSQL = sSQL & " FROM egov_dm_data dmd "
    sSQL = sSQL &      " INNER JOIN egov_dm_types dmt ON dmd.dm_typeid = dmt.dm_typeid "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON dmt.assignedto = u.userid "
    sSQL = sSQL & " WHERE dmd.dmid = " & lcl_dmid

   	set oGetDMEmailInfo = Server.CreateObject("ADODB.Recordset")
  	 oGetDMEmailInfo.Open sSQL, Application("DSN"), 3, 1

    if not oGetDMEmailInfo.eof then
       lcl_dm_typeid            = oGetDMEmailInfo("dm_typeid")
       lcl_description          = oGetDMEMailInfo("description")
       lcl_accountInfoSectionID = oGetDMEmailInfo("accountInfoSectionID")
       lcl_feature_maintain     = oGetDMEmailInfo("feature_maintain")
       lcl_feature_owners       = oGetDMEmailInfo("feature_owners")
       lcl_approvedeniedbydate  = oGetDMEmailInfo("approvedeniedbydate")
       lcl_categoryid           = oGetDMEmailInfo("categoryid")
       lcl_assignedto           = oGetDMEmailInfo("assignedto")
       lcl_assignedto_email     = oGetDMEmailInfo("assignedto_email")

       if oGetDMEmailInfo("assignedto_firstname") <> "" OR oGetDMEmailInfo("assignedto_lastname") <> "" then
          lcl_assignedto_name = oGetDMEmailInfo("assignedto_firstname") & " " & oGetDMEmailInfo("assignedto_lastname")
          lcl_assignedto_name = trim(lcl_assignedto_name)
       end if
    end if
   'END: Get DM Type and User Info --------------------------------------------

   'BEGIN: Get the Sender Info ------------------------------------------------
    lcl_sender_name  = ""
    lcl_sender_email = ""

    sSQL = "SELECT "
    sSQL = sSQL & " userfname, "
    sSQL = sSQL & " userlname, "
    sSQL = sSQL & " useremail "
    sSQL = sSQL & " FROM egov_users "
    sSQL = sSQL & " WHERE userid = " & lcl_userid

   	set oGetDMSenderInfo = Server.CreateObject("ADODB.Recordset")
  	 oGetDMSenderInfo.Open sSQL, Application("DSN"), 3, 1

    if not oGetDMSenderInfo.eof then
       if oGetDMSenderInfo("userfname") <> "" OR oGetDMSenderInfo("userlname") <> "" then
          lcl_sender_name = oGetDMSenderInfo("userfname") & " " & oGetDMSenderInfo("userlname")
          lcl_sender_name = trim(lcl_sender_name)
       end if

       lcl_sender_email = oGetDMSenderInfo("useremail")
    end if
   'END: Get the Sender Info --------------------------------------------------

    oGetDMEmailInfo.close
    oGetDMSenderInfo.close

    set oGetDMSenderInfo = nothing
    set oGetDMEmailInfo  = nothing

   'BEGIN: Determine which email to send --------------------------------------
   '  "REQUEST_OWNER"       - Email sent to assigned admin when citizen user requests 
   '                          to be an owner of a DM DATA record (i.e. Available Property, Business Listing, etc)
   '  "REQUEST_DM_APPROVAL" - Email sent to assigned admin when citizen user (owner) 
   '                          clicks the "Activate" button on a DM DATA record they are owner of.
    lcl_action_display  = ""
    lcl_from_email      = "(" & lcl_sender_name & ") <" & lcl_sender_email & ">"
    lcl_sendto_email    = lcl_assignedto_email
    lcl_cc_email        = ""
    lcl_subject         = lcl_description
    lcl_body_html       = ""
    lcl_body_text       = "<p>" & lcl_description & "</p>" & vbcrlf
    lcl_high_importance = "Y"
    lcl_accountinfo     = ""
    lcl_admin_url       = sEgovWebsiteURL & "/admin/datamgr"
    lcl_sendemail       = true

   'Set up Account Info data
    if lcl_action <> "NEW_SUBCATEGORY" then
       lcl_accountinfo = lcl_accountinfo & "<table border=""0"">" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &   "<tr>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &       "<td>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &           "<table border=""0"">" & vbcrlf
                                                        buildAccountInfoEmailData "FIELDNAME", lcl_dmid, lcl_dm_typeid, lcl_accountInfoSectionID, lcl_accountinfo, lcl_accountinfo
       lcl_accountinfo = lcl_accountinfo &           "</table>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &       "</td>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &       "<td>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &           "<table border=""0"">" & vbcrlf
                                                        buildAccountInfoEmailData "FIELDVALUE", lcl_dmid, lcl_dm_typeid, lcl_accountInfoSectionID, lcl_accountinfo, lcl_accountinfo
       lcl_accountinfo = lcl_accountinfo &           "</table>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &       "</td>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo &   "</tr>" & vbcrlf
       lcl_accountinfo = lcl_accountinfo & "</table>" & vbcrlf
    end if

   'Set up the HTML Body
    if lcl_action = "REQUEST_OWNER" then
       lcl_action_display = " (Owner Request: " & lcl_sender_name & ")"
       lcl_admin_url      = lcl_admin_url & "/datamgr_owners_list.asp?f=" & lcl_feature_owners

       lcl_body_html = lcl_body_html & "A request for ownership has been submitted by "
       lcl_body_html = lcl_body_html & lcl_sender_name & " for the following:<br />" & vbcrlf
       lcl_body_html = lcl_body_html & "<p>" & lcl_accountinfo & "</p>" & vbcrlf
       lcl_body_html = lcl_body_html & "<p>Click to open approval screen: <a href=""" & lcl_admin_url & """>" & lcl_admin_url & "</a></p>" & vbcrlf

      'Insert user as owner
       lcl_ownertype               = "OWNER"
       lcl_isApprovedDeniedByAdmin = false
       lcl_isApproved              = ""

       insertOwnerEditor lcl_orgid, _
                         lcl_dmid, _
                         lcl_userid, _
                         lcl_ownertype, _
                         lcl_isApprovedDeniedByAdmin, _
                         lcl_isApproved

    elseif lcl_action = "REQUEST_DM_APPROVAL" then
      'There is no need to resend an "approve/deny" email to the admin again because if there is already an "approvedeniedbydate"
      'then the email has already been sent once.  The owner has simply re-activated the dm data record.  We can determine if
      'a dm data record has been approved or not if "approvedeniedbydate" is NOT NULL.
       if lcl_approvedeniedbydate <> "" then
          lcl_sendemail = false
       else
          lcl_action_display = " (Request for Approval Submitted by " & lcl_sender_name & ") "
          lcl_admin_url      = lcl_admin_url & "/datamgr_list.asp?f=" & lcl_feature_maintain

          lcl_body_html = lcl_body_html & "A request for approval has been submitted by "
          lcl_body_html = lcl_body_html & lcl_sender_name & " for the following:<br />" & vbcrlf
          lcl_body_html = lcl_body_html & "<p>" & lcl_accountinfo & "</p>" & vbcrlf
          lcl_body_html = lcl_body_html & "<p>Click to open approval screen: <a href=""" & lcl_admin_url & """>" & lcl_admin_url & "</a></p>" & vbcrlf
       end if

    else  'NEW_SUBCATEGORY
       lcl_action_display = ": New Sub-Category (Request for Approval) "
       lcl_admin_url      = lcl_admin_url & "/datamgr_categories_maint.asp"
       lcl_admin_url      = lcl_admin_url & "?f="          & lcl_feature_maintain
       lcl_admin_url      = lcl_admin_url & "&dm_typeid="  & lcl_dm_typeid
       lcl_admin_url      = lcl_admin_url & "&categoryid=" & lcl_categoryid

       lcl_body_html = lcl_body_html & "At least one new sub-category has been created "
       lcl_body_html = lcl_body_html & "and a request for approval has been submitted by " & lcl_sender_name & "<br />" & vbcrlf
       lcl_body_html = lcl_body_html & "<p>Click to open approval screen: <a href=""" & lcl_admin_url & """>" & lcl_admin_url & "</a></p>" & vbcrlf       
    end if

    if lcl_sendemail then
       lcl_subject = lcl_subject & lcl_action_display

       sendEmail lcl_from_email, _
                 lcl_sendto_email, _
                 lcl_cc_email, _
                 lcl_subject, _
                 lcl_body_html, _
                 lcl_body_text, _
                 lcl_high_importance
    end if
   'END: Determine which email to send ----------------------------------------

    if lcl_isAjax = "Y" then
       response.write "sent"
    end if

 else
    if lcl_isAjax = "Y" then
       response.write "Failed to update section order - Error in AJAX Routine"
    'else
    '   response.write "datamgr_types_maint.asp?dm_typeid=" & lcl_dm_typeid & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub buildAccountInfoEmailData(ByVal iDisplayType, ByVal iDMID, ByVal iDM_TypeID, ByVal iAccountInfoSectionID, _
                              ByVal iAccountInfo, ByRef lcl_accountinfo)

  lcl_displayType = "FIELDVALUE"
  lcl_accountinfo = iAccountInfo

  if iDisplayType <> "" then
     lcl_displayType = ucase(iDisplayType)
  end if

  sSQLc = "  SELECT dmtf.dm_fieldid, "
  sSQLc = sSQLc & "  dmtf.dm_sectionid, "
  sSQLc = sSQLc & "  dmtf.section_fieldid, "
  sSQLc = sSQLc & "  dmsf.fieldname, "
  sSQLc = sSQLc & "  dmsf.fieldtype, "
  sSQLc = sSQLc & "  dmv.dm_valueid, "
  sSQLc = sSQLc & "  dmv.fieldvalue "
  sSQLc = sSQLc & "  FROM egov_dm_types_fields dmtf "
  sSQLc = sSQLc & "     LEFT OUTER JOIN egov_dm_values dmv "
  sSQLc = sSQLc & "                  ON dmtf.dm_fieldid = dmv.dm_fieldid "
  sSQLc = sSQLc & "                 AND dmv.dmid = " & iDMID
  sSQLc = sSQLc & "                 AND dmv.dm_typeid = " & iDM_TypeID
  sSQLc = sSQLc & "     LEFT OUTER JOIN egov_dm_sections_fields dmsf "
  sSQLc = sSQLc & "                  ON dmtf.section_fieldid = dmsf.section_fieldid "
  sSQLc = sSQLc & "                 AND dmsf.sectionid = " & iAccountInfoSectionID
  sSQLc = sSQLc & "  WHERE dmtf.dm_sectionid IN (SELECT dmts.dm_sectionid "
  sSQLc = sSQLc & "                              FROM egov_dm_types_sections dmts "
  sSQLc = sSQLc & "                              WHERE dmts.dm_typeid = " & iDM_TypeID
  sSQLc = sSQLc & "                                AND dmts.sectionid IN (SELECT dms.sectionid "
  sSQLc = sSQLc & "                                                       FROM egov_dm_sections dms "
  sSQLc = sSQLc & "                                                       WHERE dms.isAccountInfoSection = 1 "
  sSQLc = sSQLc & "                                                       AND dms.isActive = 1 "
  sSQLc = sSQLc & "                                                       AND dms.sectionid = " & iAccountInfoSectionID
  sSQLc = sSQLc & "                                                      ) "
  sSQLc = sSQLc & "                             ) "
  sSQLc = sSQLc & "  AND dmtf.displayInResults = 1 "
  sSQLc = sSQLc & "  ORDER BY dmtf.dm_sectionid, dmtf.dm_fieldid "

  set oAccountInfoColumns = Server.CreateObject("ADODB.Recordset")
  oAccountInfoColumns.Open sSQLc, Application("DSN"), 3, 1

  if not oAccountInfoColumns.eof then
     do while not oAccountInfoColumns.eof

        if lcl_displayType = "FIELDNAME" then
           lcl_display_value = oAccountInfoColumns("fieldname") & ":"
           lcl_isLabelStyle  = " style=""font-weight:bold"""
        else
           lcl_display_value = oAccountInfoColumns("fieldvalue")
           lcl_isLabelStyle  = ""
        end if

        lcl_accountinfo = lcl_accountinfo & "<tr><td" & lcl_isLabelStyle & ">" & lcl_display_value & "</td></tr>"

        oAccountInfoColumns.movenext
     loop
  end if

  'oAccountInfoColumns.close
  set oAccountInfoColumns = nothing

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>