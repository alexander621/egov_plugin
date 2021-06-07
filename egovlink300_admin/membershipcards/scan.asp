<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="../poolpass/poolpass_global_functions.asp" //-->
<!-- #include file="membership_card_functions.asp" -->
<%
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../"  'Override of value from common.asp

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo = request("demo")

 if lcl_demo = "Y" then
    If Not UserHasPermission( Session("UserId"), "demo_scan_membership_cards" ) Then
   	   response.redirect sLevel & "permissiondenied.asp"
    End If
 else
    If Not UserHasPermission( Session("UserId"), "scan membership cards" ) Then
   	   response.redirect sLevel & "permissiondenied.asp"
    End If
 end if

 set oMembership = New classMembership

'Check for org features
 lcl_orghasfeature_pool_attendance_view             = orghasfeature("pool_attendance_view")
 lcl_orghasfeature_customreports_membership_scanlog = orghasfeature("customreports_membership_scanlog")
 lcl_orghasfeature_memberships_usekeycards          = orghasfeature("memberships_usekeycards")

'Check for user features
 lcl_userhaspermission_customreports_membership_scanlog = userhaspermission(session("userid"),"customreports_membership_scanlog")

'set up demo variables
 if lcl_demo = "Y" then
    lcl_demo_page_title = " (DEMO)"
    lcl_demo_url        = "&demo=" & lcl_demo
 else
    lcl_demo_page_title = ""
    lcl_demo_url        = ""
 end if

'Determine the layout based on the printer assigned to the org
 lcl_layout_id    = getPrinter_CardLayout(session("orgid"))
 lcl_layout_style = "layout" & lcl_layout_id & "_"

'Determine if the membership id has been populated.  If so then process it.
 dim lcl_card_status, lcl_member_id, lcl_status_msg, lcl_status_msg_class, lcl_display_card, lcl_play_warning
 dim lcl_checkin_result, lcl_expiration_date_label

 lcl_card_status = ""

 if request.querystring("memberid") <> "" then
    lcl_member_id = request.querystring("memberid")
 else
    lcl_member_id = request.form("p_member_id")
 end if

'If the "barcode" feature is enabled, then the "memberid" we currently have is actually,
'the barcode.  We need to perform an extra step to ensure that the barcode is actually associated
'to a memberid.
if lcl_member_id <> "" then
   if lcl_orghasfeature_memberships_usekeycards then

     'Custom barcodes start with "X" (i.e. punchcards, group passs, daily rates)
      if LEFT(UCASE(lcl_member_id),1) <> "X" then
         lcl_barcode   = lcl_member_id
         lcl_member_id = getMemberIDByBarcode(session("orgid"), _
                                              lcl_barcode)

         'if lcl_member_id = 0 then
         '   lcl_member_id = ""
         'end if
      end if
   end if
end if

if session("orgid") = 60 then
	lcl_scan_card_id = lcl_member_id
    	sSQL = "SELECT P.startdate, P.expirationdate,ppm.memberid  " _
		& " FROM egov_poolpasspurchases P " _
		& " INNER JOIN egov_poolpassmembers ppm ON P.poolpassid = ppm.poolpassid " _
		& " INNER JOIN egov_familymembers fm ON fm.familymemberid = ppm.familymemberid " _
		& " INNER JOIN egov_membership_periods MP ON P.periodid = MP.periodid" _
		& " WHERE fm.userid = '" & lcl_member_id & "' AND P.orgid = '" & session("orgid") & "'  " _
		& " AND P.startdate <= '" & now() & "' and P.expirationdate > '" & now() & "' "  _
		& " ORDER BY P.poolpassid DESC"
	set oExp = Server.CreateObject("ADODB.RecordSet")
	oExp.Open sSQL, Application("DSN"), 3, 1
	if NOT oExp.EOF then 
		lcl_member_id = oExp("memberid") & ""
	end if
	oExp.Close
	Set oExp = Nothing
end if

'Set up Session variable for DISPLAY include file.
 session("CARD_STATS") = "N"
 session("CARD_PRINT") = "N"
 session("MEMBERID")   = lcl_member_id
 'lcl_additional_info   = "N"

 if lcl_demo = "Y" then

    setupDemoFields request("sim"), _
                    lcl_layout_style, _
                    lcl_status_msg, _
                    lcl_status_msg_class, _
                    lcl_expiration_date_label, _
                    lcl_display_card, _
                    lcl_play_warning, _
                    lcl_hide_scan
 else
   '---------------------------------------------------------------------------
   'Validate the member_id and determine if a card exists and display proper message
   '---------------------------------------------------------------------------
    if lcl_member_id <> "" then
      '------------------------------------------------------------------------
      'Check to see the ATTENDANCE TYPE of the card scanned.
      '------------------------------------------------------------------------
       if LEFT(UCASE(lcl_member_id),1) = "X" AND ucase(lcl_member_id) <> "X001" then
          'lcl_checkin_result = getCheckInResult( lcl_member_id )
         '---------------------------------------------------------------------
         'Get the attendancetypeid from the value scanned on the card.
         '---------------------------------------------------------------------
          lcl_card_value = REPLACE(UCASE(lcl_member_id),"X","")

          if LEFT(CStr(lcl_card_value),1) = "0" then
            '------------------------------------------------------------------
            'Strip off any preceeding zeros (0)
            '------------------------------------------------------------------
             do while LEFT(CStr(lcl_card_value),1) = "0"
                lcl_card_value = RIGHT(lcl_card_value,(LEN(lcl_card_value)-1))
             loop
          end if

          lcl_checkin_result = checkAttendanceTypeExists(lcl_card_value)

         '---------------------------------------------------------------------
         'Display the INVALID screen if the attendancetypeid does not exist and/or is actually an invalid value
         '---------------------------------------------------------------------
          if lcl_checkin_result = "N" then
             setupCardFields lcl_checkin_result, _
                             lcl_layout_style, _
                             lcl_status_msg, _
                             lcl_status_msg_class, _
                             lcl_expiration_date_label, _
                             lcl_display_card, _
                             lcl_play_warning, _
                             lcl_hide_scan, _
                             lcl_additional_info

          else
            '------------------------------------------------------------------
            'Check to see if the info has already been found and posted
            '------------------------------------------------------------------
             lcl_addinfo = "N"

             if request("addinfo") = "Y" then
                lcl_addinfo = "Y"
             end if

             if lcl_addinfo = "Y" then
                lcl_rateid         = request("rateid")
                lcl_group_num      = request("groupnum")
                lcl_attendancetype = getAttendanceType(lcl_card_value)

               '---------------------------------------------------------------
               'Populate egov_pool_attendance_log to track the scan for attendance reporting
               '---------------------------------------------------------------
                insertAttendanceLog session("orgid"), _
                                    1, _
                                    lcl_rateid, _
                                    lcl_card_value, _
                                    lcl_group_num, "1"

               '---------------------------------------------------------------
               'Create a Pool Info record for current date IF...
               'This is the first scan of the current date AND...
               'NO Pool Info record exists.
               '---------------------------------------------------------------
                checkCreatePoolInfoRec session("orgid")

               'Display VALID Card
                setupCardFields lcl_checkin_result, _
                                lcl_layout_style, _
                                lcl_status_msg, _
                                lcl_status_msg_class, _
                                lcl_expiration_date_label, _
                                lcl_display_card, _
                                lcl_play_warning, _
                                lcl_hide_scan, _
                                lcl_additional_info

             else
               '---------------------------------------------------------------
               'Depending on the attendance type there may be further input fields to collect information about the card.
               'First check the egov_poolpassrates table (attendancetypeid) to see if there are any rates that exist
               '   for this attendance type for this org.  We need to get the rateid.  If there is only one record then
               '   automatically select that rate.  Otherwise, display the screen to allow the user to select the correct
               '   pool pass rate.
               '---------------------------------------------------------------
                sSQL = "SELECT count(rateid) AS total_rate_count "
                sSQL = sSQL & " FROM egov_poolpassrates "
                sSQL = sSQL & " WHERE orgid = " & session("orgid")
                sSQL = sSQL & " AND attendancetypeid = " & lcl_card_value

              		set oRateIDCnt = Server.CreateObject("ADODB.Recordset")
              		oRateIDCnt.Open sSQL, Application("DSN"), 3, 1

                if not oRateIDCnt.eof then
                   lcl_total_rate_count = oRateIDCnt("total_rate_count")
                else
                   lcl_total_rate_count = 0
                end if

                set oRateIDCnt = nothing

               '---------------------------------------------------------------
               'Check the total count to determine what action to take
               '---------------------------------------------------------------
                if lcl_total_rate_count > 0 then
                  '------------------------------------------------------------
                  'If there is only a single rate set up for this org for this attendance type then use that rateid.
                  '------------------------------------------------------------
                   if lcl_total_rate_count = 1 AND cstr(lcl_card_value) <> "4" then
                      sSQL = "SELECT r.rateid, "
                      sSQL = sSQL & " a.attendancetype, "
                      sSQL = sSQL & " r.maxsignups  "
                      sSQL = sSQL & " FROM egov_poolpassrates r, "
                      sSQL = sSQL &      " egov_pool_attendancetypes a "
                      sSQL = sSQL & " WHERE r.attendancetypeid = a.attendancetypeid "
                      sSQL = sSQL & " AND r.orgid = " & session("orgid")
                      'sSQL = sSQL & " AND publiccanpurchase = 1 "
                      sSQL = sSQL & " AND r.attendancetypeid = " & lcl_card_value

                    		set oRateInfo = Server.CreateObject("ADODB.Recordset")
                    		oRateInfo.Open sSQL, Application("DSN"), 3, 1

                      if not oRateInfo.eof then
                         lcl_rateid         = oRateInfo("rateid")
                         lcl_attendancetype = oRateInfo("attendancetype")

                        'If this is a group attendance type then pull the max sign ups.
                        'Otherwise all other attendance types are set to (1) no matter what the maxsignup is set to.
                         if CLng(lcl_card_value) = CLng(4) then
                            lcl_maxsignups  = oRateInfo("maxsignups")
                         else
                            lcl_maxsignups  = 1
                         end if

                      else
                         lcl_rateid         = ""
                         lcl_attendancetype = ""
                         lcl_maxsignups     = 1
                      end if

                      set oRateInfo = nothing

                     '---------------------------------------------------------
                     'Populate egov_pool_attendance_log to track the scan for attendance reporting
                     '---------------------------------------------------------
                      insertAttendanceLog session("orgid"), _
                                          1, _
                                          lcl_rateid, _
                                          lcl_card_value, _
                                          lcl_maxsignups, "2"

                     '---------------------------------------------------------
                     'Create a Pool Info record for current date IF...
                     'This is the first scan of the current date AND...
                     'NO Pool Info record exists.
                     '---------------------------------------------------------
                      checkCreatePoolInfoRec session("orgid")

                     'Display VALID Card
                      setupCardFields lcl_checkin_result, _
                                      lcl_layout_style, _
                                      lcl_status_msg, _
                                      lcl_status_msg_class, _
                                      lcl_expiration_date_label, _
                                      lcl_display_card, _
                                      lcl_play_warning, _
                                      lcl_hide_scan, _
                                      lcl_additional_info

                      lcl_attendancetype = getAttendanceType(lcl_card_value)
                   else
                     'Display ADDITIONAL INFO Card
                      setupCardFields "ADDITIONAL_INFO", _
                                      lcl_layout_style, _
                                      lcl_status_msg, _
                                      lcl_status_msg_class, _
                                      lcl_expiration_date_label, _
                                      lcl_display_card, _
                                      lcl_play_warning, _
                                      lcl_hide_scan, _
                                      lcl_additional_info

                      lcl_attendancetype = getAttendanceType(lcl_card_value)

                     '---------------------------------------------------------
                     'If we made it here then that means there is more than one rate available for the public to purchase,
                     'for this org, for this attendance type.  Cycle through them to build the options for the user to select.
                     'we are trying to get the RATEID so that we can update EGOV_POOL_ATTENDANCE_LOG
                     '---------------------------------------------------------
                      sSQL = "SELECT rateid, "
                      sSQL = sSQL & " description, "
                      sSQL = sSQL & " maxsignups "
                      sSQL = sSQL & " FROM egov_poolpassrates "
                      sSQL = sSQL & " WHERE orgid = " & session("orgid")
                      sSQL = sSQL & " AND attendancetypeid = " & lcl_card_value
                      'sSQL = sSQL & " AND publiccanpurchase = 1 "
                      'periodid ** need to check the year???
                      sSQL = sSQL & " ORDER BY displayorder "

                      set oRates = Server.CreateObject("ADODB.Recordset")
                      oRates.Open sSQL, Application("DSN"), 3, 1

                      if not oRates.eof then
                         lcl_max_value = 1

                        '------------------------------------------------------
                        'Build an extended status msg
                        '------------------------------------------------------
                         lcl_status_msg = lcl_status_msg & "</div><p>" & vbcrlf
                         lcl_status_msg = lcl_status_msg & "<strong>Select the rate:</strong>&nbsp;<select name=""p_rateid"" id=""p_rateid"">" & vbcrlf

                        '------------------------------------------------------
                        'Cycle through and display the rate options
                        '------------------------------------------------------
                         while not oRates.eof
                            if oRates("maxsignups") > lcl_max_value then
                               lcl_max_value = oRates("maxsignups")
                            end if

                            lcl_max_value_msg = ""

                            if cstr(lcl_card_value) = "4" then  'Attendance Type: GROUPS
                               lcl_max_value_msg = "&nbsp;[Max on pass: " & lcl_max_value & "]"
                            end if

                            lcl_status_msg = lcl_status_msg & "<option value=""" & oRates("rateid") & """>" & oRates("description") & lcl_max_value_msg & "</option>" & vbcrlf

                            oRates.movenext
                         wend

                         lcl_status_msg = lcl_status_msg & "</select><p>" & vbcrlf

                        '------------------------------------------------------
                        'If this is a GROUP attendance type then we also need to know the number of people IN the group.
                        '------------------------------------------------------
                         if cstr(lcl_card_value) = "4" then  'Attendance Type: GROUPS
                            lcl_status_msg = lcl_status_msg & "<strong>Number of people in group:&nbsp;</strong>"
                            lcl_status_msg = lcl_status_msg & "<select name=""p_group_num"" id=""p_group_num"">"

                           '---------------------------------------------------
                           'Build the dropdown to only allow the max on pass amount
                           '---------------------------------------------------
                            for i = 1 to lcl_max_value
                                lcl_status_msg = lcl_status_msg & "  <option value=""" & i & """>" & i & "</option>" & vbcrlf
                            next

                            lcl_status_msg = lcl_status_msg & "</select><p>" & vbcrlf
                         else
                            lcl_status_msg = lcl_status_msg & "<input type=""hidden"" name=""p_group_num"" id=""p_group_num"" value=""1"" size=""5"" maxlength=""5"" />" & vbcrlf
                         end if

                         oRates.close
                         set oRates = nothing

                      end if
                   end if
                else
                  'Nothing exists for the attendancetypeid.  ID is INVALID
                   setupCardFields "NOT_EXISTS", _
                                   lcl_layout_style, _
                                   lcl_status_msg, _
                                   lcl_status_msg_class, _
                                   lcl_expiration_date_label, _
                                   lcl_display_card, _
                                   lcl_play_warning, _
                                   lcl_hide_scan, _
                                   lcl_additional_info

                   lcl_attendancetype = getAttendanceType(lcl_card_value)

                end if  'end if for total_rate_count
             end if  'end if for additional info
          end if  'end if for checkin result

      'If LEFT(UCASE(lcl_member_id)) <> "X" -----------------------------------
       else
         'If a barcode has been entered and there isn't a value for the memberid, meaning an invalid barcode,
         'we need to clear the value from the barcode so that we can get the proper "does not exist" result.
          if ucase(lcl_member_id) = "X001" then
             lcl_barcode   = "invalid"
             lcl_member_id = 0
          end if

          if lcl_barcode <> "" AND lcl_member_id = "0" then
             lcl_member_id = ""
          end if

         'Make sure that the ID is numeric
	       'response.write "HERE" & IsNumeric(lcl_member_id & "")
         if IsNumeric(lcl_member_id) then
             lcl_checkin_result = getCheckInResult(session("orgid"), lcl_member_id)
             
	     
	     lcl_isExpired = false
	     sSQL = "SELECT P.startdate, P.expirationdate  " _
		& " FROM egov_poolpasspurchases AS P, egov_poolpassmembers AS ppm, egov_membership_periods AS MP  " _
		& " WHERE P.poolpassid = ppm.poolpassid AND P.periodid = MP.periodid AND ppm.memberid = '" & lcl_member_id & "' AND P.orgid = '" & session("orgid") & "'  " _
		& " AND P.startdate <= '" & now() & "' and P.expirationdate >= '" & now() & "' "  _
		& " ORDER BY P.poolpassid DESC"
		set oExp = Server.CreateObject("ADODB.RecordSet")
		on error resume next
		oExp.Open sSQL, Application("DSN"), 3, 1
		if oExp.EOF then lcl_isExpired = true
		oExp.Close
		Set oExp = Nothing
		on error goto 0



	     'response.write "<H1>" & lcl_checkin_result & "</h1>"
             lcl_rateid         = getRateID(session("orgid"), lcl_member_id)
             lcl_rate_desc      = getRateDescription(session("orgid"), lcl_rateid)

            'Determine if the org is using barcodes and if so, get the barcode status
             if lcl_orghasfeature_memberships_usekeycards then
                sBarcodeStatus = getBarcodeStatus(lcl_member_id, _
                                                  lcl_barcode)

                sIsBarcodeStatusActive = isBarcodeStatusActive(session("orgid"), _
                                                               lcl_member_id, _
                                                               lcl_barcode)
		if lcl_isExpired and (session("orgid") = "26" or session("orgid") = "198") then 
			sIsBarcodeStatusActive = false
			sBarcodeStatus = "Not Active"
		end if

             	  lcl_status_msg_class = lcl_layout_style & "card_status_invalid"
               	lcl_status_msg            = "Keycard Status: " & sBarcodeStatus
             	  lcl_hide_scan             = "Y"
             	  lcl_play_warning          = "Y"
             			lcl_expiration_date_label = ""
             	  lcl_display_card          = "Y"
             	  lcl_additional_info       = "N"

                if sIsBarcodeStatusActive then
                	  lcl_status_msg_class   = lcl_layout_style & "card_status_valid"
                	  lcl_hide_scan          = "N"
                	  lcl_play_warning       = "N"
                else
                   lcl_checkin_result = "none"
                end if
             else
                setupCardFields lcl_checkin_result, _
                                lcl_layout_style, _
                                lcl_status_msg, _
                                lcl_status_msg_class, _
                                lcl_expiration_date_label, _
                                lcl_display_card, _
                                lcl_play_warning, _
                                lcl_hide_scan, _
                                lcl_additional_info
             end if

             if lcl_checkin_result <> "none" AND lcl_checkin_result <> "expired" then
               'Check to see if the attendance tracking feature is turned on for the org
                if lcl_orghasfeature_pool_attendance_view then
                  'Get the attendance type
                   lcl_attendancetypeid = getAttendanceTypeId(session("orgid"), lcl_rateid)

                  'If the rate associated to this memberid is set to be a "punchcard" then reduce the "punchcard count".
                  'If the "punchcard count" IS zero BEFORE this reduction then the membership card is INVALID.
                   getPunchcardInfo lcl_member_id, _
                                    lcl_isPunchcard, _
                                    lcl_punchcard_limit, _
                                    lcl_punchcard_remain_cnt

                  'Check to see if this memberid has a rate that is set to be a "punchcard".
                  'If "yes" then get the "total remaining punchcard uses".
                  'If "total remaining punchcard uses" = (0) then this membership is EXPIRED.
                  'If "total remaining punchcard uses" > (0) then decrease the total count by (1).
                   if lcl_isPunchcard then

                     'Check to see if this memberid has a rate that is set to be a "punchcard".
                     'If "yes" then get the "total remaining punchcard uses".
                     'If "total remaining punchcard uses" = (0) then this membership is EXPIRED.
                     'If "total remaining punchcard uses" > (0) then decrease the total count by (1).
                      if lcl_punchcard_remain_cnt < 1 then

                        'Display EXPIRED Card
                         setupCardFields "expired", _
                                         lcl_layout_style, _
                                         lcl_status_msg, _
                                         lcl_status_msg_class, _
                                         lcl_expiration_date_label, _
                                         lcl_display_card, _
                                         lcl_play_warning, _
                                         lcl_hide_scan, _
                                         lcl_additional_info

                      else
                        'Populate egov_pool_attendance_log to track the scan for attendance reporting
                         insertAttendanceLog session("orgid"), _
                                             lcl_member_id, _
                                             lcl_rateid, _
                                             lcl_attendancetypeid, _
                                             1, "3"

                        'Create a Pool Info record for current date IF...
                        'This is the first scan of the current date AND...
                        'NO Pool Info record exists.
                         checkCreatePoolInfoRec session("orgid")

                         lcl_punchcard_remain_cnt = updatePunchcardInfo(session("orgid"), lcl_member_id, lcl_punchcard_limit)
                      end if

                  'Not a punchcard --------------------------------------------
                   else

                     'Populate egov_pool_attendance_log to track the scan for attendance reporting
                      insertAttendanceLog session("orgid"), _
                                          lcl_member_id, _
                                          lcl_rateid, _
                                          lcl_attendancetypeid, _
                                          1, "4"

                     'Create a Pool Info record for current date IF...
                     'This is the first scan of the current date AND...
                     'NO Pool Info record exists.
                      checkCreatePoolInfoRec session("orgid")

                   end if
                end if
             end if
          else
             if lcl_orghasfeature_memberships_usekeycards AND lcl_barcode <> "" then
                lcl_status_msg_class = lcl_layout_style & "card_status_invalid"
                lcl_status_msg            = "Keycard Does Not Exist"
                lcl_hide_scan             = "Y"
                lcl_play_warning          = "Y"
                lcl_expiration_date_label = ""
                lcl_display_card          = "N"
                lcl_additional_info       = "N"
             else
               'Display NUMERIC ONLY Card
                setupCardFields "NUMERIC_ONLY", _
                                lcl_layout_style, _
                                lcl_status_msg, _
                                lcl_status_msg_class, _
                                lcl_expiration_date_label, _
                                lcl_display_card, _
                                lcl_play_warning, _
                                lcl_hide_scan, _
                                lcl_additional_info
             end if
          end if
       end if
    else
      'Display Scan Field - Initial screen
       setupCardFields "", _
                       lcl_layout_style, _
                       lcl_status_msg, _
                       lcl_status_msg_class, _
                       lcl_expiration_date_label, _
                       lcl_display_card, _
                       lcl_play_warning, _
                       lcl_hide_scan, _
                       lcl_additional_info
   end if
 end if
%>
<html>
<head>
	<title>E-Gov Administration Console {Membership Card Scan}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="membership_card.css" />	

<style>
  #caption_membershiptype {
     text-align:  center;
     font-family: Verdana,Tahoma,Arial;
  }

  #membershiptype_ratedesc {
     color:     #800000;
     font-size: 20pt;
  }

  #membershiptype_punchcard {
     color: #800000;
  }

  #punchcard_usage_info {
     color:       #800000;
     font-size:   10pt;
     text-align:  center;
     padding-top: 10px;
  }

  .lastScannedMsg {
     text-align:  center;
     color:       #800000;
     padding-top: 10px;
  }

  .fieldset {
     border-radius: 5px;
  }
</style>

	<script language="javascript" src="validator.js"></script>
	<script language="javascript">
		function checkInput() {
			var scan = document.scan_form;
			if (Trim(scan.p_member_id.value)!="") {
			  //document.scan_form.submit();      
			}
		}

		function init() {
			var scan = document.scan_form;
			<%
     if lcl_hide_scan = "Y" then
        response.write "  document.getElementById(""reset_button"").focus();" & vbcrlf
     else
        response.write "		if(scan.p_member_id!=null) {" & vbcrlf
        response.write "   	 scan.p_member_id.select();" & vbcrlf
        response.write "  		 scan.p_member_id.focus();" & vbcrlf
        response.write "  }" & vbcrlf
     end if
   %>
		}

		function retake_picture() {
		<%
			session("RedirectPage") = "../MembershipCards/scan.asp?memberid=" & lcl_member_id & lcl_demo_url
			session("RedirectLang") = "Return to Scan Membership ID"
		%>
		location.href="image_takepic.asp?memberid=<%=lcl_member_id%>&reload_pic=Y";
		}

 function checkAdditionalFields() {
   lcl_rateid    = document.getElementById("p_rateid").value;
   lcl_group_num = document.getElementById("p_group_num").value;

   location.href='scan.asp?memberid=<%=lcl_member_id%>&addinfo=Y&rateid='+lcl_rateid+'&groupnum='+lcl_group_num+'<%=lcl_url_demo%>';
 }

function openCustomReports(p_report,p_memberid,p_rateid) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  var lcl_additional_parameters;

  if(p_memberid != "") {
     lcl_additional_parameters = "&memberid=" + p_memberid;
  }

  if(p_memberid != "") {
     if(lcl_additional_parameters != "") {
        lcl_additional_parameters = lcl_additional_parameters + "&rateid=" + p_rateid;
     }else{
        lcl_additional_parameters = "&rateid=" + p_rateid;
     }
  }

  eval('window.open("../customreports/customreports.asp?cr='+p_report+lcl_additional_parameters+'", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

	</script>
</head>
<body onLoad="init()">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""scan_form"" id=""scan_form"" method=""post"" action=""scan.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""p_start"" id=""p_start"" value=""1"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""demo"" id=""demo"" value=""" & lcl_demo & """ size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <h3>Scan Membership Card" & lcl_demo_page_title & "</h3>" & vbcrlf
  response.write "          <p>" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

  'Check for a Description for the rate.  If it exists this display it.
   if lcl_rate_desc <> "" then
      response.write "            <caption id=""caption_membershiptype"">" & vbcrlf
      response.write "              <strong>Membership Type: </strong>" & vbcrlf
      response.write "              <span id=""membershiptype_ratedesc"">" & lcl_rate_desc & "</span>" & vbcrlf

     'Display the label if this is a punchcard
      if lcl_isPunchcard then
         response.write "&nbsp;&nbsp;&nbsp;<span id=""membershiptype_punchcard"">[PUNCHCARD]</span>" & vbcrlf
      end if

      response.write "            </caption>" & vbcrlf
   end if

   response.write "            <tr valign=""top"">" & vbcrlf

  'BEGIN: Display the Membership Card -----------------------------------------
  'Displaying the card - This is the non-initial scan screen
   if lcl_display_card = "Y" then
      dim lcl_memberid, lcl_fname, lcl_lname, lcl_expiration_date, lcl_pathname, lcl_printed_count

      response.write "                <td>" & vbcrlf

     'If this is a non-member attendance type
      if LEFT(UCASE(lcl_member_id),1) = "X" then
         lcl_pathname        = ""
         lcl_memberid        = lcl_member_id
         lcl_fname           = lcl_attendancetype
         lcl_lname           = ""
        	lcl_expiration_date = ""
        	lcl_printed_count   = ""
      else
        'If this is a DEMO then we can default the variables
         if lcl_demo = "Y" then
            lcl_pathname = "../images/MembershipCard_Photos/demo/demo.jpg"
            lcl_memberid        = lcl_member_id
            lcl_fname           = "John"
            lcl_lname           = "Smith"
            lcl_expiration_date = "01/01/2001"
         else
            lcl_pathname   = "../images/MembershipCard_Photos/" & lcl_member_id & ".jpg"
	    if session("orgid") = 60 then lcl_pathname = "../images/MembershipCard_Photos/users/60_" & lcl_scan_card_id & ".jpg"
            lcl_poolpassid = getCurrentPoolPassID(lcl_member_id)

            sSQL = "SELECT P.poolpassid, "
            sSQL = sSQL & " P.rateid, "
            sSQL = sSQL & " M.memberid, "
            sSQL = sSQL & " u.userfname, "
            sSQL = sSQL & " u.userlname, "
            sSQL = sSQL & " M.printed_count, "
            sSQL = sSQL & " P.paymentdate, "
            sSQL = sSQL & " P.startdate, "
            sSQL = sSQL & " P.expirationdate, "
            sSQL = sSQL & " P.periodid, "
            sSQL = sSQL & " MP.period_interval, "
            sSQL = sSQL & " MP.period_qty, "
            sSQL = sSQL & " MP.period_type "
            sSQL = sSQL & " FROM egov_poolpassmembers M, "
            sSQL = sSQL &      " egov_familymembers F, "
            sSQL = sSQL &      " egov_users U, "
            sSQL = sSQL &      " egov_poolpasspurchases P, "
            sSQL = sSQL &      " egov_membership_periods MP "
            sSQL = sSQL & " WHERE M.familymemberid = F.familymemberid "
            sSQL = sSQL & " AND F.userid = U.userid "
            sSQL = sSQL & " AND p.orgid = u.orgid "
            sSQL = sSQL & " AND P.periodid = MP.periodid "
            sSQL = sSQL & " AND M.poolpassid = P.poolpassid "
            sSQL = sSQL & " AND M.memberid = " & CLng(lcl_member_id)

            if lcl_poolpassid > 0 then
               sSQL = sSQL & " AND M.poolpassid = " & lcl_poolpassid
            end if

            set oMemberid = Server.CreateObject("ADODB.Recordset")
            oMemberid.Open sSQL, Application("DSN"), 3, 1

            if not oMemberid.EOF then
               lcl_memberid        = oMemberid("memberid") 
               lcl_fname           = oMemberid("userfname")
               lcl_lname           = oMemberid("userlname")
               lcl_expiration_date = datevalue(oMemberid("expirationdate"))
               lcl_rateid          = oMemberid("rateid")
               'lcl_expiration_date = oMembership.getExpirationDate(oMemberid("periodid"), oMemberid("startdate"))

               'if UCASE(oMemberid("period_type")) = "SEASON" then
                  'lcl_expiration_date = FormatDateTime(CDate("12/31/" & Year(oMemberid("paymentdate"))), vbshortdate)
               'else
                  'lcl_expiration_date = FormatDateTime(DateAdd(oMemberid("period_interval"),clng(oMemberid("period_qty")),DateValue(oMemberid("paymentdate"))), vbshortdate)
               'end if

              	lcl_printed_count   = oMemberid("printed_count")
            else
               lcl_memberid        = lcl_member_id
               lcl_fname           = ""
               lcl_lname           = ""
              	lcl_expiration_date = ""
              	lcl_printed_count   = ""
               lcl_poolpassid      = 0
            end if

            oMemberid.close
            set oMemberid = nothing 
         end if
      end if

      lcl_watermark_class = "card_logo_display"

     '-------------------------------------------------------------------------
     'THIS TEMPORARY FIX HAS BEEN REMOVED FOR MONTGOMERY AS PER REQUEST BY JULIE MACHON on 4/24/2013
     '
     'temporary fix for Montgomery!
      'if session("orgid") <> 26 then
     '-------------------------------------------------------------------------
        'If the pathname has been entered then show the picture
         if lcl_pathname <> "" then
            response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""300"" height=""180"" class=""" & lcl_layout_style & "card_outline"">" & vbcrlf
            response.write "  <tr valign=""top"">" & vbcrlf
            response.write "      <td align=""center"" width=""124"" class=""" & lcl_layout_style & "card_text"">" & vbcrlf
            response.write "          <img id=""img_outline"" src=""" & lcl_pathname & """ />" & vbcrlf
            response.write "      </td>" & vbcrlf
            response.write "  </tr>" & vbcrlf
            response.write "</table>" & vbcrlf
            response.write "<div align=""center""># times Membership Card has been printed: " & lcl_printed_count & "</div>" & vbcrlf
         else
            response.write "&nbsp;" & vbcrlf  'This is done as a filler because we are currently in the middle of a <TD>.
         end if
     '-------------------------------------------------------------------------
     'temporary fix for Montgomery.  do NOT show image
      'else
      '   response.write "&nbsp;" & vbcrlf
      'end if
     '-------------------------------------------------------------------------

     'BEGIN: MEMBER IS VALID - Sound File -------------------------------------
      if lcl_play_warning = "N" then
        'Check to see if org has set up a custom sound file for the rate associated to this memberid.
         lcl_soundfile = getSoundFileByRate(session("orgid"), lcl_rateid)

					    response.write "<embed src=""" & lcl_soundfile & """ autostart=""true"" loop=""false"" width=""0"" height=""0"" align=""center"" style=""display: none""></embed>" & vbcrlf
      end if
     'END: MEMBER IS VALID - Sound File ---------------------------------------

      response.write "                </td>" & vbcrlf
     'END: Display the Membership Card ----------------------------------------

     'BEGIN: Membership Information -------------------------------------------
     'Display the Last Scanned Date/Time
      lcl_lastscanned = getLastScannedDate(session("orgid"), lcl_member_id, now())

      response.write "                <td>" & vbcrlf
      response.write "                    <p>" & vbcrlf
      response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
      response.write "                      <tr>" & vbcrlf
      response.write "                          <td>" & vbcrlf
      response.write "                              <p>" & vbcrlf
      response.write "                              <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%"" class=""" & lcl_layout_style & "card_text"">" & vbcrlf
      response.write "                                <tr>" & vbcrlf
      response.write "                                    <td valign=""top"" align=""center"">" & vbcrlf
      response.write "                                        <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""" & lcl_layout_style & "card_text"" bgcolor=""#efefef"" style=""border: 1pt solid #000000"">" & vbcrlf
      response.write "                                          <tr align=""center"">" & vbcrlf
      response.write "                                              <td>" & vbcrlf
      response.write "                                                  <p><font class="""    & lcl_layout_style & "scan_member_name_text"">"   & lcl_fname                 & "&nbsp;" & lcl_lname & "</font></p>" & vbcrlf
      response.write "                                                  <p><font class="""    & lcl_layout_style & "scan_expires_date_label"">" & lcl_expiration_date_label & "</font>" & vbcrlf
      response.write "                                                  <br /><font class=""" & lcl_layout_style & "scan_expiredate_text"">"    & lcl_expiration_date       & "</font></p>" & vbcrlf
      response.write "                                              </td>" & vbcrlf
      response.write "                                          </tr>" & vbcrlf

      if lcl_lastscanned <> "" then
        'Get the number of times this ID has been scanned for today
         lcl_total_scans = getScanCount(session("orgid"), lcl_member_id, now())

         response.write "                                          <tr>" & vbcrlf
         response.write "                                              <td class=""lastScannedMsg""><p>Last Scanned on: " & lcl_lastscanned & "</p></td>" & vbcrlf
         response.write "                                          </tr>" & vbcrlf

        'Display the total number of times the card has been scan "today"
         if CLng(lcl_total_scans) > CLng(1) then
            response.write "                                          <tr>" & vbcrlf
            response.write "                                              <td class=""lastScannedMsg"">" & vbcrlf
            response.write "                                                  This card has been scanned <strong>(" & lcl_total_scans & ")</strong> times today!" & vbcrlf
            response.write "                                              </td>" & vbcrlf
            response.write "                                          </tr>" & vbcrlf
         end if
      end if

     'If this is a punchcard then display the number of remaining punchcard uses
      if lcl_isPunchcard then
         response.write "                                          <tr>" & vbcrlf
         response.write "                                              <td id=""punchcard_usage_info"">" & vbcrlf
         response.write "                                                  Scanned <strong>(" & lcl_punchcard_limit - lcl_punchcard_remain_cnt & ")</strong> times out of <strong>(" & lcl_punchcard_limit & ")</strong><br />" & vbcrlf
         response.write "                                                  Remaining Punchcard Uses: <strong>(" & lcl_punchcard_remain_cnt & ")</strong><br /><br />" & vbcrlf
         response.write "                                              </td>" & vbcrlf
         response.write "                                          </tr>" & vbcrlf

        'Set up scan log history query.
        'Only show the "View Log" button if:
        '  1. The org has the "View/Edit Pool Daily Attendance" feature turned on .
        '  2. The org and user has the "Membership Scan Log" custom report feature assigned.
        '  3. The memberid has at least one scan history record.
         if  lcl_orghasfeature_pool_attendance_view _
         AND lcl_orghasfeature_customreports_membership_scanlog _
         AND lcl_userhaspermission_customreports_membership_scanlog then
             if checkMemberIDScanned(lcl_memberid,lcl_rateid) then
                response.write "  <tr>" & vbcrlf
                response.write "      <td align=""center"">" & vbcrlf
                response.write "          <p><input type=""button"" value=""View Scan History"" class=""button"" onclick=""openCustomReports('membershipcard_scans_by_memberid','" & lcl_memberid & "','" & lcl_rateid & "')"" /></p>" & vbcrlf
                response.write "      </td>" & vbcrlf
                response.write "  </tr>" & vbcrlf
             else
                response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
             end if
         end if

      end if

      response.write "                                        </table>" & vbcrlf
      response.write "                                    </td>" & vbcrlf
      response.write "                                </tr>" & vbcrlf
      response.write "                              </table>" & vbcrlf
      response.write "                              </p>" & vbcrlf
      response.write "                          </td>" & vbcrlf
      response.write "                      </tr>" & vbcrlf

  'This sets up the initial scan screen
   else
      response.write "                <td>&nbsp;</td>" & vbcrlf
      response.write "                <td>" & vbcrlf
      response.write "                    <p>" & vbcrlf
      response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
   end if

 'Display the scan status
   if lcl_status_msg <> "" then
      response.write "                      <tr>" & vbcrlf
      response.write "                          <td>" & vbcrlf

     'PLAY WARNING Sound file.
      if lcl_play_warning = "Y" then
         response.write "<embed src=""sound_warning1.wav"" autostart=""true"" loop=""false"" width=""0"" height=""0"" align=""center"" style=""display: none""></embed>" & vbcrlf
      end if

      response.write "                              <p><div align=""center"" class=""" & lcl_status_msg_class & """>" & lcl_status_msg & "</div></p>" & vbcrlf

     'If this IS a punchcard AND there are no more "punches" available then show the screen message.
      if lcl_isPunchcard AND lcl_punchcard_remain_cnt = 0 then
         response.write "                              <div id=""punchcard_usage_info"">" & vbcrlf
         response.write "                                *** No remaining uses available on punchcard ***<br />" & vbcrlf
         response.write "                                Scanned " & lcl_punchcard_limit & " times out of " & lcl_punchcard_limit & "." & vbcrlf
         response.write "                              </div>" & vbcrlf
      end if

      response.write "                          </td>" & vbcrlf
      response.write "                      </tr>" & vbcrlf
   end if

  '----------------------------------------------------------------------------
  'If this is the initial scan and/or the ID previously scan was NOT invalid/expired
  'then show the input box to accept a member id.
  '----------------------------------------------------------------------------
   if lcl_hide_scan <> "Y" then
      response.write "                      <tr>" & vbcrlf
      response.write "                          <td align=""center"" class=""" & lcl_layout_style & "heading"">" & vbcrlf
      response.write "                              <fieldset class=""fieldset"">" & vbcrlf
      response.write "                                Scan the Barcode on the Membership Card<br />" & vbcrlf
      response.write "                                <input type=""text"" name=""p_member_id"" id=""p_member_id"" size=""20"" maxlength=""50"" onChange=""javascript:checkInput();"" />&nbsp;&nbsp;" & vbcrlf
      response.write "                                <input type=""submit"" name=""B1"" id=""B1"" value=""Search"" class=""button"" />" & vbcrlf
      response.write "                              </fieldset>" & vbcrlf
      response.write "                          </td>" & vbcrlf
      response.write "                      </tr>" & vbcrlf
   else
      response.write "                      <tr><td>&nbsp;</td></tr>" & vbcrlf

      if lcl_additional_info = "Y" then
         response.write "                      <tr><td><input type=""button"" id=""reset_button"" name=""p_continue"" class=""button"" value=""Continue"" onClick=""checkAdditionalFields()"" /></td></tr>" & vbcrlf
      else
         response.write "                      <tr><td><input type=""button"" id=""reset_button"" name=""p_reset_scan"" class=""button"" value=""Reset Scan"" onClick=""location.href='scan.asp?" & replace(lcl_demo_url,"&","") & "'"" /></td></tr>" & vbcrlf
      end if
   end if

   response.write "                    </table>" & vbcrlf
   response.write "                    </p>" & vbcrlf

  'Display if this is the scan DEMO feature
   if lcl_demo = "Y" then
      response.write "                      <p>" & vbcrlf
      response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
      response.write "                        <tr>" & vbcrlf
      response.write "                            <td>" & vbcrlf
      response.write "                                <fieldset class=""fieldset"">" & vbcrlf
      response.write "                                  <legend>Simulate Scanner</legend>" & vbcrlf
      response.write "                                  <input type=""button"" id=""simulate_scan_valid"" name=""p_simulate_scan_valid"" value=""Simulate Scan of a Valid ID"" class=""button"" onClick=""location.href='scan.asp?memberid=0000&sim=V"         & lcl_demo_url & "'"" /><br />" & vbcrlf
      response.write "                                  <input type=""button"" id=""simulate_scan_invalid"" name=""p_simulate_scan_invalid"" value=""Simulate Scan of an Invalid ID"" class=""button"" onClick=""location.href='scan.asp?memberid=1234&sim=IV" & lcl_demo_url & "'"" /><br />" & vbcrlf
      response.write "                                  <input type=""button"" id=""simulate_scan_expired"" name=""p_simulate_scan_expired"" value=""Simulate Scan of an Expired ID"" class=""button"" onClick=""location.href='scan.asp?memberid=9874&sim=E"  & lcl_demo_url & "'"" />" & vbcrlf
      response.write "                                </fieldset>" & vbcrlf
      response.write "                            </td>" & vbcrlf
      response.write "                        </tr>" & vbcrlf
      response.write "                      </table>" & vbcrlf
      response.write "                      </p>" & vbcrlf
   end if

   response.write "                </td>" & vbcrlf
   response.write "            </tr>" & vbcrlf
   response.write "          </table>" & vbcrlf
   response.write "          </p>" & vbcrlf
   response.write "      </td>" & vbcrlf
   response.write "  </tr>" & vbcrlf
   response.write "</table>" & vbcrlf
   response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'-------------------------------------------------
function getRateID(iOrgID, p_member_id)
  lcl_return = ""

  if p_member_id <> "" then
     sSQL = "SELECT distinct p.rateid "
     sSQL = sSQL & " FROM egov_poolpassmembers m, "
     sSQL = sSQL &      " egov_poolpasspurchases p "
     sSQL = sSQL & " WHERE p.poolpassid = m.poolpassid "
     sSQL = sSQL & " AND p.orgid = "    & iOrgID
     sSQL = sSQL & " AND m.memberid = " & p_member_id

   		set oRateID = Server.CreateObject("ADODB.Recordset")
   		oRateID.Open sSQL, Application("DSN"), 3, 1

     if not oRateID.eof then
        lcl_return = oRateID("rateid")
     end if

     oRateID.close
     set oRateID = nothing

  end if

  getRateID = lcl_return

end function

'-----------------------------------------------------------------
function getAttendanceTypeId(iOrgID, p_rateid)
  lcl_return = ""

  if p_rateid <> "" then
     sSQL = "SELECT attendancetypeid "
     sSQL = sSQL & " FROM egov_poolpassrates "
     sSQL = sSQL & " WHERE orgid = " & iOrgID
     sSQL = sSQL & " AND rateid = "  & p_rateid

   		set oGetAttendTypeID = Server.CreateObject("ADODB.Recordset")
   		oGetAttendTypeID.Open sSQL, Application("DSN"), 3, 1

     if not oGetAttendTypeID.eof then
        lcl_return = oGetAttendTypeID("attendancetypeid")
     end if

     oGetAttendTypeID.close
     set oGetAttendTypeID = nothing

  end if

  getAttendanceTypeID = lcl_return
end function

'----------------------------------------------------------------
sub insertAttendanceLog(iOrgID, _
                        p_member_id, _
                        p_rate_id, _
                        p_attendancetypeid, _
                        p_num_of_people, p_location)

  if p_num_of_people = "" then
     lcl_num_of_people = 1
  else
     lcl_num_of_people = p_num_of_people
  end if

  if p_member_id > 0 then
     lcl_poolpassid = getCurrentPoolPassID(p_member_id)

     sSQL = "INSERT INTO egov_pool_attendance_log ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "memberid, "
     sSQL = sSQL & "rateid, "
     sSQL = sSQL & "scan_datetime, "
     sSQL = sSQL & "attendancetypeid, "
     sSQL = sSQL & "people_count, "
     sSQL = sSQL & "poolpassid"
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iOrgID             & ", "
     sSQL = sSQL & p_member_id        & ", "
     sSQL = sSQL & p_rate_id          & ", "
     sSQL = sSQL & "'" & now()        & "', "
     sSQL = sSQL & p_attendancetypeid & ", "
     sSQL = sSQL & lcl_num_of_people  & ", "
     sSQL = sSQL & lcl_poolpassid
     sSQL = sSQL & ")"

   		set oInsertAttend = Server.CreateObject("ADODB.Recordset")
   		oInsertAttend.Open sSQL, Application("DSN"), 3, 1

     set oInsertAttend = nothing
  end if

end sub

'----------------------------------------------------------------
function getLastScannedDate(iOrgID, p_member_id, p_currentdate)
  lcl_return = ""

  if p_member_id <> "" AND LEFT(UCASE(p_member_id),1) <> "X" then
     if isnumeric(p_member_id) then

        lcl_poolpassid = getCurrentPoolPassID(p_member_id)

        sSQL = "SELECT max(scan_datetime) as scan_datetime "
        sSQL = sSQL & " FROM egov_pool_attendance_log "
        sSQL = sSQL & " WHERE orgid = "    & iOrgID
        sSQL = sSQL & " AND memberid = "   & p_member_id
        ssQL = sSQL & " AND poolpassid = " & lcl_poolpassid

        if p_currentdate <> "" then
           sSQL = sSQL & " AND scan_datetime <> '" & p_currentdate & "'"
        end if

        set oLastScan = Server.CreateObject("ADODB.Recordset")
        oLastScan.Open sSQL, Application("DSN"), 3, 1

        if not oLastScan.eof then
           lcl_return = oLastScan("scan_datetime")
        end if

        oLastScan.close
        set oLastScan = nothing

     end if
  end if

  getLastScannedDate = lcl_return

end function

'----------------------------------------------------------------
function getScanCount(iOrgID, p_member_id,p_currentdate)
  lcl_return = 0

  if p_member_id <> "" then
     sSQL = "SELECT count(scan_datetime) as total_scans "
     sSQL = sSQL & " FROM egov_pool_attendance_log "
     sSQL = sSQL & " WHERE orgid = "  & iOrgID
     sSQL = sSQL & " AND memberid = " & p_member_id
     sSQL = sSQL & " AND DATEDIFF(d,scan_datetime,'" & p_currentdate & "') = 0 "

     set oScanCnt = Server.CreateObject("ADODB.Recordset")
     oScanCnt.Open sSQL, Application("DSN"), 3, 1

     if not oScanCnt.eof then
        lcl_return = oScanCnt("total_scans")
     else
        lcl_return = 0
     end if

     oScanCnt.close
     set oScanCnt = nothing

  end if

  getScanCount = lcl_return

end function

'--------------------------------------------------------------------------
sub checkCreatePoolInfoRec(iOrgID)
  sSQL = "SELECT distinct 'x' AS lcl_exists "
  sSQL = sSQL & " FROM egov_pool_info "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND cast(CONVERT(varchar(10), pool_date, 101) AS datetime) = '" & date() & "'"

  set oCheck = Server.CreateObject("ADODB.Recordset")
  oCheck.Open sSQL, Application("DSN"), 3, 1

  if oCheck.eof then
     sSQL = "INSERT INTO egov_pool_info ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "pool_date"
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iOrgID & ", "
     sSQL = sSQL & "'" & date() & "'"
     sSQL = sSQL & ")"

     set oInsertInfo = Server.CreateObject("ADODB.Recordset")
     oInsertInfo.Open sSQL, Application("DSN"), 3, 1

     'oInsertInfo.close
     set oInsertInfo = nothing
  end if

  oCheck.close
  set oCheck = nothing

end sub

'-----------------------------------------------------------------
function checkAttendanceTypeExists(p_value)
  lcl_return = "N"

  if p_value <> "" then
     sSQL = "SELECT distinct 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_pool_attendancetypes"
     sSQL = sSQL & " WHERE attendancetypeid = " & p_value
     sSQL = sSQL & " AND isactive = 1 "

   		set oCheckAttendType = Server.CreateObject("ADODB.Recordset")
   		oCheckAttendType.Open sSQL, Application("DSN"), 3, 1

     if not oCheckAttendType.eof then
        lcl_return = oCheckAttendType("lcl_exists")
     end if

     oCheckAttendType.close
     set oCheckAttendType = nothing

  end if

  checkAttendanceTypeExists = lcl_return
end function

'-----------------------------------------------------------------
function getAttendanceType(p_value)
  lcl_return = "N"

  if p_value <> "" then
     sSQL = "SELECT attendancetype "
     sSQL = sSQL & " FROM egov_pool_attendancetypes"
     sSQL = sSQL & " WHERE attendancetypeid = " & p_value

   		set oAttendType = Server.CreateObject("ADODB.Recordset")
   		oAttendType.Open sSQL, Application("DSN"), 3, 1

     if not oAttendType.eof then
        lcl_return = oAttendType("attendancetype")
     end if

     oAttendType.close
     set oAttendType = nothing

  end if

  getAttendanceType = lcl_return
end function

'------------------------------------------------------------------------------
function getRateDescription(iOrgID, iRateID)
  lcl_return = ""

  if iRateID <> "" then
  	sSQL =  "SELECT ISNULL(r.description,'') as description, period_desc "
  	sSQL = sSQL & "FROM egov_poolpassrates r "
  	sSQL = sSQL & "LEFT JOIN egov_membership_periods p ON p.periodid = r.periodid "
     sSQL = sSQL & " WHERE r.orgid = " & iOrgID
     sSQL = sSQL & " AND rateid = " & iRateID

   		set oRateDesc = Server.CreateObject("ADODB.Recordset")
   		oRateDesc.Open sSQL, Application("DSN"), 3, 1

     if not oRateDesc.eof then
        lcl_return = oRateDesc("description")
	if session("orgid") = "26" then lcl_return = lcl_return & " - " & oRateDesc("period_desc")
     end if

     oRateDesc.close
     set oRateDesc = nothing

  end if

  getRateDescription = lcl_return

end function

'------------------------------------------------------------------------------
function getSoundFileByRate(iOrgID, iRateID)
  lcl_return = "sound_confirm1.wav"

  if iRateID <> "" then
     sSQL = "SELECT distinct m.confirmsoundfile "
     sSQL = sSQL & " FROM egov_membershipcard_layout m, egov_poolpassrates r "
     sSQL = sSQL & " WHERE m.cardid = r.cardid "
     sSQL = sSQL & " AND r.orgid = " & iOrgID
     sSQL = sSQL & " AND r.rateid = " & iRateID

   		set oSoundFile = Server.CreateObject("ADODB.Recordset")
   		oSoundFile.Open sSQL, Application("DSN"), 3, 1

     if not oSoundFile.eof then
        if oSoundFile("confirmsoundfile") <> "" then
           lcl_return = "../custom/pub/" & session("sitename") & "/unpublished_documents"
           lcl_return = lcl_return & oSoundFile("confirmsoundfile")
        end if
     end if

     oSoundFile.close
     set oSoundFile = nothing

  end if

  getSoundFileByRate = lcl_return

end function


'------------------------------------------------------------------------------
sub setupDemoFields(ByVal iSimulationMode, ByVal iLayoutStyle, ByRef lcl_status_msg, ByRef lcl_status_msg_class, _
                    ByRef lcl_expiration_date_label, ByRef lcl_display_card, ByRef lcl_play_warning, ByRef lcl_hide_scan)

    if iSimulationMode = "V" then
       lcl_status_msg            = "Membership Card<br />is VALID" 
       lcl_status_msg_class      = iLayoutStyle & "card_status_valid"
       lcl_expiration_date_label = "EXPIRES ON"
 	     lcl_display_card          = "Y"
   		  lcl_play_warning          = "N"
   		  lcl_hide_scan             = "N"
    elseif iSimulationMode = "IV" then
       lcl_status_msg            = "Membership Card<br />is INVALID" 
       lcl_status_msg_class      = iLayoutStyle & "card_status_invalid"
      	lcl_expiration_date_label = ""
    	  lcl_display_card          = "N"
       lcl_play_warning          = "Y"
		     lcl_hide_scan             = "Y"
    elseif iSimulationMode = "E" then
       lcl_status_msg            = "Membership Card<br />has EXPIRED" 
       lcl_status_msg_class      = iLayoutStyle & "card_status_invalid"
      	lcl_expiration_date_label = "EXPIRED ON"
       lcl_display_card          = "Y"
       lcl_play_warning          = "Y"
       lcl_hide_scan             = "Y"
    end if

end sub

'------------------------------------------------------------------------------
sub setupCardFields(ByVal iCheckInResult, _
                    ByVal iLayoutStyle, _
                    ByRef lcl_status_msg, _
                    ByRef lcl_status_msg_class, _
                    ByRef lcl_expiration_date_label, _
                    ByRef lcl_display_card, _
                    ByRef lcl_play_warning, _
                    ByRef lcl_hide_scan, _
                    ByRef lcl_additional_info)

 	lcl_status_msg            = "" 
  lcl_status_msg_class      = iLayoutStyle & "card_status_invalid"
		lcl_expiration_date_label = ""
  lcl_display_card          = "N"
  lcl_play_warning          = "N"
  lcl_hide_scan             = "N"
  lcl_additional_info       = "N"

  if iCheckInResult <> "" then
    'Display the INVALID screen if the attendancetypeid does not exist and/or is actually an invalid value
     if UCASE(iCheckInResult) = "N" OR UCASE(iCheckInResult) = "NONE" then
    		  lcl_status_msg            = "Membership Card is<br />INVALID" 
        lcl_play_warning          = "Y"
        lcl_hide_scan             = "Y"

     elseif UCASE(iCheckInResult) = "Y" OR UCASE(iCheckInResult) = "VALID" OR UCASE(iCheckInResult) = "VALID_MULTIPLESCAN" then
    		  lcl_status_msg            = "Membership Card<br />is VALID"
        lcl_status_msg_class      = iLayoutStyle & "card_status_valid"
  		    lcl_expiration_date_label = "EXPIRES ON"
        lcl_display_card          = "Y"

     elseif UCASE(iCheckInResult) = "ADDITIONAL_INFO" then
        lcl_status_msg            = "Additional information<br />is required"
 	      lcl_display_card          = "Y"
        lcl_play_warning          = "Y"
        lcl_hide_scan             = "Y"
        lcl_additional_info       = "Y"

     elseif UCASE(iCheckInResult) = "NOT_EXISTS" then
        lcl_status_msg            = "Membership Card is INVALID" 
 	      lcl_display_card          = "Y"
        lcl_play_warning          = "Y"
		      lcl_hide_scan             = "Y"

     elseif UCASE(iCheckInResult) = "EXPIRED" then
        lcl_status_msg            = "Membership Card<br />has EXPIRED"
        lcl_expiration_date_label = "EXPIRED ON"
        lcl_display_card          = "Y"
        lcl_play_warning          = "Y"
		      lcl_hide_scan             = "Y"

     elseif UCASE(iCheckInResult) = "NUMERIC_ONLY" then
        lcl_status_msg            = "Membership Card has not been<br />entered correctly.  Numeric values only." 
      	 lcl_play_warning          = "Y"

     end if
  end if

end sub

'------------------------------------------------------------------------------
sub getPunchcardInfo(ByVal iMemberID, ByRef lcl_isPunchcard, ByRef lcl_punchcard_limit, ByRef lcl_punchcard_remain_cnt)

  lcl_isPunchcard          = False
  lcl_punchcard_limit      = 0
  lcl_punchcard_remain_cnt = 0
  lcl_currentPoolPassID    = getCurrentPoolPassID(iMemberID)

 'Check egov_poolpassmembers to see if this memberid is considered a "punchcard" and get additional punchcard info
  sSQL = "SELECT isPunchcard, "
  sSQL = sSQL & " punchcard_limit, "
  sSQL = sSQL & " pcard_remaining_cnt "
  sSQL = sSQL & " FROM egov_poolpassmembers "
  sSQL = sSQL & " WHERE poolpassid = " & lcl_currentPoolPassID
  sSQL = sSQL & " AND memberid = "     & iMemberID

  set oIsPCard = Server.CreateObject("ADODB.Recordset")
  oIsPCard.Open sSQL, Application("DSN"), 3, 1

  if not oIsPCard.eof then
     lcl_isPunchcard          = oIsPCard("isPunchcard")
     lcl_punchcard_limit      = oIsPCard("punchcard_limit")
     lcl_punchcard_remain_cnt = oIsPCard("pcard_remaining_cnt")
  end if

  oIsPCard.close
  set oIsPCard = nothing

end sub

'------------------------------------------------------------------------------
function updatePunchcardInfo(iOrgID, iMemberID, iPunchcardLimit)
  lcl_return       = 0
  lcl_totalscanned = 0

  lcl_currentPoolPassID = getCurrentPoolPassID(iMemberID)

 'Sum all of the scans for this MemberID and PoolPassID
  sSQL = "SELECT count(a.memberid) as Total_Scanned "
  sSQL = sSQL & " FROM egov_pool_attendance_log a, "
  sSQL = sSQL &      " egov_poolpassmembers m "
  sSQL = sSQL & " WHERE a.memberid = m.memberid "
  sSQL = sSQL & " AND a.poolpassid = m.poolpassid "
  sSQL = sSQL & " AND a.orgid = "      & iOrgID
  sSQL = sSQL & " AND m.poolpassid = " & lcl_currentPoolPassID
  sSQL = sSQL & " AND a.memberid = "   & iMemberID

  set oTotalScanned = Server.CreateObject("ADODB.Recordset")
  oTotalScanned.Open sSQL, Application("DSN"), 3, 1

  if not oTotalScanned.eof then
     lcl_totalscanned = oTotalScanned("Total_Scanned")
  end if

  oTotalScanned.close
  set oTotalScanned = nothing

 'Calculate the new pcard_remaining_cnt
  if lcl_totalscanned > 0 then
     lcl_return = iPunchcardLimit - lcl_totalscanned
  else
     lcl_return = 0
  end if

 'Update the Punchcard Remaining Uses
  sSQL = "UPDATE egov_poolpassmembers "
  sSQL = sSQL & " SET pcard_remaining_cnt = " & lcl_return
  sSQL = sSQL & " WHERE memberid = " & iMemberID
  sSQL = sSQL & " AND poolpassid = " & lcl_currentPoolPassID

  set oUpdatePCard = Server.CreateObject("ADODB.Recordset")
  oUpdatePCard.Open sSQL, Application("DSN"), 3, 1

  set oUpdatePCard = nothing

  updatePunchcardInfo = lcl_return

end function
%>
