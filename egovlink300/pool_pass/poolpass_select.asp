<!DOCTYPE HTML>
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="poolpass_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_select.asp
' AUTHOR: Steve Loar
' CREATED: 01/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/27/06 Steve Loar - Code added to template
' 1.1  09/11/08 David Boyer - Added Membership Renewals
' 1.2  03/06/08 David Boyer - Added "Alternate Layout"
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim oMembership, iPeriodId, iMembershipId

 set oMembership = New classMembership 

 session("RedirectPage") = "pool_pass/poolpass_select.asp?mtype=" & request("mtype")
 session("RedirectLang") = "Return to Membership Purchase"

 Dim sUserType, iUserid, sResidentDesc
 sUserType = "P"
 iUserid   = request.cookies("userid")

'If they do not have a userid set, take them to the login page automatically
 if request.cookies("userid") = "" or request.cookies("userid") = "-1" then
	   session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
   	response.redirect "../user_login.asp"
 end if

'First find out what resident type they are
 if iUserid = "-1" then
	   iUserid = ""
 end if

 sUserType     = getUserResidentType(iUserid)
 'response.write "<h1>" & sUserType & "</h1>"
 sResidentDesc = getResidentTypeDesc(sUserType)

 if request("mtype") = "" then
	  mtype = "pool" 
 else
	   mtype = request("mtype")
 end if


 iMembershipId = oMembership.GetMembershipId( mtype )

 if request("periodid") = "" then
	   iPeriodId = oMembership.GetInitialPeriod( iMembershipId )
 else
	   iPeriodId = CLng(request("periodid"))
 end if

 if request("rateDesc") = "" then
   	iRateDesc = oMembership.getFirstMembershipRateOption_AltLayout( sUserType, iMembershipId)
 else
 	  iRateDesc = request("rateDesc")
 end if

'Determine if the membership period selected is/isn't seasonal
 lcl_isSeasonal = checkIsSeasonal(iPeriodID)

'Set up the TITLE tag
 if iorgid = 7 then
    lcl_title = sOrgName
 else
    lcl_title = "E-Gov Services " & sOrgName & " Membership Purchase"
 end if

'Check for org features
 lcl_orghasfeature_membership_renewals            = orghasfeature(iorgid,"membership_renewals")
 lcl_orghasfeature_purchase_membership_alt_layout = orghasfeature(iorgid,"purchase_membership_alt_layout")
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
<title><%=lcl_title%></title>

<link rel="stylesheet" type="text/css" href="../css/styles.css" />
<link rel="stylesheet" type="text/css" href="../global.css" />
<link rel="stylesheet" type="text/css" href="style_pool.css" />
<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

<style type="text/css">
.fieldset {
   border-radius: 5px;
}
</style>

<script type="text/javascript" src="../scripts/modules.js"></script>
<script type="text/javascript" src="../scripts/easyform.js"></script>
<script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
<script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<!--

function EditUser(iUserId) {
		location.href='../register.asp?userid=' + iUserId;
}

function checkTermsRead() {
  var lcl_termsRequiredToContinue = 'N';
  var lcl_canContinue             = true;

  if($('#termsRequiredToContinue').val() != '') {
     lcl_termsRequiredToContinue = $('#termsRequiredToContinue').val();
  }

  if(lcl_termsRequiredToContinue == 'Y') {
     if(! document.getElementById('termsRead').checked) {
        lcl_canContinue = false;
     }
  }

  return lcl_canContinue;
}

function ContinuePurchase() {
  var lcl_canContinue = checkTermsRead();

  if(lcl_canContinue) {
     document.getElementById('purchaseMembership').action = 'select_members.asp';
     document.getElementById('purchaseMembership').submit();
  } else {
     $('#termsRead').focus();
					inlineMsg(document.getElementById("termsRead").id,'<strong>Required Field Missing: </strong>You must agree to the terms by checking the box next to \'I agree\'.',10,'termsRead');
					return false
  }
}

function renewPass(iPassId) {
  var lcl_canContinue = checkTermsRead();

  if(lcl_canContinue) {
     location.href = 'select_members.asp?poolpassid=' + iPassId;
  } else {
     $('#termsRead').focus();
					inlineMsg(document.getElementById("termsRead").id,'<strong>Required Field Missing: </strong>You must agree to the terms by checking the box next to \'I agree\'.',10,'termsRead');
					return false
  }
}

function updateRateID(iRateID) {
  document.getElementById("rateid").value = iRateID;
}

function submitPoolPassForm() {
  document.getElementById("purchaseMembership").action="poolpass_select.asp";
  document.getElementById("purchaseMembership").submit();
}

function enableDisableContinueButton(iRowValue) {
<%
  if lcl_orghasfeature_purchase_membership_alt_layout then
     lcl_radio_id          = "periodid_"
     lcl_onload_rateperiod = iPeriodID
  else
     lcl_radio_id          = "rateid_"
     lcl_onload_rateperiod = request("rateid")
  end if

  response.write "  if(iRowValue=='') {" & vbcrlf
  response.write "     document.getElementById(""continueButton"").disabled=true;" & vbcrlf
  response.write "  }else{" & vbcrlf
  response.write "     if(document.getElementById(""" & lcl_radio_id & """+iRowValue)) {"& vbcrlf
  response.write "        if(document.getElementById(""" & lcl_radio_id & """+iRowValue).checked==true) {" & vbcrlf
  response.write "           document.getElementById(""continueButton"").disabled=false;" & vbcrlf
  response.write "        }else{" & vbcrlf
  response.write "           document.getElementById(""continueButton"").disabled=true;" & vbcrlf
  response.write "        }" & vbcrlf
  response.write "     }else{" & vbcrlf
  response.write "        document.getElementById(""continueButton"").disabled=true;" & vbcrlf
  response.write "     }" & vbcrlf
  response.write "  }" & vbcrlf
%>
}
//-->
</script>
</head>

<!--#Include file="../include_top.asp"-->
<%
  RegisteredUserDisplay( "../" )

 'BEGIN: Registration or Login ------------------------------------------------
  if sOrgRegistration AND (request.cookies("userid") = "" OR request.cookies("userid") = "-1") then
     response.write "<div class=""reserveformtitle"">Contact Information</div>" & vbcrlf
     response.write "  <div class=""reserveforminputarea"">" & vbcrlf
     response.write "    <p>+ <strong><font class=""reserveforminstructions"">You need to register now or sign in to complete your purchase.</font></strong></p>" & vbcrlf
     response.write "    <p>+ <strong>" & vbcrlf
     response.write "    <input type=""button"" name=""loginButton"" id=""loginButton"" value=""Login"" class=""reserveformbutton"" onclick=""GotoLogin()"" />" & vbcrlf
     response.write "    or" & vbcrlf
     response.write "    <input type=""button"" name=""registerButton"" id=""registerButton"" value=""Register Now!"" class=""reserveformbutton"" onclick=""GotoRegister()"" />" & vbcrlf
     response.write "    </strong>" & vbcrlf
     response.write "    </p>" & vbcrlf
     response.write "  </div>" & vbcrlf
  end if
  if iorgid = "228" and sUserType = "Z" then
	  'response.write "Click ""MANAGE ACCOUNT"" above to complete your registration"
	  response.redirect "../manage_account.asp"
  end if

  response.write "<div id=""content"" class=""poolpasssel"">" & vbcrlf
  response.write "  <div id=""leftcontent"">" & vbcrlf
  response.write "	   <div id=""poolformtitle"">Membership Purchase</div>" & vbcrlf
  response.write "   	<div id=""poolforminputarea"">" & vbcrlf
  response.write "    		<p>Registration Information&nbsp;&nbsp;&nbsp;</p>" & vbcrlf
                          ShowUserInfo iUserId, _
                                       sUserType, _
                                       sResidentDesc

                          showMembershipTerms iorgid, _
                                              iMembershipID
  response.write "   	</div>" & vbcrlf
  response.write "    <p>" & vbcrlf
  response.write "  </p>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <form name=""purchaseMembership"" id=""purchaseMembership"" method=""post"" action=""poolpass_form.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""results"" id=""results"" value="""" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""searchstart"" id=""searchstart"" value="""     & sSearchStart   & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""iMembershipId"" id=""iMembershipId"" value=""" & iMembershipId  & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""isSeasonal"" id=""isSeasonal"" value="""       & lcl_isSeasonal & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""usertype"" id=""usertype"" value="""           & sUserType      & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""iuserid"" id=""iuserid"" value="""             & iUserId        & """ />" & vbcrlf

 'BEGIN: Check to see which layout the org has turned-on ----------------------
  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "    <input type=""hidden"" name=""rateid"" id=""rateid"" value=""" & request("rateid") & """ />" & vbcrlf
     response.write "  <div id=""membershipDropdownOptions"">" & vbcrlf

    'Membership Rates Dropdown
     response.write "    <p id=""membershipTypes"">" & vbcrlf
     response.write "      Membership Types:<br />" & vbcrlf
                           oMembership.ShowResidentRates_AltLayout iMembershipId, _
                                                                   sUserType, _
                                                                   iRateDesc
     response.write "    </p>" & vbcrlf

    'Membership Periods Checkboxes
     response.write "    <div id=""rightpicks_altlayout"" align=""left"">" & vbcrlf
                           oMembership.ShowPeriodPicksForMembership_AltLayout iMembershipId, _
                                                                              iPeriodID, _
                                                                              sUserType, _
                                                                              iRateDesc
     response.write "    </div>" & vbcrlf
  else
    'Membership Periods Dropdown
     response.write "  <div>" & vbcrlf
   	 response.write "    <p id=""periodpick"">" & vbcrlf
     response.write "      <strong>Membership Period:</strong>&nbsp;" & vbcrlf
                           oMembership.ShowPeriodPicksForMembership iMembershipId, _
                                                                    iPeriodId
     response.write "    </p>" & vbcrlf

    'Membership Rates Checkboxes
     response.write "    <div align=""center"">" & vbcrlf
     response.write "      <div id=""rightpicks"" align=""left"">" & vbcrlf
                          	 	oMembership.ShowMembershipResidencyRates iMembershipId, _
                                                                      sUserType, _
                                                                      iPeriodId 
     response.write "      </div>" & vbcrlf
  end if

  response.write "    <br />" & vbcrlf
  response.write "    <br />" & vbcrlf
  response.write "    <br />" & vbcrlf

 'Continue with Purchase button
  if request.cookies("userid") <> "" AND request.cookies("userid") <> "-1" AND oMembership.RatesAreVisible( mtype, sUserType) then
	  %>
	  <style>
	  #continueButton:disabled
	  {
		  border-color: #eee !important;
		  background-color: #ccc !important;
		  cursor: not-allowed !important;
	  }
	  </style>

	  <%
     response.write "    <input type=""button"" name=""continueButton"" id=""continueButton"" value=""Continue with Purchase"" class=""reserveformbutton"" onclick=""ContinuePurchase();"" />" & vbcrlf
  end if

  response.write "  </div>" & vbcrlf
 'END: Check to see which layout the org has turned-on ------------------------

  response.write "</form>" & vbcrlf

 'Show the renewal membership record if one exists for the user selected
  if lcl_orghasfeature_membership_renewals then
     showRenewalMembership iUserID, _
                           iMembershipID, _
                           iPeriodID
  end if

  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------

 'BEGIN: Spacing Code ---------------------------------------------------------
  response.write "<p><br />&nbsp;<br />&nbsp;</p>" & vbcrlf
 'END: Spacing Code -----------------------------------------------------------

 'Check for javascripts
  response.write "<script language=""javascript"">" & vbcrlf
  response.write "  enableDisableContinueButton('" & lcl_onload_rateperiod & "');" & vbcrlf
  response.write "</script>" & vbcrlf
%>
<!--#Include file="../include_bottom.asp"-->  
<%
 set oMembership = nothing

'------------------------------------------------------------------------------
sub showRenewalMembership(p_userid, p_membershipid, p_periodid)

		iPassCount = 0
  iRowCount  = 0

 'Build query to check for a renewal for user selected
		sSQL = "SELECT P.poolpassid, "
  sSQL = sSQL & " P.rateid, "
  sSQL = sSQL & " P.periodid, "
  sSQL = sSQL & " U.userfname, "
  sSQL = sSQL & " U.userlname, "
  sSQL = sSQL & " P.paymentamount, "
  sSQL = sSQL & " P.paymenttype, "
  sSQL = sSQL & " P.paymentdate, "
  sSQL = sSQL & " P.startdate, "
  sSQL = sSQL & " P.expirationdate, "
		sSQL = sSQL & " P.paymentlocation, "
  sSQL = sSQL & " P.paymentresult, "
  sSQL = sSQL & " M.membershipdesc, "
  sSQL = sSQL & " MP.period_desc, "
  sSQL = sSQL & " P.previous_poolpassid "
sSQL = sSQL & " FROM egov_users U, "
  sSQL = sSQL &      " egov_memberships M, "
  sSQL = sSQL &      " egov_membership_periods MP, "
  sSQL = sSQL &      " egov_poolpasspurchases P "
sSQL = sSQL & " WHERE P.orgid = " & iorgid
  sSQL = sSQL & " AND P.paymentresult <> 'Pending' "
sSQL = sSQL & " AND P.paymentresult <> 'Declined' "
  sSQL = sSQL & " AND U.userid = P.userid "
  sSQL = sSQL & " AND P.membershipid = M.membershipid "
  sSQL = sSQL & " AND P.periodid = MP.periodid "
  sSQL = sSQL & " AND U.userid = "       & p_userid
  sSQL = sSQL & " AND P.membershipid = " & p_membershipid
  sSQL = sSQL & " AND P.periodid = "     & p_periodid
		sSQL = sSQL & " ORDER BY P.poolpassid "

 	set oRequests = Server.CreateObject("ADODB.Recordset")
	 oRequests.Open sSQL, Application("DSN"), 3, 1

  if not oRequests.eof then
     bgcolor            = "#eeeeee"
     lcl_showheader_row = "Y"
     lcl_close_renewal  = "N"
     lcl_td_styles      = " class=""box_header"" id=""renewalCell"""

   		while not oRequests.eof
  	   		iPassCount = iPassCount + 1
   		  	iRowCount  = iRowCount + 1

       'Retrieve the RATE info
        getRateInfo oRequests("rateid"), _
                    nAmount, _
                    sMessage, _
                    sDescription, _
                    iMaxsignups, _
                    iAttendanceTypeID, _
                    lcl_rate_residenttype, _
                    lcl_rate_residenttypedesc, _
                    lcl_isPunchcard, _
                    lcl_punchcard_limit

       'Determine if the Renewal column/button are displayed
        if lcl_orghasfeature_membership_renewals then
           lcl_showHideRenewalButton = showHideRenewalButton(oRequests("poolpassid"))
        else
           lcl_showHideRenewalButton = "N"
        end if

       'Display only the record that is to be renewed.
        if lcl_showHideRenewalButton = "Y" then

          'Show the first row of column headers.
           if lcl_showheader_row = "Y" then
              lcl_showheader_row = "N"
              lcl_close_renewal  = "Y"

              response.write "<fieldset id=""membershipRenewals"" class=""fieldset"">" & vbcrlf
              response.write "  <legend><strong>Membership(s) Available for Renewal</strong>&nbsp;</legend>" & vbcrlf
              response.write "<br />" & vbcrlf
              response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""border:solid 1pt #000000;"" class=""liquidtable"">" & vbcrlf
              response.write "  <thead><tr align=""left"">" & vbcrlf
              response.write "      <td" & lcl_td_styles & " align=""center"">Pass<br />ID</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & ">Membership Type</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & " align=""center"">Purchase<br />Date</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & " align=""center"">Start<br />Date</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & " align=""center"">Expiration<br />Date</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & ">Purchaser</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & " align=""center"" id=""column_renewal"">Renewal<br />of Pass ID</td>" & vbcrlf
              response.write "      <td" & lcl_td_styles & ">&nbsp;</td>" & vbcrlf
              response.write "  </tr></thead>" & vbcrlf
           else
              lcl_showheader_row = lcl_showheader_row
              lcl_close_renewal  = lcl_close_renewal
           end if

           lcl_row_mouseover = ""
           lcl_row_mouseout  = ""
           lcl_td_mouseover  = ""
           lcl_td_mouseout   = ""
           'lcl_onclick       = " onClick=""location.href='poolpass_details.asp?iPoolPassId=" & oRequests("poolpassid") & "';"""

        			response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & bgcolor & """" & lcl_row_mouseover & lcl_row_mouseout & ">" & vbcrlf
				response.write "<td class=""repeatheaders"">Pass ID</td>"
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & oRequests("poolpassid") & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Membership Type</td>"
     	   		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("membershipdesc") & " &ndash; " & oRequests("period_desc") & "<br />" & lcl_rate_residenttypedesc & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Purchase Date</td>"
      		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & DateValue(oRequests("paymentdate"))    & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Start Date</td>"
      		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & DateValue(oRequests("startdate"))      & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Expiration Date</td>"
      		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & DateValue(oRequests("expirationdate")) & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Purchaser</td>"
  	      		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("userfname") & " " & oRequests("userlname")   & "</td>" & vbcrlf
				response.write "<td class=""repeatheaders"">Renewal of Pass ID</td>"
				strPrevPoolPassID = oRequests("previous_poolpassid")
				if strPrevPoolPassID = "" or isnull(strPrevPoolPassID) then strPrevPoolPassID = "N/A"
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" &  strPrevPoolPassID      & "</td>" & vbcrlf
     	   		response.write "      <td><input type=""button"" name=""renew"" id=""button_renew_" & oRequests("poolpassid") & """ value=""Renew"" onclick=""renewPass('" & oRequests("poolpassid") & "');"" class=""button"" /></td>" & vbcrlf
   		     	response.write "</tr>" & vbcrlf

           bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

        end if

     			oRequests.movenext
   		wend

     if lcl_close_renewal = "Y" then
      		response.write "</table>" & vbcrlf
        response.write "</fieldset>" & vbcrlf
     end if

	 end if

  oRequests.Close
  set oRequests = nothing 

end sub

'------------------------------------------------------------------------------
 function checkIsSeasonal(p_periodid)
   lcl_return = False

   if p_periodid <> "" then
      sSQL = "SELECT is_seasonal "
      sSQL = sSQL & " FROM egov_membership_periods "
      sSQL = sSQL & " WHERE orgid = " & iorgid
      sSQL = sSQL & " AND periodid = " & p_periodid

     	set oPeriod = Server.CreateObject("ADODB.Recordset")
    	 oPeriod.Open sSQL, Application("DSN"), 3, 1

      if not oPeriod.eof then
         if oPeriod("is_seasonal") = "" then
            lcl_return = False
         else
            lcl_return = oPeriod("is_seasonal")
         end if
      end if

      set oPeriod = nothing

   end if

   checkIsSeasonal = lcl_return

 end function

'------------------------------------------------------------------------------
sub showMembershipTerms(iOrgID, iMembershipID)

  dim sOrgID, sMembershipID, sTermsRequired

  sOrgID         = 0
  sMembershipID  = 0
  sTermsRequired = "N"

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iMembershipID <> "" then
     sMembershipID = clng(iMembershipID)
  end if

  response.write "<p>" & vbcrlf

  sSQL = "SELECT showterms, "
  sSQL = sSQL & " membershipterms "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND membershipid = " & sMembershipID
  sSQL = sSQL & " AND showterms = 1 "

 	set oMembershipTerms = Server.CreateObject("ADODB.Recordset")
	 oMembershipTerms.Open sSQL, Application("DSN"), 3, 1

  if not oMembershipTerms.eof then

     lcl_showterms = oMembershipTerms("showTerms")

     if lcl_showTerms then
        lcl_membershipterms = oMembershipTerms("membershipterms")
        sTermsRequired      = "Y"

        response.write "  <fieldset id=""membershipTerms"" class=""fieldset"">" & vbcrlf
        response.write "    <legend>Membership Terms&nbsp;</legend>" & vbcrlf
        response.write "    <p>" & lcl_membershipterms & "</p>" & vbcrlf
        response.write "  </fieldset>" & vbcrlf
        response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
        response.write "    <tr valign=""top"">" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <input type=""checkbox"" name=""termsRead"" id=""termsRead"" value=""Y"" onclick=""clearMsg('termsRead');"" />" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            I agree. You must check here to indicate that you agree to the above terms and conditions before continuing registration." & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "  </table>" & vbcrlf
     end if
  end if

  oMembershipTerms.close
  set oMembershipTerms = nothing

  response.write "  <input type=""hidden"" name=""termsRequiredToContinue"" id=""termsRequiredToContinue"" value=""" & sTermsRequired & """ size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "</p>" & vbcrlf

end sub
%>
