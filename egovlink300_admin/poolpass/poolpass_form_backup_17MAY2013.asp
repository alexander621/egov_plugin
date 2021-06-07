<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_rates.asp
' AUTHOR: Steve Loar
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/31/06 Steve Loar - Code added
' 1.1	 10/05/06	Steve Loar - Security, Header and nav changed
' 1.2  09/08/08 David Boyer - Added Membership Renewals
' 1.3  03/09/09 David Boyer - Added Alternate Layout feature
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iMembershipId, oMembership, iPeriodId

 sLevel     = "../"  'Override of value from common.asp
 lcl_onload = ""

 if not userhaspermission(session("userid"),"purchase membership") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 set oMembership = New classMembership

'Set the membershipid to the one for pools, later this will be a dropdown
 oMembership.SetMembershipId( "pool" )
 iMembershipId = oMembership.MembershipId

 iUserId       = getFirstUserID()
 sUserType     = ""
 sResidentDesc = ""
 iPeriodId     = clng(request("periodid"))
 iRateDesc     = request("rateDesc")
 sSearchName   = ""
 sResults      = ""
 sSearchStart  = -1

 if request("userid") <> "" then
   	iUserId = request("userid")
 end if

 sUserType     = GetUserResidentType(iUserid)
 sResidentDesc = GetResidentTypeDesc(sUserType)

 if request("periodid") = "" then
   	iPeriodId = oMembership.GetFirstMembershipPeriodId( iMembershipId )
 end if

 if request("rateDesc") = "" then
   	iRateDesc = oMembership.getFirstMembershipRateOption_AltLayout( sUserType, iMembershipId)
 end if

'Determine if the membership period selected is/isn't seasonal
 lcl_isSeasonal = checkIsSeasonal(iPeriodID)

 session("RedirectPage") = "../poolpass/poolpass_form.asp?userid=" & iUserId & "&iMembershipId=" & iMembershipId & "&periodid=" & iPeriodId
 session("RedirectLang") = "Return to Membership Purchase"

'See if a search term was passed
 if request("searchname") <> "" then
   	sSearchName = request("searchname")
 end if

 if request("results") <> "" then
   	sResults = request("results")
 end if

 if request("searchStart") <> "" then
 	  sSearchStart = request("searchStart")
 end if

'Check for org features
 lcl_orghasfeature_membership_renewals            = orghasfeature("membership_renewals")
 lcl_orghasfeature_purchase_membership_alt_layout = orghasfeature("purchase_membership_alt_layout")

'Check for user permissions
 lcl_userhaspermission_membership_renewals = userhaspermission(session("userid"),"membership_renewals")

'Check to see if the org has the feature turned-on and the user has it assigned
 lcl_membershiprenewals_feature = "N"

 if lcl_orghasfeature_membership_renewals AND lcl_userhaspermission_membership_renewals then
    lcl_membershiprenewals_feature = "Y"
 end if
%>
<html>
<head>
	<title>E-Gov Administration Console {Membership Purchase}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="style_pool.css" />

<style type="text/css">
.fieldset {
  border-radius: 5px;
}
</style>

<script type="text/javascript">
  <!--

	function SearchCitizens() {
   var optiontext;
   var optionchanged;
   var searchtext      = document.getElementById('searchname').value;
   var searchchanged   = searchtext.toLowerCase();
   var lcl_searchstart = document.getElementById('searchStart').value;

   iSearchStart = parseInt(lcl_searchstart) + 1;

		for (x=iSearchStart; x < document.getElementById('userid').length ; x++)	{
 			optiontext    = document.getElementById('userid').options[x].text;
 			optionchanged = optiontext.toLowerCase();
 			if (optionchanged.indexOf(searchchanged) != -1) {
    				document.getElementById('userid').selectedIndex    = x;
				    document.getElementById('results').value           = 'Possible Match Found.';
    				document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
				    document.getElementById('searchStart').value       = x;
    				document.getElementById('purchaseMembership').submit();
				    return;
			 }
		}
		document.getElementById('results').value           = 'No Match Found.';
		document.getElementById('searchresults').innerHTML = 'No Match Found.';
		document.getElementbyId('searchStart').value       = -1;
	}

 function ClearSearch() {
   document.getElementById('searchStart').value = -1;
 }

	function UserPick() {
 		document.getElementById('searchname').value        = '';
	 	document.getElementById('results').value           = '';
		 document.getElementById('searchresults').innerHTML = '';
 		document.getElementById('searchStart').value       = -1;
<%
  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "document.getElementById('rateDesc').value = '';" & vbcrlf
  end if
%>
	 	document.getElementById('purchaseMembership').submit();
	}

	function ContinuePurchase() {
   document.getElementById('purchaseMembership').action = 'select_members.asp';
   document.getElementById('purchaseMembership').submit();
	}

	function EditUser(iUserId) {
		 location.href='../dirs/update_citizen.asp?userid=' + iUserId;
	}

	function NewUser()	{
 		location.href='../dirs/register_citizen.asp';
	}

function renewPass(iPassId) {
//  if (confirm("Delete Pass #" + iPassId + "?")) {
   			location.href='select_members.asp?poolpassid=' + iPassId;
//		}
//  inlineMsg(document.getElementById("button_renew_"+iPassId).id,'<strong>Coming Soon: </strong>Renew Membership Option for PoolPassID: '+iPassId,8,'button_renew_'+iPassId);
}

function updateRateID(iRateID) {
  document.getElementById('rateid').value = iRateID;
}

function submitPoolPassForm() {
  document.getElementById('purchaseMembership').action = 'poolpass_form.asp';
  document.getElementById('purchaseMembership').submit();
}

function enableDisableContinueButton(iRowValue) {
<%
  if lcl_orghasfeature_purchase_membership_alt_layout then
     lcl_radio_id          = "periodid_"
     'lcl_onload_rateperiod = iPeriodID
     lcl_onload = lcl_onload & "enableDisableContinueButton('" & iPeriodID & "');"
  else
     lcl_radio_id          = "rateid_"
     'lcl_onload_rateperiod = request("rateid")
     lcl_onload = lcl_onload & "enableDisableContinueButton('" & request("rateid") & "');"
  end if

  response.write "  if(iRowValue=='') {" & vbcrlf
  response.write "     document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "  }else{" & vbcrlf
  response.write "     if(document.getElementById('" & lcl_radio_id & "'+iRowValue)) {"& vbcrlf
  response.write "        if(document.getElementById('" & lcl_radio_id & "'+iRowValue).checked==true) {" & vbcrlf
  response.write "           document.getElementById('continueButton').disabled=false;" & vbcrlf
  response.write "        }else{" & vbcrlf
  response.write "           document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "        }" & vbcrlf
  response.write "     }else{" & vbcrlf
  response.write "        document.getElementById('continueButton').disabled=true;" & vbcrlf
  response.write "     }" & vbcrlf
  response.write "  }" & vbcrlf
%>
}
 //-->
 </script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<body onload="<%=lcl_onload%>">
<%
  response.write "<form name=""purchaseMembership"" id=""purchaseMembership"" method=""post"" action=""poolpass_form_backup_17MAY2013.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""results"" id=""results"" value="""" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""searchStart"" id=""searchStart"" value="""     & sSearchStart   & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iMembershipId"" id=""iMembershipId"" value=""" & iMembershipId  & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""isSeasonal"" id=""isSeasonal"" value="""       & lcl_isSeasonal & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""usertype"" id=""usertype"" value="""           & sUserType      & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iuserid"" id=""iuserid"" value="""             & iUserId        & """ />" & vbcrlf

  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "  <input type=""hidden"" name=""rateid"" id=""rateid"" value=""" & request("rateid") & """ />" & vbcrlf
  end if

  response.write "<div id=""content"">" & vbcrlf
  response.write "<div id=""poolcentercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "  <tr><th colspan=""2"" align=""left"">Purchase a " & session("sOrgName") & "&nbsp;" & oMembership.GetMembershipName() & "&nbsp;Membership</th></tr>" & vbcrlf
  response.write "	 <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "       			Select a registered Citizen from the drop down list, select the pass they want and then press the " & vbcrlf
  response.write "          <strong>Continue with Purchase</strong> button.  If their name is not on the list then select " & vbcrlf
  response.write "          <strong>New User</strong> to add them to the list. If their information is incorrect then select " & vbcrlf
  response.write "          <strong>Edit User Profile</strong>." & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "		<tr>" & vbcrlf
  response.write "      <td colspan=""2"" nowrap=""nowrap"">" & vbcrlf
  response.write "      				<p>" & vbcrlf
  response.write "          		Name Search:" & vbcrlf
  response.write "            <input type=""text"" name=""searchname"" id=""searchname"" value=""" & sSearchName & """ size=""25"" maxlength=""50"" onchange=""ClearSearch();"" />" & vbcrlf
  response.write "          		<input type=""button"" name=""searchButton"" id=""searchButton"" class=""button"" value=""Search"" onclick=""SearchCitizens();"" />" & vbcrlf
  response.write "					     		<span id=""searchresults"">" & sResults & "</span>" & vbcrlf
  response.write "   					  		<br /><div id=""searchtip"">(last name, first name)</div>" & vbcrlf
  response.write "   		 				</p>" & vbcrlf
  response.write "   		 				Select Name:" & vbcrlf
  response.write "          <select name=""userid"" id=""userid"" onchange=""UserPick();"">" & vbcrlf
                              showUserDropdown(iUserId)
  response.write "    						</select>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "		<tr>" & vbcrlf
  response.write "      <td colspan=""2"" nowrap=""nowrap"">" & vbcrlf
  response.write "      				<input type=""button"" name=""newUserButton"" id=""newUserButton"" class=""button"" onclick=""NewUser();"" value=""New User"" />&nbsp;&nbsp;" & vbcrlf
  response.write "          <input type=""button"" name=""editUserButton"" id=""editUserButton"" class=""button"" onclick=""EditUser(" & iUserId & ");"" value=""Edit User Profile"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
                    showUserInfo iUserId, sUserType, sResidentDesc
                    'showMembershipTerms session("orgid"), iMembershipID
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

 'BEGIN: Check to see which layout the org has turned-on ----------------------
  response.write "<div id=""purchasePicks"">" & vbcrlf

  if lcl_orghasfeature_purchase_membership_alt_layout then
     response.write "<div id=""rightpicks_altlayout"">" & vbcrlf

    'Membership Rates Dropdown
     response.write "  <p id=""membershipRatesDropdown"">" & vbcrlf
     response.write "    Membership Types:<br />" & vbcrlf
                         oMembership.ShowMembershipRates_AltLayout sUserType, _
                                                                   iMembershipId, _
                                                                   iRateDesc
     response.write "  </p>" & vbcrlf

    'Membership Periods Checkboxes
     response.write "  <div class=""shadow"">" & vbcrlf
                         oMembership.ShowPeriodPicksForMembership_AltLayout iMembershipId, _
                                                                            iPeriodID, _
                                                                            sUserType, _
                                                                            iRateDesc
     response.write "  </div>" & vbcrlf
  else
     response.write "<div id=""rightpicks"">" & vbcrlf

    'Membership Periods Dropdown
     response.write "  <p>" & vbcrlf
     response.write "    Membership Period:&nbsp;" & vbcrlf
                         oMembership.ShowPeriodPicksForMembership iMembershipId, _
                                                                  iPeriodId
     response.write "  </p>" & vbcrlf

    'Membership Rates Checkboxes
     response.write "  <div class=""shadow"">" & vbcrlf
                         oMembership.ShowMembershipResidentRates sUserType, _
                                                                 iMembershipId, _
                                                                 iPeriodId
     response.write "  </div>" & vbcrlf
  end if

  response.write "  <br />" & vbcrlf

 'Determine if "Continue with Purchase" button is displayed.
 'Is the citizen selected is a "non-resident"?
 'If "yes" then check to see if the org has the "nonresidentlimit" feature assigned to it.
 'If "yes" then total the number records, for non-residents, on egov_poolpasspurchases for the CURRENT "membership period" selected.
 'If the total records on egov_poolpasspurchases is LESS THAN the "Non-Resident Cap" set 
  response.write "  <input type=""button"" name=""continue"" id=""continueButton"" class=""button"" value=""Continue with Purchase"" onclick=""ContinuePurchase();"" />" & vbcrlf

  response.write "</div>" & vbcrlf
 'END: Check to see which layout the org has turned-on ------------------------

  response.write "</div>" & vbcrlf

 'BEGIN: Renewal Memberships --------------------------------------------------
  if lcl_membershiprenewals_feature = "Y" then

    'Show the renewal membership record if one exists for the user selected
     showRenewalMembership iUserID, _
                           iMembershipID, _
                           iPeriodID
  end if
 'END: Renewal Memberships ----------------------------------------------------

  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

  set oMembership = nothing 

'------------------------------------------------------------------------------
sub showUserInfo(iUserID, sUserType, sResidentDesc)

  sSQL = "SELECT userfname, "
  sSQL = sSQL & " userlname, "
  sSQL = sSQL & " useraddress, "
  sSQL = sSQL & " useraddress2, "
  sSQL = sSQL & " userunit, "
  sSQL = sSQL & " usercity, "
  sSQL = sSQL & " userstate, "
  sSQL = sSQL & " userzip, "
  sSQL = sSQL & " usercountry, "
  sSQL = sSQL & " useremail, "
  sSQL = sSQL & " userhomephone, "
  sSQL = sSQL & " userworkphone, "
  sSQL = sSQL & " userfax, "
  sSQL = sSQL & " userbusinessname, "
  sSQL = sSQL & " userpassword, "
  sSQL = sSQL & " userregistered, "
  sSQL = sSQL & " residenttype, "
  sSQL = sSQL & " registrationblocked, "
  sSQL = sSQL & " blockeddate, "
  sSQL = sSQL & " blockedadminid, "
  sSQL = sSQL & " blockedexternalnote, "
  sSQL = sSQL & " blockedinternalnote "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & iUserID

 	set oUser = Server.CreateObject("ADODB.Recordset")
	 oUser.Open sSQL, Application("DSN"), 3, 1

  if not oUser.eof then
    'Build the City, State, and Zip display value
     lcl_cityStateZip = ""

     if oUser("usercity") <> "" then
        lcl_cityStateZip = oUser("usercity")
     end if

     if oUser("userstate") <> "" then
        if lcl_cityStateZip <> "" then
           lcl_cityStateZip = lcl_cityStateZip & ", " & oUser("userstate")
        else
           lcl_cityStateZip = oUser("userstate")
        end if
     end if

     if oUser("userzip") <> "" then
        if lcl_cityStateZip <> "" then
           lcl_cityStateZip = lcl_cityStateZip & " " & oUser("userzip")
        else
           lcl_cityStateZip = oUser("userzip")
        end if
     end if

     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Name:</td>" & vbcrlf
     response.write "      <td width=""60%"">" & oUser("userfname") & " " & oUser("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Email:</td>" & vbcrlf
     response.write "      <td>" & oUser("useremail") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Phone:</td>" & vbcrlf
     response.write "      <td>" & FormatPhone(oUser("userhomephone")) & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Address:</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write            oUser("useraddress") & "<br />" 

     if oUser("useraddress2") = "" then
      		response.write oUser("useraddress2") & "<br />" & vbcrlf
     end if

     response.write            lcl_cityStateZip & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Business:</td>" & vbcrlf
     response.write "      <td>" & oUser("userbusinessname") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 	oUser.close
 	set oUser = nothing

end sub

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

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf

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

        response.write "          <fieldset id=""membershipTerms"" class=""fieldset"">" & vbcrlf
        response.write "            <legend>Membership Terms&nbsp;</legend>" & vbcrlf
        response.write "            <p>" & lcl_membershipterms & "</p>" & vbcrlf
        response.write "            <p><input type=""checkbox"" name=""termsRead"" id=""termsRead"" value=""Y"" /> I have read terms</p>" & vbcrlf
        response.write "          </fieldset>" & vbcrlf
     end if
  end if

  oMembershipTerms.close
  set oMembershipTerms = nothing

  response.write "          <input type=""text"" name=""termsRequiredToContinue"" id=""termsRequiredToContinue"" value=""" & sTermsRequired & """ size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub showUserDropdown(iUserID)

  sSQL = "SELECT userid, userfname, userlname, useraddress "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND isdeleted = 0 "
  sSQL = sSQL & " AND userregistered = 1 "
  sSQL = sSQL & " AND headofhousehold = 1 "
  sSQL = sSQL & " AND userfname IS NOT NULL "
  sSQL = sSQL & " AND userlname IS NOT NULL "
  sSQL = sSQL & " AND userfname <> '' "
  sSQL = sSQL & " AND userlname <> '' "
  sSQL = sSQL & " ORDER BY userlname, userfname "

 	set oResident = Server.CreateObject("ADODB.Recordset")
	 oResident.Open sSQL, Application("DSN"), 3, 1

  if not oResident.eof then
     while not oResident.eof
        if CLng(iUserId) = CLng(oResident("userid")) then
           lcl_selected = " selected=""selected"" "
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & oResident("userid") & """" & lcl_selected & ">" & oResident("userlname") & ", " & oResident("userfname") & " &ndash; " & oResident("useraddress") & "</option>" & vbcrlf

        oResident.movenext
     wend
  end if

 	oResident.close
	 set oResident = nothing

end sub

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
		sSQL = sSQL & " P.paymentlocation, "
  sSQL = sSQL & " P.paymentresult, "
  sSQL = sSQL & " M.membershipdesc, "
  sSQL = sSQL & " MP.period_desc, "
  sSQL = sSQL & " P.previous_poolpassid "
		sSQL = sSQL & " FROM egov_users U, "
  sSQL = sSQL &      " egov_memberships M, "
  sSQL = sSQL &      " egov_membership_periods MP, "
  sSQL = sSQL &      " egov_poolpasspurchases P "
		sSQL = sSQL & " WHERE P.orgid = " & session("orgid")
  sSQL = sSQL & " AND P.paymentresult <> 'Pending' "
		sSQL = sSQL & " AND P.paymentresult <> 'Declined' "
  sSQL = sSQL & " AND U.userid = P.userid "
  sSQL = sSQL & " AND P.membershipid = M.membershipid "
  sSQL = sSQL & " AND P.periodid = MP.periodid "
  'sSQL = sSQL & " AND (P.paymentdate >= '" & fromDate & "' AND P.paymentdate < '" & toDate & "') "
  sSQL = sSQL & " AND U.userid = "       & p_userid
  sSQL = sSQL & " AND P.membershipid = " & p_membershipid
  sSQL = sSQL & " AND P.periodid = "     & p_periodid
		sSQL = sSQL & " ORDER BY P.poolpassid "
response.write sSQL & "<br />"
 	set oRequests = Server.CreateObject("ADODB.Recordset")
	 oRequests.Open sSQL, Application("DSN"), 3, 1

  if not oRequests.eof then
     bgcolor            = "#eeeeee"
     lcl_showheader_row = "Y"
     lcl_close_renewal  = "N"
     lcl_isRateEnabled  = true

   		while not oRequests.eof
  	   		iPassCount = iPassCount + 1
   		  	iRowCount  = iRowCount + 1

       'Retrieve the RATE info
        getRateInfo oRequests("rateid"), _
                    lcl_rate_description, _
                    lcl_rate_residenttype

       'Determine if the Renewal column/button are displayed
        if lcl_membershiprenewals_feature = "Y" then
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

              lcl_style_width = " style=""width:600px"""

              response.write "<div id=""membershipRenewals"">" & vbcrlf
              response.write "<fieldset class=""fieldset"">" & vbcrlf
              response.write "  <legend>Membership(s) Available for Renewal&nbsp;</legend>" & vbcrlf
              response.write "<br />" & vbcrlf
              response.write "<div class=""shadow""" & lcl_style_width & ">" & vbcrlf
              response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist""" & lcl_style_width & ">" & vbcrlf
              response.write "  <tr class=""tablelist"" align=""left"">" & vbcrlf
              response.write "      <th>&nbsp;</th>" & vbcrlf
              response.write "      <th style=""text-align: center"">Pass<br />ID</th>" & vbcrlf
              response.write "      <th>Membership Type</th>" & vbcrlf
              response.write "      <th>Purchase Date</th>" & vbcrlf
              response.write "      <th>Purchaser</th>" & vbcrlf
              response.write "      <th>Payment<br />Amount</th>" & vbcrlf
              response.write "      <th>Payment Method</th>" & vbcrlf
              response.write "      <th>Status</th>" & vbcrlf
              response.write "      <th id=""column_renewal"">Renewal<br />of Pass ID</th>" & vbcrlf
              'response.write "      <th colspan=""2"">&nbsp;</th>" & vbcrlf
              response.write "      <th>&nbsp;</th>" & vbcrlf
              response.write "  </tr>" & vbcrlf
           else
              lcl_showheader_row = lcl_showheader_row
              lcl_close_renewal  = lcl_close_renewal
           end if

'           lcl_row_mouseover = " onMouseOver=""mouseOverRow(this);"""
'           lcl_row_mouseout  = " onMouseOut=""mouseOutRow(this);"""
'           lcl_td_mouseover  = " onMouseOver=""tooltip.show('click to edit');"""
'           lcl_td_mouseout   = " onMouseOut=""tooltip.hide();"""
           lcl_row_mouseover = ""
           lcl_row_mouseout  = ""
           lcl_td_mouseover  = ""
           lcl_td_mouseout   = ""
           lcl_onclick       = " onClick=""location.href='poolpass_details.asp?iPoolPassId=" & oRequests("poolpassid") & "';"""

        			response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & bgcolor & """" & lcl_row_mouseover & lcl_row_mouseout & ">" & vbcrlf
  	      		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">&nbsp;</td>" & vbcrlf
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""left"">" & oRequests("poolpassid") & "</td>" & vbcrlf
     	   		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("membershipdesc") & " &ndash; " & oRequests("period_desc") & "<br />" & lcl_rate_description & "</td>" & vbcrlf
      		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & DateValue(oRequests("paymentdate")) & "</td>" & vbcrlf
  	      		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("userfname") & " " & oRequests("userlname") & "</td>" & vbcrlf
        			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & formatcurrency(oRequests("paymentamount"),2) & "</td>" & vbcrlf

        			cTotalAmount = cTotalAmount + CDbl(oRequests("paymentamount"))

        			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & MakeProper(oRequests("paymentlocation")) & " &mdash; " & MakeProper(oRequests("paymenttype")) & "</td>" & vbcrlf
		        	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" &  oRequests("paymentresult") & "</td>" & vbcrlf
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" &  oRequests("previous_poolpassid") & "</td>" & vbcrlf
     	   		response.write "      <td><input type=""button"" name=""renew"" id=""button_renew_" & oRequests("poolpassid") & """ value=""Renew"" onclick=""renewPass('" & oRequests("poolpassid") & "');"" class=""button"" /></td>" & vbcrlf
   		     	response.write "</tr>" & vbcrlf

           bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

        end if

     			oRequests.movenext
   		wend

     if lcl_close_renewal = "Y" then
      		response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</fieldset>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</p>" & vbcrlf
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
      sSQL = sSQL & " WHERE orgid = " & session("orgid")
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
%>