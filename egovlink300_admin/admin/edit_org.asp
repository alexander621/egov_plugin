<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: edit_org.asp
' AUTHOR: Steve Loar
' CREATED: 2/1/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the organization properties
'
' MODIFICATION HISTORY
' 1.0	2/01/07		Steve Loar - INITIAL VERSION
' 1.1	10/22/08	David Boyer - Added "UserBids (postings) - Notify Email"
' 1.2	8/28/2009	Steve Loar - Added Facility Rental Periods
' 1.3	7/13/2010	Steve Loar - Added the deactivation display code
' 1.4	01/17/2011	Steve Loar - Added View Full Site URL for Mobile
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim OrgId, sOrgname, sOrgCity, sOrgState, sOrgPublicWebsiteURL, sOrgEgovWebsiteURL, sOrgTopGraphicRightURL, sOrgTopGraphicLeftURL
Dim sOrgWelcomeMessage, sOrgActionLineDescription, sOrgHeaderSize, sOrgPaymentDescription, sOrgPaymentGateway, sOrgVirtualSiteName1
Dim sOrgRequestCalOn, sOrgRequestCalForm, sOrgActionName, sOrgPaymentName, sOrgDocumentName, sOrgCalendarName
Dim sOrgCustomButtonsOn, sOrgRegistration, sOrgTimeZoneID, sOrgDisplayFooter, sOrgDisplayMenu, sOrgCustomMenu
Dim sOrgFacilityHoldDays, sDefaultEmail, sDefaultPhone, sdefaultcity, sdefaultstate, sdefaultzip, sinternal_default_contact
Dim sinternal_default_phone, sinternal_default_email, sLatitude, sLongitude, sSeparate_index_catalog, sUsesAfter5Adjustment
Dim sUsesWeekDays, sAllowedUnresolvedDays, sdefaultareacode, sOrgGoogleAnalyticAccnt, spostings_email, sEvaluationFormId
Dim sFacilitySurveyFormId, smembershipcard_printer, sGoogleMapApiKey, sGoogleSearchID_documents, spostings_userbids_notifyemail
Dim sOrgPaymentOn, iResidentReservePeriod, iNonResidentReservePeriod, sPublicMenuOptCityHome_isEnabled, sPublicMenuOptEGovHome_isEnabled
Dim sPublicMenuOptCityHome_Label, sPublicMenuOptEGovHome_Label, lcl_checked_cityhome, lcl_checked_egovhome
Dim sRentalSurveyFormId, sOrgIsDeactivated, sShowMobileNavagation, sCitizenPaysFee
Dim sHasMobilePages, sOrgMobileWebsiteURL, sMobileLogo, sPublicDocumentsRoot, sViewFullSiteURL, sViewFullSiteLabel
Dim sPrivacyPolicyEgov, sPrivacyPolicyMobile

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin(session("UserID")) Then 
	response.redirect "../default.asp"
End If 

OrgId = CLng(session("orgid"))
GetOrgProperties( OrgId )

'Setup the BODY onload
lcl_onload = ""
lcl_onload = lcl_onload & "enableDisableOptionLabel('CITY');"
lcl_onload = lcl_onload & "enableDisableOptionLabel('EGOV');"
%>
<html>
<head>
	<title>E-Gov Administration Console {Maintain Organization Properties}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="javascript" src="tablesort.js"></script>

	<script language="javascript">
	<!--

		function Validate()	
		{
			document.frmOrganization.submit();
		}

		function openWin(page,p_wintype,p_width,p_height) 
		{
			if ((p_wintype=="")||(p_wintype==undefined)) 
			{
				s_wintype="new";
			}
			else
			{
				s_wintype=p_wintype;
			}
			if ((p_width=="")||(p_width==undefined)) 
			{
				s_width=600;
			}
			else
			{
				s_width=p_width;
			}
			if ((p_height=="")||(p_height==undefined)) 
			{
				s_height=700;
			}
			else
			{
				s_height=p_height;
			}

			lcl_screen_width  = screen.availWidth;
			lcl_screen_height = screen.availHeight;

			lcl_x_position = (lcl_screen_width/2) - (s_width/2);
			lcl_y_position = (lcl_screen_height/2) - (s_height/2);

			OpenWin = window.open(page, s_wintype, "status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width="+s_width+",height="+s_height+",left="+lcl_x_position+",top="+lcl_y_position);
			if (document.images) 
			{
				OpenWin.focus();
			}
		}

		function enableDisableOptionLabel(iField) 
		{
			var lcl_fieldid  = ""
			var lcl_disabled = false;

			if(iField != "") 
			{
				lcl_fieldid = iField.toLowerCase();
			}

			if(! document.getElementById("public_menuopt_" + lcl_fieldid + "home_enabled").checked) 
			{
				lcl_disabled = true;
			}

			document.getElementById("public_menuopt_" + lcl_fieldid + "home_label").disabled = lcl_disabled;
		}

	//-->
	</script>
</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>

<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
		<!--BEGIN: PAGE TITLE-->
 <%
'Set up page variables based on orgid.
lcl_page_title   = GetOrgName(orgid)
lcl_button_label = "Update"

lcl_page_title   = lcl_page_title & " Properties"
lcl_button_label = lcl_button_label & " Properties"

'Determine if there is a message to display
lcl_message = "&nbsp;"

If request("success") = "SU" Then 
	lcl_message = "<strong style=""color: #ff0000"">*** Successfully Updated ***</strong>"
ElseIf request("success") = "SA" then
	lcl_message = "<strong style=""color: #ff0000"">*** Successfully Created... ***</strong>"
End If 

 %>
		<p>
		 	<font size="+1"><strong><%=lcl_page_title%></strong></font><br />
		</p>
		<!--END: PAGE TITLE-->

		<table border="0" cellspacing="0" cellpadding="0" style="width: 900px">
			<tr>
				<td>
					<% displaybuttons "TOP", lcl_button_label %>
				</td>
				<td align="right">
					<%=lcl_message%>
				</td>
			</tr>
		</table>

		<!--BEGIN: EDIT FORM-->
		<form name="frmOrganization" action="org_update.asp" method="post">
		  <input type="hidden" name="orgid" value="<%=orgid%>" />

		<div id="tableorgshadow" class="shadow">
		<table id="tableorg" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
  			<tr><th>Properties</th></tr>
			<tr>
    			<td>
				<table>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" id="isdeactivated" name="isdeactivated" <%=sOrgIsDeactivated%> /> <strong>This Client is Deactivated</strong><br /><br /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Org Name:</td>
					  <td colspan="2"><input type="text" name="orgname" value="<%=sOrgname%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Org City:</td>
					  <td colspan="2"><input type="text" name="orgcity" value="<%=sOrgCity%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Virtual Directory:</td>
					  <td colspan="2"><input type="text" name="OrgVirtualSiteName" value="<%=sOrgVirtualSiteName1%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default City:</td>
					  <td colspan="2"><input type="text" name="defaultcity" value="<%=sdefaultcity%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Org State:</td>
					  <td colspan="2"><input type="text" name="orgstate" value="<%=sOrgState%>" size="10" maxlength="10" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default State:</td>
					  <td colspan="2"><input type="text" name="defaultstate" value="<%=sdefaultstate%>" size="10" maxlength="10" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default Zip:</td>
					  <td colspan="2"><input type="text" name="defaultzip" value="<%=sdefaultzip%>" size="15" maxlength="15" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default Area Code:</td>
					  <td colspan="2"><input type="text" name="defaultareacode" value="<%=sdefaultareacode%>" size="3" maxlength="3" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Org URL:</td>
					  <td><input type="text" name="OrgPublicWebsiteURL" value="<%=sOrgPublicWebsiteURL%>" size="50" maxlength="50" /></td>
					  <td><input type="checkbox" name="public_menuopt_cityhome_enabled" id="public_menuopt_cityhome_enabled" value="on" onclick="enableDisableOptionLabel('CITY');"<%=lcl_checked_cityhome%> />Enable Public Menu Option&nbsp;&nbsp;<input type="text" name="public_menuopt_cityhome_label" id="public_menuopt_cityhome_label" value="<%=sPublicMenuOptCityHome_Label%>" size="30" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Egov URL:</td>
					  <td><input type="text" name="OrgEgovWebsiteURL" value="<%=sOrgEgovWebsiteURL%>" size="50" maxlength="50" /></td>
					  <td><input type="checkbox" name="public_menuopt_egovhome_enabled" id="public_menuopt_egovhome_enabled" value="on" onclick="enableDisableOptionLabel('EGOV');"<%=lcl_checked_egovhome%> />Enable Public Menu Option&nbsp;&nbsp;<input type="text" name="public_menuopt_egovhome_label" id="public_menuopt_egovhome_label" value="<%=sPublicMenuOptEGovHome_Label%>" size="30" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Top Right Graphic:</td>
					  <td colspan="2"><input type="text" name="OrgTopGraphicRightURL" value="<%=sOrgTopGraphicRightURL%>" size="100" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Logo:</td>
					  <td colspan="2"><input type="text" name="OrgTopGraphicLeftURL" value="<%=sOrgTopGraphicLeftURL%>" size="100" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Logo Height:</td>
					  <td colspan="2"><input type="text" name="OrgHeaderSize" value="<%=sOrgHeaderSize%>" size="5" maxlength="5" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Welcome Msg:</td>
					  <td colspan="2"><input type="text" name="OrgWelcomeMessage" value="<%=sOrgWelcomeMessage%>" size="100" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Action Line Msg:</td>
					  <td colspan="2"><textarea class="orgedittext" name="OrgActionLineDescription"><%=sOrgActionLineDescription%></textarea></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Payment Msg:</td>
					  <td colspan="2"><textarea class="orgedittext" name="OrgPaymentDescription"><%=sOrgPaymentDescription%></textarea></td>
				  </tr>
				  <tr>
					  <!-- Do not remove this check box -->
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgPaymentOn" <% If sOrgPaymentOn Then response.write " checked=""checked"" "%> /> They have Payments (Do not remove this pick!)</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Payment Gateway:</td>
					  <td colspan="2">
						  <select name="OrgPaymentGateway">
						  <% ShowPaymentGateways sOrgPaymentGateway	%>
						  </select>
					  </td>
				  </tr>
				  <tr>
					<td wrap="nowrap" align="right">&nbsp;</td>
					<td colspan="2"><input type="checkbox" name="citizenpaysfee" <% = sCitizenPaysFee %> /> The Citizen Pays The Processing Fee</td>
				  </tr>
				  <tr>
					  <td>&nbsp;</td>
					  <td colspan="2"><strong>Credit Cards Accepted</strong></td>
				  </tr>
				  <% ShowCreditCardsSelections OrgId %>
				  <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgRequestCalOn" <% If sOrgRequestCalOn Then response.write " checked=""checked"" "%> /> They have a Calendar Request Form</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Calendar Req Form #:</td>
					  <td colspan="2"><input type="text" name="OrgRequestCalForm" value="<%=sOrgRequestCalForm%>" size="10" maxlength="10" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Class Evaluation Form:</td>
					  <td colspan="2"><input type="text" name="EvaluationFormId" value="<%=sEvaluationFormId%>" size="10" maxlength="10" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Facility Survey Form:</td>
					  <td colspan="2"><input type="text" name="facilitysurveyformid" value="<%=sFacilitySurveyFormId%>" size="10" maxlength="10" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Rental Survey Form:</td>
					  <td colspan="2"><input type="text" name="rentalsurveyformid" value="<%=sRentalSurveyFormId%>" size="10" maxlength="10" /></td>
				  </tr>
	<!--              <tr>
					  <td nowrap="nowrap" align="right">Action Line:</td><td colspan="2"><input type="text" name="OrgActionName" value="<%=sOrgActionName%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Payments:</td><td colspan="2"><input type="text" name="OrgPaymentName" value="<%=sOrgPaymentName%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Documents:</td><td colspan="2"><input type="text" name="OrgDocumentName" value="<%=sOrgDocumentName%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Calendar</td><td colspan="2"><input type="text" name="OrgCalendarName" value="<%=sOrgCalendarName%>" size="50" maxlength="50" /></td>
				  </tr>
	-->
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgCustomButtonsOn" <% If sOrgCustomButtonsOn Then response.write " checked=""checked"" "%> /> Custom Buttons</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgRegistration" <% If sOrgRegistration Then response.write " checked=""checked"" "%> /> Citizen Registration</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Time Zone:</td>
					  <td colspan="2">
						  <select name="OrgTimeZoneID">
							<% ShowTimeZones sOrgTimeZoneID	%>
						  </select>
					  </td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgDisplayFooter" <% If sOrgDisplayFooter Then response.write " checked=""checked"" "%> /> Display Footer Nav on Public</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgDisplayMenu" <% If sOrgDisplayMenu Then response.write " checked=""checked"" "%> /> Display Header and Nav on Public</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="OrgCustomMenu" <% If sOrgCustomMenu Then response.write " checked=""checked"" "%> /> Custom Menu</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Facility Hold Days:</td>
					  <td colspan="2"><input type="text" name="OrgFacilityHoldDays" value="<%=sOrgFacilityHoldDays%>" size="5" maxlength="5" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Resident Reserve<br />Period(Facilities):</td>
					  <td colspan="2"><input type="text" name="residentreserveperiod" value="<%=iResidentReservePeriod%>" size="3" maxlength="3" />&nbsp;Months</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Nonresident Reserve<br />Period(Facilities):</td>
					  <td colspan="2"><input type="text" name="nonresidentreserveperiod" value="<%=iNonResidentReservePeriod%>" size="3" maxlength="3" />&nbsp;Months</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default Email:</td>
					  <td colspan="2"><input type="text" name="DefaultEmail" value="<%=sDefaultEmail%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Default Phone:</td>
					  <td colspan="2"><input type="text" name="DefaultPhone" value="<%=sDefaultPhone%>" size="20" maxlength="20" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Internal Default Contact:</td>
					  <td colspan="2"><input type="text" name="internal_default_contact" value="<%=sinternal_default_contact%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Internal Default Phone:</td>
					  <td colspan="2"><input type="text" name="internal_default_phone" value="<%=sinternal_default_phone%>" size="20" maxlength="20" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Internal Default Email:</td>
					  <td colspan="2"><input type="text" name="internal_default_email" value="<%=sinternal_default_email%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Google Analytics Account:</td>
					  <td colspan="2"><input type="text" name="OrgGoogleAnalyticAccnt" value="<%=sOrgGoogleAnalyticAccnt%>" size="50" maxlength="50" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Google Map API Key:</td>
					  <td colspan="2"><input type="text" name="googlemapapikey" value="<%=sGoogleMapApiKey%>" size="110" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Google Search ID<br />(documents):</td>
					  <td colspan="2"><input type="text" name="googlesearchid_documents" value="<%=sGoogleSearchID_documents%>" size="50" maxlength="1000" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right" valign="top">Job/Bid Postings:<br />(Email your resume)</td>
					  <td align="top" colspan="2"><input type="text" name="postings_email" value="<%=spostings_email%>" size="50" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">UserBids Upload Email:</td>
					  <td colspan="2"><input type="text" name="postings_userbids_notifyemail" value="<%=spostings_userbids_notifyemail%>" size="50" maxlength="100" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right" valign="top">
						  Membership Card Printer:<br />
						  <a href="javascript:openWin('edit_org_printers.asp?orgid=<%=orgid%>','new',800,400);">[maintain printers]</a>
					  </td>
					  <td colspan="2">
						  <select name="membershipcard_printer">
							<% getPrinterOptions smembershipcard_printer %>
						  </select>
					  </td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="separate_index_catalog" <% If sSeparate_index_catalog Then response.write " checked=""checked"" "%> /> Separate Index Catalog</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Latitude:</td>
					  <td colspan="2"><input type="text" name="latitude" value="<%=sLatitude%>" size="13" maxlength="13" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">Longitude:</td>
					  <td colspan="2"><input type="text" name="longitude" value="<%=sLongitude%>" size="13" maxlength="13" /></td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="usesafter5adjustment" <% If sUsesAfter5Adjustment Then response.write " checked=""checked"" "%> /> Uses the &quot;After 5PM is Next Day&quot; Logic on Action Line</td>
				  </tr>
				  <tr>
					  <td nowrap="nowrap" align="right">&nbsp;</td>
					  <td colspan="2"><input type="checkbox" name="usesweekdays" <% If sUsesWeekDays Then response.write " checked=""checked"" "%> /> Counts Week Days Only on Action Line Calculations</td>
				  </tr>
				<tr>
					<td nowrap="nowrap" align="right">&nbsp;</td>
					<td colspan="2"><input type="checkbox" id="hasmobilepages" name="hasmobilepages" <%=sHasMobilePages%> /> Has Mobile Pages (This activates rerouting on E-Gov pages)</td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">Mobile URL:</td>
					<td colspan="2"><input type="text" id="orgmobilewebsiteurl" name="orgmobilewebsiteurl" value="<%=sOrgMobileWebsiteURL%>" size="100" maxlength="100" /></td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">Mobile Logo:</td>
					<td colspan="2"><input type="text" id="mobilelogo" name="mobilelogo" value="<%=sMobileLogo%>" size="100" maxlength="100" /></td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">Mobile Doc Path:</td>
					<td colspan="2"><input type="text" id="publicdocumentsroot" name="publicdocumentsroot" value="<%=sPublicDocumentsRoot%>" size="100" maxlength="100" /></td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">Mobile View Full Site URL:</td>
					<td colspan="2"><input type="text" id="viewfullsiteurl" name="viewfullsiteurl" value="<%=sViewFullSiteURL%>" size="100" maxlength="100" /></td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">Mobile View Full Site Label:</td>
					<td colspan="2"><input type="text" id="viewfullsitelabel" name="viewfullsitelabel" value="<%=sViewFullSiteLabel%>" size="50" maxlength="50" />
					 (Default: View Full Site)
					</td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right">&nbsp;</td>
					<td colspan="2"><input type="checkbox" id="showmobilenavagation" name="showmobilenavagation" <%=sShowMobileNavagation%> /> Show Mobile Navagation Bars</td>
				</tr>
				<tr>
					<td nowrap="nowrap" align="right" valign="top">Privacy Policy:</td>
					<td colspan="2">
         <fieldset style="border-radius: 6px; background-color: #eeeeee;">
<%
             displayPrivacyPolicyFields "EDIT", "EGOV", sPrivacyPolicyEgov
             displayPrivacyPolicyFields "EDIT", "MOBILE", sPrivacyPolicyMobile
%>
<!--
             Footer - E-Gov:<br />
             <input type="text" name="privacypolicy_egov" id="privacypolicy_egov" value="<% 'sPrivacyPolicyEgov%>" size="100" /><br />
             Footer - Mobile:<br />
             <input type="text" name="privacypolicy_mobile" id="privacypolicy_mobile" value="<% 'sPrivacyPolicyMobile%>" size="100" />
-->
         </fieldset>
     </td>
				</tr>
			</table>
        </td>
      </tr>
	</table>
			</div>
		</form>
		<!--END: EDIT FORM-->

		<!--BEGIN: FUNCTION LINKS-->
		<% displaybuttons "BOTTOM", lcl_button_label %>
		<!--END: FUNCTION LINKS-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>
<%
'------------------------------------------------------------------------------
' void GetOrgProperties iOrgId 
'------------------------------------------------------------------------------
Sub GetOrgProperties( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT OrgName, "
	sSql = sSql & " orgCity, "
	sSql = sSql & " OrgState, "
	sSql = sSql & " OrgPublicWebsiteURL, "
	sSql = sSql & " OrgEgovWebsiteURL, "
	sSql = sSql & " OrgTopGraphicRightURL, "
	sSql = sSql & " OrgTopGraphicLeftURL, "
	sSql = sSql & " OrgWelcomeMessage, "
	sSql = sSql & " OrgActionLineDescription, "
	sSql = sSql & " OrgHeaderSize, "
	sSql = sSql & " OrgPaymentDescription, "
	sSql = sSql & " googlemapapikey, "
 sSql = sSql & " googlesearchid_documents, "
	sSql = sSql & " ISNULL(OrgPaymentGateway,0) as OrgPaymentGateway, "
	sSql = sSql & " citizenpaysfee, "
	sSql = sSql & " OrgVirtualSiteName, "
	sSql = sSql & " OrgRequestCalOn, "
	sSql = sSql & " OrgRequestCalForm, "
	sSql = sSql & " OrgActionName, "
	sSql = sSql & " OrgTimeZoneID, "
	sSql = sSql & " OrgPaymentName, "
	sSql = sSql & " OrgDocumentName, "
	sSql = sSql & " OrgCalendarName, "
	sSql = sSql & " OrgCustomButtonsOn, "
	sSql = sSql & " OrgRegistration, "
	sSql = sSql & " OrgDisplayFooter, "
	sSql = sSql & " OrgDisplayMenu, "
	sSql = sSql & " OrgCustomMenu, "
	sSql = sSql & " OrgFacilityHoldDays, "
	sSql = sSql & " DefaultEmail, "
	sSql = sSql & " DefaultPhone, "
	sSql = sSql & " defaultcity, "
	sSql = sSql & " defaultstate, "
	sSql = sSql & " defaultzip, "
	sSql = sSql & " internal_default_contact, "
	sSql = sSql & " internal_default_phone, "
	sSql = sSql & " internal_default_email, "
	sSql = sSql & " separate_index_catalog, "
	sSql = sSql & " latitude, "
	sSql = sSql & " longitude, "
	sSql = sSql & " usesafter5adjustment, "
	sSql = sSql & " usesweekdays, "
	sSql = sSql & " defaultareacode, "
	sSql = sSql & " OrgGoogleAnalyticAccnt, "
	sSql = sSql & " postings_email, "
	sSql = sSql & " membershipcard_printer, "
	sSql = sSql & " EvaluationFormId, "
	sSql = sSql & " facilitysurveyformid, "
	sSql = sSql & " postings_userbids_notifyemail, "
	sSql = sSql & " OrgPaymentOn, "
	sSql = sSql & " residentreserveperiod, "
	sSql = sSql & " nonresidentreserveperiod, "
	sSql = sSql & " public_menuopt_cityhome_enabled, "
	sSql = sSql & " public_menuopt_egovhome_enabled, "
	sSql = sSql & " isdeactivated, "
	sSql = sSql & " public_menuopt_cityhome_label, "
	sSql = sSql & " public_menuopt_egovhome_label, "
	sSql = sSql & " rentalsurveyformid, "
	sSql = sSql & " hasmobilepages, "
	sSql = sSql & " ISNULL(orgmobilewebsiteurl,'') AS orgmobilewebsiteurl, "
	sSql = sSql & " ISNULL(mobilelogo,'') AS mobilelogo, "
	sSql = sSql & " ISNULL(publicdocumentsroot,'') AS publicdocumentsroot, "
	sSql = sSql & " ISNULL(viewfullsiteurl,'') AS viewfullsiteurl, "
	sSql = sSql & " ISNULL(viewfullsitelabel,'') AS viewfullsitelabel, "
	sSql = sSql & " showmobilenavagation, "
 sSql = sSql & " privacypolicy_egov, "
 sSql = sSql & " privacypolicy_mobile "
	sSql = sSql & " FROM Organizations "
	sSql = sSql & " WHERE orgid = " & iOrgId

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sOrgname                         = oRs("OrgName")
		sOrgCity                         = oRs("orgCity")
		sOrgState                        = oRs("OrgState")
		sOrgPublicWebsiteURL             = oRs("OrgPublicWebsiteURL")
		sOrgEgovWebsiteURL               = oRs("OrgEgovWebsiteURL")
		sOrgTopGraphicRightURL           = oRs("OrgTopGraphicRightURL")
		sOrgTopGraphicLeftURL            = oRs("OrgTopGraphicLeftURL")
		sOrgWelcomeMessage               = oRs("OrgWelcomeMessage")
		sOrgActionLineDescription        = oRs("OrgActionLineDescription")
		sOrgHeaderSize                   = oRs("OrgHeaderSize")
		sOrgPaymentDescription           = oRs("OrgPaymentDescription")
		sOrgPaymentGateway               = oRs("OrgPaymentGateway")
		sOrgVirtualSiteName1             = oRs("OrgVirtualSiteName")
		sOrgRequestCalOn                 = oRs("OrgRequestCalOn")
		sOrgPaymentOn                    = oRs("OrgPaymentOn")
		sOrgRequestCalForm               = oRs("OrgRequestCalForm")
		sOrgActionName                   = oRs("OrgActionName")
		sOrgPaymentName                  = oRs("OrgPaymentName")
		sOrgDocumentName                 = oRs("OrgDocumentName")
		sOrgCalendarName                 = oRs("OrgCalendarName")
		sOrgCustomButtonsOn              = oRs("OrgCustomButtonsOn")
		sOrgRegistration                 = oRs("OrgRegistration")
		sOrgTimeZoneID                   = oRs("OrgTimeZoneID")
		sOrgDisplayFooter                = oRs("OrgDisplayFooter")
		sOrgDisplayMenu                  = oRs("OrgDisplayMenu")
		sOrgCustomMenu                   = oRs("OrgCustomMenu")
		sOrgFacilityHoldDays             = oRs("OrgFacilityHoldDays")
		sDefaultEmail                    = oRs("DefaultEmail")
		sDefaultPhone                    = oRs("DefaultPhone")
		sdefaultcity                     = oRs("defaultcity")
		sdefaultstate                    = oRs("defaultstate")
		sdefaultzip                      = oRs("defaultzip")
		sdefaultareacode                 = oRs("defaultareacode")
		sinternal_default_contact        = oRs("internal_default_contact")
		sinternal_default_phone          = oRs("internal_default_phone")
		sinternal_default_email          = oRs("internal_default_email")
		sLatitude                        = oRs("latitude")
		sLongitude                       = oRs("longitude")
		sSeparate_index_catalog          = oRs("separate_index_catalog")
		sUsesAfter5Adjustment            = oRs("usesafter5adjustment")
		sUsesWeekDays                    = oRs("usesweekdays")
		'sAllowedUnresolvedDays           = oRs("allowedunresolveddays")
		sOrgGoogleAnalyticAccnt          = oRs("OrgGoogleAnalyticAccnt")
		spostings_email                  = oRs("postings_email")
		smembershipcard_printer          = oRs("membershipcard_printer")
		sEvaluationFormId                = oRs("EvaluationFormId")
		sFacilitySurveyFormId            = oRs("facilitysurveyformid")
		sRentalSurveyFormId	       		    = oRs("rentalsurveyformid")
		sGoogleMapApiKey                 = oRs("googlemapapikey")
  sGoogleSearchID_documents        = oRs("googlesearchid_documents")
		spostings_userbids_notifyemail   = oRs("postings_userbids_notifyemail")
		iResidentReservePeriod           = oRs("residentreserveperiod")
		iNonResidentReservePeriod        = oRs("nonresidentreserveperiod")
		sPublicMenuOptCityHome_isEnabled = oRs("public_menuopt_cityhome_enabled")
		sPublicMenuOptEGovHome_isEnabled = oRs("public_menuopt_egovhome_enabled")
		sPublicMenuOptCityHome_Label     = oRs("public_menuopt_cityhome_label")
		sPublicMenuOptEGovHome_Label     = oRs("public_menuopt_egovhome_label")
		sViewFullSiteURL                 = oRs("viewfullsiteurl")
		sViewFullSiteLabel	              = oRs("viewfullsitelabel")
  sPrivacyPolicyEgov               = oRs("privacypolicy_egov")
  sPrivacyPolicyMobile             = oRs("privacypolicy_mobile")

		If oRs("hasmobilepages") Then
			sHasMobilePages = " checked=""checked"" "
		Else
			sHasMobilePages = ""
		End If 

		sOrgMobileWebsiteURL = oRs("orgmobilewebsiteurl")
		sMobileLogo = oRs("mobilelogo")
		sPublicDocumentsRoot = oRs("publicdocumentsroot")
		'response.write "sPublicDocumentsRoot = " & sPublicDocumentsRoot & "<br />"

		If sPublicMenuOptCityHome_isEnabled Then 
			lcl_checked_cityhome = " checked=""checked"""
		Else 
			lcl_checked_cityhome = ""
		End If 

		If sPublicMenuOptEGovHome_isEnabled Then 
			lcl_checked_egovhome = " checked=""checked"""
		Else 
			lcl_checked_egovhome = ""
		End If 

		If sPublicMenuOptCityHome_Label = "" Or IsNull(sPublicMenuOptCityHome_Label) Then 
			sPublicMenuOptCityHome_Label = "City Home"
		End If 

		If sPublicMenuOptEGovHome_Label = "" Or IsNull(sPublicMenuOptEGovHome_Label) Then 
			sPublicMenuOptEGovHome_Label = "E-Gov Home"
		End If 

		If oRs("isdeactivated") Then
			sOrgIsDeactivated = " checked=""checked"" "
		Else
			sOrgIsDeactivated = ""
		End If 

		If oRs("showmobilenavagation") Then
			sShowMobileNavagation = " checked=""checked"" "
		Else
			sShowMobileNavagation = ""
		End If 

		If oRs("citizenpaysfee") Then 
			sCitizenPaysFee = " checked=""checked"" "
		Else
			sCitizenPaysFee = ""
		End If

	End If 

	oRs.Close
	set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowTimeZones( iOrgTimeZoneID )
'------------------------------------------------------------------------------
Sub ShowTimeZones( ByVal iOrgTimeZoneID )
	Dim sSql, oRs

	sSql = "SELECT TimeZoneID, TZName FROM timezones ORDER BY TZName"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF  
		response.write vbcrlf & "<option value=""" & oRs("TimeZoneID") & """"
		If clng(iOrgTimeZoneID) = clng(oRs("TimeZoneID")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("TZName") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' void ShowPaymentGateways( iPaymentGateway )
'------------------------------------------------------------------------------
Sub ShowPaymentGateways( ByVal iPaymentGateway )
	Dim sSql, oRs

	sSql = "SELECT paymentgatewayid, paymentgatewayname, admingatewayname FROM egov_payment_gateways ORDER BY paymentgatewayid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<option value=""0""" 
	If clng(iPaymentGateway) = clng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">None</option>"

	Do While Not oRs.EOF  
		response.write vbcrlf & "<option value=""" & oRs("paymentgatewayid") & """"
		If clng(iPaymentGateway) = clng(oRs("paymentgatewayid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("paymentgatewayid") & ". " & oRs("paymentgatewayname") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

End Sub 


'----------------------------------------------------------------------------------------
' void ShowCreditCardsSelections( OrgId )
'----------------------------------------------------------------------------------------
Sub ShowCreditCardsSelections( ByVal OrgId )
	Dim sSql, oRs

	sSql = "SELECT creditcardid, creditcard FROM creditcards ORDER BY creditcard"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<tr><td>&nbsp;</td><td colspan=""2""><input type=""checkbox"" name=""creditcardid"" value=""" & oRs("creditcardid") & """"
		If OrgHasCreditCard( OrgId, oRs("creditcardid") ) Then
			response.write " checked=""checked"" "
		End If 
		response.write " /> &nbsp; " & oRs("creditcard") & "</td></tr>"
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'----------------------------------------------------------------------------------------
' boolean OrgHasCreditCard( OrgId, iCreditCardId )
'----------------------------------------------------------------------------------------
Function OrgHasCreditCard( ByVal OrgId, ByVal iCreditCardId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(creditcardid) AS hits FROM egov_organizations_to_creditcards WHERE orgid = " & OrgId
	sSql = sSql & " AND creditcardid = " & iCreditCardId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > CLng(0) Then 
			OrgHasCreditCard = True 
		Else
			OrgHasCreditCard = False 
		End If 
	Else
		OrgHasCreditCard = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
Sub getPrinterOptions( ByVal p_value )
	Dim sSql, oRs 

	sSql = "SELECT printerid, printer_name, default_printer, active_flag "
	sSql = sSql & " FROM egov_membershipcard_printers "
	sSql = sSql & " WHERE active_flag = 1 "
	sSql = sSql & " ORDER BY UPPER(printer_name), printerid "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	if not oRs.eof then
		lcl_selected             = ""
		lcl_default_printer_text = ""

		While Not  oRs.eof 
			'Determine if this is the option selected in the list.
			if p_value <> "" then
				if CLng(oRs("printerid")) = CLng(p_value) then
					lcl_selected = " selected"
				else
					lcl_selected = ""
				end if
			end if

			'If this is the default printer then display text in option name
			if oRs("default_printer") then
		  		lcl_default_printer_text = " [DEFAULT PRINTER]"
			else
  				lcl_default_printer_text = ""
			end if

			response.write "<option value=""" & oRs("printerid") & """" & lcl_selected & ">" & oRs("printer_name") & lcl_default_printer_text & "</option>" & vbcrlf

			oRs.movenext
		Wend 
	End If 

End Sub 

'------------------------------------------------------------------------------
sub displayButtons( ByVal p_section, ByVal p_label )

	response.write "<div id=""functionlinks"">" & vbcrlf

	If UCase(p_section) = "TOP" Then 
		response.write "<input type=""button"" class=""button"" name=""return"" id=""return"" value=""<< Return to Feature Selection"" onclick=""location.href='featureselection.asp?orgid=" & request("orgid") & "'"" />" & vbcrlf
	End If 

	response.write "<input type=""button"" class=""button"" name=""sAction"" id=""sAction"" value=""" & p_label & """ onclick=""Validate();"" />" & vbcrlf
	response.write "</div>" & vbcrlf

end sub
%>
