<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: new_org.asp
' AUTHOR: Steve Loar
' CREATED: 2/23/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the new organization setup screen
'
' MODIFICATION HISTORY
' 1.0   2/23/07   Steve Loar - INITIAL VERSION
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
Dim sUsesWeekDays, sAllowedUnresolvedDays, sdefaultareacode, sOrgGoogleAnalyticAccnt, sGoogleSearchID_documents
Dim sPublicMenuOptCityHome_isEnabled, sPublicMenuOptEGovHome_isEnabled, sPublicMenuOptCityHome_Label, sPublicMenuOptEGovHome_Label
Dim lcl_checked_cityhome, lcl_checked_egovhome, sShowMobileNavagation
Dim sHasMobilePages, sOrgMobileWebsiteURL, sMobileLogo, sPublicDocumentsRoot, sViewFullSiteURL, sViewFullSiteLabel
Dim sPrivacyPolicyEgov, sPrivacyPolicyMobile

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then 
	response.redirect "../default.asp"
End If 

OrgId = CLng(0) ' New Org
GetOrgProperties
%>
<html>
<head>
 	<title>E-Gov Administration Console {New Organization}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
 	<link rel="stylesheet" type="text/css" href="admin.css" />

 	<script language="javascript" src="tablesort.js"></script>

<script language="javascript">
<!--

	function Validate() {
		 document.frmOrganization.submit();
	}

 function enableDisableOptionLabel(iField) {
   var lcl_fieldid  = ""
   var lcl_disabled = false;

   if(iField != "") {
      lcl_fieldid = iField.toLowerCase();
   }

   if(! document.getElementById("public_menuopt_" + lcl_fieldid + "home_enabled").checked) {
      lcl_disabled = true;
   }

   document.getElementById("public_menuopt_" + lcl_fieldid + "home_label").disabled = lcl_disabled;
 }
//-->
</script>
</head>
<body>
<%
	  ShowHeader sLevel
%>
<!--#Include file="../menu/menu.asp"--> 
<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
 <%
  'Page Title

	response.write "<p>" & vbcrlf
	response.write "<font size=""+1""><strong>New Organization Properties</strong></font><br />" & vbcrlf
	response.write "</p>" & vbcrlf

  'Buttons
   display_buttons "TOP", "Create Organization"
 %>
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
                  <td wrap="nowrap" align="right">City Name:</td>
                  <td colspan="2"><input type="text" name="orgname" value="<%=sOrgname%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">City:</td>
                  <td colspan="2"><input type="text" name="orgcity" value="<%=sOrgCity%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Virtual Directory:</td>
                  <td colspan="2"><input type="text" name="OrgVirtualSiteName" value="<%=sOrgVirtualSiteName1%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default City:</td>
                  <td colspan="2"><input type="text" name="defaultcity" value="<%=sdefaultcity%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default State:</td>
                  <td colspan="2"><input type="text" name="defaultstate" value="<%=sdefaultstate%>" size="10" maxlength="10" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default Zip:</td>
                  <td colspan="2"><input type="text" name="defaultzip" value="<%=sdefaultzip%>" size="15" maxlength="15" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default Area Code:</td>
                  <td colspan="2"><input type="text" name="defaultareacode" value="<%=sdefaultareacode%>" size="3" maxlength="3" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default Email:</td>
                  <td colspan="2"><input type="text" name="DefaultEmail" value="<%=sDefaultEmail%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Default Phone:</td>
                  <td colspan="2"><input type="text" name="DefaultPhone" value="<%=sDefaultPhone%>" size="20" maxlength="20" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Org URL:</td>
                  <td><input type="text" name="OrgPublicWebsiteURL" value="<%=sOrgPublicWebsiteURL%>" size="50" maxlength="50" /></td>
                  <td><input type="checkbox" name="public_menuopt_cityhome_enabled" id="public_menuopt_cityhome_enabled" value="on" onclick="enableDisableOptionLabel('CITY');"<%=lcl_checked_cityhome%> />Enable Public Menu Option&nbsp;&nbsp;<input type="text" name="public_menuopt_cityhome_label" id="public_menuopt_cityhome_label" value="<%=sPublicMenuOptCityHome_Label%>" size="30" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Egov URL:</td>
                  <td><input type="text" name="OrgEgovWebsiteURL" value="<%=sOrgEgovWebsiteURL%>" size="50" maxlength="50" /></td>
                  <td><input type="checkbox" name="public_menuopt_egovhome_enabled" id="public_menuopt_egovhome_enabled" value="on" onclick="enableDisableOptionLabel('EGOV');"<%=lcl_checked_egovhome%> />Enable Public Menu Option&nbsp;&nbsp;<input type="text" name="public_menuopt_egovhome_label" id="public_menuopt_egovhome_label" value="<%=sPublicMenuOptEGovHome_Label%>" size="30" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Top Right Graphic:</td>
                  <td colspan="2"><input type="text" name="OrgTopGraphicRightURL" value="<%=sOrgTopGraphicRightURL%>" size="100" maxlength="100" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Logo:</td>
                  <td colspan="2"><input type="text" name="OrgTopGraphicLeftURL" value="<%=sOrgTopGraphicLeftURL%>" size="100" maxlength="100" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Logo Height:</td>
                  <td colspan="2"><input type="text" name="OrgHeaderSize" value="<%=sOrgHeaderSize%>" size="5" maxlength="5" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Welcome Msg:</td>
                  <td colspan="2"><input type="text" name="OrgWelcomeMessage" value="<%=sOrgWelcomeMessage%>" size="100" maxlength="100" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Action Line Msg:</td>
                  <td colspan="2"><textarea class="orgedittext" name="OrgActionLineDescription"><%=sOrgActionLineDescription%></textarea></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Payment Msg:</td>
                  <td colspan="2"><textarea class="orgedittext" name="OrgPaymentDescription"><%=sOrgPaymentDescription%></textarea></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgPaymentOn" <% If OrgPaymentOn Then response.write " checked=""checked"" "%> /> They have Payments (Do not remove this pick!)</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Payment Gateway:</td>
                  <td colspan="2">
                      <select name="OrgPaymentGateway">
                        <% ShowPaymentGateways %>
                      </select>
                  </td>
              </tr>
			  <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="citizenpaysfee" /> The Citizen Pays The Processing Fee</td>
              </tr>
              <tr>
                  <td>&nbsp;</td>
                  <td><strong>Credit Cards Accepted</strong></td>
              </tr>
              <% ShowCreditCardsSelections	%>
              <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgRequestCalOn" <% If sOrgRequestCalOn Then response.write " checked=""checked"" "%> /> They have a Calendar Request Form</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Calendar Req Form #:</td>
                  <td colspan="2"><input type="text" name="OrgRequestCalForm" value="<%=sOrgRequestCalForm%>" size="10" maxlength="10" /></td>
              </tr>
								
<!--              <tr>
                  <td wrap="nowrap" align="right">Action Line:</td><td><input type="text" name="OrgActionName" value="<%=sOrgActionName%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Payments:</td><td><input type="text" name="OrgPaymentName" value="<%=sOrgPaymentName%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Documents:</td><td><input type="text" name="OrgDocumentName" value="<%=sOrgDocumentName%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Calendar</td><td><input type="text" name="OrgCalendarName" value="<%=sOrgCalendarName%>" size="50" maxlength="50" /></td>
              </tr>
-->
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgCustomButtonsOn" <% If sOrgCustomButtonsOn Then response.write " checked=""checked"" "%> /> Custom Buttons</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgRegistration" <% If sOrgRegistration Then response.write " checked=""checked"" "%> /> Citizen Registration</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Time Zone:</td>
                  <td colspan="2">
                      <select name="OrgTimeZoneID">
                        <% ShowTimeZones sOrgTimeZoneID	%>
                  </select>
                  </td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgDisplayFooter" <% If sOrgDisplayFooter Then response.write " checked=""checked"" "%> /> Display Footer Nav on Public</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgDisplayMenu" <% If sOrgDisplayMenu Then response.write " checked=""checked"" "%> /> Display Header and Nav on Public</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="OrgCustomMenu" <% If sOrgCustomMenu Then response.write " checked=""checked"" "%> /> Custom Menu</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Facility Hold Days:</td>
                  <td colspan="2"><input type="text" name="OrgFacilityHoldDays" value="<%=sOrgFacilityHoldDays%>" size="5" maxlength="5" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact First Name:</td>
                  <td colspan="2"><input type="text" name="firstname" value="" size="24" maxlength="24" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact Last Name:</td>
                  <td colspan="2"><input type="text" name="lastname" value="" size="25" maxlength="25" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact Username:</td>
                  <td colspan="2"><input type="text" name="username" value="" size="32" maxlength="32" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact Password:</td>
                  <td colspan="2"><input type="text" name="password" value="" size="16" maxlength="16" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact Phone:</td>
                  <td colspan="2"><input type="text" name="internal_default_phone" value="<%=sinternal_default_phone%>" size="20" maxlength="20" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Contact Email:</td>
                  <td colspan="2"><input type="text" name="internal_default_email" value="<%=sinternal_default_email%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Google Analytics Account:</td>
                  <td colspan="2"><input type="text" name="OrgGoogleAnalyticAccnt" value="<%=sOrgGoogleAnalyticAccnt%>" size="50" maxlength="50" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Google Map API Key:</td>
                  <td colspan="2"><input type="text" name="googlemapapikey" value="<%= GetDefaultGoogleMapApiKey() %>" size="110" maxlength="100" /></td>
              </tr>
        				  <tr>
					             <td nowrap="nowrap" align="right">Google Search ID<br />(documents):</td>
           					  <td colspan="2"><input type="text" name="googlesearchid_documents" value="<%=sGoogleSearchID_documents%>" size="50" maxlength="1000" /></td>
        				  </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="separate_index_catalog" <% If sSeparate_index_catalog Then response.write " checked=""checked"" "%> /> Separate Index Catalog</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Latitude:</td>
                  <td colspan="2"><input type="text" name="latitude" value="<%=sLatitude%>" size="13" maxlength="13" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">Longitude:</td>
                  <td colspan="2"><input type="text" name="longitude" value="<%=sLongitude%>" size="13" maxlength="13" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="usesafter5adjustment" <% If sUsesAfter5Adjustment Then response.write " checked=""checked"" "%> /> Uses the &quot;After 5PM is Next Day&quot; Logic on Action Line</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right">&nbsp;</td>
                  <td colspan="2"><input type="checkbox" name="usesweekdays" <% If sUsesWeekDays Then response.write " checked=""checked"" "%> /> Counts Week Days Only on Action Line Calculations</td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right" valign="top">Job/Bid Postings<br>(Email your resume):</td>
                  <td colspan="2"><input type="text" name="postings_email" value="<%=spostings_email%>" size="50" maxlength="100" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right" valign="top">UserBids Upload Email:</td>
                  <td colspan="2"><input type="text" name="postings_userbids_notifyemail" value="<%=spostings_userbids_notifyemail%>" size="50" maxlength="100" /></td>
              </tr>
              <tr>
                  <td wrap="nowrap" align="right" valign="top">
                      Membership Card Printer:<br>
                      <a href="javascript:openWin('edit_org_printers.asp?orgid=<%=orgid%>','new',800,400);">[maintain printers]</a>
                  </td>
                  <td colspan="2">
                      <select name="membershipcard_printer">
                        <% getPrinterOptions %>
                      </select>
                  </td>
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
					<td colspan="2"><input type="text" id="viewfullsitelabel" name="viewfullsitelabel" value="<%=sViewFullSiteLabel%>" size="50" maxlength="50" /></td>
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
             displayPrivacyPolicyFields "NEW", "EGOV", sPrivacyPolicyEgov
             displayPrivacyPolicyFields "NEW", "MOBILE", sPrivacyPolicyMobile
%>
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

  <% display_buttons "", "Create Organization" %>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------
Sub GetOrgProperties( )

	sOrgname                  = "Demo City"
	sOrgCity                  = "Demo"
	sOrgPublicWebsiteURL      = "http://www.cityofeclink.com"
	sOrgEgovWebsiteURL        = "http://www.egovlink.com/eclink"
	sOrgTopGraphicRightURL    = "http://www.egovlink.com/eclink/custom/images/eclink/top_image_bg.jpg"
	sOrgTopGraphicLeftURL     = "http://www.egovlink.com/eclink/custom/images/eclink/logo.jpg"
	sOrgWelcomeMessage        = "Welcome to the City of EC Link e-Government Services Website!"

	'sOrgActionLineDescription = "<p>The OnLine Service Request system allows visitors to request information, submit requests for service, or submit comments for review. The Service Request system covers a wide variety of city departments and services.</p>"
	sOrgActionLineDescription = "<br /><br /><p>The OnLine Service Request system allows visitors to request information, submit requests for "
	sOrgActionLineDescription = sOrgActionLineDescription & "service, or submit comments for review.</p><p>A tracking number is assigned to "
	sOrgActionLineDescription = sOrgActionLineDescription & "each new Service Request. Using this tracking number, the submitter can check "
	sOrgActionLineDescription = sOrgActionLineDescription & "back on the website at any time to view the current status of the request.</p>"
	sOrgActionLineDescription = sOrgActionLineDescription & "<p>NOTE: Requests entered after hours will be received by city staff by 8:30 a.m. "
	sOrgActionLineDescription = sOrgActionLineDescription & "the next business day.</p><p>It is not necessary to register to submit a "
	sOrgActionLineDescription = sOrgActionLineDescription & "Service Request.  However if you do register then you will be able to track all "
	sOrgActionLineDescription = sOrgActionLineDescription & "the Service Requests submitted under your registration ID.</p><p>Submitting an "
	sOrgActionLineDescription = sOrgActionLineDescription & "email address with your Service Request will result in confirmation that the "
	sOrgActionLineDescription = sOrgActionLineDescription & "Request has been received and will allow you to receive status updates for "
	sOrgActionLineDescription = sOrgActionLineDescription & "your Request</p>"

	sOrgHeaderSize                   = 80

	sOrgPaymentDescription           = "<p>The Payment Center allows visitors to make payments online. The Payment Center system covers a wide "
	sOrgPaymentDescription           = sOrgPaymentDescription & "variety of city departments and services as indicated to the left. To enter a "
	sOrgPaymentDescription           = sOrgPaymentDescription & "payment, please click a link on the left and follow the instructions on the "
	sOrgPaymentDescription           = sOrgPaymentDescription & "payment form.</p>"

	sOrgPaymentGateway               = 2
	sOrgVirtualSiteName1             = "eclink"
	sOrgRequestCalOn                 = 0
	sOrgRequestCalForm               = ""
	sOrgCustomButtonsOn              = 0
	sOrgRegistration                 = 1
	sOrgTimeZoneID                   = 1
	sOrgDisplayFooter                = 1
	sOrgDisplayMenu                  = 1
	sOrgCustomMenu                   = 1
	sOrgFacilityHoldDays             = ""
	sDefaultEmail                    = ""
	sDefaultPhone                    = ""
	sdefaultcity                     = "Demo"
	sdefaultstate                    = "OH"
	sdefaultzip                      = "45223"
	sdefaultareacode                 = "513"
	sinternal_default_contact        = "Peter Selden"
	sinternal_default_phone          = "5136814030"
	sinternal_default_email          = "pselden@eclink.com"
	sLatitude                        = "39.241077"
	sLongitude                       = "-84.347748"
	sSeparate_index_catalog          = 1
	sUsesAfter5Adjustment            = 0
	sUsesWeekDays                    = 0
	OrgGoogleAnalyticAccnt           = ""
	spostings_email                  = ""
	spostings_userbids_notifyemail   = ""
	sPublicMenuOptCityHome_isEnabled = True
	sPublicMenuOptEGovHome_isEnabled = True
	sPublicMenuOptCityHome_Label     = "City Home"
	sPublicMenuOptEGovHome_Label     = "E-Gov Home"

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
  
	sHasMobilePages       = ""
	sOrgMobileWebsiteURL  ="http://m.egovlink.com/newcity"
	sMobileLogo           = "http://m.egovlink.com/newcity/custom/images/newcity/logo.png"
	sPublicDocumentsRoot  = "http://m.egovlink.com/public_documents300/newcity/"
	sViewFullSiteURL      = "http://demo2.besavvy2.newcity.com?full=true"
	sViewFullSiteLabel    = "View Full Site"
	sShowMobileNavagation = " checked=""checked"" "

 sPrivacyPolicyEgov   = ""
 sPrivacyPolicyMobile = ""

End Sub 


'------------------------------------------------------------------------------
Sub ShowTimeZones( ByVal iOrgTimeZoneID )
	Dim sSql, oRs

	sSql = "Select TimeZoneID, TZName from timezones order by TZName"

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
		
	oRs.close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
Sub ShowPaymentGateways( )
	Dim sSql, oRs

	sSql = "SELECT paymentgatewayid, paymentgatewayname, admingatewayname FROM egov_payment_gateways ORDER BY paymentgatewayid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<option value=""0"">None</option>"

	Do While Not oRs.EOF  
		response.write vbcrlf & "<option value=""" & oRs("paymentgatewayid") & """"
		response.write ">" & oRs("paymentgatewayid") & ". " & oRs("paymentgatewayname") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' Function GetOrgName()
'------------------------------------------------------------------------------
Function GetOrgName( ByVal iorgid )
	Dim sSql, oRs

	sSql = "SELECT orgname FROM organizations WHERE orgid = " & iorgid
'		response.write sSql
'		response.end

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetOrgName = oRs("orgname")
	End If
		
	oRs.close
	Set oRs = Nothing

End Function 


'---------------------------------------------------------------------------------------
sub getPrinterOptions()
	Dim sSql, oRs, lcl_default_printer_text

	sSql = "SELECT printerid, printer_name, default_printer, active_flag "
	sSql = sSql & " FROM egov_membershipcard_printers "
	sSql = sSql & " WHERE active_flag = 1 "
	sSql = sSql & " ORDER BY UPPER(printer_name), printerid "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		lcl_selected             = ""
		lcl_default_printer_text = ""

		Do While Not oRs.EOF
			'If this is the default printer then display text in option name
			If oRs("default_printer") Then 
				lcl_default_printer_text = " [DEFAULT PRINTER]"
			Else 
				lcl_default_printer_text = ""
			End If 
			response.write "<option value=""" & oRs("printerid") & """>" & oRs("printer_name") & lcl_default_printer_text & "</option>" & vbcrlf
			oRs.movenext
		Loop 
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' string GetDefaultGoogleMapApiKey()
'------------------------------------------------------------------------------
Function GetDefaultGoogleMapApiKey()
	Dim sSql, oRs

	sSql = "SELECT googlemapapikey FROM organizations WHERE orgid = 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultGoogleMapApiKey = oRs("googlemapapikey")
	Else
		GetDefaultGoogleMapApiKey = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void ShowCreditCardsSelections
'------------------------------------------------------------------------------
Sub ShowCreditCardsSelections()
	Dim sSql, oRs

	sSql = "SELECT creditcardid, creditcard FROM creditcards ORDER BY creditcard"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<tr><td>&nbsp;</td><td><input type=""checkbox"" name=""creditcardid"" value=""" & oRs("creditcardid") & """ /> &nbsp; " & oRs("creditcard") & "</td></tr>"
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
Sub display_buttons( ByVal p_topbottom, ByVal p_button_label )

	If UCase(p_topbottom) = "TOP" Then 
		response.write "<input type=""button"" class=""button"" name=""Return"" value=""<< Return to Feature Selection"" onclick=""location.href='featureselection.asp?orgid=" & request("orgid") & "'"" /> &nbsp; "
	End If 

	response.write "<input type=""button"" class=""button"" name=""sAction"" value=""" & p_button_label & """ onclick=""Validate();"" />" & vbcrlf

End Sub 



%>
