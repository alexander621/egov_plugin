<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
' FILENAME: select_members.asp
' AUTHOR: Steve Loar
' CREATED: 01/31/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0 01/31/06  Steve Loar - Code added
' 1.1 09/09/08  David Boyer - Added Membership Renewals
' 1.2 01/30/09  David Boyer - Now create a memberid for purchases of a rate that is set to be a "punchcard" so a photo ID can be created.
' 1.3 03/04/09  David Boyer - Added Membership Rate "message" field.
'
'------------------------------------------------------------------------------
	Dim iUserId,  sUserType, iRateId, sMessage, sRateName, iMaxSignUps, nAmount, iFamilyCount, oMembership
	Dim iMembershipId, iPeriodId
	
	sLevel = "../" ' Override of value from common.asp

	If Not UserHasPermission( Session("UserId"), "purchase membership" ) Then
		  response.redirect sLevel & "permissiondenied.asp"
	End If 

	Set oMembership = New classMembership

'Determine if this is a purchase or a renewal.
'  - If the "poolpassid" is NULL then it is a "purchase".
'  - Otherwise, if there is a value then it is a "renewal".
 if trim(request("poolpassid")) <> "" then
    lcl_poolpassid       = trim(request("poolpassid"))
    lcl_renewal_text     = "Renewal"
    lcl_renewaldesc_text = "to renew"

    getPoolPassInfo lcl_poolpassid, iUserId, iRateId, iMembershipId, iPeriodID, lcl_isSeasonal

   'Using this global function gives up the ResidentType which we need for our "sUserType" variable.
    getRateInfo iRateId, lcl_rate_description, lcl_rate_residenttype

    sUserType = lcl_rate_residenttype

   'If this is a renewal membership purchase then the startdate will be:
   '  1. If the current date <= expiration date of the previous membership then the new start date will be defaulted to the the previous expirationdate+1(day)
   '  2. If the current date > expireation date of the previous membership then the new start date will be defaulted to the current date
   'If this is a new membership purchase then startdate will be the current date.
    lcl_startdate      = datevalue(oMembership.getMembershipStartDate(lcl_poolpassid))
    lcl_expirationdate = datevalue(getPoolPassExpirationDate(lcl_poolpassid))

 else
    lcl_poolpassid       = ""
    lcl_renewal_text     = "Purchase"
    lcl_renewaldesc_text = ""
   	iUserId              = request("userid")
   	sUserType            = request("usertype")
   	iRateId              = request("rateid")
   	iMembershipId        = CLng(request("iMembershipId"))
   	iPeriodID            = request("periodid")
    lcl_isSeasonal       = request("isseasonal")
    lcl_startdate        = date()
    lcl_expirationdate   = oMembership.getExpirationDate(iPeriodID, lcl_startdate)
 end if

 oMembership.MembershipId = iMembershipId 
	sRateName     = ""
	sMessage      = ""
	iMaxSignUps   = 1
	nAmount       = 001.00
	iFamilyCount  = 0

 if DATEDIFF("d",lcl_expirationdate,date()) > 0 then
    lcl_startdate = date()
 end if

'Get the rate name and message
 GetRateData iRateId, sRateName, sMessage, nAmount, iMaxSignUps, iAttendanceTypeID, iIsPunchcard, iPunchcardLimit

'Default the attendancetypeid to Member (1) if for whatever reason it is empty
 if iAttendanceTypeID = "" OR isnull(iAttendanceTypeID) then
    iAttendanceTypeID = 1
 end if
	
'Set the return to url
	session("RedirectPage") = "../poolpass/select_members.asp?userid=" & iUserId & "&usertype=" & sUserType & "&rateid=" & iRateId & "&iMembershipId=" & iMembershipId & "&periodid=" & iPeriodId
	session("RedirectLang") = "Return to Pool Pass Purchase"
%>
<html>
<head>
	<title>E-Gov Pool Membership <%=lcl_renewal_text%>: Select Members (<%=lcl_expirationdate%> - <%=date()%>)</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="style_pool.css" />
	<link rel="stylesheet" type="text/css" href="poolpass.css" />
 <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

 <script src="../scripts/isvaliddate.js"></script>
 <script src="../scripts/tooltip_new.js"></script>
 <script src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
  <!--

  	function UpdateFamily( iUserId )
	{
		//location.href='../dirs/family_members.asp?userid=' + sUserId;
		location.href='../dirs/family_list.asp?userid=' + iUserId;
	}

	function GoBack(sUserId) {
   <%
     if UCASE(lcl_renewal_text) = "RENEWAL" then
       'Retrieve the search criteria
        orderBy       = session("orderBy")
        subTotals     = session("subTotals")
        showDetail    = session("showDetail")
        fromDate      = session("fromDate")
        toDate        = session("toDate")
        sUserlname    = session("userlname")
        iMembershipId = session("imembershipid")
       	iPeriodId     = session("iperiodid")

        lcl_search_criteria = "?"
        lcl_search_criteria = lcl_search_criteria & "orderBy="           & orderBy
        lcl_search_criteria = lcl_search_criteria & "&subTotals="        & subTotals
        lcl_search_criteria = lcl_search_criteria & "&showDetail="       & showDetail
        lcl_search_criteria = lcl_search_criteria & "&fromDate="         & fromDate
        lcl_search_criteria = lcl_search_criteria & "&toDate="           & toDate
        lcl_search_criteria = lcl_search_criteria & "&userlname=' + escape(""" & sUserlname & """) + '"
        lcl_search_criteria = lcl_search_criteria & "&membershipid="     & iMembershipId
        lcl_search_criteria = lcl_search_criteria & "&periodid="         & iperiodid

        response.write "location.href='poolpass_list.asp" & lcl_search_criteria & "';" & vbcrlf
     else
        response.write "location.href='poolpass_form.asp?userid=' + sUserId + '&periodid=" & iPeriodId & "';" & vbcrlf
     end if
   %>
 }

	function AddFamilyMember()
	{
		//var rege = /(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d/;
		var rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var Ok = rege.test(document.addFamily.birthdate.value);

		if (document.addFamily.firstname.value == "")
		{
			alert("Please input a first name.");
			document.addFamily.firstname.focus();
			return;
		}
		else if (document.addFamily.lastname.value == "")
		{
			alert("Please input a last name.");
			document.addFamily.lastname.focus();
			return;
		}
		else if (document.addFamily.birthdate.value == "")
		{
			alert("Please input a birth date.");
			document.addFamily.birthdate.focus();
			return;
		}
		else if (! Ok)
		{
			alert("Birth date should be in the format of MM/DD/YYYY.  Please enter it again.");
			document.addFamily.birthdate.focus();
			return;
		}
		else
		{
			document.addFamily.submit();
		}
	}

	function ContinuePurchase() {
   //set var checkbox_choices to zero
		 var checkbox_choices = 0;

 		if (document.addPoolPass.passIncl.length > 1) {
    	//Loop from zero to the one minus the number of checkbox button selections
    	for (counter = 0; counter < document.addPoolPass.passIncl.length; counter++)	{	
      				//If a checkbox has been selected it will return true
     				 //(If not it will return false)
      				if (document.addPoolPass.passIncl[counter].checked) {
              checkbox_choices = checkbox_choices + 1;
          }
     }
   }	else	{
  			if (document.addPoolPass.passIncl.checked)	{
     				checkbox_choices = checkbox_choices + 1;
  			}
 		}

 		if (checkbox_choices == 0) {
   	 		alert( "Please select at least one name to include on the pass.");
		 	   return;
 		} else if (checkbox_choices > document.addPoolPass.imaxsignups.value) {
    			alert("You are limited to " + document.addPoolPass.imaxsignups.value + " person(s) for this type of pass. \n Please deselect some, or choose another type of pass.");
		 	   return;
 		} else if (document.addPoolPass.imaxsignups.value < 5 && checkbox_choices != document.addPoolPass.imaxsignups.value) {
    			alert("This pass is for " + document.addPoolPass.imaxsignups.value + " person(s). \n Please select more, or choose another type of pass.");
		 	   return;
 		} else {

       lcl_valid   = "Y";
       lcl_message = "";

       //Validate the start date
       lcl_start_date = document.getElementById("startdate").value;
       if (lcl_start_date!="") {
           	var Ok = isValidDate(lcl_start_date);
		<% if session("orgid") = 60 then %>
			if (lcl_start_date.indexOf("/1/") < 0)
			{
				Ok = false;
			}
		<% end if %>


           	if(! Ok)	{
		<% if session("orgid") = 60 then %>
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Membership Start Date\" must be in a date format and <b>on the first of the month</b>.<br /><span style=\"color:#800000;\">(i.e. mm/dd/yyyy)</span>";
	       <% else %>
               lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Membership Start Date\" must be in a date format.<br /><span style=\"color:#800000;\">(i.e. mm/dd/yyyy)</span>";
		<% end if %>
               lcl_valid = "N";
               lcl_expiration_date = "<%=lcl_expirationdate%>";
            <% if lcl_poolpassid <> "" then %>
            } else {
               lcl_expiration_date = "<%=lcl_expirationdate%>";
               if (Date.parse(lcl_start_date) <= Date.parse('<%=lcl_expirationdate%>')) {
                   lcl_message = lcl_message + "<strong>Invalid Value: </strong>The \"Membership Start Date\" (" + lcl_start_date + ") can not be less than or equal to the Expiration Date (<%=lcl_expirationdate%>) of the previous membership.";
                   lcl_valid = "N";
               }
            <% end if %>
            }
       } else {
           lcl_message = lcl_message + "<strong>Required Field Missing: </strong>Membership Start Date.";
           lcl_valid = "N";
       }

       if (lcl_valid=="Y") {
        			//alert("Pass is OK.");
		 	       document.addPoolPass.submit();
       } else {
           inlineMsg(document.getElementById("startdate").id,lcl_message,8,'startdate');
           return false;
       }
 		}
	}

 function doCalendar(ToFrom)  {
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
	}
  //-->
</script>
</head>

 <% ShowHeader sLevel %>
 <!--#Include file="../menu/menu.asp"--> 

<body>
<div id="content">
  <div id="centercontent"> <!-- This makes the tables skinny, take out to widen -->

<input type="button" class="button" onclick="javascript:GoBack('<%=iUserId%>');" value="<< Back" />
<br /><br />

<div class="shadow">
<table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
 	<tr>
    <%
      if iAttendanceTypeID <> 4 then
         lcl_colspan = " valign=""top"" colspan=""2"""
      else
         lcl_colspan = ""
      end if

      response.write "      <th" & lcl_colspan & ">" & vbcrlf
      response.write "          You have selected " & lcl_renewaldesc_text &  " the "
      response.write            oMembership.GetMembershipPeriodName( iPeriodId )
      response.write            "&nbsp;" & sRateName & "&nbsp;"
      response.write            oMembership.GetMembershipName()
      response.write            "&nbsp;Membership "

     'If this has been set up with an Attendance Type of GROUP (egov_pool_attendancetypes.attendancetypeid = 4)
     'Then allow the user to be able to enter a "head-count" and a "per head-count price" to calculate the total
      'if iAttendanceTypeID = 4 then
      '   response.write "          for:"
      '   response.write "      </th>" & vbcrlf
      '   response.write "      <th>" & vbcrlf
      '   response.write "          <fieldset>" & vbcrlf
      '   response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""background-color: #93bee1;"">" & vbcrlf
      '   response.write "            <tr>" & vbcrlf
      '   response.write "                <td><strong>Number of People: </strong></td>" & vbcrlf
      '   response.write "                <td><input type=""text"" name=""p_group_count"" size=""5"" maxlength=""5"" /></td>" & vbcrlf
      '   response.write "            </tr>" & vbcrlf
      '   response.write "            <tr>" & vbcrlf
      '   response.write "                <td><strong>with a price per person of: $</strong></td>" & vbcrlf
      '   response.write "                <td><input type=""text"" name=""p_group_price"" size=""8"" maxlength=""10"" /></td>" & vbcrlf
      '   response.write "            </tr>" & vbcrlf
      '   response.write "            <tr>" & vbcrlf
      '   response.write "                <td align=""right""><strong>Final Price: </strong>" & vbcrlf
      '   response.write "                <td id=""final_group_price""></td>" & vbcrlf
      '   response.write "            </tr>" & vbcrlf
      '   response.write "          </table>" & vbcrlf
      '   response.write "          </fieldset>" & vbcrlf
      'else
         response.write "          at " & FormatCurrency(nAmount,2) & vbcrlf
         response.write "          <input type=""hidden"" name=""p_group_count"" value=""1"" size=""4"" maxlength=""5"" />" & vbcrlf
         response.write "          <input type=""hidden"" name=""p_group_price"" value=""" & nAmount & """ size=""5"" maxlength=""10"" />" & vbcrlf
      'end if

      response.write "      </th>" & vbcrlf
    %>
  </tr>
	 <tr><td colspan="2"><%=sMessage%></td></tr>
		<tr><td colspan="2">Please select family members to be included in the membership from the list below.</td></tr>
</table>
</div>

<form method="post" name="addPoolPass" action="poolpass_add.asp">
 	<input type="hidden" name="userid" value="<%=iUserId%>" />
	 <input type="hidden" name="usertype" value="<%=sUserType%>" />
	 <input type="hidden" name="rateid" value="<%=iRateId%>" />
	 <input type="hidden" name="imaxsignups" value="<%=iMaxSignUps%>" />
	 <input type="hidden" name="imembershipid" value="<%=oMembership.MembershipId%>" />
	 <input type="hidden" name="iperiodid" value="<%=iPeriodId%>" />
  <input type="hidden" name="poolpassid" value="<%=lcl_poolpassid%>" />
  <input type="hidden" name="isSeasonal" value="<%=lcl_isSeasonal%>" />
  <input type="hidden" name="isPunchcard" id="isPunchcard" value="<%=iIsPunchcard%>" />
  <input type="hidden" name="punchcard_limit" id="punchcard_limit" value="<%=iPunchcardLimit%>" />

<p><input type="button" class="button" onclick="UpdateFamily('<%=iUserId%>');" value="Update Family Members" /></p>

<div class="shadow">
<table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
  <tr>
      <th>Include</th>
      <th>First Name</th>
      <th>Last Name</th>
    <%
      if lcl_poolpassid <> "" then
         response.write "<th>Member ID</th>" & vbcrlf
      end if
    %>
      <th>Relation</th>
      <th>Birthdate</th>
      <th>&nbsp;</th>
  </tr>
 			<% 
	 	  'GetPurchaserInfo(iUserId)
				  iFamilyCount = GetFamilyMembers(iUserId, iRateId, lcl_poolpassid)

      if not lcl_isSeasonal then
         lcl_startdate_title = "Start Date"
      else
         lcl_startdate_title = "Season"
      end if
	 		%>
</table>
</div>
<% if session("orgid") = "26" then %>
<div class="shadow">
<table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
  <tr>
      <th>Purchase Notes</th>
  </tr>
		<tr>
      <td> 
      	<textarea name="purchasenotes" id="purchasenotes"></textarea>
      </td>
     </tr>
 </table>
</div>
<% end if %>
<div class="shadow">
<table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
  <tr>
      <th>Payment Type</th>
      <th>Payment Location</th>
      <th style="text-align:center">Membership<br /><%=lcl_startdate_title%></th>
      <th>&nbsp;</th>
  </tr>
		<tr>
      <td> 
			       <select name="paymenttype" id="paymenttype" size="1">
			         <option value="CCScan">CCScan</option>
			         <option value="Check">Check</option>
			         <option value="Cash">Cash</option>
			       </select>
			   </td>
			   <td>
			       <select name="paymentlocation" id="paymentlocation" size="1">
			         <option value="walkin">Walk In</option>
			         <option value="phone">Phone Call</option>
			       </select>
			   </td>
			   <td align="center">
     <%
       if not lcl_isSeasonal then
	       stdate = lcl_startdate
	       if session("OrgID") = 60 then stdate = month(stdate) & "/1/" & year(stdate)
          response.write "<input type=""text"" name=""startdate"" id=""startdate"" value=""" & datevalue(stdate) & """ size=""10"" maxlength=""10"" onchange=""clearMsg('startdate');"" />" & vbcrlf
          response.write "<img src=""../images/calendar.gif"" border=""0"" style=""cursor: hand"" onMouseOver=""tooltip.show('Click to View Calendar');"" onMouseOut=""tooltip.hide();"" onclick=""doCalendar('startdate');"" />" & vbcrlf
       else
          response.write "<select name=""startdate"" id=""startdate"" onchange=""clearMsg('startdate');"">" & vbcrlf

         'Can only renew seasons for the current season and a set number of years after the current year.
          lcl_additional_years = 1

          for i = year(date()) to year(date()) + lcl_additional_years
             response.write "  <option value=""1/1/" & i & """>" & i & "</option>" & vbcrlf
          next

          response.write "</select>" & vbcrlf
          'response.write "<input type=""hidden"" name=""startdate"" id=""startdate"" value=""" & datevalue(lcl_startdate) & """ size=""10"" maxlength=""10"" onchange=""clearMsg('startdate');"" />" & vbcrlf
       end if

       response.write "</td>" & vbcrlf
       response.write "<td width=""200"">" & vbcrlf

       if iFamilyCount > 0 then
          response.write "<input type=""button"" class=""button"" id=""continuebutton"" name=""continue"" value=""Complete This " & lcl_renewal_text & """ onclick=""ContinuePurchase();"" />" & vbcrlf
       end if
     %>
      </td>
  </tr>
</table>
</div>
</form>

  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
Set oMembership = Nothing
'------------------------------------------------------------------------------
sub GetRateData(ByVal iRateId, ByRef sRateName, ByRef sMessage, ByRef nAmount, ByRef iMaxSignUps, ByRef iAttendanceTypeID, _
                ByRef iIsPunchcard, ByRef iPunchcardLimit)
  sRateName         = ""
  sMessage          = ""
  nAmount           = ""
  iMaxSignUps       = ""
  iAttendanceTypeID = ""
  iIsPunchcard      = False
  iPunchcardLimit   = 0

  sSQL = "SELECT description, amount, message, maxsignups, attendancetypeid, isPunchcard, punchcard_limit "
  sSQL = sSQL & " FROM egov_poolpassrates "
  sSQL = sSQL & " WHERE rateid = '" & iRateID & "'"

 	set oRate = Server.CreateObject("ADODB.Recordset")
	 oRate.Open sSQL, Application("DSN"), 3, 1

  if not oRate.eof then
     sRateName         = oRate("description")
     sMessage          = oRate("message")
     nAmount           = oRate("amount")
     iMaxSignUps       = oRate("maxsignups")
     iAttendanceTypeID = oRate("attendancetypeid")
     iIsPunchcard      = oRate("isPunchcard")
     iPunchcardLimit   = oRate("punchcard_limit")
  end if

  set oRate = nothing

end sub

'------------------------------------------------------------------------------
Sub GetPurchaserInfo( iUserId )
	Dim sSQL

	sSQL = "SELECT userfname, userlname, userbirthdate "
 sSQL = sSQL & " FROM egov_users "
 sSQL = sSQL & " WHERE userid = " & iUserId

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN") , 3, 1

 if NOT oUser.eof then
 	 	response.write "  <tr>" & vbcrlf
    response.write "      <td align=""left""><input type=""checkbox"" checked=""checked"" name=""passIncl"" value=""" & iUserId & """ /></td>" & vbcrlf
    response.write "      <td>" & oUser("userfname") & "</td>" & vbcrlf
    response.write "      <td align=""left"">" & oUser("userlname") & "</td>" & vbcrlf
    response.write "      <td align=""left"">Yourself</td>" & vbcrlf
    response.write "      <td align=""left"">" & oUser("userbirthdate") & "</td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
	end if
		
	oUser.close
	Set oUser = Nothing

End Sub 

'------------------------------------------------------------------------------
Function GetFamilyMembers( iUserId, iRateId, iPoolPassID )
	Dim sSQL, iCount, sPreselected

'Get the preseelcted family member types
	sPreselected = GetPreselected( iRateId )

	iCount      = 0
 lcl_bgcolor = "#eeeeee"

 sSQL = "SELECT familymemberid, firstname, lastname, birthdate, relationship "
 sSQL = sSQL & " FROM egov_familymembers "
 sSQL = sSQL & " WHERE isdeleted = 0 "
 sSQL = sSQL & " AND belongstouserid = " & iUserID
 sSQL = sSQL & " ORDER BY birthdate ASC "

	set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN") , 3, 1

 if not oUser.eof then
   	while not oUser.eof
      	iCount = iCount + 1

      'If this is a renewal then show the memberids.
       if iPoolPassID <> "" then
          sSQLm = "SELECT memberid FROM egov_poolpassmembers "
          sSQLm = sSQLm & " WHERE poolpassid = " & iPoolPassID
          sSQLm = sSQLm & " AND familymemberid = " & oUser("familymemberid")

         	set rsm = Server.CreateObject("ADODB.Recordset")
         	rsm.Open sSQLm, Application("DSN") , 3, 1

          if not rsm.eof then
             lcl_member_id = rsm("memberid")
          else
             lcl_member_id = "&nbsp;"
          end if

          set rsm = nothing
       end if

     	'If the family members relation is preselected, check it
     		'if InStr(sPreselected, oUser("relationship")) > 0 then
       if UCASE(sPreselected) = UCASE(oUser("relationship")) then
       			lcl_checked = " checked=""checked"""
       else
          lcl_checked = ""
       end if

       response.write "  <tr align=""left"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf		
     		response.write "      <td><input type=""checkbox"" name=""passIncl"" value=""" & oUser("familymemberid") & """" & lcl_checked & " /></td>" & vbcrlf
     		response.write "      <td>" & oUser("firstname") & "</td>" & vbcrlf
       response.write "      <td>" & oUser("lastname")  & "</td>" & vbcrlf

      'Show the memberid if this is a renewal
       if iPoolPassID <> "" then
          response.write "      <td>" & lcl_member_id & "</td>" & vbcrlf
       end if

     		if oUser("relationship") = "Yourself" then
       			response.write "      <td>Purchaser</td>" & vbcrlf
     		else
       			response.write "      <td>" & oUser("relationship") & "</td>" & vbcrlf
     		end if

     		response.write "      <td>" & oUser("birthdate") & "</td>" & vbcrlf
       response.write "      <td>&nbsp;</td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

       lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     		oUser.movenext
    wend
 end if
		
	oUser.close
	set oUser = nothing

	GetFamilyMembers = iCount

End Function  

'------------------------------------------------------------------------------
Function GetPreselected( iRateId )
	' This builds a string that can be searched to see if the family member is preselected for that rate
	Dim oCmd, oRelation
	GetPreselected = ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetPoolPassPreselectedList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iRateId", 3, 1, 4, iRateId)
	    Set oRelation = .Execute
	End With

	Do While Not oRelation.eof 
		GetPreselected = GetPreselected & oRelation("relation")
		oRelation.movenext
	Loop 
		
	oRelation.close
	set oRelation = nothing
	set oCmd      = nothing

end function

'------------------------------------------------------------------------------
function getPoolPassExpirationDate(p_poolpassid)
  lcl_return = ""

  if p_poolpassid <> "" then
     sSQL = "SELECT expirationdate "
     sSQL = sSQL & " FROM egov_poolpasspurchases "
     sSQL = sSQL & " WHERE poolpassid = " & p_poolpassid
     sSQL = sSQL & " AND orgid = " & session("orgid")

    	set oExpDate = Server.CreateObject("ADODB.Recordset")
   	 oExpDate.Open sSQL, Application("DSN"), 3, 1

     if not oExpDate.eof then
        lcl_return = oExpDate("expirationdate")
     end if

     set oExpDate = nothing

  end if

  getPoolPassExpirationDate = lcl_return

end function
%>
