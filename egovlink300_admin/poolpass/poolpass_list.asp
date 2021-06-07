<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: poolpass_list.asp
' AUTHOR: Steve Loar
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the report of membership
'
' MODIFICATION HISTORY
' 1.0  05/09/06	Steve Loar - INITIAL VERSION
' 1.1	 10/05/06	Steve Loar - Security, Header and nav changed
' 1.2  09/05/08 David Boyer - Added Membership Renewals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 dim iMembershipId, iPeriodId

 sLevel              = "../"  'Override of value from common.asp
 lcl_isRootAdmin     = false

 if not userhaspermission(session("UserId"), "membership detail" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if


'Set the membershipid the the type selected.  If this is the initial screen opening then default it to "pool".
 set oMembership = New classMembership
 if request("sMembershipType") <> "" then
    lcl_membership_type = request("sMembershipType")
 else
    'lcl_membership_type = "pool"
    lcl_membership_type = oMembership.GetFirstMembershipType()
 end if
' oMembership.SetMembershipId( "pool" )
 oMembership.SetMembershipId(lcl_membership_type)

'Retrieve the search criteria
 orderBy    = request("orderBy")
 subTotals  = request("subTotals")
 showDetail = request("showDetail")
 fromDate   = request("fromDate")
 toDate     = request("toDate")
 today      = Date()
 sUserlname = request("userlname")

'Setup the search criteria variables
 If request("membershipid") = "" Then
   	iMembershipId = GetFirstMembershipId()
 Else
   	iMembershipId = CLng(request("membershipid"))
 End If 

 If request("periodid") = "" Then
	   iPeriodId = CLng(0)
 Else
   	iPeriodId = CLng(request("periodid"))
 End If 

 If orderBy = "" or IsNull(orderBy) Then
    orderBy = "date"
 End If

 if toDate = "" OR IsNull(toDate) then
    toDate = dateAdd("d",0,today)
 else
   'If the value entered is not a valid date then use the default if the field was blank
    if not dbready_date(toDate) then
       toDate = dateAdd("d",0,today)
    end if
 end if

 if fromDate = "" OR IsNull(fromDate) then
    fromDate = DateSerial(Year(Now()),1,1)
 else
   'If the value entered is not a valid date then use the default if the field was blank
    if not dbready_date(fromDate) then
       fromDate = DateSerial(Year(Now()),1,1)
    end if
 end if

 toDate = dateAdd("d",1,toDate)

'Set the session variables
 session("orderBy")       = orderBy
 session("subTotals")     = subTotals
 session("showDetail")    = showDetail
 session("fromDate")      = fromDate
 session("toDate")        = toDate
 session("userlname")     = suserlname
	session("iMembershipId") = imembershipid
	session("iPeriodId")     = iperiodid

'Check to see if the org has the feature turned-on and the user has it assigned
 lcl_membershiprenewals_feature = "N"
 if orghasfeature("membership_renewals") AND userhaspermission(session("userid"),"membership_renewals") then
    lcl_membershiprenewals_feature = "Y"
 end if

'Style width for tables
 lcl_style_width = "width:1000px;"

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = true
 end if

'Check for org features
 lcl_orghaspermission_deletePoolPassPurchases = orghasfeature("delete_poolpass_purchases")

'Check for user permissions
 lcl_userhaspermission_deletePoolPassPurchases = userhaspermission(session("userid"),"delete_poolpass_purchases")
%>
<html>
<head>
  <title>E-GovLink {Membership Purchase Details}</title>
  
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="style_pool.css" />
	<link rel="stylesheet" type="text/css" href="poolpass.css" />
	<link rel="stylesheet" type="text/css" media="print" href="receiptprint.css" />
 <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />


	<script src="../scripts/selectAll.js"></script>
	<script src="../scripts/modules.js"></script>
 <script src="../scripts/tooltip_new.js"></script>
 <script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript">
	<!--
		function checkStat() 
		{
			if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) 
			{
				alert("You must select the status.");
				form1.statusPending.focus();
				return false;
			}
		}
		
		function CheckAllStatus() 
		{
			if (document.form1.CheckAllStat.checked) 
			{
				document.form1.statusPending.checked   = true;
				document.form1.statusCompleted.checked = true;
				document.form1.statusDenied.checked    = true;
			} 
			else 
			{
				document.form1.statusPending.checked   = false;
				document.form1.statusCompleted.checked = false;
				document.form1.statusDenied.checked    = false;
			}
		}

		function doCalendar(ToFrom) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

function removePass(iPassId) {
  if (confirm("Delete Pass #" + iPassId + "?")) {
   			location.href='poolpass_remove.asp?passid=' + iPassId;
		}
}

function renewPass(iPassId) {
//  if (confirm("Delete Pass #" + iPassId + "?")) {
   			location.href='select_members.asp?poolpassid=' + iPassId;
//		}
//  inlineMsg(document.getElementById("button_renew_"+iPassId).id,'<strong>Coming Soon: </strong>Renew Membership Option for PoolPassID: '+iPassId,8,'button_renew_'+iPassId);
}

	//-->
	</script>

</head>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<div id="content">
 <div id="centercontent">

<table border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td>
          <font size="+1"><strong><%=Session("sOrgName")%> Membership Purchase Details</strong></font>
      </td>
      <td align="right">
      <%
        lcl_message = ""

        if request("success") = "SN" then
           lcl_message = "<strong style=""color:#FF0000"">*** Successfully Created... ***</strong>"
        elseif request("success") = "SU" then
           lcl_message = "<strong style=""color:#FF0000"">*** Successfully Updated... ***</strong>"
        elseif request("success") = "SD" then
           lcl_message = "<strong style=""color:#FF0000"">*** Successfully Deleted... ***</strong>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
</table>

<p>
<fieldset id="search">
  <legend><strong>Search/Sorting Option(s)</strong></legend><br />
<table border="0" cellpadding="3" cellspacing="0" id="searchtable">
  <form action="poolpass_list.asp" method="post" name="searchform">
  <tr>
		  		<td colspan="3"><strong>Membership Type:</strong>&nbsp;<% ShowMembershipPicks iMembershipId %></td>
		</tr>
  <tr>
  				<td colspan="3"><strong>Membership Period:</strong>&nbsp;<% ShowPeriodPicks iPeriodId %></td>
		</tr>
  <tr>
 			  <td valign="top">
	     		  <strong>Purchase From:</strong> <input type="text" name="fromDate" value="<%=fromDate%>" />
				      <img src="../images/calendar.gif" border="0" style="cursor: hand" onMouseOver="tooltip.show('Click to View Calendar');" onMouseOut="tooltip.hide();" onclick="doCalendar('From');" />
 			  </td>
	 		  <td>&nbsp;</td>
			   <td valign="top">
		      	 <strong>To:</strong> <input type="text" name="toDate" value="<%=dateAdd("d",-1,toDate)%>" />
				      <img src="../images/calendar.gif" border="0" style="cursor: hand" onMouseOver="tooltip.show('Click to View Calendar');" onMouseOut="tooltip.hide();" onclick="doCalendar('To');" />
			   </td>
  </tr>
  <tr>
  				<td colspan="3"><strong>Purchaser Last Name:</strong> <input type="text" name="userlname" value="<%=sUserlname%>" size="20" maxlength="50" /></td>
		</tr>
  <tr>
      <td colspan="3"><input type="submit" value="Search" class="button" /></td>
  </tr>
</form>
</table>
</fieldset>
</p>
<button type="button" onClick="window.location='../export/csv_export.asp'">Export to CSV</button>
<br />
<br />

<% List_Payments sSortBy, iMembershipId, iPeriodId %>

 </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
Sub List_Payments( sSortBy, iMembershipId, iPeriodId )
	Dim cTotalAmount, iPassCount

		iPasscount = 0

 'Pull the StartDate now instead of the PaymentDate
  varWhereClause = " AND (P.paymentdate >= '" & fromDate & "' AND P.paymentdate <= '" & toDate & "') "
  'varWhereClause = " AND (P.startdate >= '" & fromDate & "' AND P.startdate < '" & toDate & "') "

		If sUserlname <> "" Then
  			varWhereClause = varWhereClause & " AND U.userlname like '%" & REPLACE(sUserlname,"'","''") & "%' "
		End If 

		If CLng(iMembershipId) > CLng(0) Then
		  	varWhereClause = varWhereClause & " AND P.membershipid = " & iMembershipId
		End If 

		If CLng(iPeriodId) > CLng(0) Then
  			varWhereClause = varWhereClause & " AND P.periodid = " & iPeriodId
		End If 

		sSQL = "SELECT P.poolpassid as [Pass ID], " & vbcrlf
		'lcl_rate_description
		sSQL = sSQL & " ISNULL(m.membershipdesc,'') + ' - ' + ISNULL(mp.period_desc,'') + '<br />' + ISNULL(ppr.description,'') + CASE WHEN ppr.description IS NOT NULL AND pprt.description IS NOT NULL THEN ' - ' ELSE '' END + ISNULL(pprt.description,'') as [Membership Type], " & vbcrlf
  		sSQL = sSQL & " P.paymentdate as [Purchase Date], " & vbcrlf
  		sSQL = sSQL & " isnull(P.startdate,'') as [Membership Start Date], " & vbcrlf
  		sSQL = sSQL & " isnull(P.expirationdate,'') as [Expiration Date], " & vbcrlf
    		if lcl_membershiprenewals_feature = "Y" then
			'CALCULATE RENEW BY DATE
			sSQL = sSQL & " CASE ppr.isRenewable WHEN 1 THEN DATEADD(DD, ppr.renewalTimeAfterExpire, p.expirationdate) ELSE '' END as [Renew By Date], " & vbcrlf
  		end if
  		sSQL = sSQL & " ISNULL(U.userfname,'') + ' ' + ISNULL(U.userlname,'') as Purchaser, " & vbcrlf
  		sSQL = sSQL & " P.paymentamount as [Payment Amount], " & vbcrlf
		sSQL = sSQL & " UPPER(LEFT(p.paymentlocation,1))+LOWER(SUBSTRING(p.paymentlocation,2,LEN(p.paymentlocation))) + ' - ' + UPPER(LEFT(p.paymenttype,1))+LOWER(SUBSTRING(p.paymenttype,2,LEN(p.paymenttype))) as [Payment Method], " & vbcrlf
  		sSQL = sSQL & " P.paymentresult as Status " & vbcrlf
        	if lcl_membershiprenewals_feature = "Y" then
  			sSQL = sSQL & " ,P.previous_poolpassid as [Renewal of Pass ID] " & vbcrlf
		end if

  	'	sSQL = sSQL & " ,P.rateid " & vbcrlf
  		if session("orgid") = "26" then
  			sSQL = sSQL & " ,P.note as Notes,ISNULL(au.firstname,'') + ' ' + ISNULL(au.lastname,'') as [Admin Processor] " & vbcrlf
		end if

		sSQL = sSQL & " FROM egov_users U " & vbcrlf
		sSql = sSql & " INNER JOIN egov_poolpasspurchases P ON U.userid = p.userid " & vbcrlf
		sSql = sSql & " INNER JOIN egov_memberships M ON M.membershipid = P.membershipid " & vbcrlf
		sSql = sSql & " INNER JOIN egov_membership_periods MP ON mp.periodid = p.periodid AND MP.orgid = P.orgid " & vbcrlf
		sSQL = sSQL & " INNER JOIN egov_poolpassrates ppr ON ppr.rateid = p.rateid " & vbcrlf
		sSQL = sSQL & " LEFT JOIN egov_poolpassresidenttypes pprt ON pprt.orgid = p.orgid and UPPER(pprt.resident_type) = ppr.residenttype " & vbcrlf
		sSql = sSql & " LEFT JOIN users au ON au.userid = P.adminid" & vbcrlf
		sSQL = sSQL & " WHERE P.orgid = " & session("orgid") & vbcrlf
  		sSQL = sSQL & " AND P.paymentresult <> 'Pending' " & vbcrlf
		sSQL = sSQL & " AND P.paymentresult <> 'Declined' " & vbcrlf
  		sSQL = sSQL & varWhereClause & vbcrlf
		sSQL = sSQL & " ORDER BY P.poolpassid " & vbcrlf
		'response.write sSQL

		'lastTitle = "Test"
		'lastDate  = "1/1/02"

		set oRequests = Server.CreateObject("ADODB.Recordset")

		oRequests.PageSize       = 25
		oRequests.CacheSize      = 5
		oRequests.CursorLocation = 3

		session("DISPLAYQUERY") = sSQL
		oRequests.Open sSQL, Application("DSN"), 0, 1

  response.write "<div class=""shadow"" style=""" & lcl_style_width & """>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" style=""" & lcl_style_width & """>" & vbcrlf
  response.write "  <tr class=""tablelist"" align=""left"" valign=""bottom"">" & vbcrlf
  response.write "      <th style=""text-align: center"" nowrap=""nowrap"">Pass<br />ID</th>" & vbcrlf
  response.write "      <th>Membership Type</th>" & vbcrlf
  response.write "      <th style=""text-align: center"">Purchase<br />Date</th>" & vbcrlf
  response.write "      <th style=""text-align: center"">Membership<br />Start Date</th>" & vbcrlf
  response.write "      <th style=""text-align: center"">Expiration<br />Date</th>" & vbcrlf

    if lcl_membershiprenewals_feature = "Y" then
     response.write "      <th style=""text-align: center"" nowrap>Renew By<br />Date</th>" & vbcrlf
  end if

  response.write "      <th>Purchaser</th>" & vbcrlf
  response.write "      <th style=""text-align: center"">Payment<br />Amount</th>" & vbcrlf
  response.write "      <th>Payment Method</th>" & vbcrlf
  response.write "      <th>Status</th>" & vbcrlf

  if lcl_membershiprenewals_feature = "Y" then
     response.write "      <th style=""text-align: center"" id=""column_renewal"" nowrap>Renewal of<br />Pass ID</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
  end if
  if session("orgid") = "26" then
  	response.write "      <th>Notes</th>" & vbcrlf
  	response.write "      <th>Admin Processor</th>" & vbcrlf
  end if

  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 	If oRequests.EOF then
	  		response.write "  <tr><td colspan=""10"" bgcolor=""#ffffff""><strong>No records found</strong></td></tr>" & vbcrlf
		Else 

     Dim abspage, pagecnt

     bgcolor                   = "#ffffff"
     lcl_expiredate_textcolor  = ""
     lcl_renewbydate_textcolor = ""
     lcl_renewbydate           = ""

   		while not oRequests.eof
  	   		iPasscount = iPasscount + 1
   		  	iRowCount  = iRowCount + 1
        bgcolor    = changeBGColor(bgcolor,"#eeeeee","#ffffff")

       'Start Date - Format
        if datevalue(oRequests("Membership Start Date")) = "1/1/1900" OR oRequests("Membership Start Date") = "" then
           lcl_startdate = datevalue(oRequests("Purchase Date"))
        else
           lcl_startdate = datevalue(oRequests("Membership Start Date"))
        end if

       'Expiration Date - Format
        if datevalue(oRequests("Expiration Date")) = "1/1/1900" OR oRequests("Expiration Date") = "" then
           lcl_expirationdate = ""
        else
           lcl_expirationdate = datevalue(oRequests("Expiration Date"))
        end if

       'Determine if the current date is GREATER THAN or EQUAL TO the expiration date.
       'If it is then change the text color to RED.
        if datevalue(date()) >= lcl_expirationdate then
           lcl_expiredate_textcolor = " style=""color: #ff0000"""
        else
           lcl_expiredate_textcolor = ""
        end if

       ''Retrieve the RATE info
        'getRateInfo oRequests("rateid"), lcl_rate_description, lcl_rate_residenttype

       ''Get the Days to Renewal After Expiration Date and isRenewable
        'sSQL3 = "SELECT renewalTimeAfterExpire, isRenewable "
        'sSQL3 = sSQL3 & " FROM egov_poolpassrates "
        'sSQL3 = sSQL3 & " WHERE orgid = " & session("orgid")
        'sSQL3 = sSQL3 & " AND rateid = " & oRequests("rateid")

       	'set rs3 = Server.CreateObject("ADODB.Recordset")
       	'rs3.Open sSQL3, Application("DSN"), 3, 1
'
        'if not rs3.eof then
           'lcl_renewalTimeAfterExpire = rs3("renewalTimeAfterExpire")
           'lcl_isRenewable            = rs3("isRenewable")
        'else
           'lcl_renewalTimeAfterExpire = 0
           'lcl_isRenewable            = False
        'end if
'
        'set rs3 = nothing

       'Determine if the Renewal column/button are displayed
        if lcl_membershiprenewals_feature = "Y" then
           lcl_showHideRenewalButton = showHideRenewalButton(oRequests("Pass ID"))
        else
           lcl_showHideRenewalButton = "N"
        end if

       'Renew by Date - Format
        'if lcl_membershiprenewals_feature = "Y" and lcl_showHideRenewalButton = "Y" then
              'lcl_renewbydate = datevalue(DATEADD("d",lcl_renewalTimeAfterExpire,oRequests("Expiration Date")))
        'end if

       'Determine if the current date is GREATER THAN or EQUAL TO the Renew by Date.
       'If it is then change the text color to RED.
        if lcl_renewbydate <> "" then
           if CDate(date()) >= CDate(lcl_renewbydate) then
              lcl_renewbydate_textcolor = " style=""color: #ff0000"""
           end if
        end if

        lcl_row_mouseover = " onMouseOver=""mouseOverRow(this);"""
        lcl_row_mouseout  = " onMouseOut=""mouseOutRow(this);"""
        'lcl_td_mouseover  = " onMouseOver=""tooltip.show('click to edit');"""
        'lcl_td_mouseout   = " onMouseOut=""tooltip.hide();"""
        lcl_td_mouseover  = ""
        lcl_td_mouseout   = ""
        lcl_onclick       = " onClick=""location.href='poolpass_details.asp?iPoolPassId=" & oRequests("Pass ID") & "';"""

     			response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & bgcolor & """" & lcl_row_mouseover & lcl_row_mouseout & ">" & vbcrlf
   		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & oRequests("Pass ID") & "</td>" & vbcrlf
  	   		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("Membership Type") & "</td>" & vbcrlf
   		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & datevalue(oRequests("Purchase Date"))    & "</td>" & vbcrlf
   		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & lcl_startdate                          & "</td>" & vbcrlf
   		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center""" & lcl_expiredate_textcolor & ">" & lcl_expirationdate                     & "</td>" & vbcrlf

    	'if lcl_membershiprenewals_feature = "Y" and lcl_showHideRenewalButton = "Y" then
    	if lcl_membershiprenewals_feature = "Y" then
           if oRequests("Renew By Date") <> "" then
      		  	   response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center""" & lcl_renewbydate_textcolor & ">" & oRequests("Renew By Date") & "</td>" & vbcrlf
           else
           			response.write "      <td>&nbsp;</td>" & vbcrlf
           end if
        end if

  	   		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("Purchaser") & "</td>" & vbcrlf
     			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" & formatcurrency(oRequests("Payment Amount"),2) & "</td>" & vbcrlf

     			cTotalAmount = cTotalAmount + CDbl(oRequests("Payment Amount"))

     			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("Payment Method") & "</td>" & vbcrlf
		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" &  oRequests("Status") & "</td>" & vbcrlf

        if lcl_membershiprenewals_feature = "Y" then
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""center"">" &  oRequests("Renewal of Pass ID") & "</td>" & vbcrlf
        end if

       'Check to see if the Membership Renewal feature is to be displayed.
       'Also check to see if any renewals exist for the org and the rates.  If not then hide the column.
        if lcl_membershiprenewals_feature = "Y" then
           if lcl_showHideRenewalButton = "Y" then
           			response.write "      <td><input type=""button"" name=""renew"" id=""button_renew_" & oRequests("Pass ID") & """ value=""Renew"" onclick=""renewPass('" & oRequests("Pass ID") & "');"" class=""button"" /></td>" & vbcrlf
           else
              response.write "      <td>&nbsp;</td>" & vbcrlf
           end if
        end if
  			if session("orgid") = "26" then
				response.write "<td>" & oRequests("note") & "</td>" & vbcrlf
				response.write "<td>" & oRequests("firstname") & " " & oRequests("lastname") & "</td>" & vbcrlf
			end if

       'This button has been disabled per Jerry's/Peter's decision in task meeting so that it can be determine if anyone
       'complains about it being disabled and actually need use of it.  When a membership purchase is deleted, EVERYTHING
       'related to that purchase is deleted when actually it needs to follow the class format and "refunds" so that
       'money can be tracked in budgets.  3/15/2012 - David Boyer
       'The "delete" button is now only available to ROOT ADMIN via suggestion by David Boyer to counter Jerry's suggestion
       'to delete the membership manually.  The delete button completely removes the membership purcahse record as well as the
       'membership id(s) associated to that pool pass purchase.  If the user is NOT a "root admin" then the button will be
       'disabled.  3/19/2012 - David Boyer
       'The "delete" button is now only available to users that have been assigned the feature permission.
       'This was decided and confirmed by Peter on 4/4/2012 as a change needed since Christina and I were picking up additional
       'work maintaining orgs' records instead of giving them the ability to do so.  4/9/2012 - David Boyer
        lcl_onclick_deleteButton = ""
        lcl_tooltip_deleteButton = " onMouseOver=""tooltip.show('Please contact E-Gov Support to have this record deleted');"" onMouseOut=""tooltip.hide();"""

        'if lcl_isRootAdmin then
        if lcl_orghaspermission_deletePoolPassPurchases AND lcl_userhaspermission_deletePoolPassPurchases then
           lcl_onclick_deleteButton = " onclick=""removePass('" & oRequests("Pass ID") & "');"""
           lcl_tooltip_deleteButton = ""
        end if

     			response.write "      <td align=""center""><input type=""button"" name=""remove"" value=""Delete"" class=""button""" & lcl_onclick_deleteButton & lcl_tooltip_deleteButton & " /></td>" & vbcrlf
		     	response.write "  </tr>" & vbcrlf

     			oRequests.movenext
				response.flush
   		wend

	 end if

  oRequests.Close
  Set oRequests = Nothing 

		iRowCount = iRowCount + 1

		response.write "  <tr id=""" & iRowCount & """ bgcolor=""#C0C0C0"">" & vbcrlf
		response.write "      <td colspan=""100"" style=""border-top: 1pt solid #000000"" align=""right"">" & vbcrlf
  response.write "          Passes Sold: <strong>"  & iPasscount & "</strong>&nbsp;&nbsp;"
  response.write "          Total Amount: <strong>" & FormatCurrency(cTotalAmount) & "</strong>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

		response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
 
end sub

'------------------------------------------------------------------------------
Sub ShowMembershipPicks( iMembershipId )
	Dim sSql, oMember

	sSQL = "SELECT membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & session("orgid") 
	
	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 3, 1
	
	if NOT oMember.EOF then
  		response.write "<select name=""membershipid"">" & vbcrlf
  		response.write "  <option value=""0"">All</option>" & vbcrlf

  		Do While Not oMember.EOF
    			if CLng(oMember("membershipid")) = CLng(iMembershipId) then
	      			lcl_selected = " selected=""selected"" "
       else
          lcl_selected = ""
    			end if

    			response.write "  <option value=""" & oMember("membershipid") & """" & lcl_selected & ">" & oMember("membershipdesc") & "</option>" & vbcrlf

    			oMember.MoveNext
  		Loop 

  		response.write "</select>" & vbcrlf
 end if
	
	oMember.close
	Set oMember = Nothing
end sub

'------------------------------------------------------------------------------
Sub ShowPeriodPicks( iPeriodId )
	Dim sSql, oPeriod

	sSQL = "SELECT periodid, period_desc FROM egov_membership_periods WHERE orgid = " & session("orgid") 
	
	Set oPeriod = Server.CreateObject("ADODB.Recordset")
	oPeriod.Open sSQL, Application("DSN"), 3, 1
	
	if not oPeriod.eof then
  		response.write "<select name=""periodid"">" & vbcrlf
		  response.write "  <option value=""0"">All</option>" & vbcrlf

  		while not oPeriod.eof
       if clng(oPeriod("periodid")) = clng(iPeriodID) then
          lcl_selected = " selected=""selected"""
       else
          lcl_selected = ""
       end if

     		response.write "  <option value=""" & oPeriod("periodid") & """" & lcl_selected & ">" & oPeriod("period_desc") & "</option>" & vbcrlf

    			oPeriod.movenext
		  wend

  		response.write "</select>" & vbcrlf
 end if
	
	oPeriod.close
	set oPeriod = nothing
End Sub 

'------------------------------------------------------------------------------
function GetFirstMembershipId()
	Dim sSql, oMember

	sSQL = "SELECT MIN(membershipid) AS membershipid FROM egov_memberships WHERE orgid = " & session("orgid") 
	
	set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 3, 1
	
	If IsNull(oMember("membershipid")) Then
  		GetFirstMembershipId = 0
	Else
		  GetFirstMembershipId = oMember("membershipid")
	End If 
	
	oMember.close
	set oMember = nothing
end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
	 set oDTB = Server.CreateObject("ADODB.Recordset")
 	oDTB.Open sSQLi, Application("DSN"), 3, 1

  set oDTB = nothing


end sub
%>
