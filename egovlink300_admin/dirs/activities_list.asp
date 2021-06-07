<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: activities_list.asp
' AUTHOR: Steve Loar
' CREATED: 01/24/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a list of classes by attendee
'
' MODIFICATION HISTORY
' 1.0   01/24/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sName, oRs, iActivityCount, lcl_PurchaseTotal, lcl_RefundTotal, lcl_sc_from_startdate
Dim lcl_sc_to_startdate, purchaseToDate, startToDate

sLevel     = "../"    'Override of value from common.asp
lcl_hidden = "HIDDEN" 'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

iUserId = CLng(request("u"))
sName   = GetCitizenName( iUserId )
iActivityCount = 0
lcl_PurchaseTotal = CDbl(0.00)
lcl_RefundTotal = CDbl(0.00)

lcl_head_of_household = getHeadofHousehold(iUserid)

lcl_sc_from_date = request("frompurchasedate")
lcl_sc_to_date   = request("topurchasedate")
lcl_sc_from_startdate = request("fromstartdate")
lcl_sc_to_startdate = request("tostartdate")
If lcl_sc_from_startdate <> "" Then
	'response.write "lcl_sc_to_startdate = " & lcl_sc_to_startdate & "<br />"
	startToDate = CDate(lcl_sc_to_startdate)
	startToDate = dateAdd("d",1,startToDate)
End If 

lcl_sc_orderby   = request("sc_orderby")
today            = Date()

if lcl_sc_from_date = "" OR IsNull(lcl_sc_from_date) then
   lcl_sc_from_date = DateSerial(Year(Now())-5,1,1)
end if

If lcl_sc_to_date = "" Or IsNull(lcl_sc_to_date) Then 
	lcl_sc_to_date = dateAdd("d",0,today)
Else
	lcl_sc_to_date = CDate(lcl_sc_to_date)
End If 
purchaseToDate = dateAdd("d",1,lcl_sc_to_date)

%>

<html>
<head>
	<title><%=langBSPayments%></title>
	<meta http-equiv="Content-type" content="text/html;charset=UTF-8">

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="./reservationliststyles.css" />

	<script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="javascript">
	<!--
		// create the egov NameSpace
		var eGovLink = eGovLink || {}; 

		// create the sub-NameSpace with the methods inside   
		eGovLink.CitizenActivities = (function() 
		{
			var doCalendar = function ( ToFrom ) 
			{
				var w = (screen.width - 350)/2;
				var h = (screen.height - 350)/2;
				eval('window.open("../classes/calendarpicker.asp?p=1&updateform=activity_list&updatefield=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			}

			var validate = function()
			{
				// check the purchase from date
				if ($("#frompurchasedate").val() != '')
				{
					if (! isValidDate($("#frompurchasedate").val()))
					{
						//alert("The Purchase 'From' Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
						inlineMsg(document.getElementById("frompurchasedate").id,'<strong>Invalid Value: </strong> This must be a valid date in the format of MM/DD/YYYY.',5,'frompurchasedate');
						$("#frompurchasedate").focus();
						return;
					}
				}
				// check the purchase to date
				if ($("#topurchasedate").val() != '')
				{
					if (! isValidDate( $("#topurchasedate").val() ))
					{
						//alert("The Purchase 'To' Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
						inlineMsg(document.getElementById("topurchasedate").id,'<strong>Invalid Value: </strong> This must be a valid date in the format of MM/DD/YYYY.',5,'topurchasedate');
						$("#topurchasedate").focus();
						return;
					}
				}
				// check the start from date
				if ($("#fromstartdate").val() != '')
				{
					if (! isValidDate($("#fromstartdate").val()))
					{
						//alert("The Start 'From' Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
						inlineMsg(document.getElementById("fromstartdate").id,'<strong>Invalid Value: </strong> This must be a valid date in the format of MM/DD/YYYY.',5,'fromstartdate');
						$("#fromstartdate").focus();
						return;
					}
					if ($("#tostartdate").val() == '')
					{
						//alert("The Start 'To' Date cannot be blank when a 'From' Date is specified.  \nPlease enter it again.");
						inlineMsg(document.getElementById("tostartdate").id,'<strong>Missing: </strong> The To Date cannot be blank when the From Date is specified.',5,'tostartdate');
						$("#tostartdate").focus();
						return;
					}
				}
				// check the start to date
				if ($("#tostartdate").val() != '')
				{
					if (! isValidDate($("#tostartdate").val()))
					{
						//alert("The Start 'To' Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
						inlineMsg(document.getElementById("tostartdate").id,'<strong>Invalid Value: </strong> This must be a valid date in the format of MM/DD/YYYY.',5,'tostartdate');
						$("#tostartdate").focus();
						return;
					}
					if ($("#fromstartdate").val() == '')
					{
						//alert("The Start 'From' Date cannot be blank when a 'To' Date is specified.  \nPlease enter it again.");
						inlineMsg(document.getElementById("fromstartdate").id,'<strong>Missing: </strong> The From Date cannot be blank when the To Date is specified.',5,'fromstartdate');
						$("#fromstartdate").focus();
						return;
					}
				}
				document.activity_list.submit();
			}

			// This makes the functions publicly accessible
			return {
				doCalendar: doCalendar,
				validate: validate
			};

		}());

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">
		<font size="+1"><b>Recreation Activities of <%=sName%></b><br>
      &nbsp;&nbsp;&nbsp;(Head of Household: <%=lcl_head_of_household%>)</font><p>
		<a href="javascript:history.back()"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a><br>

  <fieldset>
    <legend>Search Criteria&nbsp;</legend><p>
  <table border="0" cellspacing="0" cellpadding="2">
    <form name="activity_list" method="post" action="activities_list.asp">
      <input type="<%=lcl_hidden%>" name="u" value="<%=iUserId%>" size="10" maxlength="10">
      <input type="<%=lcl_hidden%>" name="v" value="<%=request("v")%>" size="10" maxlength="10">
    <tr>
		<td><strong>Purchase Date</strong></td>
        <td>From: </td>
        <td>
            <input type="text" id="frompurchasedate" name="frompurchasedate" value="<%=lcl_sc_from_date%>" size="15" maxlength="10">&nbsp;
            <a href="javascript:void eGovLink.CitizenActivities.doCalendar('frompurchasedate');"><img src="../images/calendar.gif" border="0"></a>
        </td>
		<td>To: </td>
        <td>
            <input type="text" id="topurchasedate" name="topurchasedate" value="<%=lcl_sc_to_date%>" size="15" maxlength="10">&nbsp;
            <a href="javascript:void eGovLink.CitizenActivities.doCalendar('topurchasedate');"><img src="../images/calendar.gif" border="0"></a>
			&nbsp;
			<%DrawDateChoices "purchasedate" %>
        </td>
	</tr>
	<tr>
		<td><strong>Start Date</strong></td>
        <td>From: </td>
        <td>
            <input type="text" id="fromstartdate" name="fromstartdate" value="<%=lcl_sc_from_startdate%>" size="15" maxlength="10">&nbsp;
            <a href="javascript:void eGovLink.CitizenActivities.doCalendar('fromstartdate');"><img src="../images/calendar.gif" border="0"></a>
        </td>
		<td>To: </td>
        <td>
            <input type="text" id="tostartdate" name="tostartdate" value="<%=lcl_sc_to_startdate%>" size="15" maxlength="10">&nbsp;
            <a href="javascript:void eGovLink.CitizenActivities.doCalendar('tostartdate');"><img src="../images/calendar.gif" border="0"></a>
			&nbsp;
			<%DrawDateChoices "startdate" %>
        </td>
	</tr>
	<tr>
        <td colspan="5">Order By:&nbsp;
            <select name="sc_orderby">
              <%
                lcl_selected_activity = ""
                lcl_selected_pdate    = ""

                if UCase(lcl_sc_orderby) = "ACTIVITY" then
                   lcl_selected_activity = " selected=""selected"" "
                Else
					lcl_selected_pdate = " selected=""selected"" "
                End If 
              %>
              <option value="ACTIVITY"<%=lcl_selected_activity%>>Activity Name</option>
              <option value="PURCHASE_DATE"<%=lcl_selected_pdate%>>Purchase Date</option>
            </select>
        </td>
    </tr>
    <tr>
        <td colspan="5">
            <input type="button" class="button" value="Search" onclick="eGovLink.CitizenActivities.validate();">
        </td>
    </tr>
    </form>
  </table>
  </fieldset>
<%

	' Get the activities that they were part of
	sSql = "SELECT P.paymentdate, C.classname, ISNULL(T.activityno,'') AS activityno, L.quantity, C.startdate, J.journalentrytype, "
	sSql = sSql & " L.paymentid, ISNULL(P.relatedpaymentid,0) AS relatedpaymentid, C.classid, L.familymemberid, L.classlistid "
	sSql = sSql & " FROM egov_class_list L, egov_class_payment P, egov_class C, egov_class_time T, egov_journal_entry_types J "
	sSql = sSql & " WHERE L.attendeeuserid = " & iUserId
	sSql = sSql & " AND L.paymentid = P.paymentid AND L.classid = C.classid "
	sSql = sSql & " AND L.classtimeid = T.timeid AND P.journalentrytypeid = J.journalentrytypeid "
	sSql = sSql & " AND (P.paymentdate >= '" & lcl_sc_from_date & "' AND P.paymentdate < '" & purchaseToDate & "') "

	If lcl_sc_from_startdate <> "" Then
		sSql = sSql & " AND (C.startdate >= '" & lcl_sc_from_startdate & "' AND C.startdate < '" & startToDate & "') "
	End If 

	'Setup the ORDER BY
	If lcl_sc_orderby = "ACTIVITY" Then 
		sSql = sSql & " ORDER BY C.classname"
	Else 
		sSql = sSql & " ORDER BY P.paymentdate DESC, C.classname"
	End If 

	'response.write sSql & "<br /><br />"
	'response.End 

	response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""citizenreport"">"
	response.write vbcrlf & "<tr class=""tablelist"">"
	response.write "<th align=""left"">Activity Name</th><th>Qty</th><th align=""left"">Purchase Date</th><th align=""left"">Start Date</th><th>Receipt #</th><th>Class #</th><th>Purchase<br />Amount</th><th>Refund<br />Amount</th>"
	response.write "</tr>"
	bgcolor = "#eeeeee"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	'LOOP AND DISPLAY THE RECORDS
	Do While Not oRs.EOF
		iActivityCount = iActivityCount + 1
		If bgcolor="#eeeeee" Then 
			bgcolor="#ffffff" 
		Else 
			bgcolor="#eeeeee"
		End If 

		'Calulate the Payment Total
'		sRefundAmount = GetPurchaseTotal( oRs("paymentid") ) - GetRefundFeeAmount( oRs("paymentid") )
     	If GetPaymentDetails( oRs("paymentid"), iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminUserId, iPriorPaymentId ) Then 
       		sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )

			If sJournalEntryType <> "" Then 
				Select Case sJournalEntryType
					Case "purchase"
						'Show purchase details
						'ShowPurchaseDetails iPaymentId, "credit", sJournalEntryType, 0
						sPaymentTotal = getPaymentTotal( oRs("paymentid"), "credit", sJournalEntryType, 0, oRs("classlistid") )
						sPaymentTotal = formatcurrency(sPaymentTotal,2)
						lcl_PurchaseTotal = lcl_PurchaseTotal + CDbl(sPaymentTotal)
					Case "refund"
						'Show refund stuff
						'ShowPurchaseDetails iPaymentId, "debit", sJournalEntryType, iPriorPaymentId
						sPaymentTotal = getPaymentTotal( oRs("paymentid"), "debit", sJournalEntryType, oPriorPaymentId, oRs("classlistid") )
						sPaymentTotal = formatcurrency(sPaymentTotal,2)
						lcl_RefundTotal = lcl_RefundTotal + CDbl(sPaymentTotal)
					Case "transfer"
						'Show citizen account transfer
						sPaymentTotal = formatcurrency(0,2)
					Case "deposit"
						'Show citizen account deposit
						sPaymentTotal = formatcurrency(0,2)
					Case "withdrawl"
						'Show citizen account withdrawl
						sPaymentTotal = formatcurrency(0,2)
				End Select 
			Else 
				sPaymentTotal = formatcurrency(0,2)
			End If 

			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""tablelist"">"
			response.write "<td>" & oRs("classname") & "</td>"
			response.write "<td align=""center"">" & oRs("quantity") & "</td>"
			response.write "<td>" & oRs("paymentdate") & "</td>"
			response.write "<td>" & oRs("startdate") & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & oRs("activityno") & "</td>"

			If oRs("journalentrytype") = "purchase" Then 
				' These are the purchases
				response.write "      <td align=""right"">" & sPaymentTotal & "</td>"
				response.write "      <td>&nbsp;</td>"
			Else
				' These are the refunds
				response.write "      <td>&nbsp;</td>"
				response.write "      <td align=""right"">" & sPaymentTotal & "</td>"
			End If 
			response.write "  </tr>"

		Else 
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""tablelist"">"
			response.write "<td>" & oRs("classname") & "</td>"
			response.write "<td align=""center"">" & oRs("quantity") & "</td>"
			response.write "<td>" & oRs("paymentdate") & "</td>"
			response.write "<td>" & oRs("startdate") & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & oRs("activityno") & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td>&nbsp;</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	' Total Row
	response.write vbcrlf & "<tr id=""activitylisttotal""><td colspan=""6"" align=""right"">Totals</td>"
	response.write "<td align=""right"">" & FormatCurrency(lcl_PurchaseTotal,2,,,0) & "</td>"
	response.write "<td align=""right"">" & FormatCurrency(lcl_RefundTotal,2,,,0) & "</td>"
	response.write "</tr>"

	response.write vbcrlf & "</table>"
	'response.write vbcrlf & "</div>"
%>				  
	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%

'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void ShowRelatedPayment  iPaymentid, sClassname, sActivityno, iQuantity, sStartdate,  sBgcolor
'------------------------------------------------------------------------------------------------------------
Sub ShowRelatedPayment( ByVal iPaymentid, ByVal sClassname, ByVal sActivityno, ByVal iQuantity, ByVal sStartdate, ByRef sBgcolor )
	Dim sSql, oRs

	sSql = "SELECT p.paymentdate, ISNULL(p.relatedpaymentid,0) AS relatedpaymentid, p.paymenttotal "
	sSql = sSql & " FROM egov_class_payment p "
	sSql = sSql & " WHERE p.paymentid = " & iPaymentid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If sBgcolor = "#eeeeee" Then
			sBgcolor = "#ffffff" 
		Else
			sBgcolor = "#eeeeee"
		End If

		'Calulate the Payment Total
		If GetPaymentDetails( iPaymentid, iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminUserId, iPriorPaymentId ) Then 
			sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )

			If sJournalEntryType <> "" Then 
				Select Case sJournalEntryType
					Case "purchase"
						'Show purchase details
						sPaymentTotal = getPaymentTotal( iPaymentid, "credit", sJournalEntryType, 0, "" )
						sPaymentTotal = FormatCurrency(sPaymentTotal,2)
					Case "refund"
						'Show refund stuff
						'ShowPurchaseDetails iPaymentId, "debit", sJournalEntryType, iPriorPaymentId
						sPaymentTotal = getPaymentTotal( iPaymentid, "debit", sJournalEntryType, oPriorPaymentId, "" )
						sPaymentTotal = FormatCurrency(sPaymentTotal,2)
					Case "transfer"
						'Show citizen account transfer
						sPaymentTotal = FormatCurrency(0,2)
					Case "deposit"
						'Show citizen account deposit
						sPaymentTotal = FormatCurrency(0,2)
					Case "withdrawl"
						'Show citizen account withdrawl
						sPaymentTotal = FormatCurrency(0,2)
				End Select 
			Else 
				sPaymentTotal = FormatCurrency(0,2)
			End If 

			response.write vbcrlf & "<tr bgcolor=""" &  sBgcolor  & """ class=""tablelist"">"
			response.write "<td>" & sClassname & "</td>"
			response.write "<td align=""center"">" & iQuantity & "</td>"
			response.write "<td>" & oRs("paymentdate") & "</td>"
			response.write "<td>" & sStartdate & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & iPaymentid & """>" & iPaymentid & "</a></td>"
			response.write "<td align=""center"">" & sActivityno & "</td>"
			response.write "<td align=""center"">" & sPaymentTotal & "</td>"
			response.write "</tr>"

			If clng(oRs("relatedpaymentid")) > clng(0) Then 
				ShowRelatedPayment oRs("relatedpaymentid"), sClassname, sActivityno, iQuantity, sStartdate, sBgcolor 
			End If 
		Else 
			response.write vbcrlf & "<tr bgcolor=""" &  sBgcolor  & """ class=""tablelist"">"
			response.write "<td>" & sClassname & " (" & sActivityno & ")</td>"
			response.write "<td align=""center"">" & iQuantity & "</td>"
			response.write "<td>" & oRs("paymentdate") & "</td>"
			response.write "<td>" & sStartdate & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & iPaymentid & """>" & iPaymentid & "</a></td>"
			response.write "<td align=""center"">&nbsp;</td>"
			response.write "<td align=""center"">&nbsp;</td>"
			response.write "</tr>"

			If clng(oRs("relatedpaymentid")) > clng(0) Then 
				ShowRelatedPayment oRs("relatedpaymentid"), sClassname, sActivityno, iQuantity, sStartdate, sBgcolor 
			End If 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' boolean GetPaymentDetails( iPaymentId, iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminUserId, iPriorPaymentId )
'------------------------------------------------------------------------------------------------------------
Function GetPaymentDetails( ByVal iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate, ByRef iAdminLocationId, ByRef iJournalEntryTypeId, ByRef sNotes, ByRef iAdminUserId, ByRef iPriorPaymentId )
	Dim sSql, oRs

	sSql = "SELECT userid, paymenttotal, paymentdate, ISNULL(adminlocationid,0) AS adminlocationid, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, journalentrytypeid, notes, relatedpaymentid "
	sSql = sSql & " FROM egov_class_payment "
	sSql = sSql & " WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND orgid = " & Session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		iUserId = oRs("userid")
		nTotal = oRs("paymenttotal")
		dPaymentDate = oRs("paymentdate")
		iAdminLocationId = oRs("adminlocationid")
		iJournalEntryTypeId = oRs("journalentrytypeid")
		sNotes = oRs("notes")
		iAdminUserId = oRs("adminuserid")
		iPriorPaymentId = oRs("relatedpaymentid")
		GetPaymentDetails = True 
	Else 
		GetPaymentDetails = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' string GetJournalEntryType( iJournalEntryTypeId )
'------------------------------------------------------------------------------------------------------------
Function GetJournalEntryType( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

	If iJournalEntryTypeId <> "" Then 
		sSql = "SELECT journalentrytype "
		sSql = sSql & " FROM egov_journal_entry_types "
		sSql = sSql & " WHERE journalentrytypeid = " & iJournalEntryTypeId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetJournalEntryType = oRs("journalentrytype")
		Else 
			GetJournalEntryType = ""
		End If 

		oRs.Close
		Set oRs = Nothing
	Else 
		GetJournalEntryType = ""
	End If 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' double getPaymentTotal( iPaymentId, sEntryType, sJournalEntryType, iPriorPaymentId, iClassListID )
'------------------------------------------------------------------------------------------------------------
Function getPaymentTotal( ByVal iPaymentId, ByVal sEntryType, ByVal sJournalEntryType, ByVal iPriorPaymentId, ByVal iClassListID )
	Dim sSql, oRs, cTotal, cAmount

	cTotal = CDbl(0.00)

	' Pull a set of items purchased
	sSql = "SELECT itemtype, itemid, L.itemtypeid, sum(amount) as amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 0 "
	sSql = sSql & " AND entrytype = '" & sEntryType & "' "
	sSql = sSql & " AND L.paymentid = " & iPaymentId

	If iClassListId <> "" Then 
		sSql = sSql & " AND L.itemid = " & iClassListID
	End If 

	sSql = sSql & " GROUP BY itemtype, itemid, L.itemtypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		'Change to a case statement as you add more things that can be bought; like gifts and lodges or memberships
		If oRs("itemtype") = "recreation activity" Then 
			If sJournalEntryType = "refund" Then 
				cAmount = GetPriorPurchaseTotal( iPriorPaymentId, oRs("itemtypeid"), oRs("itemid") )
			Else 
				cAmount = CDbl(oRs("amount"))
			End If 
			cTotal = cTotal + cAmount
		End If 
		oRs.MoveNext
	Loop 

	oRs.close 
	Set oRs = Nothing

	If sJournalEntryType = "refund" Then
		cTotal = cTotal - ShowRefundFee( iPaymentId, cTotal )
	End If 

	getPaymentTotal = formatcurrency(cTotal,2)

End Function 


'------------------------------------------------------------------------------------------------------------
' double GetPriorPurchaseTotal( iPriorPaymentId, iItemTypeId, iItemId )
'------------------------------------------------------------------------------------------------------------
function GetPriorPurchaseTotal( ByVal iPriorPaymentId, ByVal iItemTypeId, ByVal iItemId )
	Dim sSql, oRs

	If iPriorPaymentId = "" Then 
		iPriorPaymentId = 0
	End If 

	'Pull a sum of what paid for prior class
	sSql = "SELECT sum(amount) as amount "
	sSql = sSql & " FROM egov_accounts_ledger "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND entrytype = 'credit' "
	sSql = sSql & " AND paymentid = "  & iPriorPaymentId
	sSql = sSql & " AND itemtypeid = " & iItemTypeId
	sSql = sSql & " AND itemid = "     & iItemId
	sSql = sSql & " GROUP BY itemtypeid, itemid"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPriorPurchaseTotal = CDbl(oRs("amount"))
	Else 
		GetPriorPurchaseTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' double ShowRefundFee( iPaymentId, cTotal )
'------------------------------------------------------------------------------------------------------------
Function ShowRefundFee( ByVal iPaymentId, ByVal cTotal )
	Dim sSql, oRs, cRefundShortage

	cRefundShortage = CDbl(0.00) 

	cRefundShortage = cTotal - GetRefundDebit( iPaymentId )

	' Pull a the refund fee row
	sSql = "SELECT itemtype, itemid, amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 "
	sSql = sSql & " AND entrytype = 'credit' AND P.isrefunddebit = 1 "
	sSql = sSql & " AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		ShowRefundFee = CDbl(oRs("amount") + cRefundShortage)
	Else 
		ShowRefundFee = CDbl(cRefundShortage)
	End If 
	
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------------------------------------
' double GetRefundDebit( iPaymentId )
'------------------------------------------------------------------------------------------------------------
Function GetRefundDebit( ByVal iPaymentId )
	Dim sSql, oRs

	'Pull a sum of what paid for prior class
	sSql = "SELECT sum(amount) as amount "
	sSql = sSql & " FROM egov_accounts_ledger "
	sSql = sSql & " WHERE  ispaymentaccount = 0 AND entrytype = 'debit' "
	sSql = sSql & " AND paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY paymentid "
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRefundDebit = CDbl(oRs("amount"))
	Else 
		GetRefundDebit = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' double  GetPurchaseTotal( iPaymentId )
'------------------------------------------------------------------------------------------------------------
Function GetPurchaseTotal( ByVal iPaymentId )
	Dim sSql, oRs

	'Pull a the purchase total sum
	sSql = "SELECT sum(amount) AS amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 0 "
	sSql = sSql & " AND entrytype = 'debit' AND L.paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY L.paymentid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPurchaseTotal = CDbl(oRs("amount"))
	Else 
		GetPurchaseTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

end function


'------------------------------------------------------------------------------------------------------------
' double GetRefundFeeAmount( iPaymentId )
'------------------------------------------------------------------------------------------------------------
Function GetRefundFeeAmount( ByVal iPaymentId )
	Dim sSql, oRs
	
	'Pull a the refund fee row
	sSql = "SELECT amount FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 "
	sSql = sSql & " AND entrytype = 'credit' AND P.isrefunddebit = 1 "
	sSql = sSql & " AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set  oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRefundFeeAmount = CDbl(oRs("amount"))
	Else 
		GetRefundFeeAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' string getHeadofHousehold( iUserid )
'------------------------------------------------------------------------------------------------------------
Function getHeadofHousehold( ByVal iUserid )
	Dim sSql, oRs 

	If iUserid <> "" Then 
		sSql = "SELECT userfname, userlname FROM egov_users "
		sSql = sSql & "WHERE userid = (select distinct belongstouserid "
		sSql = sSql & "FROM egov_familymembers WHERE userid = " & iUserid & ")"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			getHeadofHousehold = oRs("userfname") & " " & oRs("userlname")
		Else 
			getHeadofHousehold = ""
		End If 
	Else 
		getHeadofHousehold = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' void DrawDateChoices sName
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""18"">Tomorrow</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""14"">Next Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""13"">Next Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""15"">Next Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 

%>
