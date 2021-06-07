<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()

If request.cookies("userid") = "" Then
	response.redirect "user_login.asp"
End If 

Dim iTransType, re, matches

Set re = New RegExp
re.Pattern = "^\d+$"

' trantype should only be 1 or 0
If request("trantype") <> "" Then
	iTransType = request("trantype")
	Set matches = re.Execute(iTransType)
	If matches.Count > 0 Then
		iTransType = CLng(iTransType)
	Else
		iTransType = CLng(0)
	End If 
Else
	iTransType = CLng(0)
End If 

if iTransType = CLng(1) then
sTitle = "Payment History"
  Session("RedirectLang") = "Return to Payment History"

else
  sTitle = "Action Request History"
  Session("RedirectLang") = "Return to Request History"
end if

'Check for org features
lcl_orghasfeature_issue_location    = orghasfeature(iorgid, "issue location")
lcl_orghasfeature_requestmergeforms = orghasfeature(iorgid, "requestmergeforms")

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName & " - " & sTitle %></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	<script language="javascript" src="scripts/modules.js"></script>
	<script language="javascript" src="scripts/easyform.js"></script>  

<script language="javascript">
<!--
function openWin2(url, name) {
			popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function PDFdocument(p_flid,p_action_autoid,p_orgid,p_user_id) 
{
  //var newFile = "http://secure.eclink.com/egovlink/action_line/actionline_pdf.asp?sys=DEV&iletterid=" + p_flid + "&action_autoid=" + p_action_autoid + "&orgid=" + p_orgid + "&userid=" + p_user_id;
  var newFile = "viewPDF.asp?action_autoid=" + p_action_autoid + "&orgid=" + p_orgid + "&userid=" + p_user_id;
  newWin = window.open(newFile);
  newWin.focus();
}
//-->
</script>
</head>

<!--#Include file="include_top.asp"-->

<font class="pagetitle">Welcome to <%=sOrgName%>&nbsp;<%=sTitle%></font><br />

<%	RegisteredUserDisplay( "" ) %>

<div id="content">
	 <div id="centercontent">

<!-- <div class="transactionreportshadow"> -->
<div>
<table border="0" cellspacing="0" cellpadding="2" class="transactionreport liquidtable" id="useractivitytable">
  <% List_Transactions iTransType %>
</table>
</div>

 	</div>
</div>

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
function List_Transactions( ByVal iTransType )
	Dim sSql, oRs
	lcl_public_actionline_pdf = ""

	if iTransType = CLng(0) then
		sSql = "SELECT * "
	else
		sSql = "SELECT DISTINCT paymentid, paymentdate, paymentservicename, Trans_Date "
	end if

	sSql = sSql & " FROM user_transaction_history2 "
	sSql = sSql & " WHERE orgid = " & iorgid
	sSql = sSql & " AND useremail = '" & LookUpInformation(request.cookies("userid")) & "' "
	sSql = sSql & " AND trantype = " & iTransType 
	sSql = sSql & " ORDER BY Trans_Date DESC"

	'response.write "<!-- " & sSql & " -->"

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	sBGColor = "#FFFFFF"

	if not oRs.eof then

   		blnFoundTrans = False 

  		'DISPLAY HEADER ROW
   		'if sTranType = "0" then
		strHeaderRow = ""
		If iTransType = CLng(0) Then 

			'Action Line ITEM
			strHeaderRow = strHeaderRow & "    <td align=""center"" class=""transaction_header"" nowrap=""nowrap"">Submit Date</td>" & vbcrlf
			strHeaderRow = strHeaderRow & "    <td align=""center"" class=""transaction_header"" nowrap=""nowrap"">Current Status</td>" & vbcrlf
			strHeaderRow = strHeaderRow & "    <td align=""center"" class=""transaction_header"" nowrap=""nowrap"">Request Name</td>" & vbcrlf
			strHeaderRow = strHeaderRow & "    <td align=""center"" class=""transaction_header"" nowrap=""nowrap"">Date of Last Activity</td>" & vbcrlf

			if lcl_orghasfeature_issue_location then
				lcl_featurename_issuelocation = getFeatureName("issue location")
				strHeaderRow = strHeaderRow & "    <td align=""center"" class=""transaction_header"" nowrap=""nowrap"">" & lcl_featurename_issuelocation & "</td>" & vbcrlf
			end if

			strHeaderRow = strHeaderRow & "    <td align=""right"" class=""transaction_header"">&nbsp;</td>" & vbcrlf

			if lcl_orghasfeature_requestmergeforms then
				strHeaderRow = strHeaderRow & "    <td align=""right"" class=""transaction_header"">&nbsp;</td>" & vbcrlf
			end if

			response.write "<thead><tr>" & strHeaderRow & "</tr></thead>"
		else
			'PAYMENT ITEM
			strHeaderRow = strHeaderRow & "    <td align=""center"" width=""150"" class=""transaction_header"">Payment Date</td>" & vbcrlf
			strHeaderRow = strHeaderRow & "    <td align=""center"" width=""200"" class=""transaction_header"">Payment Type</td>" & vbcrlf
			strHeaderRow = strHeaderRow & "    <td align=""right"" class=""transaction_header"">&nbsp;</td>" & vbcrlf
			response.write "<thead><tr>" & strHeaderRow & "</tr></thead>"
		end if

		sBGColor = "#ffffff"
		Do While Not oRs.EOF
			sBGColor = changeBGColor(sBGColor,"#ffffff","#e0e0e0")

			'Make sure valid transactions exist
			if iTransType = CLng(0) then
				if not isnull(oRs("category_id")) then
					lcl_tracking_number = getActionLineTrackingNumber(oRs("action_autoid"))

					'Determine if the issue/problem location is to be displayed
					'If so, then get all of the street data.
					if lcl_orghasfeature_issue_location then
						sSql = "SELECT streetnumber, "
						sSql = sSql & " streetprefix, "
						sSql = sSql & " streetaddress, "
						sSql = sSql & " streetsuffix, "
						sSql = sSql & " streetdirection "
						sSql = sSql & " FROM egov_action_response_issue_location "
						sSql = sSql & " WHERE actionrequestresponseid = " & oRs("action_autoid")

						set oGetIssueLocation = Server.CreateObject("ADODB.Recordset")
						oGetIssueLocation.Open sSql, Application("DSN"), 3, 1

						if not oGetIssueLocation.eof then
							lcl_streetnumber    = oGetIssueLocation("streetnumber")
							lcl_streetprefix    = oGetIssueLocation("streetprefix")
							lcl_streetaddress   = oGetIssueLocation("streetaddress")
							lcl_streetsuffix    = oGetIssueLocation("streetsuffix")
							lcl_streetdirection = oGetIssueLocation("streetdirection")
						else
							lcl_streetnumber    = ""
							lcl_streetprefix    = ""
							lcl_streetaddress   = ""
							lcl_streetsuffix    = ""
							lcl_streetdirection = ""
						end if

						'Build the street name
						lcl_street_name = buildStreetAddress(lcl_streetnumber, lcl_streetprefix, lcl_streetaddress, lcl_streetsuffix, lcl_streetdirection)
						if lcl_street_name = "" or isnull(lcl_street_name) then lcl_street_name = "N/A"
					end if

					set oGetIssueLocation = nothing

					'Action Line ITEM
					response.write "<tr style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
					response.write "    <td class=""transaction_header repeatheader"">Submit Date</td>" & vbcrlf
					response.write "    <td align=""left"" width=""150"" nowrap=""nowrap"">"  & oRs("submit_date") & "</td>" & vbcrlf
					response.write "    <td class=""transaction_header repeatheader"">Current Status</td>" & vbcrlf
					response.write "    <td align=""center"" width=""150"">" & oRs("status") & "</td>" & vbcrlf
					response.write "    <td class=""transaction_header repeatheader"">Request Name</td>" & vbcrlf
					response.write "    <td align=""left"" width=""200"">"   & oRs("category_title") & "</td>" & vbcrlf
					response.write "    <td class=""transaction_header repeatheader"">Date of Last Activity</td>" & vbcrlf
					response.write "    <td align=""left"" width=""150"" nowrap=""nowrap"">" & GetDateofLastActivity(oRs("action_autoid"),oRs("submit_date")) & "</td>" & vbcrlf
					'response.write "    <td align=""right"" nowrap><a href=""view_request.asp?REQUEST_ID=" & oRs("action_autoid")& "0000"">View Details...</a></td>" & vbcrlf
					'response.write "    <td align=""center""><input type=""button"" name=""sViewDetails"" id=""sViewDetails"" value=""View Details"" class=""button"" onclick=""location.href='view_request.asp?request_id=" & oRs("action_autoid")& "'"" /></td>" & vbcrlf

					if lcl_orghasfeature_issue_location then
						response.write "    <td class=""transaction_header repeatheader"">" & lcl_featurename_issuelocation & "</td>" & vbcrlf
						response.write "    <td align=""left"" width=""150"" nowrap=""nowrap"">" & lcl_street_name & "</td>" & vbcrlf
					end if

					response.write "    <td align=""center""><input type=""button"" name=""sViewDetails"" id=""sViewDetails"" value=""View Details"" class=""button bottomspace"" onclick=""location.href='action_request_lookup.asp?request_id=" & lcl_tracking_number & "'"" /></td>" & vbcrlf

					if lcl_orghasfeature_requestmergeforms then
						lcl_user_id = request.cookies("userid")

						'Check to see if a PDF has been associated with the request
						sSqlpdf = "SELECT isnull(r.public_actionline_pdf,f.public_actionline_pdf) as public_actionline_pdf "
						sSqlpdf = sSqlpdf & " FROM egov_action_request_forms f, egov_actionline_requests r "
						sSqlpdf = sSqlpdf & " WHERE f.action_form_id = r.category_id "
						sSqlpdf = sSqlpdf & " AND r.action_autoid = " & oRs("action_autoid")

						set rsp = Server.CreateObject("ADODB.Recordset")
						rsp.Open sSqlpdf, Application("DSN"), 3, 1

						if not rsp.eof then
							lcl_public_actionline_pdf = rsp("public_actionline_pdf")

							if lcl_public_actionline_pdf <> "" then
								response.write "    <td align=""right"" nowrap>" & vbcrlf
								'response.write "        <a href=""javascript:PDFdocument(" & lcl_public_actionline_pdf & "," & oRs("action_autoid") & "," & iorgid & "," & lcl_user_id & ")"">View PDF...</a>" & vbcrlf
								response.write "        <input type=""button"" name=""sViewPDF"" id=""sViewPDF"" value=""View PDF"" class=""button"" onClick=""window.open('viewPDF.asp?iRequestID=" & oRs("action_autoid") & "');"" />" & vbcrlf
								response.write "    </td>" & vbcrlf
							else
								response.write "    <td>&nbsp;</td>" & vbcrlf
							end if
						else
							response.write "    <td>&nbsp;</td>" & vbcrlf
						end if

						set rsp = nothing

					end if

					response.write "</tr>"

					blnFoundTrans = True

				end if
			else
				'Payment ITEM
				response.write "<tr>" & vbcrlf
				response.write "    <td align=""right"" width=""150"">"  & oRs("paymentdate") & "</td>" & vbcrlf
				response.write "    <td align=""left""><strong>Payment: "  & oRs("paymentservicename") & "</strong></td>" & vbcrlf
				'response.write "    <td align=""right"" nowrap=""nowrap""><A HREF=""view_receipt.asp?PAYMENT_ID=" & oRs("paymentid")& """>View Receipt...</A></td>" & vbcrlf
				response.write "    <td align=""right"" nowrap=""nowrap""><input type=""button"" name=""sViewReceipt"" id=""sViewReceipt"" value=""View Receipt"" class=""button"" onclick=""location.href='view_receipt.asp?PAYMENT_ID=" & oRs("paymentid")& "'"" /></td>" & vbcrlf
				response.write "</tr>"

				blnFoundTrans = True
			end if

			oRs.movenext
		Loop 
	end if

	'No Transactions
	if not blnFoundTrans then
		response.write "<tr><td>No transactions Found</td></tr>" & vbcrlf
	end if

	oRs.Close
	set oRs = nothing

end function


'------------------------------------------------------------------------------
Function LookUpInformation( ByVal sUserID )
	Dim sSql, oRs, sReturnValue

	sReturnValue = "UNKNOWN"

	'sSql = "SELECT useremail FROM egov_users WHERE userid = " & CLng(request("userid"))
	sSql = "SELECT useremail FROM egov_users WHERE userid = " & CLng(sUserID)
	
	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.eof Then 
		sReturnValue = oRs("useremail")
	End If 

	Set oRs = Nothing 

	LookUpInformation = sReturnValue
	
End Function 


'------------------------------------------------------------------------------
function GetDateofLastActivity(iRequestID,datSubmitDate)

	 sReturnValue = datSubmitDate

	 sSql = "SELECT TOP 1 action_editdate "
  sSql = sSql & " FROM egov_action_responses "
  sSql = sSql & " WHERE (action_autoid = '" & iRequestID & "') "
  sSql = sSql & " ORDER BY action_editdate DESC"

	 set oDate = Server.CreateObject("ADODB.Recordset")
	 oDate.Open sSql, Application("DSN"), 3, 1

	 if not oDate.eof then
	   	sReturnValue = oDate("action_editdate")
 	end if

	 set oDate = nothing

	 GetDateofLastActivity = sReturnValue

end function
%>
