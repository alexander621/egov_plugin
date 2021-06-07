<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: invoicecreate.asp
' AUTHOR: Steve Loar
' CREATED: 04/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Displays the fixture fees and allows input of quantities and to select for the permit.
'
' MODIFICATION HISTORY
' 1.0   04/29/2008	Steve Loar - INITIAL VERSION
' 1.1	10/07/2008	Steve Loar - Changed to handle up front fees for the initial invoice
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iBillToDefault, sTotalAmount, iMaxFees, iWaiveAllFees, sWaivedMsg, bIsUpFrontInvoice

iPermitId = CLng(request("permitid"))
iMaxFees = 0

If OrgHasFeature( "up front fees" ) Then
	If PermitHasNoInvoices( iPermitId ) And PermitHasUpFrontFees( iPermitId ) Then 
		bIsUpFrontInvoice = True 
	Else
		bIsUpFrontInvoice = False 
	End If 
Else
	bIsUpFrontInvoice = False 
End If 

iBillToDefault = GetBillToDefault( iPermitId )

sTotalAmount = GetInvoiceTotal( iPermitId, bIsUpFrontInvoice )

iWaiveAllFees = GetWaiveAllFeesFlag( iPermitId )

If iWaiveAllFees Then
	sWaivedMsg = " &nbsp; &nbsp; ** All fees will be waived for this invoice. **"
	iWaiveAllFees = 1
Else
	sWaivedMsg = ""
	iWaiveAllFees = 0
End If 

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/formatnumber.js"></script>
		<script language="JavaScript" src="../scripts/removespaces.js"></script>
		<script language="JavaScript" src="../scripts/removecommas.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="Javascript">
		<!--

			function doClose()
			{
				parent.hideModal(window.frameElement.getAttribute("data-close"));
			}

			function ChangeInvoiceTotal( iRow )
			{
				// validate the input
				if (document.getElementById("invoiceamount" + iRow).value != '')
				{
					// Remove any extra spaces
					document.getElementById("invoiceamount" + iRow).value = removeSpaces(document.getElementById("invoiceamount" + iRow).value);
					//Remove commas that would cause problems in validation
					document.getElementById("invoiceamount" + iRow).value = removeCommas(document.getElementById("invoiceamount" + iRow).value);

					rege = /^\d*\.?\d{0,2}$/;
					Ok = rege.test(document.getElementById("invoiceamount" + iRow).value);
					if ( ! Ok )
					{
						alert("The invoice amount must be numeric with up to two decimal places.\nPlease correct this.");
						document.getElementById("invoiceamount" + iRow).focus();
						return false;
					}
					else
					{
						document.getElementById("invoiceamount" + iRow).value = format_number(Number(document.getElementById("invoiceamount" + iRow).value),2);
					}
				}
				else
				{
					document.getElementById("invoiceamount" + iRow).value = 0.00;
				}

				// recalculate the total
				RecalcInvoiceTotal();
			}

			function RecalcInvoiceTotal()
			{
				var InvoiceTotal = 0.00;
				for (var t = 1; t <= parseInt(document.frmFee.maxFeeCount.value); t++)
				{
					if (document.getElementById("include" + t).checked == true)
					{
						if (document.getElementById("invoiceamount" + t).value != '')
						{
							InvoiceTotal += parseFloat(document.getElementById("invoiceamount" + t).value);
						}
					}
				}
				document.getElementById("invoicetotal").innerHTML = format_number(InvoiceTotal,2);
				document.getElementById("totalamount").value = format_number(InvoiceTotal,2);
			}

			function doCreate()
			{
				// Create the parameter list and do the AJAX call
				var sParameter = 'permitid=' + encodeURIComponent(document.frmFee.permitid.value);
				sParameter += '&maxFeeCount=' + encodeURIComponent(document.frmFee.maxFeeCount.value);
				sParameter += '&totalamount=' + encodeURIComponent(document.frmFee.totalamount.value);
				sParameter += '&permitcontactid=' + encodeURIComponent(document.frmFee.permitcontactid.value);
				sParameter += '&allfeeswaived=' + encodeURIComponent(document.frmFee.allfeeswaived.value);
				for (var a = 1; a <= parseInt(document.frmFee.maxFeeCount.value); a++)
				{
					sParameter += '&permitfeeid' + a + '=' + encodeURIComponent(document.getElementById("permitfeeid" + a).value);
					sParameter += '&include' + a + '=' + encodeURIComponent(document.getElementById("include" + a).checked);
					sParameter += '&invoiceamount' + a + '=' + encodeURIComponent(document.getElementById("invoiceamount" + a).value);
				}
				//alert( sParameter );

				doAjax('invoiceupdate.asp', sParameter, 'UpdateParent', 'post', '0');
				//document.frmFee.submit();

			}

			function UpdateParent( sReturn )
			{
				//alert( sReturn);
				// Add an invoice row to the parent
				var tbl = parent.document.getElementById("invoicelist");
				var lastRow = tbl.rows.length;
				var row = tbl.insertRow(lastRow);
				if ( lastRow % 2 == 0 )
				{
					row.className = 'altrow'; 
				}

				// Invoice number cell
				var cell = row.insertCell(0);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = sReturn;
				// include cell
				cell = row.insertCell(1);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = '<%=date()%>';
				// Bill To Cell
				cell = row.insertCell(2);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				name = document.frmFee.permitcontactid.options[document.frmFee.permitcontactid.selectedIndex].text;
				nameArr = name.split(" – ");
				cell.innerHTML = nameArr[0];
				// Status cell
				cell = row.insertCell(3);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = 'Due';
				// invoice total cell
				cell = row.insertCell(4);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = document.frmFee.totalamount.value;
				// Paid Date cell
				cell = row.insertCell(5);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = '&nbsp;';
				// amount paid cell
				cell = row.insertCell(6);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = '0.00';
				// void button cell
				cell = row.insertCell(7);
				cell.align = 'center';
				cell.title = "Click Save Changes to Complete this new invoice";
				cell.innerHTML = '&nbsp;';

				//parent.RefreshPageAfterVoid( "New Invoice" );
				parent.RefreshPageAfterVoid( "NewInvoice|"+sReturn );

				// Show the new invoice
				//location.href="viewinvoice.asp?invoiceid=" + sReturn;
			}

			function ShowPick()
			{
				alert(document.frmFee.permitcontactid.options[document.frmFee.permitcontactid.selectedIndex].text);
			}

			function refreshParent()
			{
				//parent.RefreshPageAfterVoid( "New Invoice" );
			}

		//-->
		</script>

	</head>
	<body onunload="refreshParent()">
		<div id="content">
			<div id="centercontent">

				<form name="frmFee" action="invoicecreate.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<input type="hidden" id="totalamount" name="totalamount" value="<%=sTotalAmount%>" />
					<input type="hidden" name="allfeeswaived" value="<%=iWaiveAllFees%>" />
					<br />
					<br />

<%					
					tooltipclass=""
					tooltip = ""
					disabled = ""
					If CDbl(sTotalAmount) = CDbl(0.00) Then
						tooltipclass="tooltip"
						disabled = " disabled "
						tooltip = "<span class=""tooltiptext"">You cannot create an invoice<br />because the total is zero.</span>"
					end if
					%>
					<button <%=disabled%> style="margin-left:30px;" type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="doCreate();">Create Invoice<%=tooltip%></button> &nbsp; &nbsp;
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
					<p>
						Bill To: <% ShowBillingChoices iPermitId, iBillToDefault %>
					</p>
					<p>
						<span id="invoicetotaldisplayline">Invoice Total: <span id="invoicetotal"><%=sTotalAmount%></span><%=sWaivedMsg%></span>
					</p>
					<p>
						<% iMaxFees = ShowFeesToInvoice( iPermitId, bIsUpFrontInvoice ) %>
					</p>
					<input type="hidden" name="maxFeeCount" id="maxFeeCount" value="<%=iMaxFees%>" />
					<button <%=disabled%> style="margin-left:30px;" type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="doCreate();">Create Invoice<%=tooltip%></button> &nbsp; &nbsp;
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetInvoiceTotal( iPermitId, bIsUpFrontInvoice )
'--------------------------------------------------------------------------------------------------
Function GetInvoiceTotal( ByVal iPermitId, ByVal bIsUpFrontInvoice )
	Dim sSql, oRs, sTotalAmount

	sTotalAmount = CDbl(0.00)

	sSql = "SELECT permitfeeid, ISNULL(feeamount,0.00) AS feeamount, ISNULL(invoicedamount,0.00) AS invoicedamount, ISNULL(upfrontamount,0.00) AS upfrontamount "
	sSql = sSql & " FROM egov_permitfees WHERE includefee = 1 AND ISNULL(invoicedamount,0.00) <= ISNULL(feeamount,0.00) AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY displayorder"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If Not bIsUpFrontInvoice Then 
			sTotalAmount = sTotalAmount + CDbl(oRs("feeamount")) - CDbl(oRs("invoicedamount"))
		Else
			sTotalAmount = sTotalAmount + CDbl(oRs("upfrontamount"))
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	GetInvoiceTotal = FormatNumber(sTotalAmount,2,,,0)

End Function


'--------------------------------------------------------------------------------------------------
' Function ShowFeesToInvoice( iPermitId, bIsUpFrontInvoice )
'--------------------------------------------------------------------------------------------------
Function ShowFeesToInvoice( iPermitId, bIsUpFrontInvoice )
	Dim sSql, oRs, iRow, sInvoiceAmount, sTotalAmount

	iRow = 0
	sTotalAmount = CDbl(0.00)
	sSql = "SELECT permitfeeid, ISNULL(permitfeeprefix,'') AS permitfeeprefix, permitfee, ISNULL(feeamount,0.00) AS feeamount, "
	sSql = sSql & " ISNULL(upfrontamount,0.00) AS upfrontamount, ISNULL(invoicedamount,0.00) AS invoicedamount "
	sSql = sSql & " FROM egov_permitfees WHERE includefee = 1 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY displayorder"
	'  AND ISNULL(invoicedamount,0.00) != ISNULL(feeamount,0.00)
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table cellpadding=""2"" cellspacing=""0"" border=""0"" class=""feetable"" id=""invoicefeelist"">"
		response.write vbcrlf & "<caption>Fee List</caption>"
		response.write vbcrlf & "<tr><th>Include</th><th>Category</th><th>Fee Description</th>"
		If bIsUpFrontInvoice Then 
			response.write "<th>Up Front Fee</th>"
		Else 
			response.write "<th>Fee Amount</th>"
		End If 
		response.write "<th>Already Invoiced</th><th>Amount to Invoice</th></tr>"
		Do While Not oRs.EOF
			If (Not bIsUpFrontInvoice And CDbl(oRs("invoicedamount")) <> CDbl(oRs("feeamount"))) Or (bIsUpFrontInvoice And CDbl(oRs("upfrontamount")) > CDbl(0.00)) Then 
				iRow = iRow + 1
				If bIsUpFrontInvoice Then 
					sInvoiceAmount = CDbl(oRs("upfrontamount"))
				Else 
					sInvoiceAmount = CDbl(oRs("feeamount")) - CDbl(oRs("invoicedamount"))
				End If 
				sTotalAmount = sTotalAmount + sInvoiceAmount
				response.write vbcrlf & "<tr"
				If iRow Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If
				response.write ">"
				response.write "<td align=""center""><input type=""checkbox"" id=""include" & iRow & """ name=""include" & iRow & """ "
				If sInvoiceAmount > CDbl(0.00) Then 
					response.write " checked=""checked"" "
				End If 
				response.write " onclick=""RecalcInvoiceTotal();"" />"
				response.write "<input type=""hidden"" id=""permitfeeid" & iRow & """ name=""permitfeeid" & iRow & """ value=""" & oRs("permitfeeid") & """ />"
				response.write "</td>"
				response.write "<td>" & oRs("permitfeeprefix") & "</td>"
				response.write "<td>" & oRs("permitfee") & "</td>"
				response.write "<td align=""center"">"
				If bIsUpFrontInvoice Then 
					response.write FormatNumber(oRs("upfrontamount"),2,,,0)
				Else 
					response.write FormatNumber(oRs("feeamount"),2,,,0)
				End If 
				response.write "<input type=""hidden"" name=""feeamount" & iRow & """ id=feeamount" & iRow & """ value="""
				If bIsUpFrontInvoice Then
					response.write oRs("upfrontamount")
				Else 
					response.write oRs("feeamount")
				End If 
				response.write """ />"
				response.write "</td>"
				response.write "<td align=""center"">" & FormatNumber(oRs("invoicedamount"),2,,,0) & "<input type=""hidden"" name=""invoicedamount" & iRow & """ id=invoicedamount" & iRow & """ value=""" & oRs("feeamount") & """ /></td>"
				response.write "<td align=""center""><input type=""text"" id=""invoiceamount" & iRow & """ name=""invoiceamount" & iRow & """ value=""" & FormatNumber(sInvoiceAmount,2,,,0) & """ size=""9"" maxlength=""9"" onblur=""ChangeInvoiceTotal('" & iRow & "')"" /></td>"
				response.write "</tr>"
			End If 

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	Else 
		response.write "There are no fees to invoice at this time."
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowFeesToInvoice = iRow

End Function 


'--------------------------------------------------------------------------------------------------
' Function PermitHasNoInvoices( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitHasNoInvoices( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(invoiceid) AS hits FROM egov_permitinvoices WHERE isvoided = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasNoInvoices = False 
		Else
			PermitHasNoInvoices = True 
		End If 
	Else
		PermitHasNoInvoices = True 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function PermitHasUpFrontFees( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitHasUpFrontFees( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(upfrontamount),0.00) AS upfronttotal FROM egov_permitfees WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CDbl(oRs("upfronttotal")) > CDbl(0.00) Then
			PermitHasUpFrontFees = True 
		Else
			PermitHasUpFrontFees = False 
		End If 
	Else
		PermitHasUpFrontFees = False  
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetBillToDefault( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetBillToDefault( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitcontactid FROM egov_permitcontacts WHERE permitid = " & iPermitId
	sSql = sSql & " AND (isapplicant = 1 OR isbillingcontact = 1) "
	sSql = sSql & " ORDER BY permitcontactid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		' Grab the first one and pass that back
		GetBillToDefault = CLng(oRs("permitcontactid"))
	Else
		' someting is wrong
		GetBillToDefault = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetWaiveAllFeesFlag( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetWaiveAllFeesFlag( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT waiveallfees FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetWaiveAllFeesFlag = oRs("waiveallfees")
	Else
		GetWaiveAllFeesFlag = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowBillingChoices( iPermitId, iBillToDefault )
'--------------------------------------------------------------------------------------------------
Sub ShowBillingChoices( iPermitId, iBillToDefault )
	Dim sSql, oRs

	sSql = "SELECT permitcontactid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname, "
	sSql = sSql & " isapplicant, isbillingcontact, isprimarycontact, isprimarycontractor, isarchitect, iscontractor, ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSql & " FROM egov_permitcontacts WHERE ispriorcontact = 0 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permitcontactid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitcontactid") & """"
			If CLng(iBillToDefault) = CLng(oRs("permitcontactid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			If oRs("firstname") <> "" Then
				response.write oRs("firstname") & " " & oRs("lastname")
				bName = True 
			Else
				bName = False 
			End If 
			If oRs("company") <> "" Then
				If bName Then 
					response.write " ("
				End If 
				response.write oRs("company")
				If bName Then
					response.write ")"
				End If 
			End If 
			If oRs("isapplicant") Then
				response.write " &ndash; Applicant"
			Else 
				If oRs("isbillingcontact") Then
					response.write " &ndash; Billing Contact"
				Else 
					If oRs("isprimarycontact") Then
						response.write " &ndash; Primary Contact"
					Else 
						If oRs("isprimarycontractor") Then
							response.write " &ndash; Primary Contractor"
						Else 
							If oRs("isarchitect") Then
								response.write " &ndash; Architect/Engineer"
							Else 
								If oRs("iscontractor") Then
									sContractorType = GetContractorType( oRs("contractortypeid") )
									If sContractorType <> "" Then 
										response.write " &ndash; " & sContractorType
									Else 
										response.write " &ndash; Contractor"
									End If 
								End If
							End If
						End If
					End If
				End If 
			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
