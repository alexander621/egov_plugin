<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: invoicedateedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/20/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits invoice dates for a permit invoice
'
' MODIFICATION HISTORY
' 1.0   08/20/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, iInvoiceid, sSql, oRs, sPaymentDate, iPermitId, sUpdateField

iInvoiceid = CLng(request("invoiceid"))
sUpdateField = request("updatefield")
iPaymentId = CLng(request("paymentid"))
iPermitId = CLng(request("permitid"))

sSql = "SELECT paymentdate FROM egov_class_payment WHERE paymentid = " & iPaymentId
sSql = sSql & " AND orgid = " & session("orgid")

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	sPaymentDate = DateValue(oRs("paymentdate"))
End If 

oRs.Close
Set oRs = Nothing 

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

		<script language="Javascript">
		<!--

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmDate", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doUpdate()
		{
			// Check the invoicedate
			if (document.frmDate.paymentdate.value == "")
			{
				alert("Please enter a date");
				document.frmDate.paymentdate.focus();
				return;
			}
			else
			{
				if (! isValidDate(document.frmDate.paymentdate.value))
				{
					alert("The date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmDate.paymentdate.focus();
					return;
				}
			}

			// Fire off AJAX call to update the date
			doAjax('paiddateupdate.asp', 'permitid=<%=iPermitId%>&paymentid=<%=iPaymentId%>&invoiceid=<%=iInvoiceid%>&originalpaymentdate=<%=sPaymentDate%>&paymentdate=' + document.frmDate.paymentdate.value, '', 'get', '0');

			// Update the parent window
			window.opener.document.getElementById("<%=sUpdateField%>").innerHTML = document.frmDate.paymentdate.value;

			// Close yourself
			doClose();
		}
		
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function init()
		{
			//document.getElementById("paymentdate").focus();
		}

		window.onload = init; 
  $( function() {
    $( "#paymentdate" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">

				<form name="frmDate" action="paiddateupdate.asp" method="post">
					<input type="hidden" name="paymentid" value="<%=iPaymentId%>" />
					<input type="hidden" name="invoiceid" value="<%=iInvoiceid%>" />
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<input type="hidden" name="originalpaymentdate" value="<%=sPaymentDate%>" />
					<p> 
						<input type="input" name="paymentdate" id="paymentdate" value="<%=sPaymentDate%>" size="10" maxlength="10" />
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" value="Save Changes" onclick="doUpdate();" />
						 &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

