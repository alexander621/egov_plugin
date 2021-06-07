<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: criticaldateedit.asp
' AUTHOR: Steve Loar
' CREATED: 05/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Edits critical dates for a permit
'
' MODIFICATION HISTORY
' 1.0   05/20/2008	Steve Loar - INITIAL VERSION
' 1.1	08/19/2010	Steve Loar - Changed from just expiration dates to critical dates
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, sCriticalDate, iCriticalDateType, sShowDateType

iPermitId = CLng(request("permitid"))
iCriticalDateType = CLng(request("criticaldatetype"))

Select Case iCriticalDateType
	Case 1
		sDateField = "applieddate"
		sShowDateType = "Applied"
	Case 2
		sDateField = "releaseddate"
		sShowDateType = "Released"
	Case 3
		sDateField = "approveddate"
		sShowDateType = "Approved"
	Case 4
		sDateField = "issueddate"
		sShowDateType = "Issued"
	Case 5
		sDateField = "expirationdate"
		sShowDateType = "Expiration"
	Case Else
		response.End	' This is not from the code and should be ended
End Select 

response.write "<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute(""data-close"")).innerHTML='" & sShowDateType & " Date';</script>"

' get the current value
sSql = "SELECT " & sDateField & " AS criticaldate FROM egov_permits WHERE permitid = " & iPermitId
sSql = sSql & " AND orgid = " & session("orgid")

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	sCriticalDate = DateValue(oRs("criticaldate"))
End If 

oRs.Close
Set oRs = Nothing 

%>

<html lang="en">
	<head>
		<meta charset="utf-8" />
		<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<style>
			div.inputfields
			{
				padding: 1em 0 0 1em;
			}

			span.calendarimg
			{
				padding-left: 1em;
				cursor: pointer;
			}
		</style>

		<script language="javascript" src="../scripts/isvaliddate.js"></script>
		<script language="javascript" src="../scripts/ajaxLib.js"></script>
		<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  		<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

		<script>
		<!--


		function doUpdate()
		{
			// Check the criticaldate
			if ($("#criticaldate").val() == "")
			{
				alert("Please enter a date");
				//$("#criticaldate").focus();
				return;
			}
			else
			{
				if (! isValidDate($("#criticaldate").val()))
				{
					alert("The date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					//$("#criticaldate").focus();
					return;
				}
			}

			//alert($("#criticaldate").val());
			//document.frmDate.submit();

			// Fire off AJAX call to update the date
			doAjax('criticaldateupdate.asp', 'permitid=<%=iPermitId%>&criticaldatetype=<%=iCriticalDateType%>&criticaldate=' + $("#criticaldate").val() + '&originalcriticaldate=' + $("#originalcriticaldate").val(), 'continueUpdate', 'get', '0');
			
		}

		function continueUpdate( data )
		{
			//alert( data );
			// Update the parent window
			parent.document.getElementById("<%=sDateField%>").innerHTML = $("#criticaldate").val();

			// Close yourself
			doClose();
		}
		
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function init()
		{
			//$("#criticaldate").focus();
		}

		$(document).ready(function() 
		{  
			//$("#criticaldate").focus();
		});

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">

				<form name="frmDate" action="criticaldateupdate.asp" method="post">
					<input type="hidden" id="permitid" name="permitid" value="<%=iPermitId%>" />
					<input type="hidden" id="originalcriticaldate" name="originalcriticaldate" value="<%=sCriticalDate%>" />
					<input type="hidden" id="criticaldatetype" name="criticaldatetype" value="<%=iCriticalDateType%>" />

<script>
  $( function() {
    $( "#criticaldate" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
  </script>
					<div class="inputfields"> 
						<input type="text" id="criticaldate" name="criticaldate" value="<%=sCriticalDate%>" size="10" maxlength="10" autocomplete="off" />
					</div>

					<div class="inputfields">
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" value="Save Changes" onclick="doUpdate();" />
						 &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
					</div>

				</form>
			</div>
		</div>
	</body>
</html>

