<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypechange.asp
' AUTHOR: Steve Loar
' CREATED: 08/17/201
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Changes the permit type for a permit
'
' MODIFICATION HISTORY
' 1.0   08/17/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitTypeId

iPermitId = CLng(request("permitid"))

iPermitTypeId = GetPermitTypeId( iPermitId ) ' in permitcommonfunctions.asp

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/jquery-1.4.2.min.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="Javascript">
		<!--

		function ChangePermitType()
		{
			if (parseInt($("#permittypeid").val()) == parseInt($("#originalpermittypeid").val()))
			{
				doClose();
			}
			else 
			{
				var sParameter = 'permitid=' + encodeURIComponent($("#permitid").val());
				sParameter += '&permittypeid=' + encodeURIComponent($("#permittypeid").val());
				sParameter += '&originalpermittypeid=' + encodeURIComponent($("#originalpermittypeid").val());
				//alert(sParameter);
				//alert( $("#permittypeid option:selected").text() );
				doAjax('permittypechangedo.asp', sParameter, 'updateAndClose', 'post', '0');

				//window.opener.document.getElementById("permittype").innerHTML = $("#permittypeid option:selected").text();
				//doClose();
				//document.frmPermitType.submit();
			}
		}

		function updateAndClose( sReturn )
		{
			parent.document.getElementById("permittype").innerHTML = $("#permittypeid option:selected").text();
			doClose();
		}

		function doClose()
		{
			//window.close();
			//window.opener.focus();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				
				<form name="frmPermitType" action="permittypechangedo.asp" method="post">
					<input type="hidden" id="permitid" name="permitid" value="<%=iPermitId%>" />
					<input type="hidden" id="originalpermittypeid" name="originalpermittypeid" value="<%=iPermitTypeId%>" />
					<p>
						<strong>Permit Types:</strong<br />
<%						ShowPermitTypes iPermitTypeId %>					
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Change Permit Type" onclick="ChangePermitType();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


%>
