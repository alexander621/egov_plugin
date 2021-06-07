<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feepicker.asp
' AUTHOR: Steve Loar
' CREATED: 04/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects permit fees
'
' MODIFICATION HISTORY
' 1.0   04/11/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId

iPermitId = CLng(request("permitid"))

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
		<script language="javascript" src="../scripts/modules.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--

		function doSelect()
		{
			if (document.frmFee.permitfeetypeid.options[document.frmFee.permitfeetypeid.selectedIndex].value > 0)
			{
				//document.frmFee.submit();
				// Save the new fee via Ajax
				doAjax('feepickeradd.asp', 'permitid=' + document.frmFee.permitid.value + '&permitfeetypeid=' + document.frmFee.permitfeetypeid.options[document.frmFee.permitfeetypeid.selectedIndex].value, 'doNewRow', 'get', '0');
			}
			else
			{
				alert("Closing");
				doClose();
			}
		}

		function doNewRow( sReturnJSON )
		{
			var json = sReturnJSON.evalJSON(true); 
			//alert( json.flag );

			if (document.frmFee.permitfeetypeid.options[document.frmFee.permitfeetypeid.selectedIndex].value > 0)
			{
				parent.document.getElementById("maxfees").value = parseInt(parent.document.getElementById("maxfees").value) + 1;
				var tbl = parent.document.getElementById("feelist");
				var lastRow = tbl.rows.length;
				var cellCount = tbl.rows[0].cells.length;
				//alert( cellCount );
				var newRow = parseInt(parent.document.getElementById("maxfees").value);
				var row = tbl.insertRow(lastRow);
				if (newRow % 2 == 0)
				{
					row.className = "altrow";
				}
				//row.onMouseOver = parent.myMouseOver.eventHandler.bindAsEventListener(parent.myMouseOver);
				//row.onMouseOut = parent.myMouseOut.eventHandler.bindAsEventListener(parent.myMouseOut);

				// Add the Remove cell
				var cell = row.insertCell(0);
				cell.title = "Click Save Changes to Complete this add";
				cell.align = 'center';
				cell.innerHTML = '<input type="hidden" name="permitfeeid' + newRow + '" value="' + json.permitfeeid + '" /><input type="checkbox" name="removefee' + newRow + '" id="removefee' + newRow + '" />';

				// Include cell
				//cell = row.insertCell(1);
				//cell.title = "Click Save Changes to Complete this add";
				//cell.innerHTML = '&nbsp;';

				// Category cell
				cell = row.insertCell(1);
				cell.title = "Click Save Changes to Complete this add";
				cell.align = 'center';
				cell.innerHTML = json.permitfeeprefix;

				// Description cell
				cell = row.insertCell(2);
				cell.title = "Click Save Changes to Complete this add";
				//cell.innerHTML = document.frmFee.permitfeetypeid.options[document.frmFee.permitfeetypeid.selectedIndex].text;
				cell.innerHTML = json.permitfee;

				// Method cell
				cell = row.insertCell(3);
				cell.title = "Click Save Changes to Complete this add";
				cell.align = 'center';
				cell.innerHTML = json.permitfeemethod;

				// Fee Amount cell or up front fee cell
				cell = row.insertCell(4);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				if (cellCount > 5)
				{
					// Fee Amount when there is an up front fee column
					cell = row.insertCell(5);
					cell.title = "Click Save Changes to Complete this add";
					cell.innerHTML = '&nbsp;';
				}
			}
			doClose();
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<form name="frmFee" action="feepickeradd.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
					<%	ShowPermitFeeTypes %>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Selected Fee" onclick="doSelect();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" /> 
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitFeeTypes
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFeeTypes()
	Dim sSql, oRs

	iFeeCount = 0

	sSql = "SELECT permitfeetypeid, ISNULL(permitfeeprefix,'') AS permitfeeprefix, permitfee "
	sSql = sSql & "FROM egov_permitfeetypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitfee, 2"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""permitfeetypeid"">"
		response.write vbcrlf & "<option value=""0"">Select a fee to add to this permit</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitfeetypeid") & """>" 
			response.write oRs("permitfee")
			If oRs("permitfeeprefix") <> "" Then
				response.write " (" & oRs("permitfeeprefix") & ")"
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
