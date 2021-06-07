<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: fixturefeeedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Displays the fixture fees and allows input of quantities and to select for the permit.
'
' MODIFICATION HISTORY
' 1.0   04/29/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, iPermitId, sFeeName, iMaxFeeCount, sAppliedFee, bPermitIsCompleted
Dim bIsOnHold

iPermitFeeId = CLng(request("permitfeeid"))

iPermitId = GetPermitIdByPermitFeeId( iPermitFeeId )

sFeeName = GetPermitFeeName( iPermitFeeId )

iMaxFeeCount = CLng(0)

sAppliedFee = GetAppliedFeeAmount( iPermitFeeId ) ' In permitcommonfunctions.asp

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="javascript" src="../scripts/modules.js"></script>
		<script language="JavaScript" src="../scripts/formatnumber.js"></script>
		<script language="JavaScript" src="../scripts/removespaces.js"></script>
		<script language="JavaScript" src="../scripts/removecommas.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
		<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

		<script language="Javascript">
		<!--

		function ViewStepTable( iPermitFixtureId )
		{
			//var w = (screen.width - 900)/2;
			//var h = (screen.height - 400)/2;
			//eval('window.open("fixturefeesteps.asp?permitfixtureid=' + iPermitFixtureId + '", "_steptable", "width=900,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			parent.showModal('fixturefeesteps.asp?permitfixtureid=' + iPermitFixtureId, 'Fixture Step Fee Table', 50, 30);
		}

		function checkInclude( oQtyField, iRow )
		{
			if (oQtyField.value != '' && oQtyField.value != '0')
			{
				document.getElementById("include" + iRow).checked = true;
			}
		}

		function doUpdate()
		{
			var rege;
			var Ok; 

			if (parseInt(document.frmFee.maxFeeCount.value) > 0)
			{
				// Check the step table values entered
				for (var t = 1; t <= parseInt(document.frmFee.maxFeeCount.value); t++)
				{
					if (document.getElementById("qty" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("qty" + t).value = removeSpaces(document.getElementById("qty" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("qty" + t).value = removeCommas(document.getElementById("qty" + t).value);
			
						// Validate the at least quantity format
						rege = /^\d+$/;
						Ok = rege.test(document.getElementById("qty" + t).value);
						if ( ! Ok )
						{
							alert("The Quantity should be blank or a whole number value.\nPlease correct this and try saving again.");
							document.getElementById("qty" + t).focus();
							return;
						}
					}
					else
					{
						document.getElementById("qty" + t).value = 0;
					}
				}

				// build the parameter list
				var sParameter = 'permitfeeid=' + encodeURIComponent(document.frmFee.permitfeeid.value);
				sParameter += '&permitid=' + encodeURIComponent(document.frmFee.permitid.value);
				sParameter += '&maxfeecount=' + encodeURIComponent(document.frmFee.maxFeeCount.value);
				for (var a = 1; a <= parseInt(document.frmFee.maxFeeCount.value); a++)
				{
					sParameter += '&permitfixtureid' + a + '=' + encodeURIComponent(document.getElementById("permitfixtureid" + a).value);
					sParameter += '&include' + a + '=' + encodeURIComponent(document.getElementById("include" + a).checked);
					sParameter += '&qty' + a + '=' + encodeURIComponent(document.getElementById("qty" + a).value);
				}
				//alert(sParameter);
				doAjax('fixturefeeupdate.asp', sParameter, 'UpdateFees', 'post', '0');
				//document.frmFee.submit();
			}
			else
			{
				// Repost this page as they may have added some fees
				location.href = "fixturefeeedit.asp?permitfeeid=" + <%=iPermitFeeId%> + "&success=1";
			}
		}

		function UpdateFees( sReturn )
		{
			//alert( sReturn)
			// Update the fee amount
			parent.document.getElementById("fee" + document.frmFee.permitfeeid.value).innerHTML = sReturn;
			// Get the new fee total for the permit
			doAjax('getpermitfeetotal.asp', 'permitid=' + document.frmFee.permitid.value, 'UpdateTotalFees', 'get', '0');
		}

		function UpdateTotalFees( sReturn )
		{
			if ( sReturn != "Failed" )
			{
				parent.document.getElementById("feetabfeetotal").innerHTML = sReturn;
				parent.document.getElementById("invoicetabfeetotal").innerHTML = sReturn;
			}
			//doClose();
			// Try to repost this page
			location.href = "fixturefeeedit.asp?permitfeeid=" + <%=iPermitFeeId%> + "&success=Changes%20Saved";
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function addFixture()
		{	
			// build the parameter list
			var sParameter = 'permitid=' + encodeURIComponent(document.frmFee.permitid.value);
			sParameter += '&permitfeeid=' + encodeURIComponent(document.frmFee.permitfeeid.value);
			sParameter += '&permitfixturetypeid=' + encodeURIComponent(document.frmFee.permitfixturetypeid.value);
			//alert( sParameter );
			doAjax('fixturetypeadd.asp', sParameter, 'AddFixtureRow', 'post', '0');
			//document.frmFee.submit();
		}

		function AddFixtureRow( sReturn )
		{
			if (sReturn == 'SUCCESS')
			{
				var tbl = document.getElementById("fixturelist");
				var lastRow = tbl.rows.length;
				var row = tbl.insertRow(lastRow);
				// Remove cell
				var cell = row.insertCell(0);
				cell.innerHTML = '&nbsp;';
				// include cell
				cell = row.insertCell(1);
				cell.innerHTML = '&nbsp;';
				// name cell
				cell = row.insertCell(2);
				cell.innerHTML = document.frmFee.permitfixturetypeid.options[document.frmFee.permitfixturetypeid.selectedIndex].text;
				// Qty cell
				cell = row.insertCell(3);
				cell.innerHTML = '&nbsp;';
				// fee cell
				cell = row.insertCell(4);
				cell.innerHTML = '&nbsp;';
			}
		}

		function RemoveFixtures()
		{
			if (confirm("Remove the selected fixtures?"))
			{
				//alert('Removing the fixtures.');
				var iRow = 1;
				var RowsLeft = 0;
				var tbl = document.getElementById("fixturelist");
				// Check the feelist rows for any selected for removal
				for (var t = 0; t <= parseInt(document.frmFee.maxFeeCount.value); t++)
				{
					// See if a row exists for this one
					if (document.getElementById("remove" + t))
					{
						// If it is marked for removal, remove it
						if (document.getElementById("remove" + t).checked == true)
						{
							//alert(document.getElementById("permitfeeid" + t).value);
							doAjax('removepermitfeefixture.asp', 'permitfixtureid=' + document.getElementById("permitfixtureid" + t).value, '', 'get', '0');
							//alert(iRow);
							tbl.deleteRow(iRow);
							//iRow--;
						}
						else
						{
							iRow++;
							RowsLeft++;
						}
					}
				}
				document.frmFee.maxFeeCount.value = RowsLeft;
			}
		}

		function init()
		{
			if (document.getElementById("qty1"))
			{
				document.getElementById("qty1").focus();
			}
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

		//-->
		</script>

	</head>
	<body onload="init();">
		<div id="content">
			<div id="centercontent">
				<form name="frmFee" action="fixturefeeupdate.asp" method="post">
				<input type="hidden" name="permitfeeid" value="<%=iPermitFeeId%>" />
				<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
						<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML='<%=sFeeName%>';</script>
					</p>
					<br />
					<br />
<%					
					tooltipclass=""
					tooltip = ""
					disabled = ""
					If bPermitIsCompleted or bIsOnHold Then
						tooltipclass="tooltip"
						disabled = " disabled "
						tooltip = "<span class=""tooltiptext"">You cannot edit fixtures because:<br />The permit is complete or on hold.</span>"
					end if
					%>
					<p> <% ShowAllFixtures %> &nbsp; &nbsp;
						<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="addFixture();">Add Selected Fixture<%=tooltip%></button> &nbsp; &nbsp;
						<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="RemoveFixtures();">Remove Selected Fixtures<%=tooltip%></button>
					</p>
					<p> 
						<% iMaxFeeCount = ShowFixtureFees( iPermitFeeId ) %>
						<input type="hidden" name="maxFeeCount" value="<%=iMaxFeeCount%>" />
					</p>
					<p>
						<strong>Total of Fees Applied: <%=sAppliedFee%></strong>
					</p>
					<p>
						<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" id="savebutton" onclick="doUpdate();">Save Changes<%=tooltip%></button> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
					</p>
				</form>
			</div>
		</div>

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function ShowFixtureFees( iPermitFeeId )
'--------------------------------------------------------------------------------------------------
Function ShowFixtureFees( iPermitFeeId )
	Dim sSql, oRs, iFeeCount, iQty

	iFeeCount = CLng(0)

	sSql = "SELECT permitfixtureid, permitfixture, ISNULL(qty,0) AS qty, ISNULL(feeamount,0.00) AS feeamount, isincluded "
	sSQl = sSql & " FROM egov_permitfixtures WHERE permitfeeid = " & iPermitFeeId & " ORDER BY displayorder, permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<table cellpadding=""2"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""fixturelist"">"
	response.write vbcrlf & "<tr><th>Remove</th><th>Include</th><th>Fixture</th><th>Qty</th><th>Fee</th></tr>"

	Do While Not oRs.EOF
		iFeeCount = iFeeCount + 1
		response.write vbcrlf & "<tr" 
		If iFeeCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
		response.write "<td align=""center""><input type=""checkbox"" id=""remove" & iFeeCount & """ name=""remove" & iFeeCount & """ />"
		response.write "<td align=""center""><input type=""checkbox"" id=""include" & iFeeCount & """ name=""include" & iFeeCount & """"
		If oRs("isincluded") Then
			response.write " checked=""checked"" "
		End If 
		response.write " />"
		response.write "<input type=""hidden"" name=""permitfixtureid" & iFeeCount & """ id=""permitfixtureid" & iFeeCount & """ value=""" & oRs("permitfixtureid") & """ /></td>"
		response.write "<td title=""click to view step table"" onclick=""ViewStepTable('" & oRs("permitfixtureid") & "');"" >" & oRs("permitfixture") & "</td>"
		If CLng(oRs("qty")) > CLng(0) Then 
			iQty = CLng(oRs("qty"))
		Else 
			iQty = ""
		End If 
		response.write "<td align=""center""><input type=""text"" tabindex=""" & iFeeCount & """ name=""qty" & iFeeCount & """ id=""qty" & iFeeCount & """ value=""" & iQty & """ size=""10"" maxlength=""10"" onchange=""checkInclude( this, " & iFeeCount & " );"" /></td>"
		response.write "<td align=""center""><span id=""feeamount" & iFeeCount & """>" & FormatNumber(oRs("feeamount"),2,,,0) & "</span></td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</table>"
	
	oRs.Close
	Set oRs = Nothing 

	ShowFixtureFees = iFeeCount

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowAllFixtures( )
'--------------------------------------------------------------------------------------------------
Sub ShowAllFixtures( )
	Dim sSql, oRs

	sSql = "SELECT permitfixturetypeid, permitfixture FROM egov_permitfixturetypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""permitfixturetypeid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitfixturetypeid") & """>" & oRs("permitfixture") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
