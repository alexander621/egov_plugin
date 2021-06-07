<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reviewpicker.asp
' AUTHOR: Steve Loar
' CREATED: 06/18/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects reviews to add to a permit
'
' MODIFICATION HISTORY
' 1.0   06/18/2008	Steve Loar - INITIAL VERSION
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

		<script language="Javascript">
		<!--

		function doSelect()
		{
			if (document.frmReview.permitreviewtypeid.options[document.frmReview.permitreviewtypeid.selectedIndex].value > 0)
			{
				//document.frmFee.submit();
				// Save the new fee via Ajax
				doAjax('reviewpickeradd.asp', 'permitid=' + document.frmReview.permitid.value + '&permitreviewtypeid=' + document.frmReview.permitreviewtypeid.options[document.frmReview.permitreviewtypeid.selectedIndex].value, 'doNewRow', 'get', '0');
			}
			else
			{
				alert("Closing");
				doClose();
			}
		}

		function doNewRow( sPermitFeeId )
		{
			if (document.frmReview.permitreviewtypeid.options[document.frmReview.permitreviewtypeid.selectedIndex].value > 0)
			{
				parent.document.getElementById("maxreviews").value = parseInt(parent.document.getElementById("maxreviews").value) + 1;
				var tbl = parent.document.getElementById("reviewlist");
				var lastRow = tbl.rows.length;
				var newRow = parseInt(parent.document.getElementById("maxreviews").value);
				var row = tbl.insertRow(lastRow);
				if (newRow % 2 == 0)
				{
					row.className = "altrow";
				}
				row.onmouseover = function() {this.style.backgroundColor = '#93bee1';this.style.cursor='pointer';};
				row.onmouseout = function() {this.style.backgroundColor = '';this.style.cursor='';};
				
				// Add the Remove cell
				var cell = row.insertCell(0);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Review cell
				cell = row.insertCell(1);
				cell.title = "Click Save Changes to Complete this add";
				cell.align = 'center';
				cell.innerHTML = document.frmReview.permitreviewtypeid.options[document.frmReview.permitreviewtypeid.selectedIndex].text;
				cell.onmouseover = function() {this.style.backgroundColor = '#93bee1';this.style.cursor='pointer';};
				cell.onmouseout = function() {this.style.backgroundColor = '';this.style.cursor='';};

				// Status cell
				cell = row.insertCell(2);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Reviewed cell
				cell = row.insertCell(3);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Reviewer cell
				cell = row.insertCell(4);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

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
				<font size="+1"><strong>Add A Review</strong></font><br /><br />
				<form name="frmReview" action="reviewpickeradd.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
					<%	ShowPermitReviewTypes %>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Selected Review" onclick="doSelect();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
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
' void ShowPermitReviewTypes
'--------------------------------------------------------------------------------------------------
Sub ShowPermitReviewTypes()
	Dim sSql, oRs

	iFeeCount = 0

	sSql = "SELECT permitreviewtypeid, permitreviewtype FROM egov_permitreviewtypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitreviewtype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""permitreviewtypeid"">"
		response.write vbcrlf & "<option value=""0"">Select a review to add to this permit</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitreviewtypeid") & """>" 
			response.write oRs("permitreviewtype")
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
