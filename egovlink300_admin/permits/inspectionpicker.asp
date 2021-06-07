<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: inspectionpicker.asp
' AUTHOR: Steve Loar
' CREATED: 07/09/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects inspections to add to a permit
'
' MODIFICATION HISTORY
' 1.0   07/09/2008	Steve Loar - INITIAL VERSION
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
			if (document.frmInspection.permitinspectiontypeid.options[document.frmInspection.permitinspectiontypeid.selectedIndex].value > 0)
			{
				//document.frmInspection.submit();
				// Save the new inspection pick via Ajax
				doAjax('inspectionpickeradd.asp', 'permitid=' + document.frmInspection.permitid.value + '&permitinspectiontypeid=' + document.frmInspection.permitinspectiontypeid.options[document.frmInspection.permitinspectiontypeid.selectedIndex].value, 'doNewRow', 'get', '0');
			}
			else
			{
				alert("Closing");
				doClose();
			}
		}

		function doNewRow( sPermitFeeId )
		{
			if (document.frmInspection.permitinspectiontypeid.options[document.frmInspection.permitinspectiontypeid.selectedIndex].value > 0)
			{
				parent.document.getElementById("maxinspections").value = parseInt(parent.document.getElementById("maxinspections").value) + 1;
				var tbl = parent.document.getElementById("inspectionlist");
				var lastRow = tbl.rows.length;
				var newRow = parseInt(parent.document.getElementById("maxinspections").value);
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

				// Inspection cell
				cell = row.insertCell(1);
				cell.title = "Click Save Changes to Complete this add";
				cell.align = 'left';
				var index = document.frmInspection.permitinspectiontypeid.selectedIndex;
				var evalue = document.frmInspection.permitinspectiontypeid.value;
				cell.innerHTML = document.frmInspection.permitinspectiontypeid.options[index].text + ' &mdash; ' + document.getElementById("ins" + evalue).innerHTML.replace("<b>Description:</b><br>","") ;
				cell.onmouseover = function() {this.style.backgroundColor = '#93bee1';this.style.cursor='pointer';};
				cell.onmouseout = function() {this.style.backgroundColor = '';this.style.cursor='';};

				// Reinspection cell
				cell = row.insertCell(2);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Status cell
				cell = row.insertCell(3);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Scheduled Date cell
				cell = row.insertCell(4);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Inspected Date cell
				cell = row.insertCell(5);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

				// Inspector cell
				cell = row.insertCell(6);
				cell.title = "Click Save Changes to Complete this add";
				cell.innerHTML = '&nbsp;';

			}
			doClose();
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function showDesc(e)
		{
			//alert("HERE");
			//alert(e.value);
			var x = document.getElementsByClassName("insdesc");
			var i;
			for (i = 0; i < x.length; i++) {
    				x[i].style.display = "none";
			}
			document.getElementById("ins" + e.value).style.display = "";
		}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent" style="width:500px;">
				<form name="frmInspection" action="inspectionpickeradd.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
					<%	ShowPermitInspectionTypes %>
					</p>
					<p>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Add Selected Inspection" onclick="doSelect();" /> &nbsp; &nbsp; 
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' Sub ShowPermitInspectionTypes()
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectionTypes()
	Dim sSql, oRs

	sSql = "SELECT permitinspectiontypeid, permitinspectiontype, inspectiondescription "
	sSql = sSql & " FROM egov_permitinspectiontypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permitinspectiontype, inspectiondescription"
	'  AND isbuildingpermittype = 1

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""permitinspectiontypeid"" onChange=""showDesc(this);"">"
		response.write vbcrlf & "<option value=""0"">Select an inspection to add to this permit</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitinspectiontypeid") & """>" 
			response.write oRs("permitinspectiontype") '& " &mdash; " & oRs("inspectiondescription")
			response.write "</option>"
			oRs.MoveNext
		Loop 
		oRs.MoveFirst
		response.write vbcrlf & "</select>"
	End If 
	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write "<div style=""display:none;"" class=""insdesc"" id=""ins" & oRs("permitinspectiontypeid") & """><b>Description:</b><br />" & oRs("inspectiondescription") & "</div>"
			oRs.MoveNext
		Loop 
	End If 


	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
