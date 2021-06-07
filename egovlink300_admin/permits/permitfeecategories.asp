<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfeecategories.asp
' AUTHOR: Steve Loar
' CREATED: 12/14/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/14/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMin, iMax

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permit fee categories", sLevel	' In common.asp

iMin = CLng(99999)
iMax = CLng(0)

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="Javascript">
	<!--
		function validate()
		{
			for (var p = parseInt(document.frmCategory.min.value); p <= parseInt(document.frmCategory.max.value); p++)
			{
				// see if the label exists
				if (document.getElementById("label" + p))
				{
					// check for apply being checked
					if (document.getElementById("apply" + p).checked)
					{
						// check the label for being blank
						if (document.getElementById("label" + p).value == '')
						{
							alert("Please provide a label for the surcharge.");
							document.getElementById("label" + p).focus();
							return;
						}
					}
				}

				// see if the rate exists
				if (document.getElementById("rate" + p))
				{
					// check it for being blank
					if (document.getElementById("rate" + p).value == '')
					{
						if (document.getElementById("apply" + p).checked)
						{
							// if apply is checked, then we need a rate
							alert("All surcharge rates need a numberic value in the format '#.##'.");
							document.getElementById("rate" + p).focus();
							return;
						}
					}
					else
					{
						// validate the rate's format, we need a format even if the apply is not checked
						rege = /^\d{0,1}\.{0,1}\d{0,2}$/;
						Ok = rege.test(document.getElementById("rate" + p).value);
						if (! Ok)
						{
							alert("All surcharge rates need a numberic value in the format '#.##'.");
							document.getElementById("rate" + p).focus();
							return;
						}
					}
				}
			}
			//alert("success.");
			document.frmCategory.submit();
		}
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Fee Categories</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<div>
				<input type="button" name="update" class="button ui-button ui-widget ui-corner-all" value="Update" onclick="validate();" />&nbsp;&nbsp;&nbsp;
				<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" /><br /><br />
			</div>

			<form name="frmCategory" action="permitfeecategoryupdate.asp" method="post">
			<div class="shadow">
				<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
					<tr><th>Fee Category</th><th>For<br />Building<br />Permits</th><th>For<br />Commercial<br />Permits</th><th>Surcharge<br />Label</th><th>Surcharge<br />Rate</th><th>Apply<br />Surcharge</th></tr>
<%					
					ShowPermitFeeCategories
%>
				</table>
			</div>
				<input type="hidden" name="min" value="<%=iMin%>" />
				<input type="hidden" name="max" value="<%=iMax%>" />
			</form>

			<div>
				<input type="button" name="update" class="button ui-button ui-widget ui-corner-all" value="Update" onclick="validate();" />&nbsp;&nbsp;&nbsp;
				<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" /><br /><br />
			</div>
		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowPermitFeeCategories()
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFeeCategories()
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT permitfeecategorytypeid, permitfeecategory, applysurcharge, surchargelabel, surchargerate, isbuildingfee, iscommercial FROM egov_permitfeecategorytypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			If iMin > CLng(oRs("permitfeecategorytypeid")) Then
				iMin = CLng(oRs("permitfeecategorytypeid"))
			End If 
			If iMax < CLng(oRs("permitfeecategorytypeid")) Then
				iMax = CLng(oRs("permitfeecategorytypeid"))
			End If 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td>&nbsp;" & oRs("permitfeecategory") & " <input type=""hidden"" name=""permitfeecategorytypeid" & oRs("permitfeecategorytypeid") & """ value=""" & oRs("permitfeecategorytypeid") & """ /></td>"
			If oRs("isbuildingfee") Then
				response.write "<td align=""center"">Yes</td>"
			Else
				response.write "<td align=""center"">No</td>"
			End If 
			If oRs("iscommercial") Then
				response.write "<td align=""center"">Yes</td>"
			Else
				response.write "<td align=""center"">No</td>"
			End If 
			response.write "<td align=""center""><input type=""text"" value=""" & oRs("surchargelabel") & """ name=""label" & oRs("permitfeecategorytypeid") & """ id=""label" & oRs("permitfeecategorytypeid") & """ size=""30"" maxlength=""50"" /></td>"
			response.write "<td align=""center""><input type=""text"" value=""" & oRs("surchargerate") & """ name=""rate" & oRs("permitfeecategorytypeid") & """ id=""rate" & oRs("permitfeecategorytypeid") & """ size=""4"" maxlength=""4"" /></td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""apply" & oRs("permitfeecategorytypeid") & """ name=""apply" & oRs("permitfeecategorytypeid") & """ "
			If oRs("applysurcharge") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
	Else
		response.write "<tr><td colspan=""5"">&nbsp;No data exists for this feature. Contact EC Link for assistance.</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


%>
