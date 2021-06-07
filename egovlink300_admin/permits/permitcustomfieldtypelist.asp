<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcustomfieldtypelist.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit review types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sSuccessFlag, sLoadMsg

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit custom field types", sLevel	' In common.asp
PageDisplayCheck "permitv2 types", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

sSuccessFlag = request("success")
If sSuccessFlag = "cftr" Then
	sLoadMsg = "displayScreenMsg('The custom field type has been removed.');"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/jquery-1.4.2.min.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--
		function refreshPage() 
		{
			document.frmReviewSearch.searchtext.value = '';
			document.frmReviewSearch.submit();
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("&nbsp;");
		}

		$(document).ready(function(){
			<%=sLoadMsg%>
		});

	//-->
	</script>
	<script>
		function commonIFrameUpdateFunction()
		{
			UpdateCustomFieldTypes();
		}
		function UpdateCustomFieldTypes()
		{
			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type=customfieldtypes', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Get elements from parent
				var pfDD = parent.document.getElementsByClassName('permitcustomfieldtypeDD');
				for (var i = 0; i < pfDD.length; i++) {
					//Get Selected Value
  					//pfDD[i].style.display = 'inline-block';
					var selVal = pfDD[i].options[pfDD[i].selectedIndex].value;
					
					//Update The Values
					pfDD[i].innerHTML = newDDVals;
	
					//Select Previous Option
					pfDD[i].value = selVal;
				}
			}

		}
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
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
				<font size="+1"><strong>Permit Custom Field Types</strong></font><br />
			</p>
			
			<p>
				<span id="screenMsg">&nbsp;</span>
				&nbsp;
			</p>

			<!--END: PAGE TITLE-->
			<form name="frmReviewSearch" method="post" action="permitcustomfieldtypelist.asp">
				<div id="functionlinks">
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
<%					If sSearch <> "" Then	%>
						&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="Show All" onclick="refreshPage();" />
<%					End If		%>
					&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Custom Field Type" onclick="location.href='permitcustomfieldtypeedit.asp?cft=0';" />
					&nbsp; &nbsp; <input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
				</div>
			</form>

			<div class="shadow">
			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Custom Field Types</th><th>PDF Field Name</th></tr>
				<%	
					ShowPermitCustomFieldTypes session("orgid"), sSearch 
				%>
			</table>
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
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitCustomFieldTypes( iOrgid, sSearch )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitCustomFieldTypes( ByVal iOrgid, ByVal sSearch )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT customfieldtypeid, fieldname, ISNULL(pdffieldname,'') AS pdffieldname "
	sSql = sSql & " FROM egov_permitcustomfieldtypes "
	sSql = sSql & " WHERE isactive = 1 AND orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND fieldname LIKE '%" & dbsafe(sSearch) & "%' "
	End If 
	sSql = sSql & " ORDER BY fieldname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			response.write "<td class=""leftcol"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='permitcustomfieldtypeedit.asp?cft=" & oRs("customfieldtypeid") & "';"">&nbsp;" 
			response.write oRs("fieldname") & "</td>"

			response.write "<td align=""center"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='permitcustomfieldtypeedit.asp?cft=" & oRs("customfieldtypeid") & "';"">&nbsp;" 
			response.write oRs("pdffieldname") & "</td>"

			response.write "</tr>"

			oRs.MoveNext 
		Loop 
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td>&nbsp;No Custom Field Types could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td>&nbsp;No Custom Field Types could be found. Click on the New Custom Field Type button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
