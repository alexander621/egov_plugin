<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectiontypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 01/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates and edits permit inspection types
'
' MODIFICATION HISTORY
' 1.0   01/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iPermitInspectionTypeid, sPermitInspectionType, sInspectionDescription, sIsBuildingPermitType

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "permit inspection types", sLevel	' In common.asp
PageDisplayCheck "permit types", sLevel	' In common.asp

iPermitInspectionTypeid = CLng(request("permitinspectiontypeid") )

If CLng(iPermitInspectionTypeid) > CLng(0) Then
	sTitle = "Edit"
	GetPermitInspectionType iPermitInspectionTypeid
Else
	sTitle = "New"
	sPermitInspectionType = ""
	sInspectionDescription = ""
'	sIsBuildingPermitType = " checked=""checked"" "
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

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function Another()
		{
			location.href="permitinspectiontypeedit.asp?permitinspectiontypeid=0";
		}

		function Validate()
		{
			var rege;
			var Ok; 

			// Check for a fixture name
			if (document.frmInspection.inspectiondescription.value == '')
			{
				alert("Please provide an inspection description, then try saving again.");
				document.frmInspection.inspectiondescription.focus();
				return;
			}
			
			//alert("All was OK");
			// All is OK so submit
			document.frmInspection.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this permit inspection type?"))
			{
				location.href="permitinspectiontypedelete.asp?permitinspectiontypeid=<%=iPermitInspectionTypeid%>";
			}
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>
	<script>
		function commonIFrameUpdateFunction()
		{
			UpdateInspectionTypes();
		}
		function UpdateInspectionTypes()
		{
			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type=inspectiontypes', false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Get elements from parent
				var pfDD = parent.document.getElementsByClassName('permitinspectiontypeDD');
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
				<font size="+1"><strong><%=sTitle%> Permit Inspection Type</strong></font><br /><br />
				<a href="permitinspectiontypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: EDIT FORM-->
		<div id="functionlinks">
<%		If CLng(iPermitInspectionTypeid) = CLng(0) Then %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />&nbsp; &nbsp; 
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
<%			If request("success") <> "" Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Another();" value="Create Another" />
<%			End If		%>
			<br />
<%		End If %>
		</div>

		<form name="frmInspection" action="permitinspectiontypeupdate.asp" method="post">
		<input type="hidden" name="permitinspectiontypeid" value="<%=iPermitInspectionTypeid%>" />
		
		<p>
			Inspection Type: &nbsp;&nbsp; <input type="text" id="permitinspectiontype" name="permitinspectiontype" value="<%=sPermitInspectionType%>" size="50" maxlength="50" />
		</p>
		<p>
			Description: &nbsp;&nbsp; 
				<!--input type="text" id="inspectiondescription" name="inspectiondescription" value="<%=sInspectionDescription%>" size="100" maxlength="150" /-->
				<textarea id="inspectiondescription" name="inspectiondescription" style="width:600px;height:50px;vertical-align:top;"><%=sInspectionDescription%></textarea>
		</p>
<!--
		<p>
			<input type="checkbox" id="isbuildingpermittype" name="isbuildingpermittype" <%'=sIsBuildingPermitType%> /> For Building Permits
		</p>
-->
		</form>
		<!--END: EDIT FORM-->

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void GetPermitInspectionType( iPermitInspectionTypeid )
'--------------------------------------------------------------------------------------------------
Sub GetPermitInspectionType( ByVal iPermitInspectionTypeid )
	Dim sSql, oRs

	sSql = "SELECT permitinspectiontypeid, ISNULL(permitinspectiontype,'') AS permitinspectiontype, "
	sSql = sSql & " inspectiondescription "
	sSql = sSql & " FROM egov_permitinspectiontypes WHERE permitinspectiontypeid = " & iPermitInspectionTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitInspectionType = Replace(oRs("permitinspectiontype"),"""","&quot;")
		sInspectionDescription = Replace(oRs("inspectiondescription"),"""","&quot;")
'		If oRs("isbuildingpermittype") Then 
'			sIsBuildingPermitType = " checked=""checked"" "
'		Else
'			sIsBuildingPermitType = ""
'		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
