<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: displayedit.asp
' AUTHOR: Steve Loar
' CREATED: 05/12/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of displays. From here you can create or edit displays. 
'				This is not the org setup list
'
' MODIFICATION HISTORY
' 1.0   04/30/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDisplayId, sButtonValue, sDisplayName, sDisplay, sDisplayDescription, iFeatureId, sisOnPublicSide
Dim sisOnAdminSide, sAdminCanEdit, sUsesDisplayName

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then 
	response.redirect "../default.asp"
End If 

If request("displayid") <> "" Then
	iDisplayId = CLng(request("displayid"))
Else
	response.redirect "displaylist.asp"
End If 

If iDisplayId > CLng(0) Then
	sButtonValue = "Save Changes"
	GetDisplayValues iDisplayId
Else
	sButtonValue = "Create Display"
	sDisplay = ""
	sDisplayName = ""
	sDisplayDescription = ""
	iFeatureId = 0
	sisOnPublicSide = ""
	sisOnAdminSide = ""
	sAdminCanEdit = ""
	sUsesDisplayName = ""
End If 

If request("msg") <> "" Then
	If request("msg") = "n" Then
		sLoadMsg = "displayScreenMsg('This Display Was Successfully Created');"
	End If
	If request("msg") = "u" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If 
End If 


%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="Javascript">
	<!--

		function SetUpPage()
		{
			setMaxLength();
			<%=sLoadMsg%>
			$("display").focus();
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

		function validate() 
		{
			if ($F("display") == '')
			{
				$("display").focus();
				inlineMsg(document.getElementById("display").id,'<strong>Missing Display </strong>Please enter a display.',5,document.getElementById("display").id);
				return;
			}

			if ($F("displayname") == '')
			{
				$("displayname").focus();
				inlineMsg(document.getElementById("displayname").id,'<strong>Missing Display Name </strong>Please enter a display name.',5,document.getElementById("displayname").id);
				return;
			}

			if ($F("displaydescription") == '')
			{
				$("displaydescription").focus();
				inlineMsg(document.getElementById("displaydescription").id,'<strong>Missing Display Description</strong>Please enter a display description.',5,document.getElementById("displaydescription").id);
				return;
			}

			//alert("All is OK");
			document.frmDisplay.submit();
		}

		function DeleteDisplay()
		{
			if (confirm('Delete this display?'))
			{
				location.href='displaydelete.asp?displayid=<%=iDisplayId%>';
			}
		}

	//-->
	</script>

</head>

<body onload="SetUpPage();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Edit Display</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
				<input type="button" class="button" value="<< Back" onclick="location.href='displaylist.asp';" /> &nbsp;
				<input type="button" class="button" value="Delete This Display" onclick="DeleteDisplay();" />
			</td></tr></table>

			<form name="frmDisplay" action="displayeditupdate.asp" method="post">
				<input type="hidden" id="displayid" name="displayid" value="<%=iDisplayId%>" />

				<p class="displaynamecontainer">
					<strong>Display: </strong> &nbsp; <input type="text" id="display" name="display" size="45" maxlength="45" value="<%=sDisplay%>" /><br />
					<span class="subnote">* This is what the code looks for</span>
				</p>

				<p class="displaynamecontainer">
					<strong>Display Name: </strong> &nbsp; <input type="text" id="displayname" name="displayname" size="90" maxlength="90" value="<%=sDisplayName%>" /><br />
					<span class="subnote">* This is what those doing the set up look for and/or the default value displayed.</span>
				</p>

				<p class="displaynamecontainer">
					<input type="checkbox" id="usesdisplayname" name="usesdisplayname" <%=sUsesDisplayName%> /> <strong>This display uses the Display Name as the default value of the display.</strong>
				</p>

				<p class="displaynamecontainer">
					<strong>Category: </strong> &nbsp; <% ShowDisplayCategoryPicks iFeatureId %>
				</p>

				<p class="displaynamecontainer">
					<strong>This is seen on...<br />
					<input type="checkbox" name="isonpublicside" id="isonpublicside" <%=sisOnPublicSide%> /> The Public Side
					<span id="isonadmincheck"><input type="checkbox" name="isonadminside" id="isonadminside" <%=sisOnAdminSide%> /> The Admin Side</span></strong>
				</p>

				<p class="displaynamecontainer">
					<input type="checkbox" id="admincanedit" name="admincanedit" <%=sAdminCanEdit%> /> <strong>Allow Admins to enter the value displayed for an organization on the setup page.</strong>
				</p>

				<p class="displaynamecontainer">
					<strong>Display Description:</strong><br />
					<span class="subnote">* Tell the person doing the set up something about this display.</span><br />
					<textarea id="displaydescription" name="displaydescription" maxlength="2000" wrap="soft"><%=sDisplayDescription%></textarea>
				</p>

				<p class="displaynamecontainer">
					<input type="button" class="button" id="savebutton" value="<%=sButtonValue%>" onclick="validate();" />
				</p>

			</form>
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
' void GetDisplayValues iDisplayId
'--------------------------------------------------------------------------------------------------
Sub GetDisplayValues( ByVal iDisplayId )
	Dim sSql, oRs

	sSql = "SELECT display, displayname, displaydescription, admincanedit, usesdisplayname, "
	sSql = sSql & "isonpublicside, isonadminside, ISNULL(featureid,0) AS featureid, usesdisplayname "
	sSql = sSql & "FROM egov_organization_displays WHERE displayid = " & iDisplayId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sDisplay = oRs("display")
		sDisplayName = Replace(oRs("displayname"), Chr(34), "&quot;")
		sDisplayDescription = oRs("displaydescription")
		iFeatureId = oRs("featureid")
		If oRs("isonpublicside") Then
			sisOnPublicSide = " checked=""checked"" "
		Else
			sisOnPublicSide = ""
		End If
		If oRs("isonadminside") Then
			sisOnAdminSide = " checked=""checked"" "
		Else
			sisOnAdminSide = ""
		End If
		If oRs("admincanedit") Then
			sAdminCanEdit = " checked=""checked"" "
		Else
			sAdminCanEdit = ""
		End If 
		If oRs("usesdisplayname") Then
			sUsesDisplayName = " checked=""checked"" "
		Else
			sUsesDisplayName = ""
		End If 
	Else 
		sDisplay = ""
		sDisplayName = ""
		sDisplayDescription = ""
		iFeatureId = 0
		sisOnPublicSide = ""
		sisOnAdminSide = ""
		sAdminCanEdit = ""
		sUsesDisplayName = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowDisplayCategoryPicks iFeatureId
'--------------------------------------------------------------------------------------------------
Sub ShowDisplayCategoryPicks( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT featureid, featurename FROM egov_organization_features "
	sSql = sSql & "WHERE parentfeatureid = 0 AND featuretype = 'N' "
	sSql = sSql & "AND LOWER(featurename) NOT IN ('log off','e-gov administration') "
	sSql = sSql & "ORDER BY admindisplayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""featureid"">"
	response.write vbcrlf & "<option value=""0"">No Category Assigned</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("featureid") & """ "
		If clng(oRs("featureid")) = clng(iFeatureId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("featurename") & "</option>"
		oRs.MoveNext 
	Loop

	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
