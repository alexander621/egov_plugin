<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: clientdisplayedit.asp
' AUTHOR: Steve Loar
' CREATED: 05/19/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of displays. From here you can create or edit displays. 
'				This is not the org setup list
'
' MODIFICATION HISTORY
' 1.0   05/19/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDisplayId, sButtonValue, sDisplayName, sDisplay, sDisplayDescription, iFeatureId, sisOnPublicSide
Dim sisOnAdminSide, sAdminCanEdit, sUsesDisplayName, sFeatureName, sClientDisplayDescription
Dim sClientDisplayName

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
	' There are no 0 displayid so go back
	response.redirect "clientdisplaylist.asp"
End If 

If request("msg") <> "" Then
	If request("msg") = "c" Then
		sLoadMsg = "displayScreenMsg('This Display Was Successfully Cleared');"
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
<%				If sUsesDisplayname = "1" Then		%>
			if ($F("clientdisplayname") == '')
			{
				$("clientdisplayname").focus();
				inlineMsg(document.getElementById("clientdisplayname").id,'<strong>Missing Client Display Value </strong>Please enter a client display value.',5,document.getElementById("clientdisplayname").id);
				return;
			}
<%				Else				%>
			if ($F("clientdisplaydescription") == '')
			{
				$("clientdisplaydescription").focus();
				inlineMsg(document.getElementById("clientdisplaydescription").id,'<strong>Missing Client Display Value </strong>Please enter a client display value.',5,document.getElementById("clientdisplaydescription").id);
				return;
			}
<%				End If				%>
			//alert("All is OK");
			document.frmDisplay.submit();
		}

		function ClearDisplay()
		{
			if (confirm('Clear this display?'))
			{
				location.href='clientdisplayclear.asp?displayid=<%=iDisplayId%>';
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
				<font size="+1"><strong>Edit Client Display</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
				<input type="button" class="button" value="<< Back" onclick="location.href='clientdisplaylist.asp';" /> &nbsp;
				<input type="button" class="button" value="Clear This Display" onclick="ClearDisplay();" />
			</td></tr></table>

			<form name="frmDisplay" action="clientdisplayeditupdate.asp" method="post">
				<input type="hidden" id="displayid" name="displayid" value="<%=iDisplayId%>" />
				<input type="hidden" id="usesdisplayname" name="usesdisplayname" value="<%=sUsesDisplayname%>" />

				<p class="displaynamecontainer">
					<strong>Display Name: </strong> &nbsp; <%=sDisplayName%><br />
				</p>

				<p class="displaynamecontainer">
					<strong>Category: </strong> &nbsp; <%=sFeatureName%>
				</p>

				<p class="displaynamecontainer">
					<strong>This is seen on...<br />
					<input type="checkbox" name="isonpublicside" id="isonpublicside" <%=sisOnPublicSide%> readonly="readonly" /> The Public Side
					<span id="isonadmincheck"><input type="checkbox" name="isonadminside" id="isonadminside" <%=sisOnAdminSide%> readonly="readonly" /> The Admin Side</span></strong>
				</p>

				<p class="displaynamecontainer">
					<strong>Display Description:</strong> &nbsp; <%=sDisplayDescription%>
				</p>

				<p>
					<strong>Client Value for this Display:</strong><br />
<%				If sUsesDisplayname = "0" Then		%>
					<textarea id="clientdisplaydescription" name="clientdisplaydescription" maxlength="4000" wrap="soft"><%=sClientDisplayDescription%></textarea>
					<input type="hidden" name="clientdisplayname" value="" />
<%				Else				%>
					<input type="text" id=clientdisplayname name="clientdisplayname" size="100" maxlength="100" value="<%=sClientDisplayName%>" /><br />
					* The Display Name value is the default for this display.
					<input type="hidden" name="clientdisplaydescription" value="" />
<%				End If				%>

				<p class="displaynamecontainer">
					<input type="button" class="button" id="savebutton" value="Save Changes" onclick="validate();" />
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

	sSql = "SELECT D.displayid, D.displayname, ISNULL(D.displaydescription,'') AS displaydescription, "
	sSql = sSql & "ISNULL(F.featurename,'') AS featurename, D.usesdisplayname, "
	sSql = sSql & "ISNULL(O.displayname,'') AS clientdisplayname, ISNULL(O.displaydescription,'') AS clientdisplaydescription, "
	sSql = sSql & "D.isonpublicside, D.isonadminside "
	sSql = sSql & "FROM egov_organization_displays D "
	sSql = sSql & "LEFT OUTER JOIN egov_organization_features F ON D.featureid = F.featureid "
	sSql = sSql & "LEFT OUTER JOIN egov_organizations_to_displays O ON D.displayid = O.displayid AND O.orgid = " & session("orgid")
	sSql = sSql & " WHERE D.admincanedit = 1 AND D.displayid = " & iDisplayId
'	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sDisplayName = Replace(oRs("displayname"), "<", "&lt;")

		sDisplayDescription = oRs("displaydescription")

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

		If oRs("usesdisplayname") Then
			sUsesDisplayname = "1"
		Else
			sUsesDisplayname = "0"
		End If 

		sFeatureName = oRs("featurename")
		sClientDisplayDescription = oRs("clientdisplaydescription")
		sClientDisplayName = oRs("clientdisplayname")
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 




%>