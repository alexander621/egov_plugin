<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: instructor_profile.asp
' AUTHOR: Steve Loar
' CREATED: 05/10/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the page on which instructor information is added or edited by instructors.
'
' MODIFICATION HISTORY
' 1.0   05/10/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' INITIALIZE VARIABLES
Dim sFirstName, sMiddle, sLastName, sEmail, sInstrPhone, sMobilePhone, sImageAlt, iUserId
Dim sWebsiteURL, sImageURL,	sBio, iInstructorID, isemailpublic, iphonepublic, iscellpublic 
Dim sSql, oInstr

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "instructor profile" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If

' GET INSTRUCTOR ID
iInstructorID = GetUserInstructorId( Session("UserId") )

If clng(iInstructorID) = clng(0) Then
	' CREATE NEW INSTRUCTOR
	sTitle = "New Instructor Profile " 
	sLinkText = "Create Instructor Profile"
Else
	' EDIT EXISTING INSTRUCTOR
	sTitle = "Instructor Profile " 
	sLinkText = "Save Instructor Profile"
End If

sSQL = "SELECT firstname, middle, lastname, imgurl, email, phone, cellphone, isemailpublic, isphonepublic,"
sSql = sSql & " imgalt, iscellpublic, websiteurl, bio, isnull(userid,0) as userid "
sSql = sSql & " FROM egov_class_instructor WHERE instructorid = " & iInstructorID 

Set oInstr = Server.CreateObject("ADODB.Recordset")
oInstr.Open sSQL, Application("DSN"), 0, 1

If NOT oInstr.EOF Then
	oInstr.movefirst 
	sFirstName = oInstr("firstname")
	sMiddle = oInstr("middle")
	sLastName = oInstr("lastname")
	sEmail = oInstr("email")
	isemailpublic = oInstr("isemailpublic")
	sInstrPhone = oInstr("phone")
	isphonepublic = oInstr("isphonepublic")
	sMobilePhone = oInstr("cellphone")
	iscellpublic = oInstr("iscellpublic")
	sWebsiteURL = oInstr("websiteurl")
	sImageAlt = oInstr("imgalt")
	sImageURL = oInstr("imgurl")
	sBio = oInstr("bio")
	iUserId = Session("UserId")
Else
	isemailpublic = False 
	isphonepublic = False 
	iscellpublic = False 
	iUserId = Session("UserId")
End If

oInstr.close
Set oInstr = Nothing 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility11.css" />

	<script language="Javascript" src="tablesort.js"></script>

<script language="Javascript">
<!--

	function Validate() 
	{
		document.frmInstructor.sInstrPhone.value = document.frmInstructor.sPhone1.value + document.frmInstructor.sPhone2.value + document.frmInstructor.sPhone3.value;
		document.frmInstructor.sMobilePhone.value = document.frmInstructor.sMobilePhone1.value + document.frmInstructor.sMobilePhone2.value + document.frmInstructor.sMobilePhone3.value;
		document.frmInstructor.submit();
	}

	function doPicker(sFormField) 
	{
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("imagepicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

	function insertAtURL (textEl, text) 
	{
		if (textEl.createTextRange && textEl.caretPos) 
		{
			var caretPos = textEl.caretPos;
			caretPos.text =
			caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
			text + ' ' : text;
		}
		else
			textEl.value  = text;
	}

	function openWin2(url, name) 
	{
		var w = (screen.width - 350)/2;
		var h = (screen.height - 550)/2;
		popupWin = eval('window.open(url, name,"resizable,width=820,height=600,left=' + 80 + ',top=' + h + '")');
	}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e) 
	{
		var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
			input.value = input.value.slice(0, len);
		var addNdx = 1;

		while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
		{
			addNdx++;
			//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
		}

		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
	}

	function containsElement(arr, ele) 
	{
		var found = false, index = 0;

		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
		return found;
	}

	function getIndex(input) 
	{
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
	}
		return true;
	}

//-->
</script>

</head>

<body>

 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: <%=sTitle%></strong></font><br />
	<!--<a href="instructor_mgmt.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: FUNCTION LINKS-->
<div id="functionlinks">
	<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;<%=sLinkText%></a>&nbsp;&nbsp;
</div>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="frmInstructor" action="instructor_profile_save.asp" method="post">
<input type="hidden" name="instructorid" value="<%=iInstructorID%>" />
<input type="hidden" name="userid" value="<%=iUserId%>" />

<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Instructor Information</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td align="right">First Name:</td><td><input type="text" name="sFirstName" value="<%=sFirstName%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right">Middle Name:</td><td><input type="text" name="sMiddle" value="<%=sMiddle%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right">Last Name:</td><td><input type="text" name="sLastName" value="<%=sLastName%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right">Email:</td>
						<td><input type="text" name="sEmail" value="<%=sEmail%>" size="50" maxlength="255" /></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><input type="checkbox" name="isemailpublic" <%	If isemailpublic Then
																				response.write "checked=""checked"""
																			End If %> /> Display email address on public site</td>
					</tr>
					<tr>
						<td align="right">Phone:<input type="hidden" name="sInstrPhone" value="<%=sInstrPhone%>"></td>
						<td>(<input type="text" name="sPhone1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sInstrPhone,1,3)%>">) <input value="<%=Mid(sInstrPhone,4,3)%>" type="text" name="sPhone2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sInstrPhone,7,4)%>" type="text" name="sPhone3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><input type="checkbox" name="isphonepublic" <%	If isphonepublic Then
																				response.write "checked=""checked""" 
																			End If %> /> Display phone number on public site</td>
					</tr>
					<tr>
						<td align="right">Cell Phone:<input type="hidden" name="sMobilePhone" value="<%=sMobilePhone%>"></td>
						<td>(<input value="<%=Mid(sMobilePhone,1,3)%>" type="text" name="sMobilePhone1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" />) <input value="<%=Mid(sMobilePhone,4,3)%>"type="text" name="sMobilePhone2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sMobilePhone,7,4)%>" type="text" name="sMobilePhone3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><input type="checkbox" name="iscellpublic" <%	If iscellpublic Then
																				response.write "checked=""checked""" 
																			End If %> /> Display cell phone number on public site</td>
					</tr>
					<tr>
						<td align="right">Website URL:</td><td><input type="text" name="sWebsiteURL" value="<%=sWebsiteURL%>" size="50" maxlength="255"></td>
					</tr>
					<tr>
						<td align="right" valign="top">Bio:</td>
						<td>
							<textarea id="instrbio" name="sBio"><%=sBio%></textarea>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

</form>
<!--END: EDIT FORM-->

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
' FUNCTION GETINSTRUCTORINFO(IINSTRUCTORID)
'--------------------------------------------------------------------------------------------------
Function GetInstructorInfo( iInstructorID )
	Dim sSql, oValues

	sSQL = "SELECT firstname, middle, lastname, bio, imgurl, imgalt, email, phone, cellphone, isemailpublic, isphonepublic, iscellpublic, websiteurl "
	sSql = sSql & " FROM egov_class_instructor WHERE instructorid = " & iInstructorID 

	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN"), 0, 1

	If NOT oValues.EOF Then
		sFirstName = oValues("firstname")
		sMiddle = oValues("middle")
		sLastName = oValues("lastname")
		sEmail = oValues("email")
		sPhone = oValues("phone")
		sMobilePhone = oValues("cellphone")
		sWebsiteURL = oValues("websiteurl")
		isemailpublic = oValues("isemailpublic")
		sImageURL = oValues("imgurl")
		sImageAlt = oValues("imgalt")
		sBio = oValues("bio")
		isphonepublic = oValues("isphonepublic")
		iscellpublic = oValues("iscellpublic")
	Else
		sPhone = "                "
		sMobilePhone = "                "  
		isemailpublic = False 
		isphonepublic = False 
		iscellpublic = False 
	End If

	oValues.close
	Set oValues = nothing

End Function



%>


