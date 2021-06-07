<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: poc_edit.ASP
' AUTHOR: Steve Loar
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/10/06	Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	sLevel = "../" ' Override of value from common.asp

	If Not UserHasPermission( Session("UserId"), "poc" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 

	Dim sName, sEmail, sPhone, sSql, oInstr

	' GET POC ID
	If request("pocid") = "" OR NOT isnumeric(request("pocid")) OR clng(request("pocid")) = 0 Then
		' CREATE NEW INSTRUCTOR
		iPOCID = 0
		sTitle = "Add New Point of Contact"
		sLinkText = "Create Point of Contact"
	Else
		' EDIT EXISTING INSTRUCTOR
		iPOCID = request("pocid")
		sTitle = "Edit Point of Contact"
		sLinkText = "Save Changes"
	End If

	sSqL = "SELECT name, email, phone FROM egov_class_pointofcontact WHERE pocid = " & iPOCID 

	Set oInstr = Server.CreateObject("ADODB.Recordset")
	oInstr.Open sSQL, Application("DSN"), 0, 1

	If NOT oInstr.EOF Then
		oInstr.movefirst 
		sName = oInstr("name")
		sEmail = oInstr("email")
		sPhone = oInstr("phone")
	Else
		'sPhone = "                "
	End If

	oInstr.close
	Set oInstr = Nothing 
%>


<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="../recreation/facility.css" />
 	<link rel="stylesheet" href="classes.css" />

 	<script src="tablesort.js"></script>

<script>
<!--

	function Validate() 
	{
		// check that a name was entered
		if (document.POCForm.sName.value == "")
		{
			alert("Please enter a name.");
			document.POCForm.sName.focus();
			return;
		}

		// check that an email was entered
		if (document.POCForm.sEmail.value == "")
		{
			alert("Please enter an email address.");
			document.POCForm.sEmail.focus();
			return;
		}
		// validate the format of the email
		//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz))$/;
		var rege = /.+@.+\..+/i;
		var Ok = rege.test(document.POCForm.sEmail.value);

		if (! Ok)
		{
			alert("The email must be in a valid format.");
			document.POCForm.sEmail.focus();
			return;
		}

		//validate the phone number
		var sPhone = document.POCForm.sPhone1.value + document.POCForm.sPhone2.value + document.POCForm.sPhone3.value;
		if (sPhone.length < 10)
		{
			alert( "The phone number must be a valid phone number, including area code.");
			document.POCForm.sPhone1.focus();
			return;
		}
		else
		{
			document.POCForm.sPhone.value = document.POCForm.sPhone1.value + document.POCForm.sPhone2.value + document.POCForm.sPhone3.value;
			var rege = /^\d+$/;
			var Ok = rege.exec(document.POCForm.sPhone.value);
			if ( ! Ok )
			{
				alert( "The phone number must be a valid phone number, including area code.");
				document.POCForm.sPhone1.focus();
				return;
			}
		}

		document.POCForm.submit();
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
		<a href="poc_mgmt.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to Point of Contact Management</a>&nbsp;&nbsp;
		<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;<%=sLinkText%></a>&nbsp;&nbsp;
</div>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="POCForm" action="poc_save.asp" method="post">
<input type="hidden" name="pocid" value="<%=iPOCID%>" >

<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Point of Contact Information</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td align="right">Name:</td><td><input type="text" name="sName" value="<%=sName%>" size="25" maxlength="25" /></td>
					</tr>
					<tr>
						<td align="right">Email:</td>
						<td><input type="text" name="sEmail" value="<%=sEmail%>" size="50" maxlength="255" /></td>
					</tr>
					<tr>
						<td align="right">Phone:<input type="hidden" name="sPhone" value="<%=sInstrPhone%>" /></td>
						<td>(<input type="text" name="sPhone1" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" value="<%=Mid(sPhone,1,3)%>">) <input value="<%=Mid(sPhone,4,3)%>" type="text" name="sPhone2" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3" /> <input value="<%=Mid(sPhone,7,4)%>" type="text" name="sPhone3" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4" /></td>
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


%>


