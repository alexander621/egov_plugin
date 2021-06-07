<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="facility_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: WAIVER_EDIT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/2006	JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/2006	Steve Loar - Code added
' 1.1	10/06/2006	Steve Loar - Security, Header and nav changed
' 1.2	01/25/2010	Steve Loar - Adding [*NEWPAGE*]
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId, iWaiverId, oRs, sSql
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit facilities" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iFacilityId = CLng(request("iFacilityId"))
iWaiverId = CLng(request("iWaiverId"))

%>

<html>
<head>
	<title>E-Gov Waivers</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="facility.css" />

<script language="Javascript">
  <!--
//	function doPicker(sFormField) {
//      w = (screen.width - 350)/2;
//      h = (screen.height - 350)/2;
//      eval('window.open("../sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=435,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
//    }

	function doPicker(sFormField) 
	{
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("waiverpicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

     function storeCaret (textEl) {
       if (textEl.createTextRange)
         textEl.caretPos = document.selection.createRange().duplicate();
     }

     function insertAtURL (textEl, text) {
       if (textEl.createTextRange && textEl.caretPos) {
         var caretPos = textEl.caretPos;
         caretPos.text =
           caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
             text + ' ' : text;
       }
       else
         textEl.value  = text;
     }


	function SaveWaiver(passForm)
	{
		if (passForm.sName.value == "") {
			alert("Please enter a name.");
			passForm.sName.focus();
			return;
		}
		//if (passForm.sUrl.value == "") {
		//	alert("Please enter a URL.");
		//	passForm.sUrl.focus();
		//	return;
		//}
		if (passForm.sDescription.value == "") {
			alert("Please enter a description.");
			passForm.sDescription.focus();
			return;
		}
		
		var str = new String(passForm.sDescription.value);

		if (str.length > 250) {
			alert("Please shorten your description. \n Descriptions are limited to 250 characters.");
			passForm.sDescription.focus();
			return;
		}

		passForm.submit();
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
	
	<p>
	<h3>Recreation: Waivers</h3>
	<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
	</p>

	<div id="functionlinks">
		<a href="javascript:history.go(-1)"><img src="../images/cancel.gif" align="absmiddle" border="0">&nbsp;Cancel</a>&nbsp;&nbsp;
		<a href="javascript:SaveWaiver(document.waiverform);" id="new_waiver"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Save</a>
	</div>

	<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">

<%
	If iWaiverId > CLng(0) then
		sSql = "SELECT orgid, name, description,body FROM egov_waivers WHERE waiverid = " & iWaiverId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		
		If Not oRs.EOF Then
				
%>				
			<tr>
				<td valign="top"><form name="waiverform" method="post" action="waiver_save.asp">
				<input type="hidden" name="iWaiverId" value="<%=iWaiverId%>" />
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
				<strong>Waiver Name:</strong></td>
				<td valign="top"><input type="text" name="sName" value="<%=oRs("name")%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td valign="top"><strong>Description:</strong><br />(250 char max)</td>
				<td><textarea name="sDescription" id="waiver_desc" rows="6" cols="50" wrap="off"><%=oRs("description")%></textarea>
				</td>
			</tr>
			<tr>
				<td valign="top"><strong>Body:</strong></td>
			    <td>
					<table>
						<tr>
							<td valign="top"><textarea name="sBody" id="waiver_body" style="width:500px;height:400px;" wrap="off"><%=oRs("body")%></textarea></td>
							<td valign="top" width="100%" align="left"><p><strong>Instructions.</strong><br />Any of the fields below may be copied/pasted into the body of your waiver. Please copy exactly as they appear including the brackets and astericks.</p>
								<p>
									[*NEWPAGE*]<br />
									[*checkindate*]<br />
									[*checkintime*]<br />	
									[*checkoutdate*]<br />	
									[*checkouttime*]<br />	
									[*datecreated*]<br />	
									[*dateapproved*]<br />
									[*amount*]<br />	
									[*firstname*]<br />
									[*middle*]<br />	
									[*lastname*]<br />	
									[*address1*]<br />	
									[*address2*]<br />	
									[*city*]<br />	
									[*state*]<br />	
									[*zip*]<br />
									[*email*]<br />
									[*organization*]<br />	
									[*pointofcontact*]<br />	
								</p>
							</td>
						</tr>
					</table>
				</td>
			</tr>
<%
		End If 

		oRs.close
		Set oRs = Nothing 
	Else
		' Set up for new waiver
%>
		<tr>
			<td valign="top">
				<form name="waiverform" method="post" action="waiver_save.asp">
				<input type="hidden" name="iWaiverId" value="<%=iWaiverId%>" />
				<input type="hidden" name="iFacilityId" value="<%=iFacilityId%>" />
				<strong>Waiver Name:</strong>
			</td>
			<td valign="top">
				<input type="text" name="sName" value="" size="20" maxlength="20" />
			</td>
		</tr>
		<tr>
			<td valign="top">
				<strong>Description:</strong><br />(250 char max)
			</td>
			<td valign="top">
				<textarea name="sDescription" id="waiver_desc" rows="6" cols="50" wrap="off"></textarea>
			</td>
		</tr>
		<tr>
			<td valign="top">
				<strong>Body:</strong>
			</td>
			<td>
				<table>
					<tr>
						<td valign="top">
							<textarea name="sBody" id="waiver_body" style="width:500px;height:400px;" wrap="off">[*NEWPAGE*]</textarea>
						</td>
						<td valign="top" width="100%" align="left">
							<p>
								<strong>Instructions.</strong><br />Any of the fields below may be copied/pasted into the body of your waiver. Please copy exactly as they appear including the brackets and astericks.
							</p>
							<p>
								[*NEWPAGE*]<br />
								[*checkindate*]<br />
								[*checkintime*]<br />	
								[*checkoutdate*]<br />	
								[*checkouttime*]<br />	
								[*datecreated*]<br />	
								[*dateapproved*]<br />
								[*amount*]<br />	
								[*firstname*]<br />
								[*middle*]<br />	
								[*lastname*]<br />	
								[*address1*]<br />	
								[*address2*]<br />	
								[*city*]<br />	
								[*state*]<br />	
								[*zip*]<br />
								[*email*]<br />
								[*organization*]<br />	
								[*pointofcontact*]<br />	
							</p>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		</form>
<%
	End If 
	
%>
	</table>
	</div>
	<div id="functionlinks2">
		<a href="javascript:history.go(-1)"><img src="../images/cancel.gif" align="absmiddle" border="0">&nbsp;Cancel</a>&nbsp;&nbsp;
		<a href="javascript:SaveWaiver(document.waiverform);" id="new_waiver"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Save</a>
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


%>


