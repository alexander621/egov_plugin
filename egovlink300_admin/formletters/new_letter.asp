<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "form letters" ) Then
  	response.redirect sLevel & "permissiondenied.asp"
End If 
%>
<html>
<head>
	<title>E-GovLink Form Letter Management</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="Javascript">
	<!--
		function doPicker(sFormField) 
		{
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function fnCheckSubject()
		{
			if (document.NewEvent.Subject.value != '') {
				return true;
			}
			else
			{
				return false;
			}
		}

		function previewMe() 
		{
			var FormData = document.NewEvent;
			if (FormData.FLtitle.value == "")
			{
				alert("Please enter a Form Letter Title");
				FormData.FLtitle.focus();
				return false;
			}
			
			if (FormData.FLbody.value == "")
			{
				alert("Please enter the Form Letter body");
				FormData.FLbody.focus();
				return false;
			}
			
			alert("This is a PREVIEW only, Form Letter has not been saved.");
			var newFile = "preview_letter.asp?FLbody=" + FormData.FLbody.value + "&FLtitle=" + FormData.FLtitle.value + "";
			newWin = window.open(newFile,'popupName','width=600,height=500,toolbars=no,left=50,top=50,scrollbars=yes,resizable=yes,status=yes')
			newWin.focus();
			return false;
		}

function addHTMLTag(p_tag) {
  var lcl_body = document.getElementById("FLbody").value;

  if(p_tag=="BOLD") {
     lcl_body = lcl_body + " <B></B>";
  }else if(p_tag=="ITALICS") {
     lcl_body = lcl_body + " <I></I>";
  }else if(p_tag=="H1") {
     lcl_body = lcl_body + " <H1></H1>";
  }else if(p_tag=="H2") {
     lcl_body = lcl_body + " <H2></H2>";
  }else if(p_tag=="H3") {
     lcl_body = lcl_body + " <H3></H3>";
  }else if(p_tag=="LINK") {
     lcl_body = lcl_body + " <A HREF=\"url goes here\"></A>";
  }else if(p_tag=="IMG") {
     lcl_body = lcl_body + " <IMG SRC=\"image filename goes here\" WIDTH=\"0\" HEIGHT=\"0\">";
  }else if(p_tag=="FONT") {
     lcl_body = lcl_body + " <FONT style=\"font-size: 10pt;\"></FONT>";
  }else if(p_tag=="BR") {
     lcl_body = lcl_body + "<BR>";
  }else if(p_tag=="P") {
     lcl_body = lcl_body + "<P>";
  }else if(p_tag=="P_LEFT") {
     lcl_body = lcl_body + " <P align=\"LEFT\"></P>";
  }else if(p_tag=="P_CENTER") {
     lcl_body = lcl_body + " <P align=\"CENTER\"></P>";
  }else if(p_tag=="P_RIGHT") {
     lcl_body = lcl_body + " <P align=\"RIGHT\"></P>";
  }

  document.getElementById("FLbody").value = lcl_body;
}
	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
    <tr>
      <td valign="top">

		<!--<div style="margin-top:20px; margin-left:20px;" >-->

		<p>
		<h3>Add Form Letter </h3>
  <img align="absmiddle" src="../../admin/images/arrow_2back.gif" /><a class="edit" href="list_letter.asp">Return to Form Letter List</a>
  <p>
			<!--<small>[<a class=edit href="copy_form.asp?task=copyme&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Copy This Form</a>]</small> -->
			<!--<small>[<a class=edit href="../action_line/edit_form.asp?task=name&control=<%=iFormID%>&iorgid=<%=iorgid%>">Manage This Form</a>]</small> -->
   <!--<small>[<a class="edit" href="list_letter.asp">Return to Form Letter List</a>]</small>-->
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
      <td width="40%" align="right" nowrap="nowrap"><b>Common HTML formatting tags:</b></td>
      <td>[<a href="javascript:addHTMLTag('BOLD');">Bold</a>]</td>
      <td>[<a href="javascript:addHTMLTag('ITALICS');">Italics</a>]</td>
      <td>[<a href="javascript:addHTMLTag('FONT');">FONT</a>]</td>
      <td>[<a href="javascript:addHTMLTag('H1');">H1</a>]</td>
      <td>[<a href="javascript:addHTMLTag('H2');">H2</a>]</td>
      <td>[<a href="javascript:addHTMLTag('H3');">H3</a>]</td>
      <td>[<a href="javascript:addHTMLTag('LINK');">Link</a>]</td>
      <td>[<a href="javascript:addHTMLTag('IMG');">Image</a>]</td>
      <td>[<a href="javascript:addHTMLTag('BR');">BR</a>]</td>
      <td>[<a href="javascript:addHTMLTag('P');">P</a>]</td>
      <td align="center">
          Alignment:<br>
          <select name="p_format_alignment" onchange="addHTMLTag(this.value);">
            <option value=""></option>
            <option value="P_LEFT">LEFT</option>
            <option value="P_CENTER">CENTER</option>
            <option value="P_RIGHT">RIGHT</option>
          </select>
      </td>
      <td nowrap="nowrap">[<a href="http://www.w3schools.com/tags/default.asp" target="_blank">Additional TAGs</a>]</td>
  </tr>
</table>
<div class="group">
<div class="orgadminboxf">
  <div class="shadow">
			<table class="tablelist" cellspacing="0" cellpadding="5">
					<form name="NewEvent" action="save_letter.asp" method="post">
  			<tr>
		    			<td>
						       <!--BEGIN: Form Letter -->
       						<p>
      							<table>		
						       		<tr>
                   <td>
                       <b>Title:</b><br>
                       <input type="text" class="question" name="FLtitle" value="" size="72" maxlength="199" />
                   </td>
               </tr>
        							<tr>
                   <td>
                       <b>Body:</b><br>
                       <textarea class="none" name="FLbody" id="FLbody" rows="40" cols="100" style="width: 550px; font-size: 10px; font-family: Verdana,Tahoma,Arial;"></textarea>
                   </td>
               </tr>
      							</table>
      							<table>	
						       		<tr><td><input class="button" type="submit" value="ADD Form Letter" /></td>
       								<!--<td><input type=submit onClick="return previewMe()" value=" PREVIEW Form Letter "></td>--></tr>
      							</table>
       						</p>
        					<!--END: Form Letter -->
    					</td>
    					<td valign="top">
        					<!--BEGIN: Form Letter Dynamic Fields-->
       						<b>Instructions</b><br><br>
        					Any of the fields below may be copied/pasted into the body of your form letters. Please copy exactly as they appear including the brackets and astericks.
      					<%
						        sSQL = "Select userfname,userlname,userbusinessname,useremail,userhomephone,userfax,useraddress,usercity,userstate,userzip From egov_users" 
        						Set oUser = Server.CreateObject("ADODB.Recordset")
        						oUser.Open sSQL, Application("DSN"), 3, 1

        						Response.Write "<br><br>Field names:<br>"
       					  For Each Field In oUser.Fields
           							if Field.Name <> "userpassword" then
             								Response.Write "<br> [*" & Field.Name & "*]"
             					end if
       					  Next
        						Set oUser = Nothing
     						%>
        						<!--END: Form Letter Dynamic Fields -->
					    </td>
 				</tr>
					</form>
			</table>
			</div>
			</div>
		</div>

     </td>
    </tr>
  </table>

	</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


