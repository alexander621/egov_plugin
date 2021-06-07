<!-- #include file="../includes/common.asp" //-->
<%
Dim objFSO

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
If Request.QueryString("Try")="yes" Then
	ChangeLanguage(Request.Form("Language"))
  Application.Lock
	Application("Language") = RemoveExtension(Request.Form("Language"))
  Application.Unlock
	Response.Write ("<script> location.href='ChangeLanguage.asp'; </script>")
End If

' FUNCTION TO CHANGE CURRENT LANGUAGE FILE
Function ChangeLanguage(sNewLanguageFile)
	' RENAME SELECTED LANGUAGE TO LANGUAGE OF CHOICE
	objFSO.CopyFile server.MapPath("../includes/Languages/"&sNewLanguageFile), Server.MapPath("../custom/includes/SelectedLanguage.asp")
End Function

%>

<HTML>
<HEAD>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <title><%=langBSCustomization%></title>
</HEAD>
<BODY>

<%DrawTabs tabAdmin,1%>

<table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_admin.jpg"></td>
      <td><font size="+1"><b><%=langAdminLinks%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../admin"><%=langBackTo%>&nbsp;<%=langAdminLinks%></a></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <% Call DrawQuicklinks("", 1) %>
      </td>
      <td colspan="2" valign="top">
        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
<tr><th align="left">Change language settings</th></tr>          
<tr><td>
<form name=frmSelect action="ChangeLanguage.asp?Try=yes" method="post">
 <% If LCase(Application("Language"))= "english" Then %>
	 <img src="../images/USA.gif" align="absmiddle">
 <% Else %>
	 <img src="../images/SPAIN.gif" align="absmiddle">
 <% End If %>

	<select name="Language">
		<option value="english.asp" <%If Application("Language")= "english" Then response.write "Selected" End IF%>>English
		<option value="spanish.asp" <%If Application("Language")= "spanish" Then response.write "Selected" End IF%>>Spanish
  </select>
  <input type=submit value="Change">
</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<P>&nbsp;</P>

</BODY>
</HTML>

<%
Private Function RemoveExtension(name)
Dim pos, temp
	pos = InStr(1, name, ".")
	temp = Left(name,pos-1)
	RemoveExtension = temp
End Function
%>