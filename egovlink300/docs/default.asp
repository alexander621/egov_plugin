<%
	response.redirect "menu/home.asp" ' This is here to handle the yahoo slurp that crashes on this otherwise unavailable page.
	Session("DSN") = "Provider=SQLOLEDB; Data Source=DEVS0001\SQL2000; User ID=sa; Password=devsql; Initial Catalog=lovelandoh_egov;"
%>

<!-- #include file="../includes/common.asp" //-->

<%
Dim strTopicFile

' GET SEARCH VALUE FROM FORM ON FIRST PAGE
If Request.Form("SearchString") <> "" Then
	session("strHomeSearch") = Request.Form("SearchString")
	strTopicFile = "search/search.asp"
Else
	strTopicFile = "main.asp"
End If


%>
<html>
<head>
  <title><%=langBSDocuments%></title>
  <script>
   // if( self == top ) location.replace( "../../default.htm" );
  </script>
</head>

<frameset rows="143px,*" border="0" name="fstRows">
  <frame src="menu/topmenuNew.asp" name="fraTopMenu" scrolling="no" frameborder="0" noresize>
  <frameset cols="488px,*" border="0" name="fstCols">
    <frame src="menu/menuLove.asp" name="fraToc" scrolling="auto" frameborder="0">
    <frame src="<%=strTopicFile%>" name="fraTopic" frameborder="0">
  </frameset>
</frameset>
</html>