<!-- #include file="../includes/common.asp" //-->
<%
Call deleteBlog(CLng(request("blogid")))

'------------------------------------------------------------------------------
sub deleteBlog(iBlogID)

  sSQL = "DELETE FROM egov_mayorsblog "
  sSQL = sSQL & " WHERE blogid = " & iBlogID
  sSQL = sSQL & " AND orgid = "    & session("orgid")

	 set oDeleteBlog = Server.CreateObject("ADODB.Recordset")
 	oDeleteBlog.Open sSQL, Application("DSN"), 3, 1

  set oDeleteBlog = nothing

  response.redirect "mayorsblog_list.asp?success=SD"

end sub
%>