<!-- #include file="../includes/common.asp" //-->
<%
Call subDeleteFaq( CLng(request("ifaqid")), request("faqtype") )

'------------------------------------------------------------------------------
sub subDeleteFaq( iFaqID, iFAQType )

  if iFAQType = "" then
     iFAQType = "FAQ"
  end if

  sSQL = "DELETE FROM faq WHERE faqid = " & ifaqID & " AND orgid=" & session("orgid")
	 set oDeleteQues = Server.CreateObject("ADODB.Recordset")
 	oDeleteQues.Open sSQL, Application("DSN"), 3, 1

  set oDeleteQues = nothing

  response.redirect "list_faq.asp?faqtype=" & iFAQType & "&success=SD"

end sub
%>