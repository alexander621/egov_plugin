<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: faq.asp
' AUTHOR: Steve Loar
' CREATED: 11/09/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays internal Frequently Asked Questions (FAQ).
'
' MODIFICATION HISTORY
' 1.0  11/09/07 Steve Loar - Initial Version
' 1.1  03/24/09 David Boyer - Added "faqtype" for new Rumor Mill feature
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Check for the faqtype
 if request("faqtype") <> "" then
    lcl_faqtype = UCASE(request("faqtype"))
 else
    lcl_faqtype = "FAQ"
 end if

'Based on the faqtype check for the proper permission
 if lcl_faqtype = "RUMORMILL" then
    lcl_userpermission = "rumormill_internal"
    lcl_pagetitle      = "Rumor Mill"
 else
    lcl_userpermission = "internal faq"
    lcl_pagetitle      = "FAQ"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Internal <%=lcl_pagetitle%>s}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<p><span class=""titletext"">Internal " & lcl_pagetitle & "s</span><br /></p>" & vbcrlf

 	sCategory = "none"

 	sSQL = "SELECT FAQ.FaqQ, FAQ.faqA, isnull(faqcategoryname,'') AS faqcategoryname "
  sSQL = sSQL & " FROM FAQ "
 	sSQL = sSQL &      " LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
  sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 	sSQL = sSQL & " WHERE faq.orgid = " & session("orgid")
  sSQL = sSQL & " AND internalonly = 1 "
 	sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
 	sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
  sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & lcl_faqtype & "' "
 	sSQL = sSQL & " ORDER BY displayorder, sequence"

	 set oPaymentServices = Server.CreateObject("ADODB.Recordset")
	 oPaymentServices.Open sSQL, Application("DSN"), 3, 1

	 if not oPaymentServices.eof then
  		'Place a top tag no matter what categories they has set up
     response.write "<a name=""TOP"">&nbsp;</a><br />" & vbcrlf

     do while not oPaymentServices.eof
        if sCategory <> oPaymentServices("faqcategoryname") then
           sCategory = oPaymentServices("faqcategoryname")

           if sCategory <> "" then
              response.write "<span class=""titletext""><a name=""" & sCategory & """>" & sCategory & "</a></span><br />" & vbcrlf
           end if
        end if

        response.write "<p><font size=""+1""><strong>" & oPaymentServices("faqQ") &  "</strong><br />" & oPaymentServices("faqA") & "</font></p>" & vbcrlf

     			oPaymentServices.movenext
     loop
  else
     response.write "<p>No Internal " & lcl_pagetitle & "s Found.</p>" & vbcrlf
  end if

 	oPaymentServices.close 
	 set oPaymentServices = nothing

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
</body>
</html>