<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp"-->
<!-- #include file="class/classOrganization.asp"-->
<!-- #include file="includes/yahooBOSS_search.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: yahooBOSS_test.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Action Line Search Results.
'
' MODIFICATION HISTORY
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim objXmlHttp, objXMLDOM, strHTML
 dim oActionOrg

 set oActionOrg = New classOrganization
 set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
 set objXMLDOM  = Server.CreateObject("Microsoft.XMLDOM")

'Show/Hide hidden fields.  To Hide = "HIDDEN", To Show = "TEXT"
 lcl_hidden = "hidden"

'Check for cookies
 lcl_cookie_userid = request.cookies("userid")

'Build the Title
 lcl_title = sOrgName

 if iorgid <> 7 then
    lcl_title = "E-Gov Services " & lcl_title
 end if

'Yahoo API Key
 lcl_yahoo_apikey = "TsSvuWTV34GmdxsGf7Jmb_2TJVQpjWAxKT75qzfGiNdEEK6rINnDwVZhuZD9K0T4Zg"

'Check for search values
 lcl_keyword = ""
 lcl_site1   = ""
 lcl_site2   = ""

 if trim(request("search_keyword")) <> "" then
    lcl_keyword = trim(request("search_keyword"))
 end if

 if trim(request("search_site1")) <> "" then
    lcl_site1 = trim(request("search_site1"))
 end if

 if trim(request("search_site2")) <> "" then
    lcl_site2 = trim(request("search_site2"))
 end if

'Check to see if we are limiting the search to a site(s) or complete web search
 lcl_parameter_sites = ""

 if lcl_site1 <> "" then
    lcl_parameter_sites = lcl_parameter_sites & "&sites=" & lcl_site1
 end if

 if lcl_site2 <> "" then
    lcl_parameter_sites = lcl_parameter_sites & "&sites=" & lcl_site2
 end if

 'Web Search
  if lcl_keyword <> "" then
     lcl_yahoo_search_url = "http://boss.yahooapis.com/ysearch/web/v1/" & lcl_keyword & "?appid=" & lcl_yahoo_apikey & "&format=xml" & lcl_parameter_sites
	 session("lcl_yahoo_search_url") = lcl_yahoo_search_url
  else
     lcl_yahoo_search_url = ""
  end if

 'Site Explorer
 'lcl_yahoo_search_url = "http://boss.yahooapis.com/ysearch/se_inlink/v1/http:%2f%2fwww.egovlink.com%2feclink?appid=" & lcl_yahoo_apikey & "&format=xml&count=1"
 'Page Data
 'lcl_yahoo_search_url = "http://boss.yahooapis.com/ysearch/se_pagedata/v1/http::%2f%2fwww.egovlink.com/eclink?appid=" & lcl_yahoo_apikey & "&format=xml"

 'Check for page setup parameters
  lcl_inIFRAME = "N"

  if request("inIFRAME") <> "" then
     lcl_inIFRAME = ucase(request("inIFRAME"))
  end if
%>
<html>
<head>

		<title><%=lcl_title%></title>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

</head>
<%
 if lcl_inIFRAME <> "Y" then
    response.write "<div><strong>Yahoo BOSS</strong></div>" & vbcrlf
    response.write "<p>" & vbcrlf
    response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
    response.write "  <form name=""search_form"" id=""search_form"" action=""yahooBOSS_test_jf.asp"" method=""post"">" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td>Keyword:</td>" & vbcrlf
    response.write "      <td><input type=""text"" name=""search_keyword"" id=""search_keyword"" value=""" & lcl_keyword & """ size=""50"" /></td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td>Site:</td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "          <input type=""text"" name=""search_site1"" id=""search_site1"" value="""     & lcl_site1   & """ size=""50"" />" & vbcrlf
    response.write "          <i>(i.e. skokie.org or www.egovlink.com/skokie)</i>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td>Site:</td>" & vbcrlf
    response.write "      <td><input type=""text"" name=""search_site2"" id=""search_site2"" value="""     & lcl_site2   & """ size=""50"" /></td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2""><input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" /></td>" & vbrlf
    response.write "  </tr>" & vbcrlf
    response.write "  </form>" & vbcrlf
    response.write "</table>" & vbcrlf
    response.write "</p>" & vbcrlf
 end if

'Jerry added this -------------------------------------------------------------
 if lcl_yahoo_search_url <> "" then
    objXmlHttp.open "GET", lcl_yahoo_search_url, False

   'Send it on it's merry way.
    objXmlHttp.send

    response.write "<p>&nbsp;</p>" & vbcrlf
    'response.write "<h1>Some Formatted Search Results:</h1>" & vbcrlf

   'load the widget.xml document
    objXMLDOM.loadXML(objXmlHttp.responseText)
	'session( "XMLText") = objXmlHttp.responseText

    for each objChild in objXMLDOM.documentElement.childNodes
       if objChild.NodeName = "nextpage" then
          'response.write "<p><a href='" & objChild.Text & "'>You could implement a 'Next 10' here (not yet implemented)</a></p>"
       else
          if objChild.NodeName = "resultset_web" then

             for each objSecondChild in objChild.childNodes
                if objSecondChild.NodeName = "result" then

                   for each objThirdChild in objSecondChild.childNodes
                      Select Case objThirdChild.NodeName
                         case "abstract"
                            strAbstract = objThirdChild.Text
                         case "clickurl" 
                            strClickurl = objThirdChild.Text
                         case "date"
                            strDate     = objThirdChild.Text
                         case "dispurl"
                            strDispurl  = objThirdChild.Text
                         case "size"
                            strSize     = objThirdChild.Text
                         case "title"
                            strTitle    = objThirdChild.Text
                         case "url"
                            strUrl      = objThirdChild.Text
                         case else
                      end select
                   next

                   response.write "<p>" & vbcrlf
                   'response.write "<a href='" & strClickurl & "' target=""_top"">" & strTitle & "</a><br />" & vbcrlf
                   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
                   response.write "  <tr>" & vbcrlf
                   response.write "      <td colspan=""2"" style=""font-size:16px; font-weight:bold; color:#800000;"">" & strTitle & "</td>" & vbcrlf
                   response.write "  </tr>" & vbcrlf
                   response.write "  <tr>" & vbcrlf
                   response.write "      <td>&nbsp;&nbsp;</td>" & vbcrlf
                   response.write "      <td>" & strAbstract & "</td>" & vbcrlf
                   response.write "  </tr>" & vbcrlf
                   response.write "  <tr>" & vbcrlf
                   response.write "      <td>&nbsp;&nbsp;</td>" & vbcrlf
                   response.write "      <td><a href='" & strClickurl & "' target=""_top"">" & strDispurl & "</a></td>" & vbcrlf
                   response.write "  </tr>" & vbcrlf
                   response.write "</table>" & vbcrlf
                   response.write "</p>" & vbcrlf
                   'response.write "&nbsp;&nbsp;&nbsp;&nbsp;Size: " & strsize  & " Bytes"        'print anything else that you want.
                end if
             next
          end if
       end if
    next
 end if
session("lcl_yahoo_search_url") = ""
'Print out the request status -------------------------------------------------
' strHTML = objXmlHttp.responseText
' response.write strHTML

' response.write "<p>&nbsp;</p>" & vbcrlf
' response.write "<h1>Here's The Code:</h1>" & vbcrlf
' response.write "<pre>" & vbcrlf
' response.write Server.HTMLEncode(strHTML)
' response.write "</pre>" & vbcrlf

'Trash our object now that I'm finished with it.
 set objXmlHttp = nothing
 set oActionOrg = nothing 
%>
</div>
</div>

<% if lcl_inIFRAME <> "Y" then %>
<!--#Include file="include_bottom.asp"-->  
<%
   end if
	'sSQL = "SELECT * "
 'sSQL = sSQL & " FROM dbo.egov_form_list_200 "
 'sSQL = sSQL & " WHERE ((orgid=" & iorgID & ")) "
 'sSQL = sSQL & " AND (form_category_id <> 6) "
 'sSQL = sSQL & " AND (action_form_internal <> 1) "
 'sSQL = sSQL & " ORDER BY form_category_Sequence, action_form_name "

	'Set oForms = Server.CreateObject("ADODB.Recordset")
	'oForms.Open sSQL, Application("DSN") , 3, 1
%>