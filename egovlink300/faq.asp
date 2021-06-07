<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: faq.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module display FAQs
'
' MODIFICATION HISTORY
' 1.0 09/12/06		Steve Loar - Changes for categories
' 1.1	11/09/07	Steve Loar - Exclude internal only categories and those not published.
' 1.2 03/24/09 David Boyer - Added "faqtype" for new Rumor Mill feature
' 1.3	05/18/2010	Steve Loar - Put Track_DBsafe() around faqtype to stop hacking attempts.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, oFaqOrg

set oFaqOrg = New classOrganization

'Check for the faqtype
 if request("faqtype") <> "" then
    lcl_faqtype = Track_DBsafe(UCASE(request("faqtype")))
 else
    lcl_faqtype = "FAQ"
 end if

'Based on the faqtype check for the proper permission
 lcl_submitbutton_label = ""

 if lcl_faqtype = "RUMORMILL" then
    'lcl_pagetitle          = "Rumor Mill"
    'lcl_screenlabel        = lcl_pagetitle
    'lcl_submitbutton_label = "Submit a Rumor"
    'lcl_submitbutton_label = "Submit a " & lcl_pagetitle '<<<--- remove per Peter's request 12/10/09 (task: 1496)
    lcl_pagetitle          = oFaqOrg.GetOrgFeatureName("rumormill")
    lcl_feature            = "rumormill"
 else
    'lcl_pagetitle          = "FAQ"
    'lcl_screenlabel        = "Frequently Asked Questions"
    'lcl_submitbutton_label = "Ask a Question"
    'lcl_submitbutton_label = "Ask a " & lcl_pagetitle '<<<--- remove per Peter's request 12/10/09 (task: 1496)
    lcl_pagetitle          = oFaqOrg.GetOrgFeatureName("faq")
    lcl_feature            = "faq"
 end if

'Get the submit button label
 lcl_submitbutton_label = getCommentsLabel(iorgid, "", lcl_feature)

 if lcl_submitbutton_label = "" OR isnull(lcl_submitbutton_label) then
    lcl_submitbutton_label = "Ask a Question"  '<<<--- "Ask a Question" hard-coded per Peter's request 12/10/09 (task: 1496)
 end if
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

 	<title>E-Gov Services - <%=sOrgName & " " & lcl_pagetitle%></title>

	 <link rel="stylesheet" type="text/css" href="css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="global.css" />
	 <link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	 <script language="javascript" src="scripts/modules.js"></script>
	 <script language="javascript" src="scripts/easyform.js"></script> 

  <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
  <script type="text/javascript">var addthis_pub="cschappacher";</script>

</head>

<!--#Include file="include_top.asp"-->

<p>
<table border="0" cellspacing="0" cellpadding="0" style="max-width=800px;">
  <tr>
      <td>
        <%
          lcl_org_name        = oFaqOrg.GetOrgName()
          lcl_org_state       = oFaqOrg.GetState()
          'lcl_org_featurename = lcl_screenlabel
          lcl_org_featurename = lcl_pagetitle

          oFaqOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
        %>
          <!--<font class="pagetitle">Welcome to the <%'sOrgName%>&nbsp;<%'lcl_screenlabel%></font>-->
          <% checkForRSSFeed iorgid, "", "", lcl_faqtype, sEgovWebsiteURL %>
      </td>
      <td align="right">
          <% displayAddThisButton iorgid %>
      </td>
  </tr>
</table>
<%	RegisteredUserDisplay( "" ) %>
</p>

<table border="0" cellpadding="0" cellspacing="0" class="respTable">
  <tr valign="top">
      <td>
          <form action="faq.asp" method="post" id="form1" name="frmSearch">
          <p>
            <font class="searchlabel">Search <%=lcl_pagetitle%>:</font><br />
              <input type="hidden" name="Action" value="Go" />
              <input type="hidden" name="faqtype" value="<%=lcl_faqtype%>" />
            <input type="text" name="SearchString" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" /><br />
            <div class="quicklink" align="right">
              <a href="#" onClick="document.frmSearch.submit()">
                <img src="images/go.gif" border="0" />
                <font class="searchlink">Search</font>
              </a>
            </div>
           <%
            'Determine if there is a form associated to this faq/rumor
            'Also check to ensure that the org has the Action Line feature
             lcl_orghasfeature_action_line = oFaqOrg.orghasfeature("action line")

             if lcl_orghasfeature_action_line then
                lcl_postcomments_formid = getCommentsFormID(iorgid, "", lcl_feature)
                lcl_postcomments_url    = sEgovWebsiteURL & "/action.asp?actionid=" & lcl_postcomments_formid

                if lcl_postcomments_formid > 0 then
                   response.write "<p align=""center"">" & vbcrlf
                   response.write "<input type=""button"" name=""postCommentButton"" id=""postCommentButton"" value=""" & lcl_submitbutton_label & """ class=""button"" onclick=""location.href='" & lcl_postcomments_url & "'"" />" & vbcrlf
                   response.write "</p>" & vbcrlf
                end if
             end if
           %>
          </p>
          </form>
      </td>
      <td>&nbsp;&nbsp;</td>
      <td>
<%
 'BEGIN: Display Page Content --------------------------------------------------
  if request.form("SearchString") <> "" then
 	   SearchString = replace(request.form("SearchString"),"'","''")
  else
 	   SearchString = ""
  end if

 'Display list
 	response.write "          <table border=""0"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td valign=""top"">" & vbcrlf
  'response.write "                    <div class=""faqTitle"">" & lcl_screenlabel & "</div>" & vbcrlf
  response.write "                    <div class=""faqTitle"">" & lcl_pagetitle & "</div>" & vbcrlf
 	response.write "                    <div class=""faqBody"">" & vbcrlf

  if SearchString = "" then
     ShowCategoryNav False 
  else
     response.write "                      <div class=""faqnav""><a href=""faq.asp?faqtype=" & lcl_faqtype & """>Clear Results</a></div>" & vbcrlf
  end if

  if SearchString <> "" then
		   response.write "                      <p align=""center"">Results for search: &quot;" & request.form("SearchString") & "&quot;" & vbcrlf
 	end if

 	fn_DisplayFaqs SearchString

  if SearchString = "" then
     ShowCategoryNav True 
  else
     response.write "                      <div class=""faqnav"">&nbsp;</div>" & vbcrlf
  end if

  response.write "                    </div>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p><br />&nbsp;<br />&nbsp;</p>" & vbcrlf
 'END: Display Page Content ---------------------------------------------------

  response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf
%>
<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
sub fn_DisplayFaqs( dat )
	 Dim sSQL, oPaymentServices, sCategory

 	sCategory = "none"

 	if dat = "" then
   		sSQL = "SELECT FAQ.FaqQ, FAQ.faqA, isnull(faqcategoryname,'') AS faqcategoryname "
     sSQL = sSQL & " FROM FAQ "
 				sSQL = sSQL &      " LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
     sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 				sSQL = sSQL & " WHERE faq.orgid = " & iorgID
     sSQL = sSQL & " AND (internalonly = 0 OR internalonly is null) "
 				sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
 				sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
     sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & lcl_faqtype & "' "
 				sSQL = sSQL & " ORDER BY displayorder, sequence"
 	else
 				sSQL = "SELECT FAQ.FaqID, FAQ.FaqQ, FAQ.faqA, isnull(faqcategoryname,'') AS faqcategoryname "
     sSQL = sSQL & " FROM FAQ "
 				sSQL = sSQL &      " LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
     sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 				sSQL = sSQL & " WHERE faq.orgid = " & iorgID
     sSQL = sSQL & " AND (internalonly = 0 OR internalonly is null) "
 				sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
 				sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
     sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & lcl_faqtype & "' "
 				sSQL = sSQL & " AND (faqQ LIKE '%" & dat & "%' OR faqA LIKE '%" & dat & "%') "
     sSQL = sSQL & " ORDER BY displayorder, sequence"
	 end if

	 set oPaymentServices = Server.CreateObject("ADODB.Recordset")
	 oPaymentServices.Open sSQL, Application("DSN"), 3, 1

	 if not oPaymentServices.eof then
 	 	'Place a top tag no matter what categories they has set up
  	 	response.write "<a name=""TOP"">&nbsp;</a>" & vbcrlf

   		do while not oPaymentServices.eof
     			if sCategory <> oPaymentServices("faqcategoryname") then
           sCategory = oPaymentServices("faqcategoryname")

           if sCategory <> "" then
              response.write "<h4><a name=""" & sCategory & """>" & sCategory & "</a></h4>" & vbcrlf
           end if
        end if

        response.write "<p><strong>" & oPaymentServices("faqQ") &  "</strong><br />" & oPaymentServices("faqA") & "</p>" & vbcrlf

        oPaymentServices.movenext
     loop
  else
     response.write "<p align=""center"">No records available</p>" & vbcrlf
     'if dat = "" then
     '   response.write "<p align=""center"">No " & lcl_pagetitle & " Found.</p>" & vbcrlf
     'else
     '   response.write "<p align=""center"">No " & lcl_pagetitle & " were found that matched your search.</p>" & vbcrlf
     'end if
  end if

  oPaymentServices.close 
  set oPaymentServices = nothing

end sub

'------------------------------------------------------------------------------
sub ShowCategoryNav( bIsFooterNav )
 	Dim sSql, oFAQCats, iLinkCount

	 iLinkCount = 0

 	sSQL = "SELECT FAQCategoryName "
  sSQL = sSQL & " FROM faq_categories "
  sSQL = sSQL & " WHERE orgid = " & iorgID
  sSQL = sSQL & " AND faqcategoryid IN (SELECT faqcategoryid "
  sSQL = sSQL &                       " FROM FAQ) "
  sSQL = sSQL & " AND internalonly = 0 "
  sSQL = sSQL & " AND UPPER(faqtype) = '" & lcl_faqtype & "' "
  sSQL = sSQL & " ORDER BY displayorder"

	 set oFAQCats = Server.CreateObject("ADODB.Recordset")
	 oFAQCats.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

 	response.write "<div class=""faqnav"">" & vbcrlf

 	if bIsFooterNav then
   		response.write "<a href=""#TOP"">Top of Page</a>" & vbcrlf
   		iLinkCount = 1
  end if

  do while not oFAQCats.eof
     if iLinkCount > 0 then
     			response.write " | "
     end if

     response.write "<a href=""#" & oFAQCats("FAQCategoryName") & """>" & oFAQCats("FAQCategoryName") & "</a>" & vbcrlf

     iLinkCount = iLinkCount + 1
     oFAQCats.movenext
  loop

 	oFAQCats.close
	 set oFAQCats = nothing

  response.write "</div>" & vbcrlf

end sub
%>
