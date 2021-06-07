 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../communitylink_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: archive.asp
' AUTHOR:   David Boyer
' CREATED:  05/12/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a list of documents in descending order by upload date (date added).
'
' MODIFICATION HISTORY
' 1.0  05/12/09	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("documents") = "Y" then
	response.redirect "outage_feature_offline.asp"
end if

Dim oDocsOrg, re, matches

set oDocsOrg = New classOrganization
Set re = New RegExp
re.Pattern = "^\d+$"

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
lcl_hidden       = "HIDDEN"
lcl_feature_name = GetFeatureName("documents")

'Retrieve the search parameters
lcl_docsMonth = request("docsMonth")
lcl_docsYear  = request("docsYear")

'Set to the current month/year if either are blank.
If lcl_docsMonth = "" Then 
	lcl_docsMonth = Month(Now)
Else
	Set matches = re.Execute(lcl_docsMonth)
	If matches.Count > 0 Then
		lcl_docsMonth = CLng(lcl_docsMonth)
	Else
		lcl_docsMonth = Month(Now)
	End If 
End If 

If lcl_docsYear = "" Then 
	lcl_docsYear = Year(Now)
Else
	Set matches = re.Execute(lcl_docsYear)
	If matches.Count > 0 Then
		lcl_docsYear = CLng(lcl_docsYear)
	Else
		lcl_docsYear = Year(Now)
	End If 
End If 

'Get the local date/time
lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

%>
<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/easyform.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>

<!--
  <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
  <script type="text/javascript">var addthis_pub="cschappacher";</script>
-->

<script language="javascript">
  function viewDocFolder(iType,iDocID) {
    lcl_width  = 900;
    lcl_height = 400;
    lcl_left   = (screen.availWidth/2) - (lcl_width/2);
    lcl_top    = (screen.availHeight/2) - (lcl_height/2);

    if(iType=="FOLDER") {
       lcl_URL = document.getElementById("folderURL_"+iDocID).value;
    }else{
       lcl_URL = document.getElementById("documentURL_"+iDocID).value;
       lcl_URL = "<%=replace(sEgovWebsiteURL, "/" & sorgVirtualSiteName, "")%>" + lcl_URL;
    }

  		popupWin = window.open(lcl_URL, "_blank"+iType+iDocID, "width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top + ",resizable=yes,scrollbars=yes,status=yes");
  }
</script>

<style>
  .docTable {
    border:           1pt solid #000000;
    background-color: #c0c0c0;
  }

  .docTableHeaders {
    border-bottom: 1pt solid #000000;
    font-weight:   bold;
    font-size:     12px;
    text-align:    left;
  }
</style>

</head>
<!--#include file="../include_top.asp"-->
<p>
<table border="0" cellspacing="0" cellpadding="0" width="800">
  <tr>
      <td>
        <%
         'Build the welcome message
          lcl_org_name        = oDocsOrg.GetOrgName()
          lcl_org_state       = oDocsOrg.GetState()
          lcl_org_featurename = lcl_feature_name

          oDocsOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
        %>
          <!--<font class="pagetitle"><%'lcl_feature_name%></font>-->
          <% checkForRSSFeed iorgid, "", "", "DOCS", sEgovWebsiteURL %>
      </td>
      <td align="right">
          <% 'displayAddThisButton iorgid %>
      </td>
  </tr>
</table>
<% RegisteredUserDisplay("../") %>
</p>

<div id="content">
  <div id="centercontent">

<table border="0" cellspacing="0" cellpadding="2" width="800">
  <form name="mayorsblog" action="mayorsblog.asp" method="post">
  <tr valign="top">
      <td width="600">
        <%
          response.write "<input type=""button"" name=""viewDocuments"" id=""viewDocuments"" class=""button"" value=""View " & lcl_feature_name & """ onclick=""location.href='menu/home.asp'"" /><br /><br />" & vbcrlf
          response.write "<div style=""font-size:16pt;"">" & monthname(lcl_docsMonth) & " " & lcl_docsYear & "</div><br />" & vbcrlf
          displayNewDocs iorgid, request.cookies("userid"), lcl_docsMonth, lcl_docsYear
        %>
      </td>
      <td>
        <%
          response.write "<div align=""center"" style=""font-size:16pt;"">ARCHIVES</div><br />" & vbcrlf
          response.write "<div align=""center"">" & vbcrlf

          displayArchives iorgid

          response.write "</div>" & vbcrlf
        %>
      </td>
  </tr>
  </form>
</table>
<p>&nbsp;</p>

  </div>
</div>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displayNewDocs(p_orgid, p_userid, p_docsMonth, p_docsYear)

  if p_docsMonth = "" then
     lcl_docsMonth = month(now)
  else
     lcl_docsMonth = clng(p_docsMonth)
  end if

  if p_docsYear = "" then
     lcl_docsYear = year(now)
     lcl_docsYear = clng(lcl_docsYear)
  else
     lcl_docsYear = clng(p_docsYear)
  end if

 'Get the total document count for the org
  lcl_totaldocs = 0

 'Setup email fields
  lcl_email_body = ""

  sSql = "SELECT count(d.documentid) as total_docs "
  sSql = sSql & " FROM documents d "
  sSql = sSql & " WHERE d.orgid = " & p_orgid
  sSql = sSql & " AND UPPER(d.documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "

  set oDocCount = Server.CreateObject("ADODB.Recordset")
  oDocCount.Open sSql, Application("DSN"), 3, 1

  if not oDocCount.eof then
     lcl_totaldocs = oDocCount("total_docs")
  end if

  oDocCount.close
  set oDocCount = nothing

  sSql = "SELECT d.documentid, "
  sSql = sSql & " d.documenturl, "
  sSql = sSql & " d.parentfolderid, "
  sSql = sSql & " d.documenttitle, "
  sSql = sSql & " d.dateadded, "
  sSql = sSql & " d.creatoruserid, "
  sSql = sSql & " d.linkurl, "
  sSql = sSql & " d.linktargetsnew, "
  sSql = sSql & " d.documentsize, "
  sSql = sSql & " (select df.folderpath "
  sSql = sSql &  " from documentfolders df "
  sSql = sSql &  " where df.folderid = d.parentfolderid) AS parentfolderurl "
  sSql = sSql & " FROM documents d "
  sSql = sSql & " WHERE UPPER(d.documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "
  sSql = sSql & " AND d.orgid = "  & p_orgid
  sSql = sSql & " AND month(d.dateadded) = " & lcl_docsMonth
  sSql = sSql & " AND year(d.dateadded) = "  & lcl_docsYear
  sSql = sSql & " ORDER BY d.dateadded DESC, d.documentid "

 	set oDocArchive = Server.CreateObject("ADODB.Recordset")
  oDocArchive.Open sSql, Application("DSN"), 3, 1

  if not oDocArchive.eof then
     response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""docTable"">" & vbcrlf
     response.write "    <tr align=""left"">" & vbcrlf
     response.write "        <td class=""docTableHeaders"" width=""150"">Date Added</td>" & vbcrlf
     response.write "        <td class=""docTableHeaders"">Document</td>" & vbcrlf
     response.write "        <td class=""docTableHeaders"" colspan=""2"">&nbsp;</td>" & vbcrlf
     response.write "    </tr>" & vbcrlf

     lcl_bgcolor = "#eeeeee"
     iTotalCnt   = 0

     do while not oDocArchive.eof

        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        iTotalCnt   = iTotalCnt + 1
        iShowDoc    = 0

       'This check will verify that the document exists in the location within the filesystem.
        lcl_valid_doc = checkDocExists(oDocArchive("documenturl"))

        if lcl_valid_doc then

          'Make sure that the folder path is to the top-most parent folder
           lcl_root_path = "/public_documents300/"
           lcl_root_path = lcl_root_path & "custom/pub/"
           lcl_root_path = lcl_root_path & sOrgVirtualSiteName
           lcl_root_path = lcl_root_path & "/published_documents/"
'dtb_debug("filepath: [" & p_filepath & "] - [" & lcl_root_path & "]")

           lcl_filepath       = oDocArchive("parentfolderurl")
           lcl_slash_position = 0

           if lcl_filepath <> "" then
              lcl_filepath       = replace(lcl_filepath,lcl_root_path,"")
              lcl_slash_position = instr(lcl_filepath,"/")
           else
             'If a ParentFolderURL cannot be found then the document needs to be deleted 
             'and the DocSync needs to be run for the org
              lcl_email_body = lcl_email_body & "DocumentID: ["     & oDocArchive("documentid")    & "]<br />"
              lcl_email_body = lcl_email_body & "Document Title: [" & oDocArchive("documenttitle") & "]<br />"
              lcl_email_body = lcl_email_body & "Document URL: ["   & oDocArchive("documenturl")   & "]"
              lcl_email_body = lcl_email_body & "<br /><br />"
           end if

'dtb_debug("path: [" & lcl_filepath & "] - slash position: [" & lcl_slash_position & "]")
           if lcl_slash_position > 0 then
              lcl_parentfolder = mid(lcl_filepath,1,lcl_slash_position)
           else
              lcl_parentfolder = lcl_filepath
           end if

           if lcl_parentfolder <> "" then
              lcl_parentfolder = replace(lcl_parentfolder,"/","")
           end if
'dtb_debug(lcl_filepath & " - " & lcl_slash_position & " - " & lcl_parentfolder)
           lcl_parentfolderurl = lcl_root_path & lcl_parentfolder

          'Determine if the any of the documents are restricted.
           lcl_restrictdocs_exist = checkDocRestrictionExists(p_orgid, lcl_parentfolderurl)

          'If restricted documents exists now verify that the the user has access to see them.
          'If no restricted documents exist then simply show the documents
           if lcl_restrictdocs_exist then
              lcl_hasaccess = checkDocumentAccess(p_orgid,p_userid,lcl_parentfolderurl)

             'Track the number of "restricted" documents were not shown.
              if lcl_hasaccess then
                 iShowDoc         = 1
                 iCountRestricted = iCountRestricted
              else
                 iShowDoc         = 0
                 iCountRestricted = iCountRestricted + 1
              end if
           else
              iShowDoc = 1
           end if

          'Display the document if there are no restrictions
           if iShowDoc = 1 then
              iLineCnt = iLineCnt + 1

              lcl_viewDocLink = sEgovWebsiteURL & "/admin"
              lcl_viewDocLink = lcl_viewDocLink & replace(oDocArchive("documenturl"),"/public_documents300","")

              if oDocArchive("documentsize") <> "" then

                 lcl_documentsize = CLng(oDocArchive("documentsize"))

           						if lcl_documentsize > CLng(1024) then
             			 			sFileSize = FormatNumber((lcl_documentsize / 1024),0) & " KB"
                 else
          		   					sFileSize =  FormatNumber(lcl_documentsize,0) & " Bytes"
                 end if

                 lcl_docsize = "<br />( " & sFileSize & " )"
              else
                 lcl_docsize = ""
              end if

              response.write "    <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
              response.write "        <td width=""150"">" & formatdatetime(oDocArchive("dateadded"),vbshortdate) & "</td>" & vbcrlf
              response.write "        <td>" & trim(oDocArchive("documenttitle")) & lcl_docsize & "</td>" & vbcrlf
              response.write "        <td align=""center"">" & vbcrlf
              response.write "            <input type=""hidden"" name=""documentURL_" & oDocArchive("documentid") & """ id=""documentURL_" & oDocArchive("documentid") & """ value=""" & replace(oDocArchive("documenturl"),"/custom/pub","") & """ />" & vbcrlf
              response.write "            <input type=""button"" name=""viewDocument" & oDocArchive("documentid") & """ id=""viewDocument" & oDocArchive("documentid") & """ value=""View Document"" class=""button"" onclick=""viewDocFolder('DOC','" & oDocArchive("documentid") & "');"" />" & vbcrlf
              response.write "        </td>" & vbcrlf
              response.write "        <td>" & vbcrlf
              response.write "            <input type=""hidden"" name=""folderURL_" & oDocArchive("documentid") & """ id=""folderURL_" & oDocArchive("documentid") & """ value=""" & sEgovWebsiteURL & "/docs/menu/home.asp?path=" & oDocArchive("documenturl") & """ />" & vbcrlf
              response.write "            <input type=""button"" name=""openDocFolder" & oDocArchive("documentid") & """ id=""openDocFolder" & oDocArchive("documentid") & """ value=""Open Folder"" class=""button"" onclick=""viewDocFolder('FOLDER','" & oDocArchive("documentid") & "');"" />" & vbcrlf
              response.write "        </td>" & vbcrlf
              response.write "    </tr>" & vbcrlf
           end if
        end if

        if iTotalCnt = lcl_totaldocs then
           exit do
        else
           oDocArchive.movenext
        end if

     loop

     response.write "  </table>" & vbcrlf

  end if

  oDocArchive.close
  set oDocArchive = nothing

 'If a ParentFolderURL cannot be found then the document needs to be deleted 
 'and the DocSync needs to be run for the org
  if lcl_email_body <> "" then
     lcl_email_from      = "egovsupport@eclink.com"
     lcl_email_sendto    = "egovsupport@eclink.com"
     lcl_email_cc        = ""
     lcl_email_subject   = "DOCUMENTS ERROR: No ParentFolder associated to Document"
     lcl_email_textbody  = ""
     lcl_high_importance = ""

     lcl_email_body = "Org: " & sOrgName & " (" & p_orgid & ")<br /><br />" & lcl_email_body
     lcl_email_body = lcl_email_body & "NOTE: This document(s) has no ParentFolderURL associated to it. "
     lcl_email_body = lcl_email_body & "Delete the document(s) from the ""Documents"" table run the DocSync "
     lcl_email_body = lcl_email_body & "to fix the issue."

     sendEmail lcl_email_from, lcl_email_sendto, lcl_email_cc, lcl_email_subject, _
               lcl_email_body, lcl_email_textbody, lcl_high_importance
  end if

end sub

'------------------------------------------------------------------------------
sub displayArchives(p_orgid)

 'Retreive a distinct list of createdbydates from egov_mayorsblog.
  sSql = "SELECT distinct DATEPART(mm,d.dateadded) AS docsMonth, "
  sSql = sSql & " DATEPART(yyyy,d.dateadded) as docsYear "
  sSql = sSql & " FROM documents d "
  sSql = sSql & " WHERE UPPER(d.documenturl) LIKE ('%/PUBLISHED_DOCUMENTS%') "
  sSql = sSql & " AND d.orgid = " & p_orgid
  sSql = sSql & " ORDER BY 2 DESC, 1 DESC "

 	set oDocDates = Server.CreateObject("ADODB.Recordset")
  oDocDates.Open sSql, Application("DSN"), 3, 1

  if not oDocDates.eof then
     do while not oDocDates.eof

        response.write "<a href=""" & sEgovWebsiteURL & "/docs/archive.asp?docsMonth=" & oDocDates("docsMonth") & "&docsYear=" & oDocDates("docsYear") & """>" & monthname(oDocDates("docsMonth")) & " " & oDocDates("docsYear") & "</a><br />" & vbcrlf

        oDocDates.movenext
     loop
  end if

  oDocDates.close
  set oDocDates = nothing

end sub

'------------------------------------------------------------------------------
sub getCurrentNewsArchive(ByVal p_orgid, ByRef lcl_docsMonth, ByRef lcl_docsYear)

  sSql = "SELECT max(isnull(publicationstart,itemdate)) AS maxDate "
  sSql = sSql & " FROM egov_news_items "
  sSql = sSql & " WHERE itemdisplay = 1 "
  sSql = sSql & " AND orgid = " & p_orgid

 	set oMaxDate = Server.CreateObject("ADODB.Recordset")
  oMaxDate.Open sSql, Application("DSN"), 3, 1

  if not oMaxDate.eof then
     lcl_docsMonth = month(oMaxDate("maxDate"))
     lcl_docsYear  = year(oMaxDate("maxDate"))
  else
     lcl_docsMonth = month(now)
     lcl_docsYear  = year(now)
  end if

  oMaxDate.close
  set oMaxDate = nothing

end sub
%>
