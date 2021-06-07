<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:search.asp
' AUTHOR: Steve Loar	
' CREATED: ???
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents search page.
'
' MODIFICATION HISTORY
' 2.0	8/10/2009	Steve Loar - Changed search to use deep searches off the root level folders to
'					get all the files in the results. The shallow search was limited by the amount
'					of folders that could be added to the search string (limit 14479 characters).
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' INITIALIZE VARIABLES
PageSize = session("PageSize")
SiteLocale = "EN-US"
Dim strDisplayTable, scopePath, iPathLength

' SET INITIAL CONDITIONS
iPathLength = 0
NewQuery = False 
UseSavedQuery = False 
SearchString = ""
QueryForm = Request.ServerVariables("PATH_INFO")

  
'DID THE USER PRESS A SUBMIT BUTTON TO EXECUTE THE FORM? IF SO GET THE FORM VARIABLES.
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	SearchString = Request.Form("SearchString")
	FreeText = Request.Form("FreeText")
	
	'NOTE: this will be true only if the button is actually pushed.
	If Request.Form("Action") = "Go" Then
		NewQuery = True 
		RankBase = 1000
	End If

ElseIf session("strHomeSearch") <> "" Then
	SearchString = session("strHomeSearch")
	session("strHomeSearch") =""
	NewQuery = True 

ElseIf Request.ServerVariables("REQUEST_METHOD") = "GET" Then
	SearchString = Request.QueryString("qu")
	FreeText = Request.QueryString("FreeText")
	FormScope = Request.QueryString("sc")
	RankBase = Request.QueryString("RankBase")

	If Request.QueryString("pg") <> "" Then
	  NextPageNumber = Request.QueryString("pg")
	  NewQuery = False 
	  UseSavedQuery = True 
	Else
	  NewQuery = SearchString <> ""
	End if

End If

'Check for a Google Custom Search Engine ID
 lcl_googleSearchID = getGoogleSearchID(iOrgID, "googlesearchid_documents")
%>
<html>
<head>

	<title>Document Search</title>

	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="search.css" /> 
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

	<script language="Javascript">
	<!--

		function ValidateSearch()
		{
			if (document.getElementById("SearchString").value == "")
			{
				document.getElementById("SearchString").focus();
				alert('Please enter some text in the box before starting a search.');
				return;
			}
			document.frmSearch.submit();
		}

	//-->
	</script>

<% if lcl_googleSearchID <> "" then %>
<!-- Put the following javascript before the closing </head> tag. -->
<script>
//  (function() {
//    var cx = '<%'lcl_googleSearchID%>';
//    var gcse = document.createElement('script'); gcse.type = 'text/javascript'; gcse.async = true;
//    gcse.src = (document.location.protocol == 'https:' ? 'https:' : 'http:') +
//        '//www.google.com/cse/cse.js?cx=' + cx;
//    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(gcse, s);
//  })();
  
  (function() {
    var cx = '<%=lcl_googleSearchID%>';
    var gcse = document.createElement('script');
    gcse.type = 'text/javascript';
    gcse.async = true;
    gcse.src = (document.location.protocol == 'https:' ? 'https:' : 'http:') +
        '//www.google.com/cse/cse.js?cx=' + cx;
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(gcse, s);
  })();
</script>
<% end if %>

</head>

<!--#Include file="../../include_top.asp"-->
<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <div style=""padding: 10px;padding-bottom:5px;"">" & vbcrlf
  response.write "            <font class=""pagetitle"">" & sOrgName & " Online Documents Search - <a href=""../menu/home.asp"">Go Back</a></font>" & vbcrlf
  response.write "          </div>" & vbcrlf

 'BEGIN: Search Form ----------------------------------------------------------
  response.write "<fieldset style=""padding:5px; margin:10px; width:400px; border: 1px solid #c0c0c0; border-radius:5px;"">" & vbcrlf
  response.write "		<legend style=""margin-left:10px""><strong>New Search:</strong></legend>" & vbcrlf

  if lcl_googleSearchID <> "" then
    'Place this tag where you want the search box to render -->
     'response.write "<p><gcse:searchbox-only></gcse:searchbox-only></p>" & vbcrlf
     response.write "<gcse:search></gcse:search>" & vbrlf
  else

     response.write "		<form name=""frmSearch"" action=""" & QueryForm & """ method=""post"">" & vbcrlf
     response.write "		  <input type=""hidden"" name=""action"" value=""Go"" />" & vbcrlf
     response.write "		<table>" & vbcrlf
     response.write "    <tr>" & vbcrlf
     response.write "		      <td><input type=""text"" id=""SearchString"" name=""SearchString"" size=""65"" maxlength=""100"" value=""" & SearchString & """ style=""background-color:#eeeeee;width:200px; height:19px; border: 1px solid #000033;"" /></td>" & vbcrlf
     response.write "  		</tr>" & vbcrlf
     response.write "		 	<tr>" & vbcrlf
     response.write "			     <td align=""right"" width=""200"">" & vbcrlf
     response.write "					       <a href=""#"" onClick=""ValidateSearch();""><img src=""../../images/go.gif"" border=""0"" />" & langGo & "</a>" & vbcrlf
     response.write "				    </td>" & vbcrlf
     response.write "				 </tr>" & vbcrlf
     response.write "		</table>" & vbcrlf
     response.write "		</form>" & vbcrlf
  end if

  response.write "</fieldset>" & vbcrlf
 'END: Search Form ------------------------------------------------------------

 'BEGIN: Adobe Code -----------------------------------------------------------
  response.write "<p>" & vbcrlf
  response.write "	<div style=""width:750px; padding-left:10px;padding-bottom:10px;"">" & vbcrlf
  response.write "   <p>" & vbcrlf
  response.write "     Some of the pages within this section link to Portable Document Format (PDF) " & vbcrlf
  response.write "     files which require a PDF reader to view. You may download a free copy of " & vbcrlf
  response.write "     Adobe&reg; Reader&reg; if you do not already have it on your computer." & vbcrlf
  response.write "   </p>" & vbcrlf
  response.write "	  <a href=""http://www.adobe.com/products/acrobat/readstep2.html"" target=""_blank"" title=""Get Adobe Acrobat Reader Plug-in Here""><img border=""0"" src=""../../images/adreader.gif"" hspace=""10"" />Get Adobe Reader.</a>" & vbcrlf
  response.write "	</div>" & vbcrlf
  response.write "</p>" & vbcrlf
 'END: Adobe Code -------------------------------------------------------------

' PROCESS SEARCH REQUEST
if lcl_googleSearchID <> "" then
  'Place this tag where you want the search results to render
   response.write "<gcse:searchresults-only></gcse:searchresults-only>" & vbcrlf
else

If NewQuery Then
	Set Session("Query") = Nothing 
	Set Session("Recordset") = Nothing 
	NextRecordNumber = 1

	'REMOVE ANY LEADING AND ENDING QUOTES FROM SEARCHSTRING
	SrchStrLen = len(SearchString)

	If Left(SearchString, 1) = chr(34) Then
		SrchStrLen = SrchStrLen-1
		SearchString = Right(SearchString, SrchStrLen)
	End If

	If Right(SearchString, 1) = chr(34) Then
		SrchStrLen = SrchStrLen-1
		SearchString = Left(SearchString, SrchStrLen)
	End If

	If FreeText = "on" Then
		CompSearch = "$contents " & chr(34) & SearchString & chr(34)
	Else
		CompSearch = SearchString
	End If

	Set Q = Server.CreateObject("ixsso.Query")
	Set util = Server.CreateObject("ixsso.Util")
	 
	'SEARCH STRING SEARCHES FOR FILENAME AND CONTENTS
	CompSearch = "#filename """ & SearchString & "*"" OR @contents """ &  SearchString & """" 
	'CompSearch = "#filename ""*.pdf"" And #filename ""*" & SearchString & "*""" 
	'response.write CompSearch & "<br />"
	Q.Query = CompSearch
	Q.SortBy = "rank[d]"
	Q.Columns = "DocTitle, path, filename, size, write, characterization, rank"
	Q.MaxRecords = 300

	If FormScope <> "/" Then
	' SECURITY LOOP TO FILTER RECORDS

		' see if they have the restricted folder access feature
		If OrgHasFeature( iorgid, "public folder" ) Then 
   
			' CHECK DATABASE FOR FOLDER PERMISSIONS
			path = "/public_documents300/custom/pub/" & GetVirtualDirectyName()
			sSql = "EXEC ListSubFolders " & iorgid & ", '" & path & "/published_documents'"
			'response.write "<!-- sSql: " & sSql  & " -->" & vbcrlf
			
			' DEBUG CODE: RESPONSE.WRITE "<!--" &  SSQL & "--><BR>" & VBCRLF
			Set oRst = Server.CreateObject("ADODB.Recordset")
			oRst.Open sSql, Application("DSN"), 3, 1
	
			If Not oRst.EOF Then
				Do While Not oRst.EOF
					scopePath = oRst("FolderPath")
					
					sSearchFolderPath = Replace(server.mappath(scopePath),"\custom\pub\custom\pub\","\custom\pub\")
					If scopePath <> FormScope Then
						'response.write "<!-- ScopePath:  " & scopePath  & " -->" & vbcrlf
						' CHECK SECURITY ACCESS ON FOLDER AND TOSS OUT ANY THAT HAVE SECURITY
						If HasAccess( iorgid, 1, scopePath ) Then
							'response.write "<!-- ScopePath added:  " & scopePath  & " -->" & vbcrlf
							' ONLY INCLUDE PUBLISHED DOCUMENTS IN THE SEARCH
							'If InStr(sSearchFolderPath,"\published_documents\") <> 0 Then
								iPathLength = iPathLength + Len(sSearchFolderPath)
								If iPathLength < 14479 Then 
									'response.write "<!-- AddScopeToQuery: " & sSearchFolderPath  & " -->" & vbcrlf
									util.AddScopeToQuery Q, sSearchFolderPath , "deep"
								Else
									'response.write "<!-- Kicked out due to character limit at " & sSearchFolderPath  & " -->" & vbcrlf
									Exit Do 
								End If 
							'End If
						Else
							'response.write "<!-- ScopePath security check failed:  " & scopePath  & " -->" & vbcrlf
						End If
					End If
					oRst.MoveNext
				Loop
			End If
			oRst.Close
			Set oRst = Nothing
		Else
			' no restricted folders so search all their documents with a deep search
			sSearchFolderPath = Replace(server.mappath("/public_documents300/custom/pub/" & GetVirtualDirectyName() & "/published_documents"), "\custom\pub\custom\pub\", "\custom\pub\" )
			'response.write "<!-- sSearchFolderPath = " & sSearchFolderPath & "<br /> -->" & vbcrlf
			util.AddScopeToQuery Q, sSearchFolderPath, "deep"
		End If 
	End If

	If SiteLocale <> "" Then
		Q.LocaleID = util.ISOToLocaleID(SiteLocale)
	End If

	' CHECK TO SEE IF USING SEPARATE CATALOG OR MAIN CATALOG
	If blnSeparateIndex Then
		' CUSTOM INDIVIDUAL CATALOG
		Q.Catalog = "egovlink600_" & iorgid
	Else
		' DEFAULT GROUP CATALOG
		Q.Catalog = Application("DEFAULT_INDEX_CATALOG")
	End If
	'response.write "<!-- Catalog = " & Q.Catalog  & " -->" & vbcrlf
	
	Set RS = Q.CreateRecordSet("nonsequential")

	RS.PageSize = PageSize
	ActiveQuery = TRUE

ElseIf UseSavedQuery then
	If IsObject( Session("Query") ) And IsObject( Session("RecordSet") ) Then
	  Set Q = Session("Query")
	  Set RS = Session("RecordSet")

	  If IsObject(RS) Then 
		  If RS.RecordCount <> -1 and NextPageNumber <> -1 then
			   RS.AbsolutePage = NextPageNumber
			   NextRecordNumber = RS.AbsolutePosition
		  End If
	
		  ActiveQuery = True
	  Else
		Response.Write "ERROR - No saved query. Please try your search again."
	  End If 
	Else
	  Response.Write "ERROR - No saved query"
	End If
End if


If ActiveQuery Then

	If Not RS.EOF Then
		' BUILD DISPLAY TABLE WITH MATCHED DOCUMENT INFORMATION

		' PAGE OF INFO
		LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
		CurrentPage = RS.AbsolutePage
		If RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage Then 
			LastRecordOnPage = RS.RecordCount
		End If 

		Response.Write vbcrlf & "<p><b><font class=""label"">&nbsp;&nbsp;&nbsp;&nbsp;Documents " & NextRecordNumber & " to " & LastRecordOnPage
		If RS.RecordCount <> -1 Then 
			Response.Write " of " & RS.RecordCount
		End If 
		Response.Write " matching the query ""<i>"
		Response.Write SearchString & "</i>"".</font></b></p>"


		' BEGIN BUILDING STRING THAT WILL BE THE DISPLAY TABLE
		If  Not RS.EOF And NextRecordNumber <= LastRecordOnPage Then 
			strDisplayTable = vbcrlf & "<table border=""0"" class=""tablelist"" cellspacing=""0"" cellpadding=""0"" width=""95%"" align=""center"">"
			strDisplayTable = strDisplayTable & "<colgroup width=""105"">"
			strDisplayTable = strDisplayTable & vbcrlf & "<tr style=""height:26px;""><th class=""searchheader"" width=""1%"">Rank</th><th class=""searchheader"">Document Information</th><th class=""searchheader"" align=""left"">Last Modified</th><th class=""searchheader"" align=""left"" >File Size</th></tr>"
		End If 

		Do While Not RS.EOF And NextRecordNumber <= LastRecordOnPage 
	   
			'strDisplayTable = strDisplayTable & "<p>"

			' Graphically indicate rank of document with list of stars (*'s).
		
			If NextRecordNumber = 1 Then 
				RankBase=RS("rank")
			End If 
		
			If RankBase>1000 Then 
				RankBase = 1000
			ElseIf RankBase<1 Then 
				RankBase = 1
			End If 
		
			NormRank = RS("rank")/RankBase
		
			If NormRank > 0.80 Then 
				stars = "images/rankbtn5.gif"
			ElseIf NormRank > 0.60 Then 
				stars = "images/rankbtn4.gif"
			ElseIf NormRank > 0.40 Then 
				stars = "images/rankbtn3.gif"
			ElseIf NormRank >.20 Then 
				stars = "images/rankbtn2.gif"
			Else 
				stars = "images/rankbtn1.gif"
			End If 

			' BEGIN BUILDING STRING THAT CONTAINS ROW DATA
			strDisplayTable = strDisplayTable & vbcrlf & "<tr class=""RecordTitle"">"
			strDisplayTable = strDisplayTable & "<td align=""right"" valign=""top"" class=""RecordTitle"">"
			strDisplayTable = strDisplayTable & NextRecordNumber & ". &nbsp;&nbsp;"
			strDisplayTable = strDisplayTable & "</td>"
			strDisplayTable = strDisplayTable & "<td valign=""top""><b class=""RecordTitle"">"
			
			' BUILD LINKS TO DOCUMENTS FOUND
			sDisplayLink = ""
			sURLLink = ""
			sPhysicalBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/published_documents/")
			sDisplayLink = replace(lcase(replace(lcase(RS("path")),lcase(sPhysicalBasePath),"")),"\","/")
			sURLLink = "/public_documents300/" & GetVirtualDirectyName() & "/published_documents/" & sDisplayLink
			' *********************************************************************************************************
			
			' WRITE LINKS
			strDisplayTable = strDisplayTable & "<a href=""" & sURLLink & """ class=""RecordTitle""> "
			strDisplayTable = strDisplayTable & "<img src=""" & GetExtentionImage(RS("filename")) & """ border=""0"" hspace=""2"" />"
			strDisplayTable = strDisplayTable & sDisplayLink & "</a>"

			strDisplayTable = strDisplayTable & "</b></td><td>&nbsp;</td><td>&nbsp;</td></tr>"
			strDisplayTable = strDisplayTable & vbcrlf & "<tr><td valign=""top"" align=""center"">&nbsp;&nbsp;&nbsp; " & FormatNumber(NormRank * 100) & "%&nbsp;&nbsp;&nbsp; <!--<img src=""" & stars & """>--><br />"

'			strDisplayTable = strDisplayTable & "<!--<a href=""summary.htw?"& WebHitsQuery & """><img src=""images/summary.gif""" 
'			strDisplayTable = strDisplayTable & "border=""0"" align=""left"""
'			strDisplayTable = strDisplayTable & "alt=""Highlight matching terms in document using Summary mode."" />Summary</a>"
'			strDisplayTable = strDisplayTable & "-->"
			strDisplayTable = strDisplayTable & "</td><td valign=""top"">"
			
			If VarType(RS("characterization")) = 8 And RS("characterization") <> "" Then 
				strDisplayTable = strDisplayTable & "<b>Abstract:  </b>" & Server.HTMLEncode(RS("characterization")) & "<br />"
			End If 
			
			If RS("size") = "" Then 
				strDisplayTable = strDisplayTable & "(size and time unknown)"
			Else 
				strDisplayTable = strDisplayTable & "</td><td valign=""top"">"& RS("write") &" </td><td valign=""top""> "&  RS("size") 
				strDisplayTable = strDisplayTable & " bytes"
			End If 
			strDisplayTable = strDisplayTable & "</td></tr>"

			RS.MoveNext
			NextRecordNumber = NextRecordNumber+1
		 Loop
		
		strDisplayTable = strDisplayTable & vbcrlf & "</table>"

	Else    ' NOT RS.EOF (line 175)
		If NextRecordNumber = 1 Then 
			strDisplayTable = "<p>No documents matched the query</p>"
		Else 
			strDisplayTable = "<p>No more documents in the query</p>"
		End If 

	End If  ' NOT RS.EOF

%>
	<table width="95%" cellspacing="0" cellpadding="0" border="0">
	  <tr>
	    <td style="padding-left:25px;">
	<%
	
	' DISPLAY TOP NAVIGATION CONTROLS
	'Previous Button
	If CurrentPage > 1 and RS.RecordCount <> -1 Then  %>
		  <a href="#" onClick="document.frmPrev.submit();return false;">
			<img src="../../images/arrow_back.gif" align="absmiddle" border="0" />&nbsp;Prev <%=RS.PageSize%></a>&nbsp;&nbsp;
<%	End If
	
	'Next Button
	If Not RS.EOF Then  %>
	      <a href="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src="../../images/arrow_forward.gif" align="absmiddle" border="0" /></a>
<%	End If	%>
	   </td>
	  </tr>
	</table>


	<% 
	' DISPLAY SEARCH RESULTS TO SCREEN
	  response.write strDisplayTable
	%>

	<table width="95%" cellspacing="0">
	  <tr>
	    <td style="padding-left:25px;">
	<%
	
	' DISPLAY BOTTOM NAVIGATION CONTROLS
	'Previous Button
	If CurrentPage > 1 And RS.RecordCount <> -1 Then  %>
		  <a href="#" onClick="document.frmPrev.submit();return false;">
			<img src='../../images/arrow_back.gif' align='absmiddle' border="0" />&nbsp;Prev <%=RS.PageSize%></a>&nbsp;&nbsp;
	<%End If
	
	'Next Button
	If Not RS.EOF Then %>
	      <a href="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src='../../images/arrow_forward.gif' align='absmiddle' border="0" /></a>
<%	End If%>
	   </td>
	  </tr>
	</table>

	<%

	If Q.QueryIncomplete Then 
	'    If the query was not executed because it needed to enumerate to
	'    resolve the query instead of using the index, but AllowEnumeration
	'    was FALSE, let the user know 
	    response.write "<p><i><b>The query is too expensive to complete.</b></i></p><br />"
	End If 
	
	If Q.QueryTimedOut Then 
	'    If the query took too long to execute (for example, if too much work
	'    was required to resolve the query), let the user know 
	    response.write "<p><i><b>The query took too long to complete.</b></i></p><br />"
	End If 

	'    This is the "previous" form.
	'    This retrieves the previous page of documents for the query.
	SaveQuery = False 
	If CurrentPage > 1 And RS.RecordCount <> -1 Then %>
		<form name="frmPrev" action="<%=QueryForm%>" method="get">
			<input type="hidden" name="qu" value="<%=SearchString%>" />
			<input type="hidden" name="FreeText" value="<%=FreeText%>" />
			<input type="hidden" name="sc" value="<%=FormScope%>" />
			<input type="hidden" name="pg" value="<%=CurrentPage-1%>" />
			<input type="hidden" name="RankBase" value="<%=RankBase%>" />
		</form>
		<%SaveQuery = True %>
<%	End If  %>
	
	<%
	'    This is the "next" form for unsorted queries.
	'    This retrieves the next page of documents for the query.
	If Not RS.EOF Then	%>
		<form name="frmNext" action="<%=QueryForm%>" method="get">
			<input type="hidden" name="qu" value="<%=SearchString%>" />
			<input type="hidden" name="FreeText" value="<%=FreeText%>" />
			<input type="hidden" name="sc" value="<%=FormScope%>" />
			<input type="hidden" name="RankBase" value="<%=RankBase%>" />
			<input type="hidden" name="pg" value="<%=CurrentPage+1%>" />
		</form>

		<% SaveQuery = True 
	End If %>

<%
	' Page of information
	' Response.write PageOf()

	'If either of the previous or back buttons were displayed, save the query
	'and the recordset in session variables.
	If SaveQuery Then
		Set Session("Query") = Q
		Set Session("RecordSet") = RS
	Else
		RS.Close
		Set RS = Nothing
		Set Q = Nothing
		Set Session("Query") = Nothing
		Set Session("RecordSet") = Nothing
	End If
End If

end if

%>


<p>&nbsp;</p>

<!--#Include file="../../include_bottom.asp"--> 

<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTION AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------
' Function GetExtentionImage( name )
'-------------------------------------------------------------------------------------------------
Function GetExtentionImage( name )
	Dim imgSrc, pos

	'url = path & "/" & name
	'url = Replace(url, " ", "%20")

	imgSrc = "images/document.gif"

	pos = InStr(1, name, ".")
	If pos > 0 Then
		Select Case Mid(name,pos+1,3)
			Case "doc"
				imgSrc = "images/msword.gif"
			Case "xls"
				imgSrc = "images/msexcel.gif"
			Case "ppt"
				imgSrc = "images/msppt.gif"
			Case "htm"
				imgSrc = "images/msie.gif"
			Case "pdf"
				imgSrc = "images/pdf.gif"
		End Select
		'temp = Left(name,pos-1)
	'Else
		'temp = name
	End If

	GetExtentionImage = imgSrc
End Function


'-------------------------------------------------------------------------------------------------
' Function PageOf()
'-------------------------------------------------------------------------------------------------
Function PageOf()
	Dim strReturn

	' Display the page number 
	strReturn = "Page " & CurrentPage 

	If RS.PageCount <> -1 Then
		strReturn = strReturn & " of " & RS.PageCount
	End If

	PageOf = strReturn

End Function


'-------------------------------------------------------------------------------------------------
' Function LocalizeDate( d, iUserOffset )
'-------------------------------------------------------------------------------------------------
Function LocalizeDate( d, iUserOffset )
	Dim localOffset 

	localOffset = 0

	If iUserOffset = "" Then
		iUserOffset = 0
	End If

	' Call this function each time you want to display a date
	LocalizeDate = DateAdd("n", localOffset - iUserOffset,d) & sPostDate

End Function 


'------------------------------------------------------------------------------------------------------------
' FUNCTION SETORGANIZATIONPARAMETERS()
'------------------------------------------------------------------------------------------------------------
Function SetOrganizationParameters33()
	' SET DEFAULT RETURN VALUE
	iReturnValue = 1

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	sSERVER = request.servervariables("SERVER_NAME")
	sCurrent = sProtocol & sSERVER & "/" & GetVirtualDirectyName()


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID WHERE OrgEgovWebsiteURL='" & sCurrent & "'"

	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOrgInfo.EOF Then
		iOrgID = oOrgInfo("OrgID")
		sOrgName = oOrgInfo("OrgName")
		sHomeWebsiteURL = oOrgInfo("OrgPublicWebsiteURL")
		sEgovWebsiteURL = oOrgInfo("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oOrgInfo("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oOrgInfo("OrgTopGraphicRightURL")
		sWelcomeMessage = oOrgInfo("OrgWelcomeMessage")
		sActionDescription = oOrgInfo("OrgActionLineDescription")
		sPaymentDescription = oOrgInfo("OrgPaymentDescription")
		iHeaderSize = oOrgInfo("OrgHeaderSize")
		sTagline = oOrgInfo("OrgTagline")
		iPaymentGatewayID = oOrgInfo("OrgPaymentGateway")
		blnOrgAction = oOrgInfo("OrgActionOn")
		blnOrgPayment = oOrgInfo("OrgPaymentOn")
		blnOrgDocument = oOrgInfo("OrgDocumentOn")
		blnOrgCalendar = oOrgInfo("OrgCalendarOn")
		blnOrgFaq = oOrgInfo("OrgFaqOn")
		sorgVirtualSiteName = oOrgInfo("orgVirtualSiteName")
		sOrgActionName =  oOrgInfo("OrgActionName")
		sOrgPaymentName =  oOrgInfo("OrgPaymentName")
		sOrgCalendarName =  oOrgInfo("OrgCalendarName")
		sOrgDocumentName =   oOrgInfo("OrgDocumentName")
		'sOrgFaqName =  oOrgInfo("OrgFaqName")
		sOrgRegistration = oOrgInfo("OrgRegistration")
		blnCalRequest = oOrgInfo("OrgRequestCalOn")
		iCalForm =  oOrgInfo("OrgRequestCalForm")
		sHomeWebsiteTag = oOrgInfo("OrgPublicWebsiteTag")
		sEgovWebsiteTag = oOrgInfo("OrgEgovWebsiteTag")
		bCustomButtonsOn = oOrgInfo("OrgCustomButtonsOn")
		iTimeOffset = oOrgInfo("gmtoffset")
	End If

	oOrgInfo.Close
	Set oOrgInfo = Nothing 

	If NOT ISNULL(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters33 = iReturnValue
	
End Function


'------------------------------------------------------------------------------------------------------------
' Function GetPageName()
'------------------------------------------------------------------------------------------------------------
Function GetPageName()
	Dim sReturnValue

	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	For Each arr in strURL 
		sReturnValue = arr 
	Next 
	
	GetPageName = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' Function GetVirtualDirectyName()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function


'-------------------------------------------------------------------------------------------------------
' Function HasAccess( iorgid, iuserid, strvpath )
'-------------------------------------------------------------------------------------------------------
Function HasAccess( iOrgId, iUserId, strvpath )
	Dim iReturnValue, sSql, oCnn, rstAccess

	On Error Resume Next

	If CLng(iUserId) > CLng(0) Then 
		iReturnValue = False

		Set oCnn = Server.CreateObject("ADODB.Connection")
		oCnn.Open Application("DSN")
		sSql = "EXEC CHECKFOLDERACCESS '" & iOrgId & "','" & iUserId & "','" & strvpath & "'"
		Set rstAccess = oCnn.Execute(sSql)

		If Not rstAccess.EOF Then
			If rstAccess("folderid") >= 0 Then
				iReturnValue = True
			End If
		End If

		oCnn.Close
		Set rstAccess = Nothing
		Set oCnn = Nothing
	Else
		iReturnValue = True
	End If 

	HasAccess = iReturnValue

End Function



%>