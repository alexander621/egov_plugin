<!-- #include file="../../includes/common.asp" //-->
<%
Response.Buffer = True

sLocation = session("sitename")

sLevel = "../../" ' Override of value from common.asp

'Check for a Google Custom Search Engine ID
 lcl_googleSearchID = getGoogleSearchID(iOrgID, "googlesearchid_documents")
%>

<html>
<head>
  <title>E-Gov Administration Console</title>

  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="search.css" />
  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../docstyles.css" />

<style type="text/css">
  #content table {
     top: 0px !important;
     left: 0px !important;
  }
</style>

<%
FormScope = Application("eCapture_ArticlesPath")
PageSize = session("PageSize")
SiteLocale = "EN-US"
Dim strDisplayTable, scopePath, iPathLength

'Set Initial Conditions
iPathLength = 0
NewQuery = FALSE
UseSavedQuery = FALSE
SearchString = ""
'QueryForm = Request.ServerVariables("PATH_INFO")


'Did the user press a SUBMIT button to execute the form? If so get the form variables.
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  
	SearchString = Request.Form("SearchString")
    FreeText = Request.Form("FreeText")
    'NOTE: this will be true only if the button is actually pushed.
    If Request("Action") = "Go" Then
      NewQuery = TRUE
			RankBase = 1000
    End If

ElseIf session("strHomeSearch") <> "" Then
     SearchString = session("strHomeSearch")
	 session("strHomeSearch") = ""
	 NewQuery = TRUE
ElseIf Request.ServerVariables("REQUEST_METHOD") = "GET" Then
	SearchString = Request("qu")
	FreeText = Request("FreeText")
	FormScope = Request("sc")
	RankBase = Request("RankBase")
	If Request("pg") <> "" Then
		NextPageNumber = Request("pg")
		NewQuery = FALSE
		UseSavedQuery = TRUE
    Else
		NewQuery = SearchString <> ""
    End if
End If

%>

	<script language="Javascript">
	<!--
		
		function ValidateSearch()
		{
			if (document.getElementById("SearchString").value == '')
			{
				document.getElementById("SearchString").focus();
				alert("Please enter some text in the search box before starting a search.");
			}
			else
			{
				document.frmSearch.submit();
			}
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

<body>
	 <% ShowHeader sLevel %>
	<!--#Include file="../../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

  <!--Begin Page Header -->
	<p>
		<font size="+1"><strong>Documents: Search Results</strong></font><br />
	</p>


	<img src="../../images/arrow_2back.gif" align="absmiddle" />&nbsp;<a href='../default.asp'>Back To Documents</a><br /><br />
	<!--End Page Header -->
  
	<!--Begin Search Form -->
<%
  if lcl_googleSearchID <> "" then
    'Place this tag where you want the search box to render -->
     'response.write "<gcse:searchbox-only></gcse:searchbox-only>" & vbcrlf
     response.write "<gcse:search></gcse:search>" & vbrlf
  else
	 response.write "<form action=""search.asp"" method=""post"" name=""frmSearch"">" & vbcrlf
     response.write "		<table cellpadding=""0"" cellspacing=""0"" border=""0"" id=""searchbox"">" & vbcrlf
	 response.write "		<tr>" & vbcrlf
     response.write "				<td nowrap=""nowrap"">" & vbcrlf
     response.write "					<strong>Search:</strong>" & vbcrlf
     response.write "					<input type=""hidden"" name=""Action"" value=""Go"" />" & vbcrlf
     response.write "					<input type=""text"" id=""SearchString"" name=""SearchString"" size=""65"" maxlength=""100"" style=""background-color:#eeeeee;width:255px; height:19px; border:1px solid #000033;"" value=""" & SearchString & """ />" & vbcrlf
     response.write "					<a href=""#"" onClick='ValidateSearch();'><img src=""../../images/go.gif"" border=""0"" />" & langGo & "</a>" & vbcrlf
     response.write "				</td>" & vbcrlf
     response.write "			</tr>" & vbcrlf
     response.write "		</table>" & vbcrlf
     response.write "	</form>" & vbcrlf
  end if

' PROCESS SEARCH REQUEST
'if lcl_googleSearchID <> "" then
  'Place this tag where you want the search results to render
'   response.write "<gcse:searchresults-only></gcse:searchresults-only>" & vbcrlf
'else
  If NewQuery Then
	Set Session("Query") = Nothing 
	Set Session("Recordset") = Nothing 
	NextRecordNumber = 1

	'Remove any leading and ending quotes from SearchString
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
	  CompSearch = "$contents """ & SearchString & """"
	Else
	  CompSearch = SearchString
	End If

	Set Q = Server.CreateObject("ixsso.Query")
	Set util = Server.CreateObject("ixsso.Util")
	 
	'Search string searches for filename and contents
	CompSearch = "#filename=""" & SearchString & "*"" OR $contents """ &  SearchString & """"
	Q.Query = CompSearch
	Q.SortBy = "rank[d]"
	Q.Columns = "DocTitle, path, filename, size, write, characterization, rank, vpath"
	Q.MaxRecords = 300

    If FormScope <> "/" Then
		' see if they have the restricted folder access feature
		If OrgHasFeature( "public folder" ) Then 
			'response.write "Has Restricted Folders"
			' CHECK DATABASE FOR FOLDER PERMISSIONS
			sSql = "EXEC ListSearchFolders " & session("OrgID") & ", " & Session("UserID") & ", '/public_documents300/custom/pub/" & sLocation & "'"
			'response.write sSql & "<br /><br />"
			'response.End 

			Set oRst = Server.CreateObject("ADODB.Recordset")
			oRst.Open sSql, Application("DSN"), 3, 1
		 
			' Security loop to filter records
			 If Not oRst.EOF Then
				Do While Not oRst.EOF
					scopePath = oRst("FolderPath")
					sSearchFolderPath = replace(server.mappath(scopePath),"\custom\pub\custom\pub\","\custom\pub\")
					If scopePath <> FormScope Then
						iPathLength = iPathLength + Len(sSearchFolderPath)
						If iPathLength < 14479 Then 
							util.AddScopeToQuery Q, sSearchFolderPath, "shallow"
							'DEBUG CODE: 				
							'response.write scopePath & " - path<br />"
						Else
							response.write "<!-- Kick out at " & sSearchFolderPath  & " -->" & vbcrlf
							Exit Do 
						End If 
					End If
					oRst.MoveNext
				Loop
			 End If
			 oRst.Close
			 Set oRst = Nothing 
			 'response.write "Scope: " & FormScope
			 'response.End 
		Else
			documentRoot = Application("DocumentsDrive") & "\" & Application("DocumentsRootDirectory") & "\custom\pub\" & sLocation
			' no restricted folders so search all their documents with a deep search
			'session("searchpath") = "e:\egovlink300_docs\custom\pub\" & sLocation & "\published_documents"
			session("searchpath") = documentRoot & "\published_documents"

			'sSearchFolderPath = server.mappath("e:\egovlink300_docs\custom\pub\" & sLocation & "\published_documents")
			'sSearchFolderPath = "e:\egovlink300_docs\custom\pub\" & sLocation & "\published_documents"
			sSearchFolderPath = documentRoot & "\published_documents"

			'response.write "sSearchFolderPath = " & sSearchFolderPath & "<br />"
			util.AddScopeToQuery Q, sSearchFolderPath, "deep"

			'session("searchpathadmin") = "e:\egovlink300_docs\custom\pub\" & sLocation & "\published_documents"
			session("searchpathadmin") = documentRoot & "\published_documents"

			'sSearchFolderPath = server.mappath("e:\egovlink300_docs\custom\pub\" & sLocation & "\unpublished_documents")
			'sSearchFolderPath = "e:\egovlink300_docs\custom\pub\" & sLocation & "\unpublished_documents"
			sSearchFolderPath = documentRoot & "\unpublished_documents"

			'response.write "sSearchFolderPath = " & sSearchFolderPath & "<br />"
			util.AddScopeToQuery Q, sSearchFolderPath, "deep"
		End If 
'	Else
		'response.write "No FromScope"
  End If

    

    If SiteLocale <> "" Then
      Q.LocaleID = util.ISOToLocaleID(SiteLocale)
    End If
    
	' CHECK TO SEE IF USING SEPARATE CATALOG OR MAIN CATALOG
	If session("blnSeparateIndex") Then
		' CUSTOM INDIVIDUAL CATALOG
		Q.Catalog = "egovlink600_" & session("orgid")
	Else
			' DEFAULT GROUP CATALOG
		Q.Catalog = Application("IndexCatalogName")
	End If
	
	Set RS = Q.CreateRecordSet("nonsequential")

    RS.PageSize = PageSize
    ActiveQuery = TRUE

ElseIf UseSavedQuery Then 

	If IsObject( Session("Query") ) And IsObject( Session("RecordSet") ) Then
		Set Q = Session("Query")
		Set RS = Session("RecordSet")

		If RS.RecordCount <> -1 and NextPageNumber <> -1 then
			RS.AbsolutePage = NextPageNumber
			NextRecordNumber = RS.AbsolutePosition
		End If

		ActiveQuery = TRUE
	Else
		Response.Write "ERROR - No saved query"
	End If
End if

If ActiveQuery Then
	If Not RS.EOF Then
 %>

		<hr width="95%" align="center" size="1" color="black" />
		<p>
<%
		' BUILD DISPLAY TABLE WITH MATCHED DOCUMENT INFORMATION
  
		' PAGE OF INFO
		LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
		CurrentPage = RS.AbsolutePage
		If RS.RecordCount <> -1 And RS.RecordCount < LastRecordOnPage Then 
			LastRecordOnPage = RS.RecordCount
		End If 

		Response.Write "<font class=""label"">&nbsp;&nbsp;&nbsp;&nbsp;Documents " & NextRecordNumber & " to " & LastRecordOnPage
		If RS.RecordCount <> -1 Then 
			Response.Write " of " & RS.RecordCount
		End If 
		Response.Write " matching the query ""<i>"
		Response.Write SearchString & "</i>"".</font><p>"


		' BEGIN BUILDING STRING THAT WILL BE THE DISPLAY TABLE
		If  Not RS.EOF And NextRecordNumber <= LastRecordOnPage Then 
			strDisplayTable = "<table id=""searchtablelist"" cellspacing=""0"" cellpadding=""2"" border=""0"">"
			strDisplayTable = strDisplayTable & "<colgroup width=105>"
			strDisplayTable = strDisplayTable & "<tr style=""height:26px;""><th class=""subheading"" width=""1%"">Rank</th><th>Document Information</th><th align=""center"">Last Modified</th><th align=""center"">Size</th></tr>"
		End If 

		Do While Not RS.EOF And NextRecordNumber <= LastRecordOnPage 
   
			strDisplayTable = strDisplayTable & "<p>"

			' Graphically indicate rank of document with list of stars (*'s).

			If NextRecordNumber = 1 Then 
				RankBase = RS("rank")
			End If 

			If RankBase > 1000 Then 
				RankBase = 1000
			ElseIf RankBase < 1 Then 
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
			Else stars = "images/rankbtn1.gif"
			End If 


			' BEGIN BUILDING STRING THAT CONTAINS ROW DATA
			strDisplayTable = strDisplayTable & "<tr class=""RecordTitle"">"
			strDisplayTable = strDisplayTable & "<td align=""right"" valign=""top"" class=""RecordTitle"">"
			strDisplayTable = strDisplayTable & NextRecordNumber & ". &nbsp;&nbsp;"
			strDisplayTable = strDisplayTable & "</td>"
			strDisplayTable = strDisplayTable & "<td valign=""top"" class=""RecordTitle""><b>"
		
			' BUILD LINKS TO DOCUMENTS FOUND
			sDisplayLink = ""
			sURLLink = ""
			' For public documents
			'sPhysicalBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/published_documents/")
			'sUnpublishedBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/unpublished_documents/")
			'sDisplayLink = Replace(LCase(Replace(LCase(RS("path")),LCase(sPhysicalBasePath),"")),"\","/")
			sDisplayLink = RS("path")
			If InStr(sDisplayLink, "\published_documents\") > 0 Then
				' Format the published documents to work
				sPhysicalBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/published_documents/")
				sDisplayLink = Replace(LCase(Replace(LCase(RS("path")),LCase(sPhysicalBasePath),"")),"\","/")
				sURLLink = "/public_documents300/" & GetVirtualDirectyName() & "/published_documents" & sDisplayLink
			Else
				' Format the unpublished documents to work
				sPhysicalBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/unpublished_documents/")
				sDisplayLink = Replace(LCase(Replace(LCase(RS("path")),LCase(sPhysicalBasePath),"")),"\","/")
				sURLLink = "/public_documents300/" & GetVirtualDirectyName() & "/unpublished_documents" & sDisplayLink
			End If 
			
		
			' WRITE LINKS
			strDisplayTable = strDisplayTable & "<a href=""" & sURLLink & """ class=""RecordTitle""> "
			strDisplayTable = strDisplayTable & "<img src=""" & GetExtentionImage(RS("filename")) & """ border=""0"" />"
			strDisplayTable = strDisplayTable & sDisplayLink & "</a>"

			strDisplayTable = strDisplayTable & "</b></td><td>&nbsp;</td><td>&nbsp;</td></tr>"
			strDisplayTable = strDisplayTable & "<tr><td align=""center"">&nbsp;&nbsp;&nbsp; " & FormatNumber(NormRank * 100) & "%&nbsp;&nbsp;&nbsp; <!--<img src=""" & stars & """>--><br>"

			strDisplayTable = strDisplayTable & "<!--<a href=""summary.htw?"& WebHitsQuery & """><img src=""images/summary.gif""" 
			strDisplayTable = strDisplayTable & "border=""0"" align=""left"" "
			strDisplayTable = strDisplayTable & "alt=""Highlight matching terms in document using Summary mode."">Summary</a>"
			strDisplayTable = strDisplayTable & "--></td><td valign=""top"">"
		
			If VarType(RS("characterization")) = 8 And RS("characterization") <> "" Then 
				strDisplayTable = strDisplayTable & "<b>Abstract:  </b>" & Server.HTMLEncode(RS("characterization")) & "<br>"
        	End If 
		
	        If RS("size") = "" Then 
				strDisplayTable = strDisplayTable & "(time and size unknown)"
			Else 
				strDisplayTable = strDisplayTable & "</td><td align=""left"" nowrap=""nowrap"">" & RS("write") & " </td><td align=""right"" nowrap=""nowrap""> " 
				If CDbl(RS("size")) > CDbl(1000) Then 
					strDisplayTable = strDisplayTable &  FormatNumber(CDbl(RS("size"))/1000,1)
					strDisplayTable = strDisplayTable & " KB"
				Else 
					strDisplayTable = strDisplayTable &  RS("size") 
					strDisplayTable = strDisplayTable & " bytes"
				End If 
			End If 
			strDisplayTable = strDisplayTable & "</td></tr>"

			RS.MoveNext
			NextRecordNumber = NextRecordNumber + 1

		Loop


		strDisplayTable = strDisplayTable & "</table>"


	Else   ' NOT RS.EOF
		If NextRecordNumber = 1 Then 
			strDisplayTable = "No documents matched the query<P>"
		Else 
			strDisplayTable = "No more documents in the query<P>"
		End If 

	End If ' NOT RS.EOF 
%>
	<table width="95%" cellspacing=0 >
		<tr>
			<td style="padding-left:25px;">
<%

	' DISPLAY TOP NAVIGATION CONTROLS
	'Previous Button
	If CurrentPage > 1 And RS.RecordCount <> -1 Then 
%> 
	  <a href="#" onClick="document.frmPrev.submit();return false;">
	  <img src="../../images/arrow_back.gif" align="absmiddle" border="0" />&nbsp;Prev <%=RS.PageSize%></a>&nbsp;&nbsp;
<%
	End If

	'Next Button
	If Not RS.EOF Then
%>
      <a href="#" onClick="document.frmNext.submit();return false;">Next <%=RS.PageSize%>&nbsp;<img src="../../images/arrow_forward.gif" align="absmiddle" border="0" /></a>
<%
	End If
%>
		   </td>
		</tr>
	</table>


<% 
	' DISPLAY SEARCH RESULTS TO SCREEN
	response.write strDisplayTable
%>

	<table width="95%" cellspacing=0 >
		<tr>
			<td style="padding-left:25px;">
<%

	' DISPLAY BOTTOM NAVIGATION CONTROLS
	'Previous Button
	If CurrentPage > 1 And RS.RecordCount <> -1 Then  %>
		<a href="#" onClick="document.frmPrev.submit();return false;">
		<img src="../../images/arrow_back.gif" align="absmiddle" border="0" />&nbsp;Prev <%=RS.PageSize%></a>&nbsp;&nbsp;
<%
	End If

	'Next Button
	If Not RS.EOF Then %>
		<a href="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;
		<img src="../../images/arrow_forward.gif" align="absmiddle" border="0" /></a>
<%	End If		%>
			</td>
		</tr>
	</table>

<%
	If Not Q.OutOfDate Then 
		' If the index is current, display the fact %>
		<p>
		<i><b>The index is up to date.</b></i><br />
<%	End If 


	If Q.QueryIncomplete Then 
		'If the query was not executed because it needed to enumerate to
		'resolve the query instead of using the index, but AllowEnumeration
		'was FALSE, let the user know %>
	    <p>
		<i><b>The query is too expensive to complete.</b></i><br />
<%
	End If 

	If Q.QueryTimedOut Then 
		'    If the query took too long to execute (for example, if too much work
		'    was required to resolve the query), let the user know %>
		<p>
		<i><b>The query took too long to complete.</b></i><br />
<%
	End If


	'    This is the "previous" form.
	'    This retrieves the previous page of documents for the query.

	SaveQuery = False 

	If CurrentPage > 1 And RS.RecordCount <> -1 Then  
%>
		<form name="frmPrev" action="search.asp" method="get" />
			<input type="hidden" name="qu" value="<%=SearchString%>" />
			<input type="hidden" name="FreeText" value="<%=FreeText%>" />
			<input type="hidden" name="sc" value="<%=FormScope%>" />
			<input type="hidden" name="pg" value="<%=CurrentPage-1%>" />
			<input type="hidden" name="RankBase" value="<%=RankBase%>" />
		</form>
<%		
		SaveQuery = True 
	End If

	'    This is the "next" form for unsorted queries.
	'    This retrieves the next page of documents for the query.
	If Not RS.EOF Then   
%>
		<form name="frmNext" action="search.asp" method="get" />
			<input type="hidden" name="qu" value="<%=SearchString%>" />
			<input type="hidden" name="FreeText" value="<%=FreeText%>" />
			<input type="hidden" name="sc" value="<%=FormScope%>" />
			<input type="hidden" name="RankBase" value="<%=RankBase%>" />
			<input type="hidden" name="pg" value="<%=CurrentPage+1%>" />
		</form>
   
<% 
		SaveQuery = True 
	End If 


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
End if 

'end if
%>
		</div>
	</div>
	<!--END: PAGE CONTENT-->

</body>
</html>


<%
'###############################FUNCTIONS###########################################

Function GetExtentionImage( ByVal name )

	url = path & "/" & name
	url = Replace(url, " ", "%20")

	imgSrc = "../images/txt.png"

	pos = InStr(1, name, ".")
	If pos > 0 Then
		Select Case Mid(name,pos+1,3)
			Case "doc"
				imgSrc = "../images/doc.png"
			Case "xls"
				imgSrc = "../images/xls.png"
			Case "ppt"
				imgSrc = "../images/ppt.png"
			Case "htm"
				imgSrc = "../images/html.png"
			Case "pdf"
				imgSrc = "../images/pdf.png"
		End Select
		temp = Left(name,pos-1)
	Else
		temp = name
	End If

	GetExtentionImage = imgSrc
End Function


Function PageOf()

	' Display the page number 
	strReturn = "Page " & CurrentPage 

	If RS.PageCount <> -1 Then
		strReturn = strReturn & " of " & RS.PageCount
	End If

	PageOf = strReturn

End Function


Function LocalizeDate( ByVal d, ByVal iUserOffset )
	localOffset = 0

	If iUserOffset = "" Then
		iUserOffset = 0
	End If

	' Call this function each time you want to display a date
	LocalizeDate = DateAdd("n", localOffset - iUserOffset, d) & sPostDate
End Function



%>
