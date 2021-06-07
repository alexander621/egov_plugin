  <%
  Dim iPathLength

  iPathLength = 0

  ' PROCESS SEARCH REQUEST
  If NewQuery Then
    Set Session("Query") = nothing
    Set Session("Recordset") = nothing
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
      CompSearch = "$contents " & chr(34) & SearchString & chr(34)
    Else
      CompSearch = SearchString
    End If

    Set Q = Server.CreateObject("ixsso.Query")
    Set util = Server.CreateObject("ixsso.Util")
     
    'Search string searches for filename and contents
	CompSearch = "#filename=" & chr(34) & SearchString & "*"& chr(34) &  " OR $contents " & chr(34) &  SearchString & chr(34) 
	Q.Query = CompSearch
    Q.SortBy = "rank[d]"
    Q.Columns = "DocTitle, path, filename, size, write, characterization, rank, vpath"
	Q.MaxRecords = 300

	    If FormScope <> "/" Then
	        ' Security loop to filter records
	   
		' CHECK DATABASE FOR FOLDER PERMISSIONS
		    path = "/public_documents300/custom/pub/" & GetVirtualDirectyName()
'		    sSql = "EXEC ListSearchFolders " & iorgid & ",162,'" & path & "'"
		    sSql = "EXEC ListSearchFolders " & iorgid & ",'" & request.cookies("userid") & "','" & path & "'"
			
			' DEBUG CODE: response.write "<!--" &  sSql & "--><br>" & vbcrlf
		    Set oRst = Server.CreateObject("ADODB.Recordset")
		    oRst.Open sSql, Application("DSN"), 3, 1
	
		 If Not oRst.EOF Then
			Do While Not oRst.EOF
				scopePath = oRst("FolderPath")
				sSearchFolderPath = replace(server.mappath(scopePath),"\custom\pub\custom\pub\","\custom\pub\")
				If scopePath <> FormScope Then
					
					' ONLY INCLUDE PUBLISHED DOCUMENTS IN THE SEARCH
					'If instr(sSearchFolderPath,"published_documents") <> 0 Then

					If instr(sSearchFolderPath,"unpublished_documents") = 0 Then
						iPathLength = iPathLength + Len(sSearchFolderPath)
						If iPathLength < 14479 Then 
							util.AddScopeToQuery Q, sSearchFolderPath , "shallow"
						Else
							response.write "<!-- Kick out at " & sSearchFolderPath  & " -->" & vbcrlf
							Exit Do 
						End If 
						'DEBUG CODE: 
						'response.write sSearchFolderPath  & "<br>" & vbcrlf
					End If

				End If
		        oRst.MoveNext
			Loop
		 End If
	    End If

    

    If SiteLocale <> "" Then
      Q.LocaleID = util.ISOToLocaleID(SiteLocale)
    End If
    
	Q.Catalog = "egovlink300"
    Set RS = Q.CreateRecordSet("nonsequential")

    RS.PageSize = PageSize
    ActiveQuery = TRUE

  ElseIf UseSavedQuery then
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

<br>
<%
' BUILD DISPLAY TABLE WITH MATCHED DOCUMENT INFORMATION
  
  ' PAGE OF INFO
  LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
  CurrentPage = RS.AbsolutePage
  if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
	LastRecordOnPage = RS.RecordCount
  end if
  
  Response.Write "<font class=""label"">&nbsp;&nbsp;&nbsp;&nbsp;Documents " & NextRecordNumber & " to " & LastRecordOnPage
  if RS.RecordCount <> -1 then
	Response.Write " of " & RS.RecordCount
  end if
  Response.Write " matching the query " & chr(34) & "<I>"
  Response.Write SearchString & "</I>" & chr(34) & ".</font><P>"
 
  
  ' BEGIN BUILDING STRING THAT WILL BE THE DISPLAY TABLE
  If  Not RS.EOF and NextRecordNumber <= LastRecordOnPage then
  		strDisplayTable = "<table border=0 class=tablelist cellspacing=0 cellpadding=0 width=""95%"" align=center>"
		strDisplayTable = strDisplayTable & "<colgroup width=105>"
		strDisplayTable = strDisplayTable & "<tr style=""height:26px;""><th class=subheading width=""1%"">Rank</th><th>Document Information</th><th align=left>Last Modified</th><th align=left >File Size</th></tr>"
 end if

 Do While Not RS.EOF and NextRecordNumber <= LastRecordOnPage 
   
	strDisplayTable = strDisplayTable & "<p>"


 ' Graphically indicate rank of document with list of stars (*'s).

	if NextRecordNumber = 1 then
		RankBase=RS("rank")
	end if

	if RankBase>1000 then
		RankBase=1000
	elseif RankBase<1 then
		RankBase=1
	end if

	NormRank = RS("rank")/RankBase

	if NormRank > 0.80 then
		stars = "images/rankbtn5.gif"
	elseif NormRank > 0.60 then
		stars = "images/rankbtn4.gif"
	elseif NormRank > 0.40 then
		stars = "images/rankbtn3.gif"
	elseif NormRank >.20 then
		stars = "images/rankbtn2.gif"
	else stars = "images/rankbtn1.gif"

	end if


    ' BEGIN BUILDING STRING THAT CONTAINS ROW DATA




	  strDisplayTable = strDisplayTable & "<tr><td valign=top align=left><IMG SRC=""" & stars & """>&nbsp;"
	strDisplayTable = strDisplayTable & NextRecordNumber & ". &nbsp;&nbsp;</td>"
	
	strDisplayTable = strDisplayTable & "<td valign=middle><b class=""RecordTitle"">"

	
	'THIS WILL PARSE THE FILE PATH, DID THIS BEFORE BECAUSE IT SEEMED TO BE EASY
	testpath = split(RS("vpath"),"/")
	pub = false
	finalpath = ""
	for i = 0 to UBOUND(testpath)
		if pub = true then
			finalpath = finalpath & "/" & testpath(i)
		end if
		if testpath(i) = "pub" then
			pub = true
		end if
	next
		sReplaceString = "/public_documents300/" & GetVirtualDirectyName() & "/published_documents/"
		'sReplaceString = "/public_documents300/" & GetVirtualDirectyName()
		strDisplayPath = LCASE(RS("vpath"))
		strDisplayPath = Replace(strDisplayPath,sReplaceString,"")
		strDots = ""
		if len(strDisplayPath) > 35 then strDots = "..."
		strDisplayPath = Right(strDisplayPath,35)
		StrDisplayPath = strDots & strDisplayPath
		
		'if len(strDisplayPath) > 20 then strDisplayPath = "..." & Right(strDisplayPath,17)

		if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then
			strDisplayTable = strDisplayTable & "<a href=""" & RS("vpath") & """ class=""RecordTitle""> "
			'strDisplayTable = strDisplayTable & "<img src=""" & GetExtentionImage(RS("filename")) & """ border=0 >"
			'strDisplayTable = strDisplayTable & UCASE(Server.HTMLEncode( RS("filename"))) & "</a>"
			'strDisplayTable = strDisplayTable & UCASE(replace(RS("vpath"),sReplaceString,"")) & "</a>"
			strDisplayTable = strDisplayTable & "NO TITLE<br><font color=black>Location:</font> <nobr>" & strDisplayPath & "</nobr></a>"
		else
			strDisplayTable = strDisplayTable & "<a href="""& RS("vpath") & """ class=""RecordTitle"">"
			strDisplayTable = strDisplayTable & UCASE(Server.HTMLEncode(RS("DocTitle"))) & "<br><font color=black>Location:</font> <nobr>" & strDisplayPath & "</nobr></a>"
		end if
	
		strDisplayTable = strDisplayTable & "</b></td>"
		
        if RS("size") = "" then
			strDisplayTable = strDisplayTable & "(size and time unknown)"
		else
			strDisplayTable = strDisplayTable & "</td><td valign=top> "& LocalizeDate(RS("write"),session("iUserOffset"))&" </td><td valign=top > "&  RS("size") 
			strDisplayTable = strDisplayTable & " bytes"
		end if
			strDisplayTable = strDisplayTable & "</td></tr>"

RS.MoveNext
NextRecordNumber = NextRecordNumber+1
Loop


strDisplayTable = strDisplayTable & "</table>"


else   ' NOT RS.EOF
	if NextRecordNumber = 1 then
          strDisplayTable = "No documents matched the query<P>"
    else
          strDisplayTable = "No more documents in the query<P>"
    end if

end if ' NOT RS.EOF
%>
<table width="95%" cellspacing=0 >
  <tr>
    <td style="padding-left:25px;">
<%if instr(QueryForm,"_") <> 0 then%>
<%

' DISPLAY TOP NAVIGATION CONTROLS
'Previous Button
if CurrentPage > 1 and RS.RecordCount <> -1 then %>
	  <A HREF="#" onClick="document.frmPrev.submit();return false;" >
	  <img src='images/arrow_back.gif' align='absmiddle' border=0>&nbsp;Prev <%=RS.PageSize%></A>&nbsp;&nbsp;
<%End If

'Next Button
if Not RS.EOF then%>
      <A HREF="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src='images/arrow_forward.gif' align='absmiddle' border=0></a>
<%End If%>
<% elseif NextRecordNumber > 1 then %>
	<a href="search_docs.asp?action=Go&searchstring=<%=searchstring%>">VIEW ALL RESULTS</a>
<% end if%>
   </td>
  </tr>
</table>

<% 
' DISPLAY SEARCH RESULTS TO SCREEN
  RESPONSE.WRITE strDisplayTable
%>
<table width="95%" cellspacing=0 >
  <tr>
    <td style="padding-left:25px;">
<%
if instr(QueryForm,"_") <> 0 then

' DISPLAY BOTTOM NAVIGATION CONTROLS
'Previous Button
if CurrentPage > 1 and RS.RecordCount <> -1 then %>
	  <A HREF="#" onClick="document.frmPrev.submit();return false;" >
	  <img src='images/arrow_back.gif' align='absmiddle' border=0>&nbsp;Prev <%=RS.PageSize%></A>&nbsp;&nbsp;
<%End If

'Next Button
if Not RS.EOF then%>
      <A HREF="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src='images/arrow_forward.gif' align='absmiddle' border=0></a>
<%End If%>
<% elseif NextRecordNumber > 1 then%>
	<a href="search_docs.asp?action=Go&searchstring=<%=searchstring%>">VIEW ALL RESULTS</a>
<%
end if%>
   </td>
  </tr>
</table>
<%



if NOT Q.OutOfDate then
' If the index is current, display the fact %>
<P>
    <I><B>The index is up to date.</B></I><BR>
<%end if


  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %>

    <P>
    <I><B>The query is too expensive to complete.</B></I><BR>
<%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %>
    <P>
    <I><B>The query took too long to complete.</B></I><BR>
<%end if%>

<%
'    This is the "previous" form.
'    This retrieves the previous page of documents for the query.
%>

<%SaveQuery = FALSE%>
<%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
        <form name="frmPrev" action="<%=QueryForm%>" method="get">
            <INPUT TYPE="HIDDEN" NAME="qu" VALUE="<%=SearchString%>">
            <INPUT TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
            <INPUT TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <INPUT TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage-1%>" >
			<INPUT TYPE="HIDDEN" NAME = "RankBase" VALUE="<%=RankBase%>">
         </form>
            <%SaveQuery = TRUE%>
<%end if%>

<%
'    This is the "next" form for unsorted queries.
'    This retrieves the next page of documents for the query.
  if Not RS.EOF then%>
         <form name="frmNext" action="<%=QueryForm%>" method="get">
            <INPUT TYPE="HIDDEN" NAME="qu" VALUE="<%=SearchString%>">
            <INPUT TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
            <INPUT TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <INPUT TYPE="HIDDEN" NAME = "RankBase" VALUE="<%=RankBase%>">
			<INPUT TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage+1%>">
         </form>
   
    <% SaveQuery = TRUE
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
    RS.close
    Set RS = Nothing
    Set Q = Nothing
    Set Session("Query") = Nothing
    Set Session("RecordSet") = Nothing
  End If
End if %>
