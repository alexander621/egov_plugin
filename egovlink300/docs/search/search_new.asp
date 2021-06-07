<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->

<%
' INITIALIZE VARIABLES
PageSize = session("PageSize")
SiteLocale = "EN-US"
Dim strDisplayTable, scopePath, iPathLength

' SET INITIAL CONDITIONS
iPathLength = 0
NewQuery = FALSE
UseSavedQuery = FALSE
SearchString = ""
QueryForm = Request.ServerVariables("PATH_INFO")

  
'DID THE USER PRESS A SUBMIT BUTTON TO EXECUTE THE FORM? IF SO GET THE FORM VARIABLES.
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	SearchString = Request.Form("SearchString")
	FreeText = Request.Form("FreeText")
	
	'NOTE: this will be true only if the button is actually pushed.
	If Request.Form("Action") = "Go" Then
			NewQuery = TRUE
		RankBase=1000
	End If

ElseIf session("strHomeSearch") <> "" Then
	SearchString = session("strHomeSearch")
	session("strHomeSearch") =""
	NewQuery = TRUE

ElseIf Request.ServerVariables("REQUEST_METHOD") = "GET" Then
	SearchString = Request.QueryString("qu")
	FreeText = Request.QueryString("FreeText")
	FormScope = Request.QueryString("sc")
	RankBase = Request.QueryString("RankBase")

	If Request.QueryString("pg") <> "" Then
	  NextPageNumber = Request.QueryString("pg")
	  NewQuery = FALSE
	  UseSavedQuery = TRUE
	Else
	  NewQuery = SearchString <> ""
	End if

End if
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

</head>

<!--#Include file="../../include_top.asp"-->


<!--BODY CONTENT-->

<TR><TD VALIGN=TOP>
 
 <div style="padding: 10px;padding-bottom:5px;">
 <font class="pagetitle"><%=sOrgName%> Online Documents Search -  <a href="../menu/home.asp">Go Back</a></font>

 </div>


  
 <!--Begin Search Form -->
	
    <fieldset style="padding:0;margin:10px;width:235px;border: 1px solid #000000;">
		<legend><b>New Search:</b></legend>
		<form name="frmSearch" action="<%=QueryForm%>" method="post">
			<input type="hidden" name="Action" value="Go" />
			<table>
				<tr>
					<td><input type="text" id="SearchString" name="SearchString" size="65" maxlength="100" value="<%= SearchString %>" style="background-color:#eeeeee;width:200px; height:19px; border: 1px solid #000033;" /></td>
				</tr>
				<tr>
					<td align="right" width="200">
						<a href="#" onClick='ValidateSearch();'><img src="../../images/go.gif" border="0" /><%=langGo%></a>
					</td>
				</tr>
			</table>
		</form>
    </fieldset>
	
 <!--End Search Form-->
 
  

 <!--BEGIN: ADOBE CODE-->
  <P>
	<div style="width:750px; padding-left:10px;padding-bottom:10px;">Some of the pages within this section link to Portable Document Format (PDF) files which require a PDF reader to view. You may download a free copy of Adobe&reg; Reader&reg; if you do not already have it on your computer.<br><br>
	<A href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=0 src="../../images/adreader.gif" hspace=10>Get Adobe Reader.</a>
	</div>
  </p>
<!--END: ADOBE CODE-->
 
  
<%
' PROCESS SEARCH REQUEST
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
		CompSearch = "#filename=""" & SearchString & "*"" OR @contents """ &  SearchString & """" 
		Q.Query = CompSearch
	    Q.SortBy = "rank[d]"
	    Q.Columns = "DocTitle, path, filename, size, write, characterization, rank, vpath"
		Q.MaxRecords = 300

	    If FormScope <> "/" Then
	    ' SECURITY LOOP TO FILTER RECORDS

			' see if they have the restricted folder access feature
			If OrgHasFeature( iorgid, "public folder" ) Then 
	   
				' CHECK DATABASE FOR FOLDER PERMISSIONS
				path = "/public_documents300/custom/pub/" & GetVirtualDirectyName()
				sSql = "EXEC ListSearchFolders " & iorgid & ", 162, '" & path & "'"
				'response.write "<!-- sSql: " & sSql  & " -->" & vbcrlf
				
				' DEBUG CODE: RESPONSE.WRITE "<!--" &  SSQL & "--><BR>" & VBCRLF
				Set oRst = Server.CreateObject("ADODB.Recordset")
				oRst.Open sSql, Application("DSN"), 3, 1
		
				If Not oRst.EOF Then
					Do While Not oRst.EOF
						scopePath = oRst("FolderPath")
						response.write "<!-- ScopePath:  " & scopePath  & " -->" & vbcrlf
						sSearchFolderPath = replace(server.mappath(scopePath),"\custom\pub\custom\pub\","\custom\pub\")
						If scopePath <> FormScope Then
							
							' CHECK SECURITY ACCESS FOR USER
							If HasAccess(iorgid,request.Cookies("userid"),scopePath) Then

								' ONLY INCLUDE PUBLISHED DOCUMENTS IN THE SEARCH
								If instr(sSearchFolderPath,"\published_documents\") <> 0 Then
									iPathLength = iPathLength + Len(sSearchFolderPath)
									If iPathLength < 14479 Then 
										'response.write "<!-- " & sSearchFolderPath  & " -->" & vbcrlf
										util.AddScopeToQuery Q, sSearchFolderPath , "shallow"
									Else
										response.write "<!-- Kick out at " & sSearchFolderPath  & " -->" & vbcrlf
										Exit Do 
									End If 
								End If

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
				response.write "<!-- sSearchFolderPath = " & sSearchFolderPath & "<br /> -->" & vbcrlf
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
		response.write "<!-- Catalog = " & Q.Catalog  & " -->" & vbcrlf
	    
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

	' BUILD DISPLAY TABLE WITH MATCHED DOCUMENT INFORMATION
	  
	  ' PAGE OF INFO
	  LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
	  CurrentPage = RS.AbsolutePage
	  if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
		LastRecordOnPage = RS.RecordCount
	  end if
	  
	  Response.Write "<b><font class=""label"">&nbsp;&nbsp;&nbsp;&nbsp;Documents " & NextRecordNumber & " to " & LastRecordOnPage
	  if RS.RecordCount <> -1 then
		Response.Write " of " & RS.RecordCount
	  end if
	  Response.Write " matching the query ""<i>"
	  Response.Write SearchString & "</i>"".</font></b><p>"
 
  
	  ' BEGIN BUILDING STRING THAT WILL BE THE DISPLAY TABLE
	  If  Not RS.EOF and NextRecordNumber <= LastRecordOnPage then
	  		strDisplayTable = "<table border=0 class=tablelist cellspacing=0 cellpadding=0 width=""95%"" align=center>"
			strDisplayTable = strDisplayTable & "<colgroup width=105>"
			strDisplayTable = strDisplayTable & "<tr style=""height:26px;""><th class=searchheader width=""1%"">Rank</th><th class=searchheader>Document Information</th><th class=searchheader align=left>Last Modified</th><th class=searchheader align=left >File Size</th></tr>"
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
		strDisplayTable = strDisplayTable & "<tr class=""RecordTitle"">"
		strDisplayTable = strDisplayTable & "<td align=""right"" valign=TOP class=""RecordTitle"">"
		strDisplayTable = strDisplayTable & NextRecordNumber & ". &nbsp;&nbsp;"
		strDisplayTable = strDisplayTable & "</td>"
		strDisplayTable = strDisplayTable & "<td valign=TOP><b class=""RecordTitle"">"
		
		' BUILD LINKS TO DOCUMENTS FOUND
		sDisplayLink = ""
		sURLLink = ""
		sPhysicalBasePath = server.mappath("/public_documents300/" & GetVirtualDirectyName() & "/published_documents/")
		sDisplayLink = replace(lcase(replace(lcase(RS("path")),lcase(sPhysicalBasePath),"")),"\","/")
		sURLLink = "/public_documents300/" & GetVirtualDirectyName() & "/published_documents/" & sDisplayLink
		
		' WRITE LINKS
		strDisplayTable = strDisplayTable & "<a href=""" & sURLLink & """ class=""RecordTitle""> "
		strDisplayTable = strDisplayTable & "<img src=""" & GetExtentionImage(RS("filename")) & """ border=0 hspace=2>"
		strDisplayTable = strDisplayTable & sDisplayLink & "</a>"

		'else
			'strDisplayTable = strDisplayTable & "<a href="""& RS("vpath") & """ class=""RecordTitle"">"
			'strDisplayTable = strDisplayTable & UCASE(Server.HTMLEncode(RS("DocTitle"))) & "<br><font color=black>Location:</font> " & LCASE(RS("vpath")) & "</a>"
		'end if
		strDisplayTable = strDisplayTable & "</b></td><td>&nbsp;</td><td>&nbsp;</td></tr>"
		strDisplayTable = strDisplayTable & "<tr><td valign=top align=center >&nbsp;&nbsp;&nbsp; " & FormatNumber(NormRank * 100) & "%&nbsp;&nbsp;&nbsp; <!--<IMG SRC=""" & stars & """>--><br>"

    		' Construct the URL for hit highlighting
			'WebHitsQuery = "CiWebHitsFile=" & Server.URLEncode( RS("vpath") )
			'WebHitsQuery = WebHitsQuery & "&CiRestriction=" & Server.URLEncode( Q.Query )
			'WebHitsQuery = WebHitsQuery & "&CiBeginHilite=" & Server.URLEncode( "<strong class=Hit>" )
			'WebHitsQuery = WebHitsQuery & "&CiEndHilite=" & Server.URLEncode( "</strong>" )
			'WebHitsQuery = WebHitsQuery & "&CiUserParam3=" & QueryForm
	        	'WebHitsQuery = WebHitsQuery & "&CiLocale=" & Q.LocaleID
        	strDisplayTable = strDisplayTable & "<!--<a href=""summary.htw?"& WebHitsQuery & """><IMG src=""images/summary.gif""" 
		strDisplayTable = strDisplayTable & "border=0 align=left"
		strDisplayTable = strDisplayTable & "alt=""Highlight matching terms in document using Summary mode."">Summary</a>"
		strDisplayTable = strDisplayTable & "--></td><td valign=top>"
		
		if VarType(RS("characterization")) = 8 and RS("characterization") <> "" then
			strDisplayTable = strDisplayTable & "<b>Abstract:  </b>" & Server.HTMLEncode(RS("characterization")) & "<br>"
        	end if
		
		'strDisplayTable = strDisplayTable & "<br><b>Path: </b><i class=""RecordStats""><a href="""&RS("vpath")&"""  class=""RecordStats"" style=""color:blue;"">http://" 
		'strDisplayTable = strDisplayTable &  Request("server_name") & RS("vpath") & "</a><I><br>"
	        if RS("size") = "" then
			strDisplayTable = strDisplayTable & "(size and time unknown)"
		else
			'strDisplayTable = strDisplayTable & "</td><td valign=top> "& LocalizeDate(RS("write"),session("iUserOffset"))&" </td><td valign=top > "&  RS("size") 
			strDisplayTable = strDisplayTable & "</td><td valign=top>"& RS("write") &" </td><td valign=top > "&  RS("size") 
			strDisplayTable = strDisplayTable & " bytes"
		end if
		strDisplayTable = strDisplayTable & "</td></tr>"

	 RS.MoveNext
	 NextRecordNumber = NextRecordNumber+1
	 Loop
	
	strDisplayTable = strDisplayTable & "</table>"

     else   ' NOT RS.EOF (line 175)
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
	<%
	
	' DISPLAY TOP NAVIGATION CONTROLS
	'Previous Button
	if CurrentPage > 1 and RS.RecordCount <> -1 then %>
		  <A HREF="#" onClick="document.frmPrev.submit();return false;" >
		  <img src='../../images/arrow_back.gif' align='absmiddle' border=0>&nbsp;Prev <%=RS.PageSize%></A>&nbsp;&nbsp;
	<%End If
	
	'Next Button
	if Not RS.EOF then%>
	      <A HREF="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src='../../images/arrow_forward.gif' align='absmiddle' border=0></a>
	<%End If%>
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
	
	' DISPLAY BOTTOM NAVIGATION CONTROLS
	'Previous Button
	if CurrentPage > 1 and RS.RecordCount <> -1 then %>
		  <A HREF="#" onClick="document.frmPrev.submit();return false;" >
		  <img src='../../images/arrow_back.gif' align='absmiddle' border=0>&nbsp;Prev <%=RS.PageSize%></A>&nbsp;&nbsp;
	<%End If
	
	'Next Button
	if Not RS.EOF then%>
	      <A HREF="#" onClick="document.frmNext.submit();return false;" >Next <%=RS.PageSize%>&nbsp;<img src='../../images/arrow_forward.gif' align='absmiddle' border=0></a>
	<%End If%>
	   </td>
	  </tr>
	</table>

	<%
	'if NOT Q.OutOfDate then
	' If the index is current, display the fact 
	'	response.write "<P><I><B>The index is up to date.</B></I><BR>"
	'end if


	if Q.QueryIncomplete then
	'    If the query was not executed because it needed to enumerate to
	'    resolve the query instead of using the index, but AllowEnumeration
	'    was FALSE, let the user know 
	    response.write "<P><I><B>The query is too expensive to complete.</B></I><BR>"
	end if
	
	
	if Q.QueryTimedOut then
	'    If the query took too long to execute (for example, if too much work
	'    was required to resolve the query), let the user know 
	    response.write "<P><I><B>The query took too long to complete.</B></I><BR>"
	end if

	'    This is the "previous" form.
	'    This retrieves the previous page of documents for the query.
	SaveQuery = FALSE
	if CurrentPage > 1 and RS.RecordCount <> -1 then %>
	        <form name="frmPrev" action="<%=QueryForm%>" method="get">
	            <INPUT TYPE="HIDDEN" NAME="qu" VALUE="<%=SearchString%>">
	            <INPUT TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
	            <INPUT TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
	            <INPUT TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage-1%>" >
				<INPUT TYPE="HIDDEN" NAME = "RankBase" VALUE="<%=RankBase%>">
	         </form>
	            <%SaveQuery = TRUE%>
	<% end if %>
	
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



<%
'###############################FUNCTIONS###########################################
Function GetExtentionImage(name)
      url = path & "/" & name
      url = Replace(url, " ", "%20")

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


function LocalizeDate(d,iUserOffset)
   localOffset = 0
 
   If iUserOffset = "" Then
       iUserOffset =0
   End If

  ' Call this function each time you want to display a date
  LocalizeDate= dateAdd("n", localOffset - iUserOffset,d) & sPostDate
end function
%>


<P>&nbsp;</P>


<!--#Include file="../../include_bottom.asp"--> 


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTION AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


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
	SetOrganizationParameters = iReturnValue
	
End Function



'------------------------------------------------------------------------------------------------------------
' FUNCTION GETPAGENAME()
'------------------------------------------------------------------------------------------------------------
Function GetPageName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	For Each arr in strURL 
		sReturnValue = arr 
	Next 
	
	GetPageName = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function


'-------------------------------------------------------------------------------------------------------
' FUNCTION HASACCESS(IORGID,IUSERID,STRVPATH)
'-------------------------------------------------------------------------------------------------------
Function HasAccess(iorgid,iuserid,strvpath)

	  On Error Resume Next

	  iReturnValue = False

	  Set oCnn = Server.CreateObject("ADODB.Connection")
	  oCnn.Open Application("DSN")
	  sSql = "EXEC CHECKFOLDERACCESS '" & iorgid & "','" & iuserid & "','" & strvpath & "'"
	  Set rstAccess = oCnn.Execute(sSql)

	  If NOT rstAccess.EOF Then
		If rstAccess("folderid") >= 0 Then
			iReturnValue = True
		End If
	  End If

	  oCnn.Close
	  Set rstAccess = Nothing
	  Set oCnn = Nothing

	  HasAccess = iReturnValue

End Function
%>