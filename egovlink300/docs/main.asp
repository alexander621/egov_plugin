<!-- #include file="../includes/common.asp" //-->
<%
Dim path, strDate, strTimeSearch, index, iCount, intNumberChanged, sSQL,oRST,subpath
Dim blnSubFound,objSubDir,url,imgsrc,pos,temp,strPrefix,strDocument,strMsg
Dim blnFoundDocuments

' SET VARIABLES
session("strResults") = ""
path = Application("eCapture_AppPath") & "pub"
index = 0
intNumberChanged = 0
iCount = 0
blnFoundDocuments = False


' GET DATE TO SEARCH
If Request.QueryString("view") = "" Then
	strDate = "today"
Else
	strDate = Request.QueryString("view")
End If

' GET DOCUMENTS SPECIFIED IN THE TIME PERIOD
strTimeSearch = IndexSearch("john",strDate)
intNumberChanged = iCount

' BUILD DISPLAY MESSAGE TO USER
' PLURAL OR SINGULAR DOCUMENT VALUE
If intNumberChanged = 1 Then
	strPrefix = langDateSearchPrefixSingular
	strDocument = langDocument
Else
	strPrefix = langDateSearchPrefixPlural
	strDocument = lcase(langDocuments)
End If

' BUILD DISPLAY MESSAGE LINE
Select Case strDate
	Case "today"
		strMsg = strPrefix & " " & intNumberChanged & " " &  langNewDocTop & " " & strDocument & " " & langtoday & "."
	Case "week"
		strMsg = strPrefix & " " & intNumberChanged & " " &  langNewDocTop & " " & strDocument & " " & langSinceLastWeek & "."
	Case "2weeks"
		strMsg = strPrefix & " " & intNumberChanged & " " &  langNewDocTop & " " & strDocument & " " & langSinceLast2Weeks & "."
	Case "lastmonth"
		strMsg = strPrefix & " " & intNumberChanged & " " &  langNewDocTop & " " & strDocument & " " & langSinceLastMonth & "."
End Select

%>


<html>
<head>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <link href="search/search.css" rel="stylesheet">
  <link href="../css/styles.css" rel="stylesheet">
</head>

<body bgcolor="#ffffff" topmargin="11" marginheight="0" leftmargin="10">
<p class=title>The City of Loveland is happy to provide you with the documentation you need. Select a category from the list at the left.


<!--
  <div style="padding-top:13px;"><font style="font-family:Verdana,Arial; font-size:18px;"><b><%=langTabDocuments%></b></font><br><font style="font-family:Verdana,Arial; font-size:11px;"><%=strMsg%></font></div>
  <br>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
   	<tr>
      <td>
   	    <table border="0" cellpadding="0" cellspacing="0" width="100%">
   	      <tr>
   	        <td width="45"><img src="menu/images/spacer.gif" width="45" height="13"></td>
   	        <td width="100%"></td>
   	      </tr>
          <tr>
            <td colspan="3" height=24 class="ecapture">&nbsp;<%=langSearchDocuments%></td>
            <td valign=bottom></td>
          </tr>
          <tr>
            <td><img src="images/spacer.gif" width="1" height="5"></td>
          </tr>
          <tr>
            <td></td>
            <td colspan="3">
              <form action="search/search.asp" method=post name="frmSearch">
                <input type="hidden" name="Action" value="Go">
                <input type=text name="SearchString" style="background-color:#eeeeee;width:255px; height:19px; border:1px solid #000033;">
            </td>
          <tr>
			  <TD></TD>	<TD ALIGN=RIGHT WIDTH=255><a href="#" onClick='document.frmSearch.submit()'><img src="../images/go.gif" border="0"><%=langGo%></a>
              </form></TD>
          </tr>
          </tr>
       	  <tr>
      	    <td><img src="images/spacer.gif" width=1 height=15></td>
      	  </tr>
      	  <tr style="height:26px;">
         		<form name="frmView" action="main.asp" method=get>
         		<td colspan="3" height=24 class='ecapture' nowrap width="100%">&nbsp;<%=langViewNewDocs%>&nbsp;
           		<select name="view" onchange="document.frmView.submit();" style="font-family:Verdana,Tahoma,Arial;font-size:10px;background-color:#ffffff;border:1px solid #000033;">
           		  <option value="today" <% If Request.QueryString("view") = "" OR Request.QueryString("view") = "today" Then Response.Write "SELECTED" %>><%=langAddedToday%>
           		  <option value="week" <% If Request.QueryString("view") = "week" Then Response.Write "SELECTED" %>><%=langAddedLastWeek%>
           		  <option value="2weeks" <% If Request.QueryString("view") = "2weeks" Then Response.Write "SELECTED" %>><%=langAddedLastTwo%>
           		  <option value="lastmonth" <% If Request.QueryString("view") = "lastmonth" Then Response.Write "SELECTED" %>><%=langAddedLastMonth%>
           		</select>
         		</td>
         		</form>
         		<td valign=bottom><!-<font color="#000033">_____________________</font>-></td>
         	</tr>
         	<tr>
         		
            <td colspan="3">
         		  <%If blnFoundDocuments = True Then %>
         		  <table border=0 cellpadding=5 cellspacing=0 class='subtablelist' width="100%">
         		  <tr style="height:26px;"><th class=subheading width="1%">Rank</th><th align="left">Document Information</th><th align="left">Last Modified</th><th align="left">File Size</th></tr>
         			<%=strTimeSearch %>
         		  </table>
         		  <%Else%>
         		  <p><br><%=langEmptyDateSearch %></p>
         		  <%End If%>
         		  </td></tr></table>
         		</td>
         	</tr>
         	

          <%
          If Session("Admin") = True Then
          %>
            <tr>
              <td width=30><img src="images/spacer.gif" width=1 height=15></td>
            </tr>
            <tr>
              <td colspan="3" bgcolor="#eeeeee" height=22 nowrap style="border:1px solid #000033; font-family:Arial,Tahoma; font-size:10px;">&nbsp;Contribute</td>
              <td valign=bottom><font color="#000033">_____________________</font></td>
            </tr>
            <tr>
              <td width=30><img src="images/spacer.gif" width=1 height=5></td>
            </tr>
            <tr>
              <td></td>
              <td colspan="3">
             
              </td>
            </tr>
          <%
          End If
          %>
        </table>
      </td>
    </tr>
  </table>
-->  

  
</body>
</html>

<%
'********************FUNCTIONS*******************************************************************
Function MapURL(path)
     dim rootPath, url
     'Convert a physical file path to a URL for hypertext links.
     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url, "\", "/")
End Function 


Function GetExtentionImage(name)
      url = path & "/" & name
      url = Replace(url, " ", "%20")

      imgSrc = "search/images/document.gif"

      pos = InStr(1, name, ".")
      If pos > 0 Then
        Select Case Mid(name,pos+1,3)
          Case "doc"
            imgSrc = "search/images/msword.gif"
          Case "xls"
            imgSrc = "search/images/msexcel.gif"
          Case "ppt"
            imgSrc = "search/images/msppt.gif"
          Case "htm"
            imgSrc = "search/images/msie.gif"
          Case "pdf"
            imgSrc = "search/images/pdf.gif"
        End Select
        temp = Left(name,pos-1)
      Else
        temp = name
      End If
      
      GetExtentionImage = imgSrc
End Function

Private Function RemoveExtension(name)
	pos = InStr(1, name, ".")
	temp = Left(name,pos-1)
	RemoveExtension = temp
End Function

Function IndexSearch(strValue, intRange)

' CHECK DATABASE FOR FOLDER PERMISSIONS
   sSql = "EXEC ListSearchFolders " & Application("OrgID") & ", " & Session("UserID") & ", '" & path & "'"
   Set oRst = Server.CreateObject("ADODB.Recordset")
   
  ' response.write sSQL & "<<Here"
  'response.end
      
   oRst.Open sSql, Application("DSN"), 3, 1

   

  FormScope = Application("eCapture_ArticlesPath")
  PageSize = session("PageSize")
  SiteLocale = "EN-US"
  Dim strDisplayTable, scopePath 

  'Set Initial Conditions
  NewQuery = FALSE
  UseSavedQuery = FALSE
  SearchString = ""
  QueryForm = Request.ServerVariables("PATH_INFO")
  SearchString = strValue
  RankBase = 1000
    
    Set Q = Server.CreateObject("ixsso.Query")
    Set util = Server.CreateObject("ixsso.Util")
      mydate = Date()
      
	  Select Case intRange
	    Case "week"
	    ' MATCH FILES ADDED FOR LAST WEEK
			  dtTimeFrame= DateAdd("ww",-1,mydate)
	    Case "today"
        ' MATCH FILES ADDED TODAY
        'dtTimeFrame = DateAdd("h",-24,mydate) returns last 24 hours
        dtTimeFrame = DateAdd("h",-(DatePart("h",myDate)),myDate)
        
        Case "2weeks"
	    ' MATCH FILES ADDED FOR LAST WEEK
		dtTimeFrame= DateAdd("ww",-2,mydate)
		
		Case "lastmonth"
	    ' MATCH FILES ADDED FOR LAST WEEK
		 dtTimeFrame = DateAdd("m",-1,mydate)
	    End Select
 
		dtTimeFrame = GetDate(dtTimeFrame)
          
    'Search string searches for filename and contents
	CompSearch = "@write > "& dtTimeFrame
	Q.Query = CompSearch
    Q.SortBy = "rank[d]"
    Q.Columns = "DocTitle, vpath, filename, size, write, characterization, rank"
	  Q.MaxRecords = 300

    If FormScope <> "/" Then
      ' Security loop to filter records
	 
	 If Not oRst.EOF Then
		Do While Not oRst.EOF
			scopePath = oRst("FolderPath")
			If scopePath <> FormScope Then
				util.AddScopeToQuery Q, scopePath, "shallow"
				'DEBUG CODE:
				'response.write scopePath & "path<br>"
			End If
	        oRst.MoveNext
		Loop
	 End If
    End If
    
		
	If SiteLocale <> "" Then
      Q.LocaleID = util.ISOToLocaleID(SiteLocale)
    End If
    
	Q.Catalog = "ecTeamlinkDemo"
    Set RS = Q.CreateRecordSet("nonsequential")

    RS.PageSize = PageSize
    ActiveQuery = TRUE

  If ActiveQuery Then
    If Not RS.EOF Then
    NextRecordNumber = 1

    ' BEGIN BUILDING STRING THAT WILL BE THE DISPLAY TABLE
		If  Not RS.EOF and NextRecordNumber <= LastRecordOnPage then
  			strDisplayTable = "<table border=0 class=tablelist cellspacing=0 cellpadding=0 width=""95%"" align=center>"
			strDisplayTable = strDisplayTable & "<colgroup width=105>"
			strDisplayTable = strDisplayTable & "<tr><th>Rank</th><th>Document Details</th></tr>"
		end if

	 Do While Not RS.EOF 
     
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
			stars = "search/images/rankbtn5.gif"
		elseif NormRank > 0.60 then
			stars = "searchimages/rankbtn4.gif"
		elseif NormRank > 0.40 then
			stars = "search/images/rankbtn3.gif"
		elseif NormRank >.20 then
			stars = "search/images/rankbtn2.gif"
		else stars = "search/images/rankbtn1.gif"
 	  end if


    ' BEGIN BUILDING STRING THAT CONTAINS ROW DATA
	strDisplayTable = strDisplayTable & "<tr class=""RecordTitle"">"
	strDisplayTable = strDisplayTable & "<td align=""right"" valign=middle class=""RecordTitle"">"
	strDisplayTable = strDisplayTable & NextRecordNumber & ". &nbsp;&nbsp;"
	strDisplayTable = strDisplayTable & "</td>"
	strDisplayTable = strDisplayTable & "<td valign=middle><b class=""RecordTitle"">"
	
	if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then
		strDisplayTable = strDisplayTable & "<a href=""" & RS("vpath") & """ class=""RecordTitle""> "
		'strDisplayTable = strDisplayTable & "<img src=""" & GetExtentionImage(RS("filename")) & """ border=0 >"
		strDisplayTable = strDisplayTable & Server.HTMLEncode( RS("filename")) & "</a>"
	else
		strDisplayTable = strDisplayTable & "<a href="""& RS("vpath") & """ class=""RecordTitle"">"
		strDisplayTable = strDisplayTable & Server.HTMLEncode(RS("DocTitle")) & "</a>"
	end if
		strDisplayTable = strDisplayTable & "</b></td><td>&nbsp;</td><td>&nbsp;</td></tr>"
	    strDisplayTable = strDisplayTable & "<tr><td valign=top align=left><!--<IMG SRC=""" & stars & """><br>"

    ' Construct the URL for hit highlighting
			WebHitsQuery = "CiWebHitsFile=" & Server.URLEncode( RS("vpath") )
			WebHitsQuery = WebHitsQuery & "&CiRestriction=" & Server.URLEncode( Q.Query )
			WebHitsQuery = WebHitsQuery & "&CiBeginHilite=" & Server.URLEncode( "<strong class=Hit>" )
			WebHitsQuery = WebHitsQuery & "&CiEndHilite=" & Server.URLEncode( "</strong>" )
			WebHitsQuery = WebHitsQuery & "&CiUserParam3=" & QueryForm
	        'WebHitsQuery = WebHitsQuery & "&CiLocale=" & Q.LocaleID
        strDisplayTable = strDisplayTable & "<a href=""search/summary.htw?"& WebHitsQuery & """><IMG src=""search/images/summary.gif""" 
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
			strDisplayTable = strDisplayTable & "</td><td valign=top> "& RS("write") & " GMT</td><td valign=top>" & RS("size") & " bytes" 
			
		end if
			strDisplayTable = strDisplayTable & "</i></td></tr>"

RS.MoveNext
NextRecordNumber = NextRecordNumber+1
iCount = iCount + 1
Loop


strDisplayTable = strDisplayTable & "</table>"
blnFoundDocuments = True

else   ' NOT RS.EOF
	if NextRecordNumber = 1 then
          strDisplayTable = "No documents matched the query<P>"
    else
          strDisplayTable = "No more documents in the query<P>"
    end if

end if ' NOT RS.EOF

end if
IndexSearch = strDisplayTable

End Function

Function GetDate(sDate)
  sDate = sDate
  sMonth  = Month(sDate)
  sDay    = Day(sDate)
  sYear   = cStr(Year(sDate))
  sYear   = Right(sYear,2)
  
  If Len(sMonth) < 2 Then sMonth = "0" & sMonth
  If Len(sDay) < 2 Then sDay = "0" & sDay

  GetDate = sYear & "/" & sMonth & "/" & sDay 
End Function
%>
