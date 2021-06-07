<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
Response.Buffer = True
%>

<%
' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Action Line"
%>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>



<link href="search.css" rel="stylesheet">
<link rel="stylesheet" href="../css/styles.css" type="text/css">
<link href="../global.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>

<%
  PageSize = 5
  SiteLocale = "EN-US"
  Dim strDisplayTable, scopePath 

  'Set Initial Conditions
  NewQuery = FALSE
  UseSavedQuery = FALSE
  SearchString = ""
  QueryForm = Request.ServerVariables("PATH_INFO")
  
  'Did the user press a SUBMIT button to execute the form? If so get the form variables.
  If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  
	SearchString = Request("SearchString")
    FreeText = Request.Form("FreeText")
    'NOTE: this will be true only if the button is actually pushed.
    If Request.Form("Action") = "Go" Then
      NewQuery = TRUE
			RankBase=1000
    End If

  Elseif session("strHomeSearch") <> "" Then
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
</head>
<!--#Include file="../include_top.asp"-->

  <!--Begin Page Header -->
  <div style="padding-left:22px; padding-top:9px;"><b><font style="font-family:Verdana,Arial; font-size:18px;">Site Search</font></b></div>
  <!--End Page Header -->
  
  <!--Begin Search Form -->
    <form name="frmSearch" action="<%=QueryForm%>" method=post>
    <input type=hidden name="Action" value="Go">
    <table width="95%" align=center>
      <tr><td><b>Enter your query below:</b></td></tr>
       <tr><td><input type=text name="SearchString" size=65 maxlength=100 value="<%= SearchString %>" style="background-color:#eeeeee;width:255px; height:19px; border:1px solid #000033;"></td></tr>
	   
	   <tr>
	   <TD ALIGN=RIGHT WIDTH=255><a href="#" onClick='document.frmSearch.submit()'><img src="images/go.gif" border="0"><%=langGo%></a>
         </form></TD>
       </tr>
    </table>
  </form>
  <!--End Search Form-->
  <% if request("searchstring") <> "" then %>
  	<center>
  	<table width=90%>
		<tr>
			<td>
				<center>
				<div class=subcategorymenu><a href=#web>Web Search</a> | <a href=#doc>Document Search</a> | <a href=#cal>Calendar Search</a></div><br>
				</center>
				<a name=web></a>
				<fieldset>
  					<legend>Web Search</legend>
					<!--#include file="inc_websearch.asp"-->
				</fieldset>
				<br>
				<center>
				<div class=subcategorymenu><a href=#web>Web Search</a> | <a href=#doc>Document Search</a> | <a href=#cal>Calendar Search</a></div><br>
				</center>
				<a name=doc></a>
				<fieldset>
  					<legend>Document Search</legend>
					<!--#include file="inc_docsearch.asp"-->
				</fieldset>
				<br>
				<center>
				<div class=subcategorymenu><a href=#web>Web Search</a> | <a href=#doc>Document Search</a> | <a href=#cal>Calendar Search</a></div><br>
				</center>
				<a name=cal></a>
				<fieldset>
  					<legend>Calendar Search</legend>
					<!--#include file="inc_calsearch.asp"-->
				</fieldset>
			</td>
		</tr>
	</table>
	</center>
	<% end if %>
	<br>
	<br>
	<br>
	<br>
<!--#Include file="../include_bottom.asp"-->

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
