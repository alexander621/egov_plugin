<!--#include file='header.asp'-->
  <table border="0" cellpadding="10" cellspacing="0" width="100%"  class="start" >
    <tr height=500>
      <td valign="top" width='151' height=500>
		 <center> <img src='../images/icon_directory.jpg'></center>
		 <br>
	 <br>
	       <!--#include file='quicklink.asp'-->   
      </td>
      <td colspan="2" valign="top">
<%
response.write "<table><tr>"& _
"<td><font size='+1'><b>Search </b></font><br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='display_committee.asp'>"&langBackToCommittee&"</td>"& _
 "</tr></table><br><br>"


const langNoRecords="No records"
Dim objSearch, myarray,myfields,leftstr
Set objSearch = new Search
objSearch.searchString="k"
objSearch.b_SearchFieldsSwitch=array(0,0,1,0,0, 0,0,0,1,0, 1,1,1,1,1, 1,1,0,0,0, 0,0,1,1,1, 0)
objSearch.searchTable="users"
objSearch.AndOR="and"
objSearch.PageSize=20

myArray=objSearch.GetKeywordsArray
myfields=objSearch.GetTableFields
objSearch.BuildQueryString
objSearch.DisplayResults
%>



<%
class Search
'--------------------
public SearchString
public SearchTable
public b_SearchFieldsSwitch
public AndOR
public PageSize
'--------------------
private OrderedKeyString
private QueryString
private KeywordsArray
private conn
private MaxFields
private TableFields
private objRS 
'------------------------------
'--------------------------------
	public sub Class_Initialize()	
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Open Application("DSN")
	Set objRS = Server.CreateObject("ADODB.Recordset")
	set objRS.ActiveConnection = conn
	objRS.CursorLocation = 3 ' the CursorLocation and CursorType must be defined firstly, otherwise, the
	objRS.CursorType = 3     ' Recordcount will be -1
	OrderedKeyString=""
	QueryString=""
	end sub

	public sub Class_Terminate()
	Destroy conn
	end sub

  public function GetTableFields
		strSQL = "SELECT syscolumns.name, syscolumns.type, syscolumns.length, " & _
         "syscolumns.isnullable FROM sysobjects " & _
         "INNER JOIN syscolumns ON sysobjects.id = syscolumns.id " & _
         "where sysobjects.name = '"&SearchTable& _
         "' ORDER BY syscolumns.colid"
	objRS.Open strSQL
	MaxFields=objRS.recordcount
	redim TableFields(MaxFields-1)
	for i=0 to MaxFields-1
	TableFields(i)=objRS("name")
	objRS.MoveNext
	next
	objRS.close
	GetTableFields=TableFields
	end function
'---------------------------------
public function  GetKeywordsArray
SearchString=replace(SearchString,","," ")
SearchString=replace(SearchString," and "," ")
OrderedKeyString=GetOrderedKeyString(SearchString)
KeywordsArray=split(trim(OrderedKeyString),"|")
GetKeywordsArray= KeywordsArray
end function

'===============================================================
public function BuildQueryString
StrSQL1 ="SELECT  * FROM ["&SearchTable&"] "
StrSQL1 = StrSql1 & " WHERE ("  '( is embrace the search keywords 
'--------------------------------------------
For Each word in keywordsArray
if word<>"" then
	if cnt=0 then linksymbol="" else linksymbol=AndOR
'------------------------
	strEachWordConditions=""
for i=0 to MaxFields-1
  if b_SearchFieldsSwitch(i)=1 then
	 temp=TableFields(i)&" Like '%"&word&"%'"
	strEachWordConditions=temp+" or "+strEachWordConditions
  end if
next
strEachWordConditions=left(strEachWordConditions,len(strEachWordConditions)-3) ' remove the last 3 letter or
'response.write "<br>each=="&strEachWordConditions
'------------------------
	StrSQL1 =StrSQL1 &" "&linksymbol&" ("&strEachWordConditions&") " 
	cnt=cnt+1
end if
next 
'---------------------------------------------
StrSQL1 = StrSql1 & " ) " ') is end of embr
QueryString=StrSQL1
'response.write "<hr>"&QueryString
BuildQueryString=QueryString
end function
'===============================================================
public sub DisplayResults
if  QueryString="" then exit sub
	objRS.PageSize=PageSize
	objRS.Open  QueryString
	if request.querystring("currentpage")<>"" then objRS.AbsolutePage=clng(request.querystring("currentpage")) else objRS.AbsolutePage=1
if objRS.recordcount=0 then 
CustomerizedTable 0
exit sub
else
CustomerizedTable 1
end if
end sub
'===============================================================
public sub CustomerizedTable(record)
call navagatorbar
response.write "<table border=0 cellpadding=5 cellspacing=0  width=650  align=center class='tablelist'>"
response.write "<tr align=left><th width=130>Full Name</th><th>Email</th><th>Matches</th>"
if record=0 then
response.write "<tr colspan=3><td>"&langNoRecords&"</td></tr>"
else
if objRS.PageCount=1 then int_Record=objRS.recordcount else int_Record=PageSize
for j=1 to int_Record
strMatches=""
'------------------------------------------
for i=0 to MaxFields-1
if b_SearchFieldsSwitch(i)=1 then 
'---- deal with the searchable field--
	temp=objRS(TableFields(i))
	b_replace=false
		for each keyword in KeywordsArray
		 start=instr(1,temp,keyword,1)
			if start then
				b_replace=true
				takeplace=mid(objRS(TableFields(i)),start,len(keyword))
				temp=replace(temp,takeplace,"<FONT COLOR=#330066><B>"&takeplace&"</B></FONT>")
			end if
		next
   if b_replace  then strMatches=strMatches&" - <I>"&TableFields(i)&"</I>:"&temp
'-------------------------------------
end if
next
'------------------------------------------
strMatches=Right(strMatches,len(strMatches)-2)
response.write "<tr align=left>"
response.write "<td><A HREF='display_individual.asp?userid="&objRS("userid")&"'>"&trim(objRS("lastname"))&", "&trim(objRS("firstname"))&"</A></td><td>"&objRS("email")&"</td><td>"&strMatches&"</td>"
response.write "</tr>"
objRS.movenext
next
end if
response.write "</table>"
end sub
'===============================================================
sub  navagatorbar
	thisname=request.servervariables("script_name")
	currentpage=objRS.AbsolutePage	
	response.write "<div style='font-size:10px; padding-bottom:5px;'>"
	 if currentpage>1 then 
		response.write "<A HREF='"&thisname&"?currentpage="&(currentpage-1)&"'>"
		else
		response.write "<!A HREF='"&thisname&"?currentpage="&(currentpage-1)&"'>"
		end if
		response.write "<img src='../images/arrow_back.gif' align='absmiddle' border=0></A>"& _
		"<font color=#999999>"&langPrev&"&nbsp;"&PageSize&"</font>"& _
		"&nbsp;&nbsp;"& _
		"<font color=#999999>"&langNext&" "&PageSize&"</font>" 
'		response.write "<br>currentpage="&currentpage&"  totalpages="&TotalPages
		if currentpage<objRS.PageCount then 
		response.write "<A HREF='"&thisname&"?currentpage="&(currentpage+1)&"'>"
		else
		response.write "<!A HREF='"&thisname&"?currentpage="&(currentpage+1)&"'>"
		end if
	    response.write 	"<img src='../images/arrow_forward.gif' align='absmiddle' border=0></a>"
end sub

end class
%>

<%



function GetOrderedKeyString(keywords)
leftstr=""
rightstr=""
'response.write "<br>"&keywords
FirstMaohao=instr(keywords,Chr(34))
if FirstMaohao>0 then
'-- if the string contain " ---
'response.write "<br><br><br>First hyphen inside:"&FirstMaohao
right_keywords=right(keywords,len(keywords)-FirstMaohao)
SecondMaohao=instr(right_keywords,Chr(34))
'============ judge if the second Mao Hao exist ===================
if SecondMaohao>0 then
'response.write "<br>second hyphen inside:"&SecondMaohao
LeftStr=leftstr+"|"+mid(keywords,FirstMaohao+1,SecondMaohao-1)
    temp=left(keywords,FirstMaohao-1)
RightStr=right(right_keywords,len(right_keywords)-secondMaohao)+" "+temp
'response.write "<hr>1  left:"&LeftStr&" right:"&RightStr
RightStr2=GetOrderedKeyString(RightStr)
'response.write "<hr>2  left:"&LeftStr&" right:"&RightStr&" right:"&RightStr2
KeyArrayString="|"+LeftStr+"|"+replace(RightStr," ","|")
else
'-- if the string contain only single "---
keywords=replace(keywords,Chr(34),"")
KeyArrayString =replace(keywords," ","|")
'-- end of if the contain  single "---
end if
'============= end of judge ==========================================
else
KeyArrayString = replace(keywords," ","|")
'response.write "<br>*1"&SearchWords
end if
while instr(KeyArrayString,"||") 
KeyArrayString=replace(KeyArrayString,"||","|")
wend
GetOrderedKeyString=KeyArrayString
end function

sub Destroy(obj)
		'A generic object desctruction function
		on error resume next
		select case lcase(typename(obj))
			case "recordset", "connection"
				obj.close
			case "variant()"	'array
				erase obj
		end select
		set obj = nothing
	end sub	
%>
</td>
  <td width='200'>&nbsp;</td>
    </tr>
 </table>
 <!--#include file='footer.asp'-->
