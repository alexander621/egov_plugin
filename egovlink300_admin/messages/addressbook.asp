<!-- #include file="../includes/common.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

Dim sPage, oRst, iCount, iTotal, iPage, sSql, sBook, arrColors(2), index, iNumMoreRecords, iType, sText

sText=Request.QueryString("searchText")

sPage = Request.QueryString("gp")
If sPage & "" <> "" Then
  iPage = clng(sPage)
Else
  iPage = 1
End If

iType= Request.QueryString("searchtype")
If iType & "" <> "" then
  iType=clng(iType)
Else
  iType=0
End If

sSql = "EXEC ListAddressBook " & Session("OrgID") & ",'" & Request.QueryString("searchtext") & "'," & iType & "," & Session("PageSize") & "," & iPage
            
Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open sSql
  .ActiveConnection = Nothing
End With

If Not oRst.EOF then
 
  iNumMoreRecords=oRst("numMorerecords")

  arrColors(0)="ffffff"
  arrColors(1)="eeeeee"
  index=0

  Do While Not oRst.EOF
    If oRst("isGroup")=0 then
      sBook=sBook & "<tr bgcolor='" & arrColors(index) & "'>"
      sBook=sBook & "<td><input type='checkbox' myNameValue='"& oRst("LastName") & "," & oRst("FirstName") &"' name='sendtypeU_" & oRst("UserID") & "' sendValue='to' class='listcheck' onClick='addTo(this,""" & oRst("LastName") & "," & oRst("FirstName") & """)'></td>"
      sBook=sBook & "<td><input type='checkbox' myNameValue='"& oRst("LastName") & "," & oRst("FirstName") &"' name='sendtypeU_" & oRst("UserID") & "' sendValue='cc' class='listcheck' onClick='addCc(this,""" & oRst("LastName") & "," & oRst("FirstName") & """)'></td>" 
      sBook=sBook & "<td>" & oRst("LastName") & ", " & oRst("FirstName") & "</td>"
      sBook=sBook & "</tr>"
    Else
      sBook=sBook & "<tr bgcolor='" & arrColors(index) & "'>"
      sBook=sBook & "<td><input type='checkbox' myNameValue='"& oRst("LastName") &"' name='sendtypeG_" & oRst("UserID") & "' sendValue='to' class='listcheck' onClick='addTo(this,""" & oRst("LastName")  & """,false)'></td>"
      sBook=sBook & "<td><input type='checkbox' myNameValue='"& oRst("LastName") &"' name='sendtypeG_" & oRst("UserID") & "' sendValue='cc' class='listcheck' onClick='addCc(this,""" & oRst("LastName")  & """,false)'></td>" 
      sBook=sBook & "<td><b>" & oRst("LastName")  & "</b></td>"
      sBook=sBook & "</tr>"
    End If
    index= 1-index
    oRst.MoveNext 
  Loop
Else
  sBook="<tr><td colspan=3>" & langNoEntriesFound & "</td></tr>"
End If

iCount = 1
iTotal = oRst.RecordCount

%>
<html>
<head>
  <title><%=langBSAddressBook%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script>
  var pWindow, toText, ccText, listElements;
  var arrTo, arrCc
  
  function init()
  {
    Array.prototype.binarySearch=binarySearch;
    Array.prototype.trimSpaces=ArrayTrimSpaces;
    Array.prototype.toString=ArrayToString;
    pWindow=window.opener;
    toText=pWindow.document.frmMailMessage.To;
    ccText=pWindow.document.frmMailMessage.Cc;
    arrElms=document.userList.elements;
    
    arrTo=toText.value.split(";");
    arrCc=ccText.value.split(";");
    
    arrTo.trimSpaces(true);
    arrCc.trimSpaces(true);
    
    arrTo.sort();
    arrCc.sort();
    
    var elmValue;
    
    for(var i=0; i<arrElms.length; i++)
    {
      if(arrElms[i].type=="checkbox")
      {
        elmValue=arrElms[i].myNameValue;
        //alert(arrTo.binarySearch(elmValue));
        if(arrElms[i].sendValue=="to")
        {   
          if(arrTo.binarySearch(elmValue) > -1)
            arrElms[i].checked=true;
        }
        else if(arrElms[i].sendValue=="cc")
        {
          if(arrCc.binarySearch(elmValue) > -1)
            arrElms[i].checked=true;
        }
      }
    }
  }
  
  function binarySearch(sSearch)
  {
    var hi, lo, mid;
    hi=this.length-1;
    lo=0;
    
    while (hi >= lo)
    {
      mid=parseInt((hi+lo)/2);
      //alert(sSearch + "   " + mid + "  " + hi + "   " + lo + "   " + this[mid]);
      if(this[mid]==sSearch)
        return mid;
      
      if(this[mid] > sSearch)
        hi=mid-1;
      if(this[mid] < sSearch)
        lo=mid+1;
    }
    
    return -1;
  }
  
  function ArrayTrimSpaces(bAllSpaces)
  {
    for(var i=0; i<this.length; i++)
    {
      if(bAllSpaces)
      {
        this[i]=this[i].replace(/ /g, "");
      }    
      while(this[i].indexOf(" ") == 0)
        this[i]=this[i].substr(1);
        
      //alert("---" + this[i] + "---");
    }
  }
  
  function ArrayToString()
  {
    var sReturn="";
    for(var i=0; i<this.length; i++)
      if(this[i] != "")
        sReturn+=this[i] + "; ";
        
    return sReturn;
  }
  
  
  function addTo(obj,name)
  {
    var index;
    if(addTo.arguments.length ==2)
      name=name.replace(/ /g, "");
    
    if(obj.checked)  
    {
      toText.value+=name + "; ";
    }
    else
    {
    arrTo=toText.value.split(";");
    arrTo.trimSpaces();
    arrTo.sort();
    index=arrTo.binarySearch(name);
    arrTo.splice(index,1);
    toText.value=arrTo;
    }
      
  }
  
  function addCc(obj,name)
  {
    if(addCc.arguments.length ==2)
      name=name.replace(/ /g, "");
    
    if(obj.checked)
    {    
      ccText.value+=name + "; ";
    }
    else
    {
    arrCc=ccText.value.split(";");
    arrCc.trimSpaces();
    arrCc.sort();
    index=arrCc.binarySearch(name);
    arrCc.splice(index,1);
    ccText.value=arrCc;
    }
  }
  
  function testScript()
  { 
    alert(pWindow); 
  }
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" marginwidth="0" marginheight="0" onload="init();">
  <form name="frmSearch" method="get" action="addressbook.asp">
  <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
    <tr bgcolor="#93bee1">

      <td colspan=3 ><%=langSearchAddressBook%>:<br>
      <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" name="searchtext" value="<%=sText%>"><input type=button value="&nbsp;Go&nbsp;" onClick="document.frmSearch.submit()"><br>
      <%=langUsers%><input type="radio" name="searchtype" value="1" class="listcheck" <%if iType = 1 then Response.Write("CHECKED")%>> <%=langGroups%><input type="radio" name="searchtype" value="2" class="listcheck" <%if iType = 2 then Response.Write("CHECKED")%>> <%=langAll%><input type="radio" name="searchtype" value="0" class="listcheck" <%if iType < 1 then Response.Write("CHECKED")%>>
    </tr>
    </form>
    <tr bgcolor="#93bee1">
      <td colspan=3><img src="../images/arrow_back.gif" align="absmiddle">
      <%
          If IPage > 1 Then
            Response.Write "<a href=""addressbook.asp?gp=" & iPage-1 & "&searchText=" & Request.QueryString("searchText") & "&searchType=" & iType &  """>" & langPrev & " " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">" & langPrev & " " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""addressbook.asp?gp=" & iPage+1 & "&searchText=" & Request.QueryString("searchText") & "&searchType=" & iType &  """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">" & langNext & " " & Session("PageSize") & "</font>"
          End If
       %>
       <img src="../images/arrow_forward.gif" align="absmiddle" onClick="testScript();">
      </td>
    </tr>
    <tr bgcolor="#93bee1">
      <td style="color:#003366;"><b><%=langTo%></b></td>
      <td style="color:#003366;"><b><%=langCc%></b></td>
      <td style="color:#003366;" width="100%"><b><%=langName%></b></td>
    </tr>
    <form name="userList">
    <%=sBook%>
    </form>
  </table>
</body>
</html>
<%
if oRst.state=1 then oRst.Close
%>
