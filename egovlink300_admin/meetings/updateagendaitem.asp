
<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, oRst,intAgendaID, index, arrColors(2), oRstItem, sName, sLink
Dim sTypes, iType, intItemID, sLinkShown, iItemID, arrItemID(1), smid

intItemID=Request.QueryString("id")
smid = Request.QueryString("mid")

If Request.Form("_task") = "updateItem" Then

  iItemID=Request.Form("itemID")
  smid=request.Form("mid")

  Set oCmd = Server.CreateObject("ADODB.Command")

  If clng(Request.Form("type"))=ITEM_TYPE_DOCUMENT then
    With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "GetDocumentIDbyPath"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("Path", adVarChar, adParamInput, 500, iItemID)
    End With
    
    set oRst=oCmd.Execute
    
    if not oRst.eof then
      iItemID=oRst(0)
      oRst.close
    end if
    set oRst=nothing
    oCmd.Parameters.Delete("Path")
    oCmd.Parameters.Delete("OrgID")
  End If
  
  'If the type is text we need the AgendaItemURL field in the database to be NULL
  If clng(Request.Form("type")) = ITEM_TYPE_TEXT then iItemID= null End If
  
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdateAgendaItem"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, Request.Form("AgendaItemID"))
    .Parameters.Append oCmd.CreateParameter("AgendaItemTypeID", adInteger, adParamInput, 4, Request.Form("type"))
    .Parameters.Append oCmd.CreateParameter("AgendaItemURL", adInteger, adParamInput, 4, iItemID)
    .Parameters.Append oCmd.CreateParameter("AgendaItemTitle", adVarChar, adParamInput, 250, Request.Form("title"))
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../meetings/edit_agendaitem.asp?aid=" & Request.Form("AgendaID") & "&mid=" & smid

End If

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListAgendaItemTypes"
  .CommandType = adCmdStoredProc
End With

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With

oCmd.CommandText = "GetAgendaItem"
oCmd.Parameters.Append oCmd.CreateParameter("ItemID", adInteger, adParamInput, 4, intItemID)

Set oRstItem = Server.CreateObject("ADODB.Recordset")
With oRstItem
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing

If not oRstItem.EOF then
  intAgendaID=oRstItem("AgendaID")  
  sName=oRstItem("AgendaItemTitle")
  sLink=oRstItem("AgendaItemURL")
  iType=oRstItem("AgendaItemTypeID")
  
  if iType = ITEM_TYPE_DOCUMENT then sLink=oRstItem("DocumentURL")
End If

Do while not oRst.EOF
  sTypes = sTypes & "<option value='" & oRst("AgendaItemTypeID") & "'"
  If oRst("AgendaItemTypeID") = iType then sTypes = sTypes & " SELECTED "
  sTypes = sTypes & ">" & oRst("AgendaItemTypeName") & "</option>"
  oRst.MoveNext
Loop

%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <script>
  function init(type)
  {
    if(type==4)
    {
      hide("linkText");
      hide("linkField");
    }
  }
  function selectLink()
  {
    var index=document.addItem.type.selectedIndex;
    var selValue=document.addItem.type.options[index].value
    
    if (selValue==1)
    {
      window.open('picker/', 'filepicker', 'width=506,height=345,scrollbars=0,toolbars=0,statusbar=0,menubar=0,left=265,top=180');
    }
    if (selValue==2)
    {
      window.open ('pollSelect/listPolls.asp','', 'width=700,height=345,scrollbars=1,toolbars=0,statusbar=0,menubar=0,left=265,top=180')
    }
    if (selValue==3)
    {
    window.open ('discSelect/listBoards2.asp','', 'width=700,height=300,scrollbars=1,toolbars=0,statusbar=0,menubar=0,left=265,top=180')
    }
  }
  
  function verifySubmit()
  {
    var index=document.addItem.type.selectedIndex;
    var selValue=document.addItem.type.options[index].value;
  
    if((document.addItem.title.value != '' && selValue==4) || (document.addItem.link.value != '' && document.addItem.title.value != ''))
    {
      if(selValue==4)document.addItem.link.value=document.addItem.title.value;
      if(document.addItem.itemID.value =='')document.addItem.itemID.value=document.addItem.link.value;
      document.addItem.submit();
    }
    else
    {
    alert("You must fill in all fields.");
    }
  }
  function eraseLink()
  {
    document.addItem.link.value='';
    document.addItem.itemID.value='';
    
    var index=document.addItem.type.selectedIndex;
    var selValue=document.addItem.type.options[index].value
    
    if(selValue==4)
    {
      hide("linkField");
      hide("linkText");
    }
    else
    {
      show("linkField");
      show("linkText");
    }
  }
  function hide(id)
  {
    var obj=eval("document.all." + id + ".style");
    obj.visibility="hidden";
  }
  function show(id)
  {
    var obj=eval("document.all." + id + ".style");
    obj.visibility="visible";
  }
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="init('<%=iType%>');">
  <%
  DrawTabs tabMeetings,1
  %>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b>Agenda Items</b></font></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" nowrap>

		<% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
        <form name="addItem" method=post action="updateagendaitem.asp" method="post">
        <input type="hidden" name="_task" value="updateItem">
        <input type=hidden name="AgendaID" value=<%=intAgendaID%>>
        <input type=hidden name="AgendaItemID" value=<%=intItemID%>>
        <input type=hidden name="itemID" value ="<%=sLink%>">
        <input type=hidden name="mid" value ="<%=smid%>">
		<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();"><%=langUpdate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th colspan=2 align=left>Agenda Item</th>
            </tr>
            <tr>
              <td>Agenda Type:</td>
              <td><select name=type onChange="eraseLink();"><%=sTypes%></select></td>
            </tr>
            <tr>
              <td><%=langDescription%>:</td>
              <td><input type="text" name="title" value="<%=sName%>"></td>
            </tr>
            <tr>
              <td><div id="linkText">Select Link:</div></td>
              <td><div id="linkField"><input type=text name="link" value="<%=sName%>" READONLY><input type=button name="btnLink" value="Select Link" onClick="selectLink();"></div></td>
            </tr>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
