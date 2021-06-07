
<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, oRst,intAgendaID, index, arrColors(2), arrItemID(1), iItemID, sTypes, smid
intAgendaID=Request.Form("AgendaID")
if intAgendaID & "" ="" then intAgendaID=Request.QueryString("aid")
' Need to preserve the meeting id to pass back to the edit_agendaitem.asp
smid = Request.Form("mid")
if smid & "" = "" then smid = Request.QueryString("mid")
'
If Request.Form("_task") = "newItem" Then
  
  iItemID=Request.Form("itemID")

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
    .CommandText = "NewAgendaItem"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, intAgendaID)
    .Parameters.Append oCmd.CreateParameter("AgendaItemTypeID", adInteger, adParamInput, 4, Request.Form("type"))
    .Parameters.Append oCmd.CreateParameter("AgendaItemURL", adInteger, adParamInput, 4, iItemID)
    .Parameters.Append oCmd.CreateParameter("AgendaItemTitle", adVarChar, adParamInput, 250, Request.Form("title"))
    .Execute
  End With
  Set oCmd = Nothing

  Response.Redirect "../meetings/edit_agendaitem.asp?aid=" & intAgendaID & "&mid=" & smid 

End If

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "ListAgendaItemTypes"
  .CommandType = adCmdStoredProc
  .Execute
End With

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With
Set oCmd = Nothing

Do while not oRst.EOF
  sTypes = sTypes & "<option value='" & oRst("AgendaItemTypeID") & "'>" & oRst("AgendaItemTypeName") & "</option>"
  oRst.MoveNext
Loop

%>

<html>
<head>
  <title><%=langBSAnnouncements%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <script>
  function selectLink()
  {
    var index=document.addItem.type.selectedIndex;
    var selValue=document.addItem.type.options[index].value;
    
    if (selValue==1)
    {
      window.open('picker/', 'filepicker', 'width=506,height=345,scrollbars=0,,toolbars=0,statusbar=0,menubar=0,left=265,top=180');
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

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%
  DrawTabs tabMeetings,1
  %>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
      <td><font size="+1"><b><%=langAgendaItems%></b>
      </font><br><img  height=16 src="../images/spacer.gif" align="absmiddle">&nbsp;</td>
      <td width="200">&nbsp;</td>
    </tr>
<!--    
    <tr>
      <td width="151" align="center"><img src="../images/icon_home.jpg"></td>
      <td><font size="+1"><b>Agenda Items</b></font></td>
      <td width="200">&nbsp;</td>
    </tr>
-->
    <tr>
      <td valign="top" nowrap>

         <%DrawQuickLinks "",1%>
      </td>
      <td colspan="2" valign="top">
        <form name="addItem" method=post action="newagendaitem.asp" method="post">
        <input type="hidden" name="_task" value="newItem">
        <input type=hidden name="AgendaID" value=<%=intAgendaID%>>
        <input type=hidden name="mid" value=<%=smid%>>
        <input type=hidden name="itemID" value="">
		<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:verifySubmit();"><%=langCreate%></a></div>
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
              <td><input type="text" name="title"></td>
            </tr>
            <tr>
              <td><div id="linkText">Select Link:</div></td>
              <td><div id="linkField"><input type=text name="link" READONLY><input type=button name="btnLink" value="Select Link" onClick="selectLink();"></div></td>
            </tr>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
