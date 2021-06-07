<!-- #include file="../includes/common.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

Dim sSql, oRst, sAgent, lPos, dVer, lID, sTo, sCc, sSubject, sMessage, sFrom, sMessageHeader, sTime
dVer = 0
sAgent = Request.ServerVariables("HTTP_USER_AGENT")
          
lPos = InStr(1, sAgent, "MSIE")
If (lPos > 0) Then
  dVer = CDbl(Mid(sAgent, lPos + 5, InStr(lPos+5, sAgent, ";") - lPos - 5))
End If

lID = Request.QueryString("pid") & ""
If lID <> "" Then

  sSql = "EXEC GetMailMessage " & Session("UserID") & "," & lID & ",1"

  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .ActiveConnection = Application("DSN")
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sSql
    .ActiveConnection = Nothing
  End With

  If Not oRst.EOF Then
    If Not oRst.EOF Then
      sFrom     = oRst("SentFrom")
      sTo       = oRst("SentTo")
      sCc       = oRst("SentCc") 
      sSubject  = oRst("Subject")
      sMessage  = oRst("Message")
      sTime     = oRst("Date")
    End If
    oRst.Close
  End If
  
  sMessageHeader="<p>&nbsp;</p>" & "<p>&nbsp;</p>"
  sMessageHeader=sMessageHeader & "<p> -----" & langOriginalMessage & "----- </p>"
  sMessageHeader=sMessageHeader & "<p>&nbsp;&nbsp;&nbsp;<b>" & langFrom & ": </b>" & sFrom & "</p>"
  sMessageHeader=sMessageHeader & "<p>&nbsp;&nbsp;&nbsp;<b>" & langSent & ": </b>" & sTime & "</p>"
  sMessageHeader=sMessageHeader & "<p>&nbsp;&nbsp;&nbsp;<b>" & langTo & ": </b>" & sTo & "</p>"
  sMessageHeader=sMessageHeader & "<p>&nbsp;&nbsp;&nbsp;<b>" & langSubject & ": </b>" & sSubject & "</p>"
  sMessageHeader=sMessageHeader & "<p>&nbsp;</p>" & "<p>&nbsp;</p>"

  Select Case clng(Request.QueryString("type"))
    Case COMPOSE_TYPE_REPLY
      sCc=""
      sTo=FlipName(sFrom)
      sSubject=langRe & " " & sSubject
      sMessage=sMessageHeader & sMessage
      lid=0 'Force it to create a new message
    Case COMPOSE_TYPE_REPLYALL
      sTo=RemoveMe(sTo)
      sCc=RemoveMe(sCc)
      sSubject=langRe & " " & sSubject
      sMessage=sMessageHeader & sMessage
      lid=0 'Force it to create a new message
    Case COMPOSE_TYPE_FOWARD
      sTo=""
      sCc=""
      sSubject=langFw & " " & sSubject
      sMessage=sMessageHeader & sMessage
      lid=0 'Force it to create a new message
    Case Default
      'Do nothing
  End Select
  
End If


'Flip Name takes First Last and makes it Last, First;
Function FlipName(sName)
  dim iSpace, sTemp
  iSpace=instr(1,sName," ")
  sTemp=Trim(mid(sName,iSpace+1))
  sTemp=sTemp & ", " & Trim(mid(sName,1,iSpace-1)) & ";"
  FlipName=sTemp
End Function

Function RemoveMe(sString)
  dim sName, sTemp, iMe
  sTemp=sString
  sName=FlipName(Session("FullName"))
  sName=mid(sName,1,len(sName)-1) 'Remove semicolon
  iMe=instr(1,sString,sName)
  
  if iMe > 0 then
    sTemp=mid(sString, 1, iMe-1)
    sTemp=sTemp & mid(sString, iMe+len(sName)+1)
  end if
  
  RemoveMe=Trim(sTemp)
End Function


%>

<html>
<head>
  <title><%=langBSMessages%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript" src="../scripts/popup_menu.js"></script>
  <style type="text/css">
  <!--
    #edit   {width:100%; height:260px; border-left:1px solid #666666; border-top:1px solid #666666; border-right:1px solid #aaaaaa; border-bottom:1px solid #aaaaaa; font-family:Arial; font-size:12px; overflow:auto; background-color:#ffffff;}

    .toolbar {background-color:#dedede; padding:3px; border-left:1px solid #336699; border-right:1px solid #336699; border-bottom:1px solid #336699; width:90%; font-size:8px;}
    .btn      {border:1px solid #dddddd; cursor:hand;}
    .btn_over {background-color:88aacc; border:1px solid #ffffff; cursor:hand;}

    .winbtn   {border:1px solid #eeeeee; cursor:hand;}
    .winbtn_o {border:1px solid #ffffff; cursor:hand; background-color:88aacc;}
    .window {background-color:#eeeeee; padding:2px; border:1px solid #333333;}

    p {margin:0;}
  //-->
  </style>
  <script language="Javascript">
  <!--
    var lastKeyCode = 0;

    function doLoad() {
      document.execCommand('FontSize', false, 2);
    }

    function doSend() {
      document.all.Message.value = document.all.edit.innerHTML;
      document.all.SendNow.value = 1;
      if(document.frmMailMessage.Subject.value.length > 50)
      {
        alert("Too long");
      }
      else
      {
        document.all.frmMailMessage.submit();
      }
    }

    function doSave() {
      document.all.Message.value = document.all.edit.innerHTML;
      document.all.SendNow.value = 0;
      if(document.frmMailMessage.Subject.value.length > 50)
      {
        alert("Too long");
      }
      else
      {
        document.all.frmMailMessage.submit();
      }
    }

    function doAddresses(sendtype) {
      window.open("addressbook.asp?t=" + sendtype, "_addbook", "width=300,height=400,toolbars=0,status=0,menubar=0,statusbar=0,scrollbars=1,location=0");
    }

    function doMouseOver() {
      if (window.event.srcElement.parentElement.className == "btn") {
        window.event.srcElement.parentElement.className = "btn_over";
        window.event.srcElement.style.position = "relative";
        window.event.srcElement.style.top = -1;
        window.event.srcElement.style.left = -1;
      }
      if (window.event.srcElement.parentElement.className == "winbtn")
        window.event.srcElement.parentElement.className = "winbtn_o";
    }

    function doMouseOut() {
      if (window.event.srcElement.parentElement.className == "btn_over") {
        window.event.srcElement.parentElement.className = "btn";
        window.event.srcElement.style.top = 0;
        window.event.srcElement.style.left = 0;
      }
      if (window.event.srcElement.parentElement.className == "winbtn_o")
        window.event.srcElement.parentElement.className = "winbtn";
    }

    function doSetFont( fontName ) {
      document.execCommand('FontName', false, fontName);
      checkIn('');
    }

    function doSetFontSize( fontSize ) {
      document.execCommand('FontSize', false, fontSize);
      checkIn('');
    }
    
    function PickImage()
    {
      window.open('picker/', 'filepicker', 'width=506,height=345,scrollbars=0,toolbars=0,statusbar=0,menubar=0,left=265,top=180');
      //document.execCommand('InsertImage',false)
    }
    
    function AddImage(sPath)
    {
      //alert(sPath);
      document.execCommand('InsertImage',false,sPath)
    }


    function AddLink()
    {//Identify selected text
    var sText = document.selection.createRange();
    if (!sText==""){
        //Create link
         document.execCommand("CreateLink");

         //document.execCommand('CreateLink', false, 'http://www.eclink.com')

         //Replace text with URL
         if (sText.parentElement().tagName == "A"){
           //sText.parentElement().innerText=sText.parentElement().href;
           //document.execCommand("ForeColor","false","#FF0033");
         }    
      }
    else{
        alert("<%=langAlertSelect%>");
      }   
    }

  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onmouseover="doMouseOver();" onmouseout="doMouseOut();" onload="doLoad();">
  <%DrawTabs tabMessages,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_messages.jpg"></td>
      <td><font size="+1"><b><%=langMessage%>: <%=langCompose%></b></font><br>
      <%
      Select Case Right(Request.ServerVariables("HTTP_REFERER"),12)
        Case "sentmail.asp"
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""sentmail.asp"">" & langBackTo & " " & langSentMail & "</a></td>"
        Case "s/drafts.asp"
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""drafts.asp"">" & langBackTo & " " & langDrafts & "</a></td>"
        Case Else
          Response.Write "<img src=""../images/arrow_2back.gif"" align=""absmiddle"">&nbsp;<a href=""../messages"">" & langBackTo & " " & langInbox & "</a></td>"
      End Select
      %>
    </tr>
    <tr>
      <td valign="top">

        <!-- START: QUICK LINKS MODULE //-->
        <div style="padding-bottom:8px;"><b><%=langMessageLinks%></b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/newmail_small.jpg" width="16" height="16" align="absmiddle">&nbsp;<a href="compose.asp"><%=langNewMessage%></a></div>
        <br>
        <div style="padding-bottom:8px;"><b>Message Boxes</b></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="../messages/"><%=langInbox%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="drafts.asp"><%=langDrafts%></a></div>
        <div class="quicklink">&nbsp;&nbsp;<img src="../images/folder_closed.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="sentmail.asp"><%=langSentMail%></a></div>
        <br>
        <div style="padding-bottom:3px;"><%=langSearchMessages%>:</div>
        <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
        <div class="quicklink" align="right"><a href="#"><img src="../images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
        <!-- END: QUICK LINKS MODULE //-->

      </td>
      <td valign="top">
        <form name="frmMailMessage" action="send.asp" method="post">
          <input type="hidden" name="OldMailID" value="<%=lID%>">
          <input type="hidden" name="Message">
          <input type="hidden" name="SendNow" value="1">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/newmail_small.jpg" align="absmiddle">&nbsp;<a href="javascript:doSend();">Send Message</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/save.gif" align="absmiddle">&nbsp;<a href="javascript:doSave();">Save Draft</a></div>
          <table width="90%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr bgcolor="#93bee1">
              <td style="color:#003366;"><b><%=langFrom%>:</b></td>
              <td width="100%"><%= Session("FullName") %></td>
            </tr>
            <tr bgcolor="#93bee1">
              <td><input type="button" value="<%=langTo%>..." style="width:60px; font-size:11px;" onclick="doAddresses('to');"></td>
              <td><input type="text" size="60" name="To" value="<%=sTo%>"></td>
            </tr>
            <tr bgcolor="#93bee1">
              <td><input type="button" value="<%=langCc%>..." style="width:60px; font-size:11px;" onclick="doAddresses('cc');"></td>
              <td><input type="text" size="60" name="Cc" value="<%=sCc%>"></td>
            </tr>
            <tr bgcolor="#93bee1">
              <td style="color:#003366;"><b><%=langSubject%>:</b></td>
              <td><input type="text" size="60" name="Subject" value="<%= AsciiToHtml(sSubject) %>"></td>
            </tr>

          <% If (dVer < 5.5) Then %>

            <tr>
              <td colspan="2" style="border-top:1px solid #336699;color:#003366;"><div style="padding-bottom:5px;"><b><%=langMessage%>:</b></div><textarea name="edit" rows="10" style="width:100%; font-size:13px; font-family:Arial,Verdana,Tahoma;"><%= sMessage %></textarea></td>
            </tr>
          </table>

          <% Else %>

          </table>
          <div class="toolbar">
            <span class="btn"><img src="../images/edit/cut.gif" onclick="document.execCommand('Cut');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/copy.gif" onclick="document.execCommand('Copy');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/paste.gif" onclick="document.execCommand('Paste');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/undo.gif" onclick="document.execCommand('Undo');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/redo.gif" onclick="document.execCommand('Redo');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/hr.gif" onclick="document.execCommand('InsertHorizontalRule');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/link.gif" onclick="AddLink();" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/image.gif" onclick="PickImage();" unselectable="on"></span>
            <br>
            <span class="btn"><img src="../images/edit/font.gif" onmouseover="checkIn('winFont');" onmouseout="checkOut('winFont');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/fontsize.gif" onmouseover="checkIn('winFontSize');" onmouseout="checkOut('winFontSize');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/bold.gif" onclick="document.execCommand('Bold');"  unselectable="on"></span>
            <span class="btn"><img src="../images/edit/italic.gif" onclick="document.execCommand('Italic');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/underline.gif" onclick="document.execCommand('Underline');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/leftalign.gif" onclick="document.execCommand('JustifyLeft');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/centeralign.gif" onclick="document.execCommand('JustifyCenter');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/rightalign.gif" onclick="document.execCommand('JustifyRight');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/numberlist.gif" onclick="document.execCommand('InsertOrderedList');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/bulletlist.gif" onclick="document.execCommand('InsertUnorderedList');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/indent.gif" onclick="document.execCommand('Indent');" unselectable="on"></span>
            <span class="btn"><img src="../images/edit/outdent.gif" onclick="document.execCommand('Outdent');" unselectable="on"></span>
            <img src="../images/edit/sep.gif">
            <span class="btn"><img src="../images/edit/fontcolor.gif" onmouseover="checkIn('winFontColor');" onmouseout="checkOut('winFontColor');" unselectable="on"></span>
            <br>
            <div id="edit" contentEditable="true"><%= sMessage %></div>
            </div>
          </div>

          <div id="winFont" class="window" style="padding:2px; position:absolute; top:341px; left:185px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
            <table border="0" cellpadding="2" cellspacing="0">
              <tr><td class="winbtn"><font face="Arial" size="3" onclick="doSetFont('Arial');" unselectable="on">Arial</font></td></tr>
              <tr><td class="winbtn"><font face="Arial Black" size="3" onclick="doSetFont('Arial Black');" unselectable="on">Arial Black</font></td></tr>
              <tr><td class="winbtn"><font face="Courier New" size="3" onclick="doSetFont('Courier New');" unselectable="on">Courier New</font></td></tr>
              <tr><td class="winbtn"><font face="Garamond" size="3" onclick="doSetFont('Garamond');" unselectable="on">Garamond</font></td></tr>
              <tr><td class="winbtn"><font face="Tahoma" size="3" onclick="doSetFont('Tahoma');" unselectable="on">Tahoma</font></td></tr>
              <tr><td class="winbtn"><font face="Times New Roman" size="3" onclick="doSetFont('Times New Roman');" unselectable="on">Times New Roman</font></td></tr>
            </table>
          </div>

          <div id="winFontSize" class="window" style="position:absolute; top:341px; left:226px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
            <table border="0" cellpadding="2" cellspacing="0">
              <tr><td class="winbtn"><font face="Arial" size="1" onclick="doSetFontSize(1);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="2" onclick="doSetFontSize(2);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="3" onclick="doSetFontSize(3);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="4" onclick="doSetFontSize(4);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="5" onclick="doSetFontSize(5);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="6" onclick="doSetFontSize(6);" unselectable="on">ABC,abc,123</font></td></tr>
              <tr><td class="winbtn"><font face="Arial" size="7" onclick="doSetFontSize(7);" unselectable="on">ABC,abc,123</font></td></tr>
            </table>
          </div>

          <div id="winFontColor" class="window" style="position:absolute; top:341px; left:504px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#000000; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#000000');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#666666; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#666666');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#cccccc; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#cccccc');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#ffffff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ffffff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#ff0000; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff0000');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#ff9900; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff9900');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#ffff00; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ffff00');checkIn('');" unselectable="on"></td>
              </tr>
              <tr>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#00ff00; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#00ff00');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#009900; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#009900');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#00ffff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#00ffff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#0000ff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#0000ff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#000099; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#000099');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#990099; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#990099');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../images/spacer.gif" width="10" height="10" style="background-color:#ff00ff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff00ff');checkIn('');" unselectable="on"></td>
              </tr>
            </table>
          </div>

          <% End If %>

        </form>
      </td>
    </tr>
  </table>
</body>
</html>
