<!-- #include file="../../includes/common.asp" //-->

<%
Dim oCmd
Dim bShown, sLinks

If Request.Form("_task") = "newt" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewDiscussionTopic"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("DiscGroupID", adInteger, adParamInput, 4, Request.Form("DiscGroupID"))
    .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 100, Request.Form("Topic"))
    .Parameters.Append oCmd.CreateParameter("Message", adVarChar, adParamInput, 5000, Request.Form("Message"))
    .Execute
  End With
  Response.Redirect "../../meetings/discSelect/topics.asp?tid=" & Request.Form("DiscGroupID") & "&tp=1&gn=" & Request.Form("gn")
End If


Dim sAgent, lPos, dVer
dVer = 0
sAgent = Request.ServerVariables("HTTP_USER_AGENT")
          
lPos = InStr(1, sAgent, "MSIE")
If (lPos > 0) Then
  dVer = CDbl(Mid(sAgent, lPos + 5, InStr(lPos+5, sAgent, ";") - lPos - 5))
End If
%>

<html>
<head>
  <title><%=langBSDiscussions%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <style type="text/css">
  <!--
    #edit   {width:100%; height:300px; border-left:1px solid #666666; border-top:1px solid #666666; border-right:1px solid #aaaaaa; border-bottom:1px solid #aaaaaa; font-family:Arial; font-size:12px; overflow:auto; background-color:#ffffff;}

    .toolbar {background-color:#dedede; padding:3px; border-left:1px solid #336699; border-right:1px solid #336699; border-bottom:1px solid #336699; width:100%x; font-size:8px;}
    .btn      {border:1px solid #dddddd; cursor:hand;}
    .btn_over {background-color:88aacc; border:1px solid #ffffff; cursor:hand;}

    .winbtn   {border:1px solid #eeeeee; cursor:hand;}
    .winbtn_o {border:1px solid #ffffff; cursor:hand; background-color:88aacc;}
    .window {background-color:#eeeeee; padding:2px; border:1px solid #333333;}

    p {margin:0;}
  //-->
  </style>
  <script language="Javascript" src="../popup_menu.js"></script>
  <script language="Javascript">
  <!--
    var lastKeyCode = 0;

    function doLoad() {
      document.execCommand('FontSize', false, 2);
      document.all.Topic.focus();
    }

    function doKeyDown() {
      if (window.event.keyCode == 13) {
        alert(document.all.edit.innerHTML);
        return false;
      }
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


    function AddLink()
    {//Identify selected text
    var sText = document.selection.createRange();
    if (!sText==""){
        //Create link
         document.execCommand("CreateLink");

         //document.execCommand('CreateLink', false, 'http://www.eclink.com')

         //Replace text with URL
         if (sText.parentElement().tagName == "A"){
           sText.parentElement().innerText=sText.parentElement().href;
           document.execCommand("ForeColor","false","#FF0033");
         }    
      }
    else{
        alert("<%=langAlertSelect%>");
      }   
    }

  //-->
  </script>
</head>

<body bgcolor="#ffffff" onmouseover="doMouseOver();" onmouseout="doMouseOut();" onload="doLoad();" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">


  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    
    <tr>
      <td valign="top">
        
       
      </td>
      <td colspan="2" valign="top">
        <form name="frmNewTopic" action="newtopic.asp" method="post">
          <input type="hidden" name="_task" value="newt">
          <input type="hidden" name="Message">
          <input type="hidden" name="DiscGroupID" value="<%= Request.QueryString("tid") %>">
          <input type="hidden" name="gn" value="<%= Request.QueryString("gn") %>">

          <div style="font-size:10px; padding-bottom:5px;"><img src="../../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.Message.value = document.all.edit.innerHTML; document.all.frmNewTopic.submit();"><%=langCreate%></a></div>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
            <tr>
              <th align="left" colspan="2"><%=langNewTopic%></th>
            </tr>
            <tr>
              <td style="color:#336699;" width="15"><br><b>Topic:&nbsp;&nbsp;&nbsp;</b><br><br></td>
              <td><br><input type="text" name="Topic" size="50"><br><br></td>
            </tr>

          <% If (dVer < 5.5) Then %>

            <tr>
              <td colspan="2"><font color="#336699"><b><%=Message%>:</b></font><br><textarea name="edit" rows="10" style="width:100%;"></textarea></td>
            </tr>
            </table>

          <% Else %>

          </table>
          <div class="toolbar">
            <span class="btn"><img src="../../images/edit/cut.gif" onclick="document.execCommand('Cut');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/copy.gif" onclick="document.execCommand('Copy');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/paste.gif" onclick="document.execCommand('Paste');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/undo.gif" onclick="document.execCommand('Undo');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/redo.gif" onclick="document.execCommand('Redo');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/hr.gif" onclick="document.execCommand('InsertHorizontalRule');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/link.gif" onclick="AddLink();" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/image.gif" onclick="document.execCommand('InsertImage');" unselectable="on"></span>
            <br>
            <span class="btn"><img src="../../images/edit/font.gif" onmouseover="checkIn('winFont');" onmouseout="checkOut('winFont');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/fontsize.gif" onmouseover="checkIn('winFontSize');" onmouseout="checkOut('winFontSize');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/bold.gif" onclick="document.execCommand('Bold');"  unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/italic.gif" onclick="document.execCommand('Italic');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/underline.gif" onclick="document.execCommand('Underline');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/leftalign.gif" onclick="document.execCommand('JustifyLeft');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/centeralign.gif" onclick="document.execCommand('JustifyCenter');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/rightalign.gif" onclick="document.execCommand('JustifyRight');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/numberlist.gif" onclick="document.execCommand('InsertOrderedList');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/bulletlist.gif" onclick="document.execCommand('InsertUnorderedList');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/indent.gif" onclick="document.execCommand('Indent');" unselectable="on"></span>
            <span class="btn"><img src="../../images/edit/outdent.gif" onclick="document.execCommand('Outdent');" unselectable="on"></span>
            <img src="../../images/edit/sep.gif">
            <span class="btn"><img src="../../images/edit/fontcolor.gif" onmouseover="checkIn('winFontColor');" onmouseout="checkOut('winFontColor');" unselectable="on"></span>
            <br>
            <div id="edit" contentEditable="true"></div>
            </div>
          </div>

          <div id="winFont" class="window" style="padding:2px; position:absolute; top:306px; left:185px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
            <table border="0" cellpadding="2" cellspacing="0">
              <tr><td class="winbtn"><font face="Arial" size="3" onclick="doSetFont('Arial');" unselectable="on">Arial</font></td></tr>
              <tr><td class="winbtn"><font face="Arial Black" size="3" onclick="doSetFont('Arial Black');" unselectable="on">Arial Black</font></td></tr>
              <tr><td class="winbtn"><font face="Courier New" size="3" onclick="doSetFont('Courier New');" unselectable="on">Courier New</font></td></tr>
              <tr><td class="winbtn"><font face="Garamond" size="3" onclick="doSetFont('Garamond');" unselectable="on">Garamond</font></td></tr>
              <tr><td class="winbtn"><font face="Tahoma" size="3" onclick="doSetFont('Tahoma');" unselectable="on">Tahoma</font></td></tr>
              <tr><td class="winbtn"><font face="Times New Roman" size="3" onclick="doSetFont('Times New Roman');" unselectable="on">Times New Roman</font></td></tr>
            </table>
          </div>

          <div id="winFontSize" class="window" style="position:absolute; top:306px; left:226px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
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

          <div id="winFontColor" class="window" style="position:absolute; top:100px; left:504px; visibility:hidden;" onmouseout="if (!this.contains(event.toElement)) checkIn('');">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#000000; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#000000');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#666666; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#666666');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#cccccc; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#cccccc');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#ffffff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ffffff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#ff0000; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff0000');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#ff9900; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff9900');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#ffff00; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ffff00');checkIn('');" unselectable="on"></td>
              </tr>
              <tr>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#00ff00; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#00ff00');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#009900; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#009900');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#00ffff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#00ffff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#0000ff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#0000ff');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#000099; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#000099');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#990099; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#990099');checkIn('');" unselectable="on"></td>
                <td class="winbtn" style="padding:1px;"><img src="../../images/spacer.gif" width="10" height="10" style="background-color:#ff00ff; border:1px solid #999999;" onclick="edit.focus();document.execCommand('ForeColor', false, '#ff00ff');checkIn('');" unselectable="on"></td>
              </tr>
            </table>
          </div>

          <% End If %>
          <div style="font-size:10px; padding-top:5px;"><img src="../../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.Message.value = document.all.edit.innerHTML; document.all.frmNewTopic.submit();"><%=langCreate%></a></div>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>