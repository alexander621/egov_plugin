<!-- #include file="includes/common.asp" //-->
<%
sLevel     = "../"     'Override of value from common.asp
lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

lcl_field_id = request("fieldid")
%>
<html> 
<head> 
<title>E-Gov Administration Console {Color Palette}</title> 

<script language="javascript">
function returnColor(iColor) {
  parent.opener.document.getElementById("<%=lcl_field_id%>").value=iColor;
  if(parent.opener.document.getElementById("<%=lcl_field_id%>_previewcolor")) {
     parent.opener.document.getElementById("<%=lcl_field_id%>_previewcolor").style.backgroundColor=iColor;
  }
  parent.close();
}
</script>

</head>
<body>
<table border="0" cellspacing="0" cellpadding="2" width="100%">
  <tr valign="top">
      <td width="70%">
<script language="javascript">
//first lets start with the table that will contain the palette
t='<table width="100%" style="font-size: 8pt;">';

//next we define the array of basic colors including black, white and gray
c=new Array('00','CC','33','66','99','FF');

//now we will iterate through colors. Notice how number 6 corresponds to number of items in the array c
for(i=0;i<6;i++){
 for(j=0;j<6;j++){

 // each row will have 6 colors 
  t +='<tr>';
   for(k=0;k<6;k++){

    //this creates hex code for each color
    lcl_hexcode=c[i]+c[j]+c[k];

    //now lets create table cell for each color
//    t+='<td bgcolor=#'+l+'>#'+l+'</td>';
    t+='<td bgcolor=#'+lcl_hexcode+' style="cursor: hand;" onmouseover=document.getElementById("p_bgcolor").value="' + lcl_hexcode + '"; onmouseout=document.getElementById("p_bgcolor").value=""; onclick=returnColor("' + lcl_hexcode + '");>&nbsp;</td>';
   }
  t +='</tr>';
 }
}
//now display the palette 
document.write(t+'</table>');
</script>
      </td>
      <td>
          <input type="text" name="p_bgcolor" id="p_bgcolor" size="10" maxlength="6" disabled><br>
          <input type="button" value="Close Window" onclick="parent.close();">
      </td>
  </tr>
</table>
</body>
</html>