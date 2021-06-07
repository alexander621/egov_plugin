
<FORM METHOD=POST ACTION="<%=thisname%>?iofaction=<%=ActNewPostSave%>&groupid=<%=request.querystring("groupid")%>">
<table border="0" width="400" class='tablelist'>

	<% 
	intcolumns=UBOUND(conTABLEFIELDS)
	for j=1 to intcolumns-1 
		if l_register(j) then 
	if l_fields(j)>=100 then
	%>
  	    <tr>
	<td width="200"><% =fields_description(j) %></td>
    <td width="351">
	<TEXTAREA name="<% =conTABLEFIELDS(j) %>" ROWS="2" COLS="50"></TEXTAREA>
</td>
	  </tr>
	<% 	elseif instr(TableDescription(j)," bit") then %>
	<tr>
    <td width="200"><% =fields_description(j) %></td>
    <td width="351"><input type=radio name="<% =conTABLEFIELDS(j)%>" value=true>Yes&nbsp;&nbsp;
	                <input type=radio name="<% =conTABLEFIELDS(j)%>" value=false>No
	</td>
    </tr>
	<%	else %>
	    <tr>
	<td width="200"><% =fields_description(j) %></td>
    <td width="351"><input type=text name="<% =conTABLEFIELDS(j) %>" value="<%=request.querystring(conTABLEFIELDS(j))%>"  size=<%=l_fields(j)%> maxlength=<%=l_fields(j)%>></td>
	  </tr>
  <%
  end if
  	end if 
  next %>
</table>
<INPUT TYPE="submit" value='Submit Now'>
</FORM>
</form>