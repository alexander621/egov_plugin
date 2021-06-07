<!-- Different from UpdateAdmin is that, i must utilize StudentID,  -->

<FORM METHOD=POST ACTION="<%=thisname%>?StudentID=<%=request.querystring("StudentID")%>">
<table border="1" width="500">

	<% 
	intcolumns=UBOUND(conTABLEFIELDS)
	for j=0 to intcolumns
	if l_modify(j) then 
	  if l_fields(j)>150 then	   %>          
	    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><textarea rows="5" cols="40" name="<% =conTABLEFIELDS(j) %>"><%=rs(j)%></textarea></td>
       </tr>    
	<% 	elseif instr(TableDescription(j)," bit") then 
	if rs(j)=true then string1="checked" else string1="" 
	if rs(j)=false then string2="checked" else string2="" 
	
	%>
	<tr>
    <td width="200"><% =fields_description(j) %></td>
    <td width="351"><input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string1%> value='true'>Yes&nbsp;&nbsp;
	                <input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string2%> value='false'>No
	</td>
	</tr>
	 <% else %>
    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><input type=text name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>" size=<%=l_fields(j)%> maxlength=<%=l_fields(j)%>></td>
    </tr>
	  <% end if %>
	<% else %>
	</tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><%=rs(j)%>
	<input type=hidden name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>">
	</td>
    </tr>
	<%
	end if 
	next
	%>
</table>
<INPUT TYPE="submit" value='Update Now'>
</FORM>
<a href='javascript:history.go(-1)'>Go Back</a>
