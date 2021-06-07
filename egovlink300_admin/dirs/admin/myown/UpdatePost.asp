<FORM METHOD=POST ACTION="<%=thisname%>?iofaction=<%=ActUpdatePostSave%>">
<table border="0"  class='tablelist' width="400">
	<% 
	intcolumns=UBOUND(conTABLEFIELDS)
	for j=0 to intcolumns
	'===================================================
	if l_modify(j) then 
	'------------------------------------------------
	  if l_length(j)>=100 then	 
	'-------------------------
	  %>          
	    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><textarea rows="2" cols="50" name="<% =conTABLEFIELDS(j) %>"><%=rs(j)%></textarea></td>
       </tr>    
 <%
	'-------------------------
	elseif l_type(j)=50 then  ' means it is bit	
	'-------------------------
	if rs(j)=true then string1="checked" else string1="" 
	if rs(j)=false then string2="checked" else string2="" 	
	%>
	<tr>
    <td width="200"><% =fields_description(j) %></td>
    <td width="351"><input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string1%> value='true'>Yes&nbsp;&nbsp;
	                <input type=radio name="<% =conTABLEFIELDS(j)%>" <%=string2%> value='false'>No	</td>
	</tr>	
	<!---------------------->
	<% else %>
	<!--------------------->
    <tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><input type=text name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>" size=<%=l_length(j)%> maxlength=<%=l_length(j)%>></td>
    </tr>
	<!-------------------->
	  <% end if %>
	<!------------------------------------------------>
	<% else %>
	<!------------------------------------------------>
	</tr>
	<td width="133"><% =fields_description(j) %></td>
    <td width="351"><%=rs(j)%>
	<input type=hidden name="<% =conTABLEFIELDS(j) %>" value="<%=rs(j)%>">
	</td>
    </tr>
	<%
    '------------------------------------------------
	end if 
	'===================================================
	next
	%>
</table>
<INPUT TYPE="submit" value='<%=langUpdateRecord%>'>
</FORM>

