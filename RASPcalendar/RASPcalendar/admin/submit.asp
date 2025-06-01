<% @EnableSessionState=True %>
<% Option Explicit %>
<%
'--------------------------------------------------
':::: RASPcalendar ::::
'--------------------------------------------------
'ASP & Microsoft Access online calendar application.
'Includes an administration interface to allow for easy updates and maintenance.
'Copyright (C) 2010  Robert Temple
'--------------------------------------------------
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'You can contact the author at rttucson@gmail.com
'Please do not remove the author, copyright, or license info.
%><%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%> 
<% If Request.Cookies("AppLoggedIn") <> "True" Then Response.Redirect("default.asp")
If Request.Cookies("Admin") <> "True" Then Response.Redirect("default.asp") %>
<!--#include file="../Connections/connRASPcalendarAdmin.asp" -->
<html>
<head>
<title>Calendar Admin</title>
<link href="../assets/RASPcalendar.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<table width="760" border="0" cellspacing="2" cellpadding="0">
<tr> 
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
        <tr>
          <td bgcolor="#FFFFFF">
 <div class = "links">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2" align="left" valign="middle" bgcolor="#FFFFFF" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td colspan="2" align="left" valign="middle" bgcolor="#000000" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
      <td height="18" align="left" valign="middle" bgcolor="#FFE5B2" class="textBold">&nbsp;
            <a href="events.asp">EVENT LISTING</a> | <a href="submit.asp">ADD NEW EVENT</a> | <a href="purge.asp">PURGE OLD EVENTS</a> | <a href="category.asp">EDIT CATEGORIES</a></td>
      <td align="right" valign="middle" bgcolor="#FFE5B2" class="textBold"><A HREF="logout.asp">LOGOUT</A>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" align="left" valign="middle" bgcolor="#333333"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
</div>

<%

Dim Flag

If Request("Flag") = "True" Then Flag = True
' If the form was submitted and required fields were left blank:
	If Request("EventName")="" Or Request("EventDated")="" Or Request("EventLocation")="" Or Request("EventDescription")="" Then
	If Flag Then
		Response.Write "<p class='bgHeader'>You must enter the date(s), title, location, and details of the event.</p>"
	End If
' If the form hasn't been submitted, show the form:
%>

<%
Dim rsType
Dim rsType_numRows

Set rsType = Server.CreateObject("ADODB.Recordset")
rsType.Open "SELECT * FROM Type ORDER BY TypeName ASC", Connect, 3, 3

rsType_numRows = 0
%>
<div class = "links"> 
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr align="center" valign="middle"> 	
      <form action="<%= Request.ServerVariables("URL") %>" method="POST" name="submit" id="submit">
        <td> <table border="0" cellpadding="0" cellspacing="1" bgcolor="#666666">
            <tr> 
              <td align="left" valign="top" bgcolor="#FFFFFF"><table align="center" cellpadding="3" cellspacing="3">
                  <% If Request.QueryString("requsername") <> "" Then %>
                  <% End If %>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold">CATEGORY:</td>
                    <td valign="middle"> <select name="EventType" class="form" id="select">
                        <%
While (NOT rsType.EOF)
%>
                        <option value="<%=(rsType.Fields.Item("TypeId").Value)%>"><%=(rsType.Fields.Item("TypeName").Value)%></option>
                        <%
  rsType.MoveNext()
Wend
If (rsType.CursorType > 0) Then
  rsType.MoveFirst
Else
  rsType.Requery
End If
%>
                      </select> </td>
                  </tr>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> TITLE:</td>
                    <td valign="middle"> <input name="EventName" type="text" class="form" value="" size="45"> 
                    </td>
                  </tr>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> DATE 1: 
                    </td>
                    <td valign="middle"><select name="EventDated" class="form" id="EventDated">
                        <%
						Dim dLoop
					dLoop = 1
					Do while dLoop < 180
					%>
                        <option value="<%= date() + dLoop %>"><%= date() + dLoop %></option>
                        <%
					dLoop = dLoop+1 
					Loop
					%>
                      </select></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> DATE 2: 
                    </td>
                    <td valign="middle"><select name="EventDated_2" class="form" id="EventDated_2">
					<option value="">[Date 2...]</option>
                        <%
					dLoop = 1
					Do while dLoop < 180
					%>
                        <option value="<%= date() + dLoop %>"><%= date() + dLoop %></option>
                        <%
					dLoop = dLoop+1 
					Loop
					%>
                      </select></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> DATE 3: 
                    </td>
                    <td valign="middle"><select name="EventDated_3" class="form" id="EventDated_3">
					<option value="">[Date 3...]</option>
                        <%
					dLoop = 1
					Do while dLoop < 180
					%>
                        <option value="<%= date() + dLoop %>"><%= date() + dLoop %></option>
                        <%
					dLoop = dLoop+1 
					Loop
					%>
                      </select></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> DATE 4: 
                    </td>
                    <td valign="middle"><select name="EventDated_4" class="form" id="EventDated_4">
					<option value="">[Date 4...]</option>
                        <%
					dLoop = 1
					Do while dLoop < 180
					%>
                        <option value="<%= date() + dLoop %>"><%= date() + dLoop %></option>
                        <%
					dLoop = dLoop+1 
					Loop
					%>
                      </select></td>
                  </tr>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> LOCATION:</td>
                    <td valign="middle"> <input name="EventLocation" type="text" class="form" id="EventLocation" size="60"> 
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" valign="top" nowrap class="textBold"> DETAIL:</td>
                    <td valign="baseline"> <textarea name="EventDescription" cols="60" rows="10" class="form"></textarea> 
                    </td>
                  </tr>
                  <tr>
                  <td><input type="hidden" name="Flag" value="True"></td>
                  <td><input type="hidden" name="EventApproved" value="1"></td>
                  </tr>
                  <tr valign="baseline"> 
                    <td nowrap align="right">&nbsp;</td>
                    <td> <input name="submit" type="submit" class="form" id="submit2" onClick="MM_validateForm('EventName','','R','EventDescription','','R');return document.MM_returnValue" value="Submit"> 
                     
                      <input type="hidden" name="MM_insert" value="submit"> </td>
                  </tr>
                </table></td>
            </tr>
          </table></td>
      </form>
    </tr>
  </table>
  
</div>


<%
rsType.Close()
Set rsType = Nothing
%>
		  </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>
      <!--#include file="inc_footer.asp" -->
    </td>
  </tr>  
</table>
<%
' If the form has been submitted and the required fields have been filled out, add the new event to the database:
Else

	Dim EventsRS
	
	Set EventsRS = Server.CreateObject("ADODB.Recordset")
	EventsRS.Open "SELECT * FROM Events", Connect, 2, 3
	
	EventsRS.AddNew
	EventsRS("EventType") = Request("EventType")
	EventsRS("EventName") = Request("EventName")
	EventsRS("EventDated") = Request("EventDated")
	If IsDate(Request("EventDated_2")) Then EventsRS("EventDated_2") = Request("EventDated_2")
	If IsDate(Request("EventDated_3")) Then EventsRS("EventDated_3") = Request("EventDated_3")
	If IsDate(Request("EventDated_4")) Then EventsRS("EventDated_4") = Request("EventDated_4")
	EventsRS("EventLocation") = Request("EventLocation")
	EventsRS("EventDescription") = Request("EventDescription")
	EventsRS("EventApproved") = Request("EventApproved")
	EventsRS.Update

	Response.Redirect("events.asp")
	
End If
%>
</body>
</html>