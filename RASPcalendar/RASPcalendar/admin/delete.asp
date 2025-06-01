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
      <div align="center">
        <%
Dim EventsRS, Query, Flag
If Request("Flag") = "True" Then Flag = True
' If the form hasn't been submitted, show the form:
If Not Flag Then
Dim rsEdit
Dim rsEdit_numRows

Set rsEdit = Server.CreateObject("ADODB.Recordset")
rsEdit.Open "SELECT * FROM EVENTS WHERE EventId = " + Request.QueryString("iEve") + "", Connect, 3, 3
		
%>
      </div>
      <p align="center" class="textBold">Do you wish to delete the following event? This cannot be undone.</p>
      <form name="form_delete" method="post" action="delete.asp">
        <table border="0" align="center" cellpadding="0" cellspacing="0" class="text">
          <tr> 
            <td colspan="5" bgcolor="#CCCCCC"><img src="../images/spacer.gif" width="1" height="1"></td>
          </tr>        
          <tr bgcolor="#F8F8F8"> 
            <td>&nbsp;</td>
            <td bgcolor="#F8F8F8" class="notrequired">&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr bgcolor="#F8F8F8"> 
            <td>&nbsp;</td>
            <td bgcolor="#F8F8F8" class="notrequired" valign="top">Title:</td>
            <td valign="top">&nbsp;</td>
            <td valign="top"><%=(rsEdit.Fields.Item("EventName").Value)%></td>
            <td>&nbsp;</td>
          </tr>
          <tr bgcolor="#F8F8F8"> 
            <td>&nbsp;</td>
            <td valign="top" class="subtitle">&nbsp;</td>
            <td>&nbsp;</td>
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="5" bgcolor="#CCCCCC"><img src="../images/spacer.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><input type="hidden" name="MM_recordId" value="<%= rsEdit.Fields.Item("EventId").Value %>"></td>
            <td><input type="hidden" name="Flag" value="True"></td>
          </tr>
          <tr> 
            <td colspan="5" align="center"><input type="submit" name="Submit" value="Delete Event" class="text"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="5" align="center"><a href="events.asp">Do not delete</a></td>
          </tr>
        </table>
      </form>
      <% 
' If the form has been submitted, delete the event record:
Else

	Set EventsRS = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM Events WHERE EventId =" & Request("MM_recordId")
	EventsRS.Open Query, Connect, 2, 3
	EventsRS.Delete

	Response.Write "<p align=""center"" class=""textBold"">Event successfully deleted.</p>"

End If

%>      </td>
  </tr>
</table>
</body>
</html>