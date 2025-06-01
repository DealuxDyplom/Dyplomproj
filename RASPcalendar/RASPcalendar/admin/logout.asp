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
' Destroy all admin cookies:
Session("Valid") = False
Session.Abandon
Response.Cookies("AppLoggedIn") = "False"
Response.Cookies("Admin") = "False"
Response.Cookies("AppUName") = ""
%>
<html>
<head>
<title>RASPcalendar</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../assets/RASPcalendar.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="760" border="0" cellspacing="2" cellpadding="0">
<tr> 
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
        <tr>
          <td bgcolor="#FFFFFF">
 <div class = "links">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="middle" bgcolor="#FFFFFF" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td align="left" valign="middle" bgcolor="#000000" class="textBold"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
      <td height="18" align="left" valign="middle" bgcolor="#FFE5B2" class="textBold">&nbsp;</td>
  </tr>
  <tr> 
    <td align="left" valign="middle" bgcolor="#333333"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
</div>
<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#333333">
<tr> 
<td bgcolor="#FFFFFF" class="text"><br>
&nbsp;
<div align="center">You have successfully logged out.<br>
<a href="default.asp">Log back in.</a></div><br>
&nbsp;</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
  </tr>
  <tr> 
    <td>
      <!--#include file="inc_footer.asp" -->
    </td>
  </tr>  
</table>
</body>
</html>