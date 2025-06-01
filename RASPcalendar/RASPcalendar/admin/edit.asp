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
Dim MM_editAction
Dim MM_abortEdit

Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

%>

<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "submit" And CStr(Request("MM_recordId")) <> "") Then

  MM_editTable = "EVENTS"
  MM_editColumn = "EventId"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "../admin/events.asp"
  MM_fieldsStr  = "EventType|value|EventName|value|EventDated|value|EventDated_2|value|EventDated_3|value|EventDated_4|value|EventLocation|value|EventDescription|value"
  MM_columnsStr = "EventType|none,none,NULL|EventName|',none,''|EventDated|',none,NULL|EventDated_2|',none,NULL|EventDated_3|',none,NULL|EventDated_4|',none,NULL|EventLocation|',none,''|EventDescription|',none,''"

' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "delete" And CStr(Request("MM_recordId")) <> "") Then

  MM_editTable = "EVENTS"
  MM_editColumn = "EventId"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "events.asp"

' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

' create the sql update statement
 Dim Query, EventsRS
	Query = "SELECT * FROM Events WHERE EventId=" & Request("MM_recordId")
	Set EventsRS = Server.CreateObject("ADODB.Recordset")
	EventsRS.Open Query, Connect, 2, 3
	
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
	EventsRS.Close()
	Response.Redirect("events.asp")
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
' execute the delete
Dim EventDeleteRS			
	Set EventDeleteRS = Server.CreateObject("ADODB.Recordset")
			EventDeleteRS.Open "SELECT * FROM Events WHERE EventId =" & Request("MM_recordId"), Connect, 2, 3
			EventDeleteRS.Delete	

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsType
Dim rsType_numRows

Set rsType = Server.CreateObject("ADODB.Recordset")
rsType.Open "SELECT * FROM Type ORDER BY TypeName ASC", Connect, 3, 3

rsType_numRows = 0
%>
<%
Dim rsEdit__MMColParam
rsEdit__MMColParam = "1"
If (Request.QueryString("iEve") <> "") Then 
  rsEdit__MMColParam = Request.QueryString("iEve")
End If
%>
<%
Dim rsEdit
Dim rsEdit_numRows

Set rsEdit = Server.CreateObject("ADODB.Recordset")
rsEdit.Open "SELECT * FROM EVENTS WHERE EventId = " + Replace(rsEdit__MMColParam, "'", "''") + "", Connect, 3, 3
rsEdit_numRows = 0
%>

<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<div class = "links"> 
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr align="center" valign="middle"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="submit" id="submit">
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
                        <option value="<%=(rsType.Fields.Item("TypeId").Value)%>" <%If (Not isNull((rsEdit.Fields.Item("EventType").Value))) Then If (CStr(rsType.Fields.Item("TypeId").Value) = CStr((rsEdit.Fields.Item("EventType").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsType.Fields.Item("TypeName").Value)%></option>
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
                    <td align="right" valign="middle" nowrap class="textBold"> 
                      TITLE:</td>
                    <td valign="middle"> <input name="EventName" type="text" class="form" value="<%=(rsEdit.Fields.Item("EventName").Value)%>" size="45"> 
                    </td>
                  </tr>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold">DATE 1: 
                    </td>
                    <td valign="middle"><input name="EventDated" type="text" class="form" id="EventDated" value="<%=(rsEdit.Fields.Item("EventDated").Value)%>" size="20" maxlength="60"></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold">DATE 2: 
                    </td>
                    <td valign="middle"><input name="EventDated_2" type="text" class="form" id="EventDated_2" value="<%=(rsEdit.Fields.Item("EventDated_2").Value)%>" size="20" maxlength="60"></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold">DATE 3: 
                    </td>
                    <td valign="middle"><input name="EventDated_3" type="text" class="form" id="EventDated_3" value="<%=(rsEdit.Fields.Item("EventDated_3").Value)%>" size="20" maxlength="60"></td>
                  </tr>
				  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold">DATE 4: 
                    </td>
                    <td valign="middle"><input name="EventDated_4" type="text" class="form" id="EventDated_4" value="<%=(rsEdit.Fields.Item("EventDated_4").Value)%>" size="20" maxlength="60"></td>
                  </tr>
                  <tr valign="baseline"> 
                    <td align="right" valign="middle" nowrap class="textBold"> 
                      LOCATION:</td>
                    <td valign="middle"> <input name="EventLocation" type="text" class="form" id="EventLocation" value="<%=(rsEdit.Fields.Item("EventLocation").Value)%>" size="60"> 
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" valign="top" nowrap class="textBold"> DETAIL: 
                    </td>
                    <td valign="baseline"> <textarea name="EventDescription" cols="60" rows="10" class="form"><%=(rsEdit.Fields.Item("EventDescription").Value)%></textarea> 
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" valign="top" nowrap class="textBold"> APPROVED: 
                    </td>
                    <td valign="top" class="form">
                      <label>
                        <input name="EventApproved" type="radio" value="1" <% If((rsEdit.Fields.Item("EventApproved").Value)= 1) Then Response.Write "checked" %>>
                        Yes</label>
                      <br>
                      <label>
                        <input type="radio" name="EventApproved" value="0" <% If((rsEdit.Fields.Item("EventApproved").Value)= 0) Then Response.Write "checked" %>>
                        No</label>
                   </td>
                  </tr>
                  <tr valign="baseline"> 
                    <td colspan="2" align="center" valign="middle" nowrap> <input name="save" type="submit" class="form" id="submit2" onClick="MM_validateForm('EventName','','R','EVE_URL','','R','EventDescription','','R');return document.MM_returnValue" value="Save Changes"> 
                    </td>
                  </tr>
                </table></td>
            </tr>
          </table></td>
        <input type="hidden" name="MM_update" value="submit">
        <input type="hidden" name="MM_recordId" value="<%= rsEdit.Fields.Item("EventId").Value %>">
      </form>
    </tr>
    <tr align="center" valign="middle">
<td class="textBold"><a href="delete.asp?iEve=<%= rsEdit.Fields.Item("EventId").Value %>">Delete this Event</a></td>
    </tr>
  </table> 
</div>
<%
rsType.Close()
Set rsType = Nothing
%>
<%
rsEdit.Close()
Set rsEdit = Nothing
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
</body>
</html>