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
      <td align="right" valign="middle" bgcolor="#FFE5B2" class="textBold"><A href="logout.asp">LOGOUT</A>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" align="left" valign="middle" bgcolor="#333333"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
</div>
<%
Dim tfm_orderby, tfm_order
tfm_orderby = "EventDated"
tfm_order = "DESC"
If(CStr(Request.QueryString("tfm_orderby")) <> "") Then
	tfm_orderby = Cstr(Request.QueryString("tfm_orderby"))
End If
If(Cstr(Request.QueryString("tfm_order")) <> "") Then
	tfm_order = Cstr(Request.QueryString("tfm_order"))
End If

Dim sql_orderby
sql_orderby = " " & tfm_orderby & " " & tfm_order
%>
<%
Dim rsEvents__sql_orderby
rsEvents__sql_orderby = "EventDated"
if (sql_orderby <> "") then rsEvents__sql_orderby = sql_orderby
%>
<%
Dim Query, rsEvents, rsEvents_numRows

Set rsEvents = Server.CreateObject("ADODB.Recordset")
rsEvents.Open "SELECT *  FROM Events, Type WHERE EventType = TypeId ORDER BY " + Replace(rsEvents__sql_orderby, "'", "''"), Connect, 3, 3

rsEvents_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 20
Dim Repeat1__index
Repeat1__index = 0
rsEvents_numRows = rsEvents_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
Dim rsEvents_total
rsEvents_total = rsEvents.RecordCount

' set the number of rows displayed on this page
If (rsEvents_numRows < 0) Then
  rsEvents_numRows = rsEvents_total
Elseif (rsEvents_numRows = 0) Then
  rsEvents_numRows = 1
End If

' set the first and last displayed record
Dim rsEvents_first, rsEvents_last
rsEvents_first = 1
rsEvents_last  = rsEvents_first + rsEvents_numRows - 1

' if we have the correct record count, check the other stats
If (rsEvents_total <> -1) Then
  If (rsEvents_first > rsEvents_total) Then rsEvents_first = rsEvents_total
  If (rsEvents_last > rsEvents_total) Then rsEvents_last = rsEvents_total
  If (rsEvents_numRows > rsEvents_total) Then rsEvents_numRows = rsEvents_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsEvents_total = -1) Then
' count the total records by iterating through the recordset
  rsEvents_total=0
  While (Not rsEvents.EOF)
    rsEvents_total = rsEvents_total + 1
    rsEvents.MoveNext
  Wend
' reset the cursor to the beginning
  If (rsEvents.CursorType > 0) Then
    rsEvents.MoveFirst
  Else
    rsEvents.Requery
  End If
' set the number of rows displayed on this page
  If (rsEvents_numRows < 0 Or rsEvents_numRows > rsEvents_total) Then
    rsEvents_numRows = rsEvents_total
  End If
' set the first and last displayed record
  rsEvents_first = 1
  rsEvents_last = rsEvents_first + rsEvents_numRows - 1
  If (rsEvents_first > rsEvents_total) Then rsEvents_first = rsEvents_total
  If (rsEvents_last > rsEvents_total) Then rsEvents_last = rsEvents_total

End If
%>
<%
Dim MM_paramName, MM_rs, MM_rsCount, MM_size, MM_uniqueCol, MM_offset, MM_atTotal, MM_paramIsDefined
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsEvents
MM_rsCount   = rsEvents_total
MM_size      = rsEvents_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then
' use index parameter if defined, otherwise use offset parameter
  Dim r
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)
' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If
' move the cursor to the selected record
  Dim i
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then
' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend
' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If
' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If
' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If
' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsEvents_first = MM_offset + 1
rsEvents_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsEvents_first > MM_rsCount) Then rsEvents_first = MM_rsCount
  If (rsEvents_last > MM_rsCount) Then rsEvents_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)

' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
Dim MM_removeList, MM_keepURL, MM_keepForm, MM_keepBoth, MM_keepNone
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
Dim Item, NextItem
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function

' *** Move To Record: set the strings for the first, last, next, and previous links
Dim MM_keepMove, MM_moveParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
Dim params
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
Dim urlStr, MM_moveFirst, MM_moveLast, MM_moveNext, prev, MM_movePrev
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)

'sort column headers for rsEvents
Dim tfm_saveParams, tfm_keepParams, tfm_orderbyURL
tfm_saveParams = ""
tfm_keepParams = ""
If tfm_order = "ASC" Then
	tfm_order = "DESC"
Else
	tfm_order = "ASC"
End If
		
If tfm_saveParams <> "" Then
	tfm_params = Split(tfm_saveParams,",")
	For i = 0 to UBound(tfm_params)
		If Cstr(Request(tfm_params(i))) <> "" Then
			tfm_keepParams = tfm_keepParams & LCase(tfm_params(i)) & "=" & Server.URLEncode(Request(tfm_params(i))) & "&"
		End If
	Next
End If
tfm_orderbyURL = Request.ServerVariables("URL") & "?" & tfm_keepParams & "tfm_order=" & tfm_order & "&tfm_orderby="
%> 
<div class = "links"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td align="right" valign="middle" class="textBold"> 
<%
' Count number of events, create new page every 20 records:
Dim TM_counter, TM_PageEndCount
TM_counter = 0
For i = 1 to rsEvents_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > rsEvents_total Then TM_PageEndCount = rsEvents_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>Page " & TM_counter & "</b>")
End if
if(TM_PageEndCount <> rsEvents_total) then Response.Write(" : ")
next
 %>              
                  </td>
                </tr>
                <tr> 
                  <td align="left" valign="top"> 
                  <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#333333">
                      <tr align="center" valign="middle" bgcolor="#CCCCCC" class="textBold"> 
                        <td height="18"><a href="<%=tfm_orderbyURL%>EventName">TITLE</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>TypeName">CATEGORY</a></td>
                        <td width="70" height="18"><a href="<%=tfm_orderbyURL%>EventDated">DATED</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>EventLocation">LOCATION</a></td>
                        <td width="60" height="18">ACTIVE</td>
                        <td width="60" height="18">EDIT/DEL</td>
                      </tr>
                      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsEvents.EOF)) 
%>
                      <tr align="center" valign="middle" class="text"> 
                        <td align="left" bgcolor="#FFFFFF" class="textBold"><a href="../detail.asp?E_Id=<%=(rsEvents.Fields.Item("EventId").Value)%>&E_Type=<%=(rsEvents.Fields.Item("EventType").Value)%>" target="_blank"><%=(rsEvents.Fields.Item("EventName").Value)%></a></td>
                        <td align="left" bgcolor="#FFFFFF"><%=(rsEvents.Fields.Item("TypeName").Value)%></td>
                        <td align="center" bgcolor="#FFFFFF"><%=(rsEvents.Fields.Item("EventDated").Value)%>
						<%
						if rsEvents("EventDated_2") <> "" then
					Response.Write "<br>" & (rsEvents.Fields.Item("EventDated_2").Value)
					end if
					if rsEvents("EventDated_3") <> "" then
					Response.Write "<br>" & (rsEvents.Fields.Item("EventDated_3").Value)
					end if
					if rsEvents("EventDated_4") <> "" then
					Response.Write "<br>" & (rsEvents.Fields.Item("EventDated_4").Value)
					end if
					%></td>
                        <td align="center" bgcolor="#FFFFFF"><%=(rsEvents.Fields.Item("EventLocation").Value)%></td>
                        <td align="center" bgcolor="#FFFFFF"><% If (rsEvents.Fields.Item("EventApproved").Value = 1) Then 
						Response.Write "<img src='../assets/icon_yes.gif' width='13' height='13'>"
						Else 
						Response.Write "<img src='../assets/icon_no.gif' width='13' height='13'>"
						End If
						%>
                          </td>
                        <td bgcolor="#FFFFFF"><a href="edit.asp?iEve=<%=(rsEvents.Fields.Item("EventId").Value)%>"><img src="../assets/folderIcon.gif" width="22" height="18" align="absmiddle" border="0"></a></td>
                      </tr>
                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsEvents.MoveNext()
Wend
%>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<%
rsEvents.Close()
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