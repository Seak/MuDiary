<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<%
strUserName = Request.QueryString("username")
Title = strUserName & "的日记本—浏览日记"
%>
<!--#include file="inc_header.asp"-->
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;">
  <tr>
    <td width="600" height="30" align="center" valign="middle" background="img/td_bg.gif"><%= Title %></td>
  </tr>
  <tr> 
    <td align="center" valign="middle">
	 <br>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0" style="Word-Break:break-all; BORDER: 1px solid #E2E2E2;">
        <tr> 
          <td> 
<%
RecordPerPage = 10
ConnectionDatabase
strSQL = "Select ID, strUserName, dtmDateTime, strWeather, strTitle, strContent, blnSecrecy, strIP From [Diary] Where strUserName = '" & strUserName & "' Order By dtmDateTime DESC"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorLocation = 3
Rs.CacheSize = RecordPerPage
Rs.PageSize = RecordPerPage
Rs.Open strSQL, Conn, 3, 1, &H0001
absPageNum = CInt(Request.QueryString("page"))
TotalPages = Rs.PageCount
If Request.QueryString("page") = "" Or absPageNum > TotalPages Then absPageNum = 1
If Not(Rs.EOF) Then
	Rs.AbsolutePage = absPageNum
	For absRecordNum = 1 To Rs.PageSize
%>
            <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER-BOTTOM: 1px solid #E2E2E2;">
              <tr> 
                <td width="20%"><%= FormatDateTime(Rs("dtmDateTime"), 1) %></td>
                <td width="10%"><%= WeekDayName(Weekday(Rs("dtmDateTime"))) %></td>
                <td width="10%"><%= Rs("strWeather") %></td>
                <td width="45%"> 
                  <% If Not(Rs("blnSecrecy")) Or Session.Contents("PN_MUDiaryMaster_" & strUserName) Then %><%= Rs("strTitle") %><% Else %>悄悄地告诉自己……<% End If %>
                </td>
                <td width="15%"><% If Session.Contents("PN_MUDiaryMaster_" & strUserName) Then %><a href="edit.asp?id=<%= Rs("ID") %>" onFocus="this.blur()">编辑</a>　<a href="delete.asp?id=<%= Rs("ID") %>" onFocus="this.blur()">删除</a><% End If %></td>
              </tr>
              <tr> 
                <td colspan="5"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="1" background="img/line.gif"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td colspan="5"> 
                  <% If Not(Rs("blnSecrecy")) Or Session.Contents("PN_MUDiaryMaster_" & strUserName) Then %><%= Replace(Rs("strContent"), vbCrLf, "<br>") %><% Else %>告诉自己的话语，自己珍惜……<% End If %>
                </td>
              </tr>
            </table>
            <br> 
<%
		Rs.MoveNext
		If Rs.EOF Then
			Exit For
		End If
	Next
End If
%>
          </td>
        </tr>
        <tr>
          <td><div align="right">第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= Rs.RecordCount %>条</font>　 
              <% If Not absPageNum = 1 Then %>
              [<a href="?username=<%= strUserName %>&page=1" onFocus="this.blur()">首页</a>][<a href="?username=<%= strUserName %>&page=<%= absPageNum - 1 %>" onFocus="this.blur()">上一页</a>]
              <%
End If
If Not absPageNum = TotalPages Then
%>
              [<a href="?username=<%= strUserName %>&page=<%= absPageNum + 1 %>" onFocus="this.blur()">下一页</a>][<a href="?username=<%= strUserName %>&page=<%= TotalPages %>" onFocus="this.blur()">尾页</a>] 
            </div>
            <% End If %></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->