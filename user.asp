<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "—用户列表" %>
<!--#include file="inc_header.asp"-->
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;">
  <tr>
    <td width="600" height="30" align="center" valign="middle" background="img/td_bg.gif"><%= Title %></td>
  </tr>
  <tr> 
    <td align="center" valign="middle">
	 <br>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E2E2E2" style="Word-Break:break-all; BORDER: 1px solid #E2E2E2;">
        <tr align="center" valign="middle" bgcolor="#F6F6F6"> 
          <td width="10%">ID</td>
          <td width="20%">昵称</td>
          <td width="35%">信箱</td>
          <td width="25%">注册时间</td>
        </tr>
<%
RecordPerPage = 20
ConnectionDatabase
strSQL = "Select ID, strUserName, strMail, dtmDateTime From [UserList] Order By ID DESC"
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
        <tr align="center" valign="middle" bgcolor="#F6F6F6"> 
          <td><%= Rs("ID") %></td>
          <td><a href="diary.asp?username=<%= Rs("strUserName") %>" onFocus="this.blur()"><%= Rs("strUserName") %></a></td>
          <td><%= Rs("strMail") %></td>
          <td><%= Rs("dtmDateTime") %></td>
        </tr>
        <%
		Rs.MoveNext
		If Rs.EOF Then
			Exit For
		End If
	Next
End If
%>
        <tr align="center" valign="middle" bgcolor="#F6F6F6"> 
          <td colspan="4">第<%= absPageNum %>/<%= TotalPages %>页　本页<%= RecordPerPage %>条　共<%= Rs.RecordCount %>条</font>　 
            <% If Not absPageNum = 1 Then %>
            [<a href="?page=1" onFocus="this.blur()">首页</a>][<a href="?page=<%= absPageNum - 1 %>" onFocus="this.blur()">上一页</a>]<%
End If
If Not absPageNum = TotalPages Then
%>[<a href="?page=<%= absPageNum + 1 %>" onFocus="this.blur()">下一页</a>][<a href="?page=<%= TotalPages %>" onFocus="this.blur()">尾页</a>] 
            <% End If %>
          </td>
        </tr>
        <%
Rs.Close
Set strSQL = Nothing
Set Rs = Nothing
CloseDatabase
%>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->