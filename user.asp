<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "���û��б�" %>
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
          <td width="20%">�ǳ�</td>
          <td width="35%">����</td>
          <td width="25%">ע��ʱ��</td>
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
          <td colspan="4">��<%= absPageNum %>/<%= TotalPages %>ҳ����ҳ<%= RecordPerPage %>������<%= Rs.RecordCount %>��</font>�� 
            <% If Not absPageNum = 1 Then %>
            [<a href="?page=1" onFocus="this.blur()">��ҳ</a>][<a href="?page=<%= absPageNum - 1 %>" onFocus="this.blur()">��һҳ</a>]<%
End If
If Not absPageNum = TotalPages Then
%>[<a href="?page=<%= absPageNum + 1 %>" onFocus="this.blur()">��һҳ</a>][<a href="?page=<%= TotalPages %>" onFocus="this.blur()">βҳ</a>] 
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