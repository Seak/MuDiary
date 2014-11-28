<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Session.Contents("strUserName") & "的日记本—删除日记" %>
<%
If Not(Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName"))) Then
	Response.Redirect("index.asp")
	Response.End()
End If

ID = Trim(Request.QueryString("id"))

ConnectionDatabase
strSQL = "Delete ID, strUserName From [Diary] Where strUserName = '" & Session.Contents("strUserName") & "' And ID = " & ID
Conn.Execute(strSQL)
Set strSQL = Nothing
CloseDatabase

Title = Session.Contents("strUserName") & "的日记本—删除成功"
strMsg = "<div align=""center""><a href=""diary.asp?username=" & Session.Contents("strUserName") & """ onFocus=""this.blur()"">查看日记本</a></div>"
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
            <%= strMsg %>
          </td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->