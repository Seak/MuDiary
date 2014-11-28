<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "—首页" %>
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
          <td>　　开辟这个版块只是想让网友随心所欲的写些什么，信手涂鸦也好，慷慨激昂也好，快乐也好，悲伤也好，写下了，记下了走过的点点滴滴，也许并没有太大意义，只是想留下一些什么，在这缤纷的网络保留一份宁静与安祥。<br>
            　　在网络中少有的净土上给自己留下一丝空间，不再感到无家可归，不再有那份彷徨。有时，需要诉说；有时，是要告诉自己；有时，只是要证明自己还是活者。够了，这些，足够了！<% If Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName")) Then %><br>
            　　看看自己的日记本，[<a href="diary.asp?username=<%= Session.Contents("strUserName") %>" onFocus="this.blur()">进入</a>]。<% End If %></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->