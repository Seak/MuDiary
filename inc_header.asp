<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%= Title %></title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#F6F6F6" leftmargin="0" topmargin="5" oncontextmenu="return false" onselectstart="return false">
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;">
  <tr> 
    <td align="center" valign="middle"><a href="#" onFocus="this.blur()"><img src="img/title.jpg" width="600" height="99" border="0"></a></td>
  </tr>
<% If Not(Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName"))) Then %>
  <tr> 
    <td align="center" valign="middle" background="img/menu_bg.gif"><a href="index.asp" onFocus="this.blur()"><img src="img/menu_home.gif" border="0"></a><a href="user.asp" onFocus="this.blur()"><img src="img/menu_user.gif" border="0"></a><a href="join.asp" onFocus="this.blur()"><img src="img/menu_join.gif" border="0"></a><a href="login.asp" onFocus="this.blur()"><img src="img/menu_login.gif" border="0"></a><a href="help.asp" onFocus="this.blur()"><img src="img/menu_help.gif" border="0"></a><a href="copyright.asp" onFocus="this.blur()"><img src="img/menu_copyright.gif" border="0"></a><img src="img/menu_blank.gif"></td>
  </tr>
<% Else %>
  <tr>
    <td align="center" valign="middle" background="img/menu_bg.gif"><a href="index.asp" onFocus="this.blur()"><img src="img/menu_home.gif" border="0"></a><a href="write.asp" onFocus="this.blur()"><img src="img/menu_write.gif" border="0"></a><a href="modify.asp" onFocus="this.blur()"><img src="img/menu_modify.gif" border="0"></a><a href="logout.asp" onFocus="this.blur()"><img src="img/menu_logout.gif" border="0"></a><a href="help.asp" onFocus="this.blur()"><img src="img/menu_help.gif" border="0"></a><a href="copyright.asp" onFocus="this.blur()"><img src="img/menu_copyright.gif" border="0"></a><img src="img/menu_blank.gif"></td>
  </tr>
<% End If %>
  <tr> 
    <td height="15" align="center" valign="middle">&nbsp;</td>
  </tr>
</table>