<!--#include file="head.asp"-->
<title>登录验证</title>
<%
pg=request("pg")
if pg="" then
%>
<div class=content>
<form method="post" action="?">
<input name="pg" type="hidden" value="yes">
请输入你的登陆用户名<br/>
<input name="kname" type="text" value=""><br/>
请输入你的登陆密码<br/>
<input name="kpass" type="password" value="">
<br/>
<input name="g" type="submit" value="确认">
</form>
</div>
<%
elseif pg="yes" then
kname=request("kname")
kpass1=request("kpass")
kpass2=UCase(MD5(kpass1))
if kname="" or kpass1="" then call ksite_err("用户名密码不能为空")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&kname&"'",conn,1,1
if rs.bof or rs.eof then call ksite_err("不存在的用户名("&kname&")")
password1=rs("password")
password2=UCase(rs("password"))
if kpass2<>password2 then call ksite_err("输的密码不正确<br/>"&kpass1&"")
if clng(rs("userid"))<>clng(rs("siteid")) then call ksite_err("将要登录的不是子站")
kluserid=rs("userid")
klsiteid=rs("siteid")
rs.close
set rs=nothing

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_log]",conn,1,2
rs.addnew
rs("userid")=kluserid
rs("siteid")=klsiteid
rs("username")=kname
rs("ktime")=now()
rs("uip")=userip
rs.update
rs.close
set rs=nothing
session("ksite")=kname&"_"&klsiteid&"_"&kluserid
session("klog")=now
Response.redirect"index.asp"


end if
call ksite_end
%>