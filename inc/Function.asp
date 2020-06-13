<%
Function kquanxian(ksite)
ktime=dateDiff("n",klog,now)
if ksite="" or ktime>=ksitetime then Response.redirect"login.asp"
if ksite=1 then
if ksiteid<>1000 and kuserid<>1000 then Response.redirect"index.asp"
end if
End Function

'定义底部
sub ksite_end
Response.Write"<div class=""content"">"
Response.Write ksite_time
Response.Write"</div>"
Response.Write"</body>"
Response.Write"</html>"
end sub

'上一页 下一页 跳转
sub kpage(mypage)
if pagecount>1 then Response.Write"<div class=line1>"
if page>1 then Response.Write"<a href='"&gopage&"siteid="&siteid&"'>[首页]</a> "
if page>1 then Response.Write"<a href='"&gopage&"page="&page-1&"&amp;siteid="&siteid&"'>[上页]</a> "
if page<pagecount then Response.Write"<a href='"&gopage&"page="&page+1&"&amp;siteid="&siteid&"'>[下页]</a> "
if page>=1 and page<pagecount then Response.Write"<a href='"&gopage&"page="&pagecount&"&amp;siteid="&siteid&"'>[尾页]</a>"
if pagecount>1 then Response.Write"</div>"
if mypage=2 then
Response.Write"<div class=line2>共<b>"&page&"</b>/"&pagecount&"页/"&Count&"条 "
if pagecount>1 then
Response.Write"<form method=""post"" action="""&gopage&""">"
Response.Write"<input name=""siteid"" type=""hidden"" value="""&siteid&""">"
if pagecount>1 and page<pagecount then page=""&page+1&"" else pgae=""&page-1&""
Response.Write"<input name=""page"" type=""text"" size=""5"" value="""&page&""">"
Response.Write"<input name=""g"" type=""submit"" value=""跳转""></form></div>"
else
Response.Write"</div>"
end if
end if
end sub

sub ksite_err(myerr)
Response.Write"<hr><b>出现一个提醒如下</b>:<br>"&myerr&"<hr>"
call ksite_end
Response.End
End Sub
%>
