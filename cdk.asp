<!--#include file="./head.asp"-->
<title>CDK后台</title>
<script type="text/javascript">
function CheckAll(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      e.checked = true;
   }
}
function ConfirmDel(message)
{
   if (confirm(message))
   {
      document.formDel.submit()
   }
}
</script>
<%call kquanxian(1)
conn.execute("select * from [ksite_cdk]")
If Err Then 
err.Clear
call ksite_err("请先安装数据库字段")
end if
pg=request("pg")
if pg="" then
Response.write "<div class=title><a href='?siteid="&siteid&"&pg=sc'>生产CDK</a><br/><a href='?siteid="&siteid&"'>全部CDK</a>/<a href='?siteid="&siteid&"&amp;lx=1'>子站到期延期</a>/<a href='?siteid="&siteid&"&amp;lx=2'>使用时延期</a>/<a href='?siteid="&siteid&"&amp;lx=3'>vip</a></div>"

lx=request("lx")
set rs=server.CreateObject("adodb.recordset")
if lx="" then
rs.open "select * from [ksite_cdk] Order by ktime desc",conn,1,1
else
rs.open "select * from [ksite_cdk] where klx="&lx&" Order by ktime desc",conn,1,1
end if
If Not rs.eof Then	
	gopage="?lx="&lx&"&amp;"
	Count=rs.recordcount
	page=int(request("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize	
   if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
%>
<form name="formDel" method="post" action="<%=gopage%>pg=sc2&amp;siteid=<%=siteid%>&amp;page=<%=page%>">
<%
call kpage(1)
	For i=1 To PageSize 
	If rs.eof Then Exit For
if (i mod 2)=0 then Response.write "<div class=""line2"">" else Response.write "<div class=""line1"">"
Response.write ""&rs("id")&".<input type=""checkbox"" name=""cid"" value="""&rs("id")&""">"
sy=clng(rs("ksy"))
if sy=1 then Response.write "(未)" else Response.write "(已"&rs("userid")&")"

Response.write "CDK:"&rs("kcdk")&""
Response.write "(<a href='"&gopage&"page="&page&"&amp;siteid="&siteid&"&amp;pg=sc1&amp;id="&rs("id")&"'>删</a>)"
Response.write "</div><div class='content'>"
	k_lx=clng(rs("klx"))
    if k_lx=1 then Response.write"类型:结束日期+"&rs("kday")&"天"
    if k_lx=2 then Response.write"类型:当前日期+"&rs("kday")&"天"
	if k_lx=3 then Response.write"类型:网站成为vip"
Response.write "</div>"
	rs.movenext
 	Next
Response.write "<div class=tip><input onClick=""CheckAll(this.form)"" type=""button"" id=""submitAllSearch"" value=""全选""> "
Response.write "<input type=""reset"" name=""submit"" value=""重置所选""> "
Response.write "<input type=""submit"" value=""选择完毕,删除"" onClick=""ConfirmDel('是否确定删除？删除后不能恢复！');return false""></form></div>"
call kpage(2)
else
   Response.write "<div class=tip>暂时没有记录！</div>"
end if
rs.close
set rs=nothing
'-----删除cdk操作
elseif pg="sc1" then
id=request("id")
lx=request("lx")
page=request("page")
       conn.Execute"delete from [ksite_cdk] Where id="&id
Response.redirect"?siteid="&siteid&"&page="&page&"&lx="&lx&""
'-----
elseif pg="sc2" then
cid=request("cid")
lx=request("lx")
page=request("page")
conn.Execute("delete from [ksite_cdk] Where id in("&cid&")")
Response.redirect"?siteid="&siteid&"&page="&page&"&lx="&lx&""

'-----生产cdk
elseif pg="sc" then
Response.Write "<form method='post' action='?siteid="&siteid&"&amp;pg=yes'>"
Response.Write "<div class=tip>生产数量</div>"
Response.Write "<div class=line2><input type='text'  name='num1' value='' ></div>"
Response.Write "<div class=tip>延期天数(单位:天)</div>"
Response.Write "<div class=line2><input type='text'  name='num2' value='' ></div>"
Response.Write "<div class=tip>延期类型</div>"
Response.Write "<div class=line2><select name='lx' value='1'><option value='1'>1.子站过期日增加</option><option value='2'>2.使用当天增加</option><option value='3'>3.设置为VIP网站</option></select></div>"
Response.Write "<div class=line1><input type='submit' value='马上生产'></form></div>"

''''''''''''''''''''''''''''''''''''''''''''''
elseif pg="yes" then
%>
<!--#include file="./inc/keycdksc.asp"-->
<%
lx=request("lx")
num1=request("num1")
num2=request("num2")
if num1="" then call ksite_err("数量不能为留空")
if not Isnumeric(num1) then call ksite_err("必须是数字")
if lx<>3 then
if num2="" then call ksite_err("天数不能为留空")
if not Isnumeric(num2)  then call ksite_err("必须是数字")
if num1<=0 then call ksite_err("错误参数")  
end if

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_cdk]",conn,1,3
for i=1 to num1
rs.addnew
rs("kcdk")=Generate_Key
rs("ksy")=1
rs("klx")=lx
rs("kday")=num2
rs("ktime")=now()
rs.update
next
rs.close
set rs=nothing
Response.Write "<div class=title>批量生产"&num1&"个cdk成功</div>"
'-----cdk使用日志
elseif pg="cdklog" then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_cdk_log] order by id desc",conn,1,1
If Not rs.eof Then						
	gopage="?pg=cdklog&amp;"
	Count=rs.recordcount
	page=int(request("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
   if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
%>
<form name="formDel" method="post" action="?pg=scrz2&amp;siteid=<%=siteid%>&amp;page=<%=page%>">
<%
call kpage(1)
	For i=1 To PageSize
	If rs.eof Then Exit For
if (i mod 2)=0 then Response.write "<div class=""line2"">" else Response.write "<div class=""line1"">"
Response.write ""&page*PageSize+i-PageSize&".<input type=""checkbox"" name=""cid"" value="""&rs("id")&""">"
Response.write ""&rs("siteid")&"使用了"&rs("kcdk")&"(在"&rs("ktime")&"/"&rs("uip")&")</div>"
	k_lx=clng(rs("klx"))
    if k_lx=1 then Response.write "CDK延期：到期时间+"&rs("kday")&"天" 
	if k_lx=2 then Response.write "CDK延期：当前时间+"&rs("kday")&"天 cdk："&rs("kcdk")&""
	if k_lx=3 then Response.write "网站成为VIP网站"
Response.write "(<a href='?page="&page&"&amp;siteid="&siteid&"&amp;id="&rs("id")&"&amp;pg=scrz1'>删</a>)</div>"
	rs.movenext
 	Next
Response.write "<div class=tip><input onClick=""CheckAll(this.form)"" type=""button"" id=""submitAllSearch"" value=""全选""> "
Response.write "<input type=""reset"" name=""submit"" value=""重置所选""> "
Response.write "<input type=""submit"" value=""选择完毕,删除"" onClick=""ConfirmDel('是否确定删除？删除后不能恢复！');return false""></form></div>"
call kpage(2)
else
   Response.write "<div class=tip>暂时没有记录！</div>"
end if
rs.close
set rs=nothing
'-----删除使用记录操作
elseif pg="scrz1" then
id=request("id")
lx=request("lx")
page=request("page")
       conn.Execute"delete from [ksite_cdk_log] Where id="&id
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=cdklog"
'-----
elseif pg="scrz2" then
cid=request("cid")
lx=request("lx")
page=request("page")
conn.Execute("delete from [ksite_cdk_log] Where id in("&cid&")")
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=cdklog"
end if

call ksite_end
%>