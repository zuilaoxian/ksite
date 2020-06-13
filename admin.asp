<!--#include file="./head.asp"-->
<title>管理后台</title>
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
pg=request("pg")
if pg="" then
Response.write "<div class='title'><a href='?siteid="&siteid&"&amp;pg=log'>登陆日志查看</a></div>"
Response.write "<div class='title'><a href='?siteid="&siteid&"&amp;pg=yanlog'>延期日志查看</a></div>"
Response.write "<div class='title'><a href='cdk.asp?siteid="&siteid&"'>cdk生产及管理</a></div>"
Response.write "<div class='title'><a href='cdk.asp?siteid="&siteid&"&pg=cdklog'>cdk使用日志</a></div>"

elseif pg="log" then
Response.write "<div class=""title""><a href=""?siteid="&siteid&"&amp;pg=logg"">清除全部登陆日志</a></div>"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_log] order by id desc",conn,1,1
If Not rs.eof Then						
	gopage="?pg=log&amp;"
	Count=rs.recordcount
	page=int(request("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
   if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
%>
<form name="formDel" method="post" action="?pg=sc2&amp;siteid=<%=siteid%>&amp;page=<%=page%>">
<%
call kpage(1)
	For i=1 To PageSize
	If rs.eof Then Exit For
if (i mod 2)=0 then Response.write "<div class=""line2"">" else Response.write "<div class=""line1"">"
Response.write ""&page*PageSize+i-PageSize&".<input type=""checkbox"" name=""cid"" value="""&rs("id")&""">"
Response.write ""&rs("username")&"/"&rs("userid")&"登录了(在"&rs("ktime")&"/"&rs("uip")&")"
Response.write "(<a href='?page="&page&"&amp;siteid="&siteid&"&amp;pg=sc1&amp;id="&rs("id")&"'>删</a>)</div>"
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

'-----删除登录记录
elseif pg="sc1" then
id=request("id")
lx=request("lx")
page=request("page")
       conn.Execute"delete from [ksite_log] Where id="&id
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=log"
'-----
elseif pg="sc2" then
cid=request("cid")
lx=request("lx")
page=request("page")
conn.Execute("delete from [ksite_log] Where id in("&cid&")")
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=log"


'-----登录记录一键删除
elseif pg="logg" then
conn.execute("Truncate table [ksite_log]")
response.redirect "?siteid="&siteid&"&pg=log"

'-----延期记录
elseif pg="yanlog" then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_addday_log] order by id desc",conn,1,1
If Not rs.eof Then						
	gopage="?pg=yanlog&amp;"
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
Response.write ""&rs("siteid")&"延期了(在"&rs("ktime")&"/"&rs("uip")&")(<a href='?page="&page&"&amp;siteid="&siteid&"&amp;pg=scrz1&amp;id="&rs("id")&"'>删</a>)</div></div><div class='content'>"
k_lx=clng(rs("klx"))
    if k_lx=0 then Response.write "&nbsp;延期方式：普通延期"&rs("kday")&"天" 
	if k_lx=1 then Response.write "CDK 过期+"&rs("kday")&"天 cdk："&rs("kcdk")&""
	if k_lx=2 then Response.write "CDK 当天+"&rs("kday")&"天 cdk："&rs("kcdk")&""
	if k_lx=3 then Response.write "网站升级为VIP网站"
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
'-----删除延期记录
elseif pg="scrz1" then
id=request("id")
lx=request("lx")
page=request("page")
       conn.Execute"delete from [ksite_addday_log] Where id="&id
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=yanlog"
'-----
elseif pg="scrz2" then
cid=request("cid")
lx=request("lx")
page=request("page")
conn.Execute("delete from [ksite_addday_log] Where id in("&cid&")")
Response.redirect"?siteid="&siteid&"&page="&page&"&pg=yanlog"

'-----安装sql字段
elseif pg="az" then
conn.Execute("CREATE TABLE [dbo].[ksite_log] (ID int IDENTITY (1,1) not null PRIMARY key)")
conn.Execute("ALTER TABLE [ksite_log] ADD [userid] bigint")
conn.Execute("ALTER TABLE [ksite_log] ADD [siteid] bigint")
conn.Execute("ALTER TABLE [ksite_log] ADD [username] varchar(50)")
conn.Execute("ALTER TABLE [ksite_log] ADD [ktime] datetime")
conn.Execute("ALTER TABLE [ksite_log] ADD [uip] ntext")

conn.Execute("CREATE TABLE [dbo].[ksite_addday_log] (ID int IDENTITY (1,1) not null PRIMARY key)")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [userid] bigint")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [siteid] bigint")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [kcdk] varchar(50)")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [ktime] datetime")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [klx] bigint")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [kday] bigint")
conn.Execute("ALTER TABLE [ksite_addday_log] ADD [uip] ntext")

conn.Execute("CREATE TABLE [dbo].[ksite_cdk] (ID int IDENTITY (1,1) not null PRIMARY key)")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [kcdk] varchar(50)")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [kday] bigint")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [klx] bigint")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [ktime] datetime")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [ksy] bigint")
conn.Execute("ALTER TABLE [ksite_cdk] ADD [userid] bigint")

conn.Execute("CREATE TABLE [dbo].[ksite_cdk_log] (ID int IDENTITY (1,1) not null PRIMARY key)")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [userid] bigint")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [kcdk] varchar(50)")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [siteid] bigint")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [klx] bigint")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [kday] bigint")
conn.Execute("ALTER TABLE [ksite_cdk_log] ADD [uip] ntext")

response.redirect "?siteid="&siteid

'-----删除sql字段
elseif pg="del" then
conn.Execute("DROP TABLE [dbo].[ksite_log]")
conn.Execute("DROP TABLE [dbo].[ksite_cdk]")
conn.Execute("DROP TABLE [dbo].[ksite_cdk_log]")
conn.Execute("DROP TABLE [dbo].[ksite_addday_log]")
response.redirect "index.asp?siteid="&siteid

end if
call ksite_end
%>