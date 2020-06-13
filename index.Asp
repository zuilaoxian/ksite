<!--#include file="head.asp"-->
<title>自助延期</title>

<script type="text/javascript">
function showHint(str)
{
var xmlhttp;
if (str.length==0)
  {
  document.getElementById("txtHint").innerHTML="";
  return;
  }
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
    document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","./inc/keycdk.asp?q="+str,true);
xmlhttp.send();
}
</script>

<%call kquanxian(0)
'数据库检测代码
conn.execute("select * from [ksite_log]")
If Err Then 
err.Clear
Response.Write "<hr>功能还未被配置"
if kuserid=1000 and ksiteid=1000 then Response.Write "  <a href='admin.asp?pg=az'>安装sql字段</a>"
Response.Write "<hr>"
end if

'-----查询本周已延期次数
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [ksite_addday_log] where userid="&kuserid&" and klx=0 and datediff(month,ktime,getdate())=0",conn,1,1
ksite_cday=rs.recordcount
rs.close
set rs=nothing
'本月延期次数
ksite_times=ksite_mday-ksite_cday
'-----查询子站到期时间
set rs=server.CreateObject("adodb.recordset")
rs.open "select userid,endtime,sitevip from [user] where userid="&kuserid,conn,1,1
endtime=rs("endtime")
sitevip=rs("sitevip")
rs.close
set rs=nothing
if sitevip=1 then k_sitevip="(vip网站)"
if isnull(endtime) or Instr(endtime,"999") or kuserid=1000 then
Response.Write "<div class='tip'>到期时间：无限期"&k_sitevip&"</div>"
else 
Response.Write "<div class='tip'>到期时间："&endtime&"("&-dateDiff("d",endtime,now)&")"&k_sitevip&"</div>"
end if
pg=request("pg")
if pg="" then
%>
<div class='title'><a href='?pg=yanqi1'>自助延期</a></div>
<div class='line1'>每月<%=ksite_mday%>次,每次<%=ksite_aday%>天,剩余<%=ksite_times%>次</div>


<div class='tip'>CDK延期</div>
<form method='post' action='?siteid=<%=siteid%>&pg=yanqi2'>
<div class='line2'><input name='cdk' value='' id='txt1' onkeyup='showHint(this.value)'><span id='txtHint'></span></div>
<div class='line1'><input type='submit' value='确定使用'></form></div>

<%
elseif pg="yanqi1" then
	'次数用完、永久、1000，不可延期
if isnull(endtime) or Instr(endtime,"999") or kuserid=1000 then ksite_err("您是永久用户，无需延期")
if ksite_times<=0 then ksite_err("本月延期次数已经用完")
	'如果不是1000正站长才可延期
	if kuserid<>1000 and kuserid<>"" and kuserid>1000 then
	'记录延期日志
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from [ksite_addday_log]",conn,1,3
	rs.addnew
	rs("userid")=kuserid
	rs("siteid")=ksiteid
	rs("ktime")=now()
	rs("klx")=0
	rs("kday")=ksite_aday
	rs("uip")=userip
	rs.update
	rs.close
	set rs=nothing
	conn.Execute("update [user] set endtime='"&DateAdd("d",ksite_aday,endtime)&"' where userid="&kuserid)
	Response.Write "<div class='tip'>延期成功<br/>结束日期:"&DateAdd("d",ksite_aday,endtime)&"</div>"
	end if
elseif pg="yanqi2" then
	cdk=request("cdk")
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from [ksite_cdk] where kcdk='"&cdk&"'",conn,1,1
	if rs.bof or rs.eof then
	ksite_err("不存在的cdk")
	else
	ksy=clng(rs("ksy"))
	klx=clng(rs("klx"))
	kday=clng(rs("kday"))
	rs.close
	set rs=nothing

  if ksy=1 then
    if klx=1 then'类型为从子站过期日追加
	if kuserid<>1000 and kuserid<>"" and kuserid>1000 then conn.Execute("update [user] set endtime='"&DateAdd("d",kday,endtime)&"' where userid="&kuserid)
	Response.Write "<div class='tip'>延期成功<br/>结束日期:"&DateAdd("d",kday,endtime)&"</div>"
    end if
	if klx=2 then'类型为从即日起追加
	if kuserid<>1000 and kuserid<>"" and kuserid>1000 then conn.Execute("update [user] set endtime='"&DateAdd("d",kday,now)&"' where userid="&kuserid)
	Response.Write "<div class='tip'>延期成功<br/>结束日期:"&DateAdd("d",kday,now)&"</div>"
	end if
	if klx=3 then'类型为升级vip网站
	if sitevip=1 then ksite_err("你的网站已经是vip网站，无需升级")
	if kuserid<>1000 and kuserid<>"" and kuserid>1000 then conn.Execute("update [user] set sitevip=1 where userid="&kuserid)
	Response.Write "<div class='tip'>升级网站为vip成功</div>"
    end if
	'记录延期日志
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from [ksite_addday_log]",conn,1,3
	rs.addnew
	rs("userid")=kuserid
	rs("siteid")=ksiteid
	rs("ktime")=now()
	rs("kcdk")=cdk
	rs("klx")=klx
	rs("kday")=kday
	rs("uip")=userip
	rs.update
	rs.close
	set rs=nothing
	'记录cdk使用日志
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from [ksite_cdk_log]",conn,1,3
	rs.addnew
	rs("userid")=kuserid
	rs("siteid")=ksiteid
	rs("ktime")=now()
	rs("kcdk")=cdk
	rs("klx")=klx
	rs("kday")=kday
	rs("uip")=userip
	rs.update
	rs.close
	set rs=nothing
	'修改cdk状态为已使用
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from [ksite_cdk] where kcdk='"&cdk&"'",conn,1,3
	rs("ksy")=2
	rs.update
	rs.close
	set rs=nothing
  else
 ksite_err("cdk"&cdk&"已被使用")
  end if
end if

elseif pg="yanlog" then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from [ksite_addday_log] where userid="&kuserid&" order by id desc",conn,1,1
If Not rs.eof Then						
	gopage="?pg=yanlog&amp;"
	Count=rs.recordcount
	page=int(request("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
call kpage(1)
	For i=1 To PageSize
	If rs.eof Then Exit For
if (i mod 2)=0 then Response.write "<div class=""line2"">" else Response.write "<div class=""line1"">"
Response.write ""&page*PageSize+i-PageSize&"."
Response.write "自助延期(在"&rs("ktime")&")</div></div><div class='content'>"
k_lx=clng(rs("klx"))
    if k_lx=0 then Response.write "&nbsp;延期方式：普通延期"&rs("kday")&"天" 
	if k_lx=1 then Response.write "CDK 过期+"&rs("kday")&"天 cdk："&rs("kcdk")&""
	if k_lx=2 then Response.write "CDK 当天+"&rs("kday")&"天 cdk："&rs("kcdk")&""
	if k_lx=3 then Response.write "网站升级为VIP网站"
    Response.write "</div>"
	rs.movenext
 	Next
call kpage(2)
else
   Response.write "<div class=tip>暂时没有记录！</div>"
end if
rs.close
set rs=nothing

end if
call ksite_end
%>