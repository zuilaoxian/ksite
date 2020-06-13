<!--#include file="../api/conn.asp"-->
<%
response.expires=-1
q=request.querystring("q")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [ksite_cdk] where kcdk='"&q&"'",conn,1,1
if rs.bof or rs.eof then
Response.write"不存在的cdk"
else
  if clng(rs("ksy"))=1 then
  Response.write"<br/>可用&nbsp;"
  k_lx=clng(rs("klx"))
    if k_lx=1 then Response.write"类型:结束日期+"&rs("kday")&"天"
    if k_lx=2 then Response.write"类型:当前日期+"&rs("kday")&"天"
    if k_lx=3 then Response.write"类型:网站升级为vip网站"
  else
  Response.write"已被使用"
  end if
end if
rs.close
set rs=nothing
%>