<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0, minimum-scale=1.0, maximum-scale=0">
<!--#include file="./inc/config.asp"-->
<!--#include file="./inc/Function.asp"-->
<!--#include file="./api/conn.asp"-->
<!--#include file="./api/md5.asp"-->
<link rel="stylesheet" href="css/css.css?v=<%=minute(now)%>" type="text/css" />
</head>
<body>
<script type="text/javascript">
function ConfirmDel(message)
{
   if (confirm(message))
   {
   document.formDel.submit();
   }else{
   }
}
</script>
<style type="text/css">
.line1{font-size:11px;}
</style> 
<%
if ksite<>"" then
%>
用户名:<%=username%> ID:<%=kuserid%> 站点ID:<%=ksiteid%> <%if kuserid=1000 then Response.write "&nbsp;<a href='admin.asp'>后台管理</a>"
Response.write "&nbsp;<a href='index.asp?pg=yanlog'>延期记录</a>"%>
<%end if%>