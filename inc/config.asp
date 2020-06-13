<%
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR")

'-----容错过滤，1是，其他否
kerrtop=1

'-----操作过期时间
ksitetime=10

'-----每月延期次数
ksite_mday=2

'-----每月延期，延期天数
ksite_aday=17

'-----页面列表数量
PageSize=10

'-----时间设置
ksite_time=year(date)&"-"&month(date)&"-"&day(date)&"　"&time()&"　"&weekdayname(weekday(now))

if kerrtop=1 then On Error Resume Next

ksite=session("ksite")
klog=session("klog")
if ksite<>"" and Instr(ksite,"_") then
kstr=Split(ksite,"_")
username=kstr(0)
kuserid=clng(kstr(2))
ksiteid=clng(kstr(1))
end if
%>