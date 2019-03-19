<!--#include file="session_cookie.asp" -->
<!--#include file="function/function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Pragma" CONTENT="no-cache"> 
<meta http-equiv="Cache-Control" CONTENT="no-cache"> 
<meta http-equiv="Expires" CONTENT="0"> 
<title>主页</title>
<link href="css/globle.css" rel="stylesheet" type="text/css" />
<link href="css/<%= userskin() %>.css" rel="stylesheet" type="text/css" />
<style>
body{ color:#333333}
.conf_table td{	line-height:23px; height:23px}
.conf_table{ margin-top:20px; margin-left:28px}
.indexstate{ float:left;}
.aboutbar td{ font-size:14px;}
.aboutbar span,.aboutbar div {white-space:nowrap}
.explain{ display:block; height:30px; line-height:30px}
.sx{ border-bottom:dotted #666 1px; margin-left:3px; margin-right:3px;color:#666}
</style>
<script language="javascript" src="js/tab2.js"></script>
<script language="javascript" src="js/FusionCharts/FusionCharts.js"></script>
</head> 

<body id="bodyc">
<%
openconn()
ztdate=DateAdd("d",-1,now()) '得到昨天的日期

year_str=cstr(year(now()))	'默认按当前年份来显示
month_str=cstr(month(now()))	'默认按当前年份、当前月份来显示。必须是两位才行。如4月，写成04
todayn=conn.execute("SELECT count(0) FROM 销售订单 WHERE Year([下单日期])="&year(now())&" and month([下单日期])="&month(now())&" and day([下单日期])="&day(now())&" "&yh&" ")(0)
yedayn=conn.execute("SELECT count(0) FROM 销售订单 WHERE Year([下单日期])="&year(ztdate)&" and month([下单日期])="&month(ztdate)&" and day([下单日期])="&day(ztdate)&" "&yh&" ")(0)
%>

<div class="aboutbar them_c">
<table style="width:100%"><tr>
<td align="left" style="padding-right:25px;"><span>全部用户 今日总订单数：<%= todayn %> 个； 昨日总订单数：<%= yedayn %> 个</span></td>
<td align="right" style="padding-right:10px">
<span style="font-size:12px;font-weight:normal"><script language="javascript" src="js/today.js"></script></span>
</td>
</tr></table>
</div>

<div class="indexstate">
<table class="conf_table">
<tr><td><strong class="them_c">今日订单详细</strong></td></tr>
<tr><td valign="top" style="padding-right:0px">
<%
sql="SELECT (select 店铺名称 from 店铺名称 where id=销售订单.店铺) as 店铺, count(0) FROM 销售订单 WHERE Year([下单日期])="&year(now())&" and month([下单日期])="&month(now())&" and day([下单日期])="&day(now())&"  "&yh&"  GROUP BY 店铺"
set rs=conn.execute(sql)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "【"&Array_str(0,i)&"】有 <b class='bigfont'>"&Array_str(1,i)&"</b> 个订单<br>"
	next
else
	response.Write "<span class=explain>无相应订单记录</span>"
end if
closers(rs)
%></td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">发生的退换货订单</strong></td></tr>
<tr><td nowrap="nowrap" style="padding-right:0px"><%
set rs=conn.execute("select (select 店铺名称 from 店铺名称 where id=销售订单.店铺) as 店铺名称,count(0),店铺 from 销售订单 where (订单步骤>=4) "&yh&" group by 店铺")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "【"&Array_str(0,i)&"】有 <span class='bigfont b'>"&Array_str(1,i)&"</span> 个退换货订单！ <br>"
	next	
else
	response.Write "<span class=explain>无相应订单记录</span>"
end if
closers(rs)
%>
</td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">等待确认的退返件</strong></td></tr>
<tr><td nowrap="nowrap" style="padding-right:0px;">
<%
sql="select "&_
" (select 店铺名称 from 店铺名称 where id=temp.店铺), "&_
" count(0), "&_
" max(收到返件日期) "&_
" from (select (select 店铺 from 销售订单 where ddid=返件.ddid) as 店铺,收到返件日期 from 返件 where 订单确认=0 ) as temp "&_
" group by 店铺 order by  max(收到返件日期) desc "
'response.Write sql
set rs=conn.execute(sql)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "【"&Array_str(0,i)&"】有 <span class='bigfont b'>"&Array_str(1,i)&"</span> <span style='color:#FF3300'>个退返件待确认.. </span><br>"
	next	
else
	response.Write "<span class=explain>无等待确认的退返件</span>"
end if
closers(rs)
%>

</td></tr>
</table>
</div>


<div class="indexstate">
<table class="conf_table">
<tr><td><strong class="them_c">等待出库的订单</strong></td></tr>
<tr><td nowrap="nowrap"><%
set rs=conn.execute("select (select 店铺名称 from 店铺名称 where id=销售订单.店铺) as 店铺名称,count(0),店铺 from 销售订单 where 订单状态=1 and 订单步骤=1 "&yh&" group by 店铺")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "【"&Array_str(0,i)&"】有 <span class='bigfont b'>"&Array_str(1,i)&"</span> 个订单配货中.. <br>"
	next
else
	response.Write "<span class=explain>无相应订单记录</span>"
end if
closers(rs)
%>
</td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">等待发货的订单</strong></td></tr>
<tr><td nowrap="nowrap"><%
set rs=conn.execute("select (select 店铺名称 from 店铺名称 where id=销售订单.店铺) as 店铺名称,count(0),店铺 from 销售订单 where 订单状态=1 and 订单步骤=2 "&yh&" group by 店铺")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "【"&Array_str(0,i)&"】有 <span class='bigfont b'>"&Array_str(1,i)&"</span> 个订单待发货.. <br>"
	next
else
	response.Write "<span class=explain>无相应订单记录</span>"
end if
closers(rs)
%>
</td></tr>
</table>

</div>

<div class="indexstate">
<table class="conf_table">
<tr>
  <td><strong class="them_c">即将缺货的部分商品</strong></td></tr>
<tr><td nowrap="nowrap">
<%
fazhi=conn.execute("select fazhi from 设置")(0)
set rs=conn.execute("select top 5 (select 商品名称 from 商品库 where 库存.上架编号=上架编号) as 商品名称,sum(库存剩余) as 库存剩余,(select 分类单位 from 商品分类 where id=max(库存.分类id)) as 分类单位 from 库存 group by 上架编号 having sum(库存剩余)<="&fazhi&" order by 库存剩余")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		spname=Array_str(0,i)
		if len(spname)>12 then spname=left(spname,12)&"..."
		response.Write "<span class=""sx"">"&spname&"</span><b>"&Array_str(1,i)&"</b> "&Array_str(2,i)&"<br>"
	next
	response.Write "... ... ..."	
else
	response.Write "<span class=explain>无缺量信息</span>"
end if
closers(rs)
%>
</table>
</div>

<%
sqllist=" select top 20 cast(month(下单日期) as varchar)+'-'+cast(day(下单日期) as varchar)+'',count(0),max(下单日期) from 销表数据 where datediff(day,下单日期,getdate())<=20 group by cast(month(下单日期) as varchar)+'-'+cast(day(下单日期) as varchar)+'' order by max(下单日期) "
set rs=conn.execute(sqllist)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		namestr=Array_str(0,i)&""
		namestr=right(namestr,len(namestr)-InStr(namestr,"-"))
		tempxml=tempxml+"<set name='"&namestr&"日' value='"&Array_str(1,i)&"' />"
		'randomize  '演示用的	
		'tempxml=tempxml+"<set name='"&namestr&"日' value='"&Array_str(1,i)&""&int(rnd*36)&"' />" '演示用的	
	next
	xmldata=tempxml
end if
if xmldata="" then xmldata="<set name=' ' value='' />"

closeconn()
%>

<div class="indexstate" style="padding-left:27px; margin-top:10px; clear:both;position:relative;">
<div class="them_c" style="padding:15px 0px 10px 6px;"><b>近20天内订单量曲线</b></div>
<div style="position:absolute; top:30px; height:120px; width:810px;filter:alpha(opacity=0);opacity:0.0; "><img src="images/charts.jpg" /></div>
<div id="chartdiv"></div>
</div>

<script>
var tttt="<graph chartTopMargin='0' chartRightMargin='15' chartLeftMargin='10' chartBottomMargin='0' bgColor='#ffffff' showBorder='0' canvasBorderThickness='1' canvasBorderColor='#cccccc'   lineThickness='2' baseFontSize='12' showYAxisValues='0' animation='1' decimalPrecision='0' yAxisMaxValue='1' formatNumberScale='0' ><%= xmldata %></graph>";
var chart1 = new FusionCharts("js/fusioncharts.com.Charts/Line.swf", "chart1Id", "800", "110", "0", "1");
chart1.addParam("wmode","Opaque") //不处于最上层
chart1.setDataXML(tttt);
chart1.render("chartdiv");
</script>

</body>
</html>