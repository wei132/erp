<!--#include file="session_cookie.asp" -->
<!--#include file="function/function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Pragma" CONTENT="no-cache"> 
<meta http-equiv="Cache-Control" CONTENT="no-cache"> 
<meta http-equiv="Expires" CONTENT="0"> 
<title>��ҳ</title>
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
ztdate=DateAdd("d",-1,now()) '�õ����������

year_str=cstr(year(now()))	'Ĭ�ϰ���ǰ�������ʾ
month_str=cstr(month(now()))	'Ĭ�ϰ���ǰ��ݡ���ǰ�·�����ʾ����������λ���С���4�£�д��04
todayn=conn.execute("SELECT count(0) FROM ���۶��� WHERE Year([�µ�����])="&year(now())&" and month([�µ�����])="&month(now())&" and day([�µ�����])="&day(now())&" "&yh&" ")(0)
yedayn=conn.execute("SELECT count(0) FROM ���۶��� WHERE Year([�µ�����])="&year(ztdate)&" and month([�µ�����])="&month(ztdate)&" and day([�µ�����])="&day(ztdate)&" "&yh&" ")(0)
%>

<div class="aboutbar them_c">
<table style="width:100%"><tr>
<td align="left" style="padding-right:25px;"><span>ȫ���û� �����ܶ�������<%= todayn %> ���� �����ܶ�������<%= yedayn %> ��</span></td>
<td align="right" style="padding-right:10px">
<span style="font-size:12px;font-weight:normal"><script language="javascript" src="js/today.js"></script></span>
</td>
</tr></table>
</div>

<div class="indexstate">
<table class="conf_table">
<tr><td><strong class="them_c">���ն�����ϸ</strong></td></tr>
<tr><td valign="top" style="padding-right:0px">
<%
sql="SELECT (select �������� from �������� where id=���۶���.����) as ����, count(0) FROM ���۶��� WHERE Year([�µ�����])="&year(now())&" and month([�µ�����])="&month(now())&" and day([�µ�����])="&day(now())&"  "&yh&"  GROUP BY ����"
set rs=conn.execute(sql)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "��"&Array_str(0,i)&"���� <b class='bigfont'>"&Array_str(1,i)&"</b> ������<br>"
	next
else
	response.Write "<span class=explain>����Ӧ������¼</span>"
end if
closers(rs)
%></td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">�������˻�������</strong></td></tr>
<tr><td nowrap="nowrap" style="padding-right:0px"><%
set rs=conn.execute("select (select �������� from �������� where id=���۶���.����) as ��������,count(0),���� from ���۶��� where (��������>=4) "&yh&" group by ����")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "��"&Array_str(0,i)&"���� <span class='bigfont b'>"&Array_str(1,i)&"</span> ���˻��������� <br>"
	next	
else
	response.Write "<span class=explain>����Ӧ������¼</span>"
end if
closers(rs)
%>
</td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">�ȴ�ȷ�ϵ��˷���</strong></td></tr>
<tr><td nowrap="nowrap" style="padding-right:0px;">
<%
sql="select "&_
" (select �������� from �������� where id=temp.����), "&_
" count(0), "&_
" max(�յ���������) "&_
" from (select (select ���� from ���۶��� where ddid=����.ddid) as ����,�յ��������� from ���� where ����ȷ��=0 ) as temp "&_
" group by ���� order by  max(�յ���������) desc "
'response.Write sql
set rs=conn.execute(sql)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "��"&Array_str(0,i)&"���� <span class='bigfont b'>"&Array_str(1,i)&"</span> <span style='color:#FF3300'>���˷�����ȷ��.. </span><br>"
	next	
else
	response.Write "<span class=explain>�޵ȴ�ȷ�ϵ��˷���</span>"
end if
closers(rs)
%>

</td></tr>
</table>
</div>


<div class="indexstate">
<table class="conf_table">
<tr><td><strong class="them_c">�ȴ�����Ķ���</strong></td></tr>
<tr><td nowrap="nowrap"><%
set rs=conn.execute("select (select �������� from �������� where id=���۶���.����) as ��������,count(0),���� from ���۶��� where ����״̬=1 and ��������=1 "&yh&" group by ����")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "��"&Array_str(0,i)&"���� <span class='bigfont b'>"&Array_str(1,i)&"</span> �����������.. <br>"
	next
else
	response.Write "<span class=explain>����Ӧ������¼</span>"
end if
closers(rs)
%>
</td></tr>
</table>

<table class="conf_table">
<tr><td><strong class="them_c">�ȴ������Ķ���</strong></td></tr>
<tr><td nowrap="nowrap"><%
set rs=conn.execute("select (select �������� from �������� where id=���۶���.����) as ��������,count(0),���� from ���۶��� where ����״̬=1 and ��������=2 "&yh&" group by ����")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		response.Write "��"&Array_str(0,i)&"���� <span class='bigfont b'>"&Array_str(1,i)&"</span> ������������.. <br>"
	next
else
	response.Write "<span class=explain>����Ӧ������¼</span>"
end if
closers(rs)
%>
</td></tr>
</table>

</div>

<div class="indexstate">
<table class="conf_table">
<tr>
  <td><strong class="them_c">����ȱ���Ĳ�����Ʒ</strong></td></tr>
<tr><td nowrap="nowrap">
<%
fazhi=conn.execute("select fazhi from ����")(0)
set rs=conn.execute("select top 5 (select ��Ʒ���� from ��Ʒ�� where ���.�ϼܱ��=�ϼܱ��) as ��Ʒ����,sum(���ʣ��) as ���ʣ��,(select ���൥λ from ��Ʒ���� where id=max(���.����id)) as ���൥λ from ��� group by �ϼܱ�� having sum(���ʣ��)<="&fazhi&" order by ���ʣ��")
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		spname=Array_str(0,i)
		if len(spname)>12 then spname=left(spname,12)&"..."
		response.Write "<span class=""sx"">"&spname&"</span><b>"&Array_str(1,i)&"</b> "&Array_str(2,i)&"<br>"
	next
	response.Write "... ... ..."	
else
	response.Write "<span class=explain>��ȱ����Ϣ</span>"
end if
closers(rs)
%>
</table>
</div>

<%
sqllist=" select top 20 cast(month(�µ�����) as varchar)+'-'+cast(day(�µ�����) as varchar)+'',count(0),max(�µ�����) from �������� where datediff(day,�µ�����,getdate())<=20 group by cast(month(�µ�����) as varchar)+'-'+cast(day(�µ�����) as varchar)+'' order by max(�µ�����) "
set rs=conn.execute(sqllist)
if not rs.eof then
	Array_str=rs.GetRows()
	for i=0 to UBound(Array_str,2)
		namestr=Array_str(0,i)&""
		namestr=right(namestr,len(namestr)-InStr(namestr,"-"))
		tempxml=tempxml+"<set name='"&namestr&"��' value='"&Array_str(1,i)&"' />"
		'randomize  '��ʾ�õ�	
		'tempxml=tempxml+"<set name='"&namestr&"��' value='"&Array_str(1,i)&""&int(rnd*36)&"' />" '��ʾ�õ�	
	next
	xmldata=tempxml
end if
if xmldata="" then xmldata="<set name=' ' value='' />"

closeconn()
%>

<div class="indexstate" style="padding-left:27px; margin-top:10px; clear:both;position:relative;">
<div class="them_c" style="padding:15px 0px 10px 6px;"><b>��20���ڶ���������</b></div>
<div style="position:absolute; top:30px; height:120px; width:810px;filter:alpha(opacity=0);opacity:0.0; "><img src="images/charts.jpg" /></div>
<div id="chartdiv"></div>
</div>

<script>
var tttt="<graph chartTopMargin='0' chartRightMargin='15' chartLeftMargin='10' chartBottomMargin='0' bgColor='#ffffff' showBorder='0' canvasBorderThickness='1' canvasBorderColor='#cccccc'   lineThickness='2' baseFontSize='12' showYAxisValues='0' animation='1' decimalPrecision='0' yAxisMaxValue='1' formatNumberScale='0' ><%= xmldata %></graph>";
var chart1 = new FusionCharts("js/fusioncharts.com.Charts/Line.swf", "chart1Id", "800", "110", "0", "1");
chart1.addParam("wmode","Opaque") //���������ϲ�
chart1.setDataXML(tttt);
chart1.render("chartdiv");
</script>

</body>
</html>