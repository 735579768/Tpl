<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Tpl.class.asp"-->
<!--#include file="Accessdb.class.asp"-->
<!--#include file="lang.asp"-->
<%
on error goto 0
set tpl=New Asptpl
set db=new AccessDb

'打开新数据库
		connstr= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.MapPath("data/gjnew.mdb")&";Jet OLEDB:Database Password="&db_pwd&";"
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open connstr

set rs=db.table("wh_info").where("btypeid=2").sel()
do while not rs.eof

		set rss=server.CreateObject("adodb.recordset")
		rss.open "select * from kl_archives",conn,1,3
		rss.addnew
		'导出新闻

		rss("arctitle")=rs("title")&""
		rss("arccontent")=rs("content")&""
		rss("arcdescr")=left(rs("content")&"",30)
		rss("cat_id")=59
		rss("type_id")=1
		rss("arcsource")=rs("nfrom")&""
		rss("archits")=rs("hits")&""
		rss("fbdate")=rs("date")&""
		rss("uddate")=rs("date1")&""
		rss("arcpic")=replace(rs("image")&"","/industry/Up/newsimage","/uploads/image/oldpic")
		rss("arcauthor")=rs("writer")&""
		rss("arctpl")="article_detail.html"
		'===============================================
		rss.update
		rss.close
		set rss=nothing

	rs.movenext
loop










response.End()


a=array("title")
b=array("网站首页")
set page=tpl.getkeyvalue(a,b)
tpl.assign "page",page
set rs=db.query("select * from kl_archives")

'连贯操作查询
db.table("kl_archives").where("id=45").fild("*").top("8").jin("kl_cats on kl_archives.cat_id=kl_cats.cat_id").sel
'变量测试
rsarr=db.rstoarr(rs)
if isarray(rsarr) then echo "is array"
echo ubound(rsarr)
for each a in rsarr
	if isobject(a) then
		if a.Exists("arctitle") then echo a("arctitle")&"<br>"
	end if
next

tpl.assign "a","a123456789"
tpl.assign "b",""'默认输出测试

'foreach数组测试
arr=array("arr1","arr2")
tpl.assign "arrr",arr


'tpl.assign "arr",arr

'组合键值对象
key=array("key1","key2","key3")
val=array("5","2222222","3")
set kvarr=tpl.getkeyvalue (key,val)


'传递一个键值对象
tpl.assign "obj",kvarr


'组合一个键值对象数组
arr=array(kvarr,kvarr)
tpl.assign "loopobjarr",rsarr

tpl.display("index.html")
'p_var_list.add "b","a"
'p_var_list.add "b","bbb"
'redim arr(10)
'arr=array("a","b")
'p_var_list.add "b",arr
'p_var_list("b")=arr
'echo isarray(p_var_list("b"))
'echo p_var_list("b")(0)


'for each i in p_var_list.item
'	echo i
'next
Function echo(str)
	response.Write str
End Function
%>