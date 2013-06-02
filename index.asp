<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="tpl.class.asp"-->
<!--#include file="lang.asp"-->
<%
set tpl=New Asptpl

'变量测试
tpl.assign "a","a123456789"
tpl.assign "b",""'默认输出测试

'foreach数组测试
arr=array("arr1","arr2")
tpl.assign "arrr",arr


'tpl.assign "arr",arr

'组合键值对象
key=array("key1","key2","key3")
val=array("5","2222222","3")
set kvarr=tpl.getkeyvalue ("kv",key,val)


'传递一个键值对象
tpl.assign "obj",kvarr


'组合一个键值对象数组
arr=array(kvarr,kvarr)
tpl.assign "loopobjarr",arr

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