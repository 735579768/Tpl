<!--#include file="tpl.class.asp"-->
<%
set tpl=New Asptpl
arr=array("a","b")
tpl.assign "a","aval"
tpl.assign "arr",arr
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