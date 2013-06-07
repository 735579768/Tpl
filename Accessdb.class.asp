<%
'==========================================================================
'文件名称：Accessdb.class.asp
'功　　能：数据库操作类
'===author 赵克立
'===blog:http://zhaokeli.com/
'===version.v1.0.0
'版权声明：可以在任意作品中使用本程序代码，但请保留此版权信息。
'          如果你修改了程序中的代码并得到更好的应用，请发送一份给我，谢谢。
'==========================================================================
'const db_path="data/#aspadmindata.mdb"
const db_pwd=""
Class Accessdb
	public kl_conn
	private kl_rs
	private kl_sql
	private kl_err
	private kl_sqlkey
	'==================================
	'初始化类函数
	'功能：初始化数据
	'==================================
	Private Sub Class_Initialize()
		connstr= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.MapPath(Sql_Data)&";Jet OLEDB:Database Password="&db_pwd&";"
		Set kl_conn = Server.CreateObject("ADODB.Connection")
		Set kl_sqlkey = server.CreateObject("Scripting.Dictionary")
		kl_Conn.Open connstr
	End Sub
	Private Sub Class_Terminate()
		Set kl_conn = Nothing
	End Sub
	Public Property Let Conn(pdbConn)
		If IsObject(pdbConn) Then
			Set kl_conn = pdbConn
		End If
	End Property
	Public Property Get conn()
		conn = kl_conn
	End Property
	'==================================
	'query函数
	'功能：查询记录集
	'==================================
	Function query(sqlstr)
		on error resume next
		err.clear
		if sqlstr<>"" then kl_sql=sqlstr
		set kl_rs=server.CreateObject("adodb.recordset")
		kl_rs.open kl_sql,kl_conn,1,3
		if err.number<>0 then 
			echoerr 0,"query SQL语句出错:"&kl_sql
		end if
		set query=kl_rs
	End Function
	'==================================
	'rsToArr函数
	'功能：把记录集转成键值对数组
	'==================================
	Function rsToArr(rss)
	on error resume next
	err.clear
		num=rss.recordcount
		redim klvalarr(num)
		if num>0 then
			for i=0 to num-1 
				if not rss.eof then 	
					set kl_keyval = server.CreateObject("Scripting.Dictionary")
					for each a in rss.fields
						kl_keyval(a.name)=rss(a.name)
					next
					set klvalarr(i)=kl_keyval
					set kl_keval=nothing
					rss.movenext
				end if
			next
		end if
		if err.number<>0 then
			echo "rsToArr convert keyval error:"&err.description
		end if
		rsToArr=klvalarr
	End Function
	'======================sql连贯操作函数start==============
	Function  where(str)
		kl_sqlkey("where")=str
		set where=Me
	End Function
	Function  table(str)
		kl_sqlkey("table")=str
		set table=Me
	End Function
	Function  top(str)
		kl_sqlkey("top")=" top "&str
		set top=Me
	End Function
	Function  jin(str)
		kl_sqlkey("jin")=str
		set jin=Me
	End Function
	Function  fild(str)
		kl_sqlkey("fild")=str
		set fild=Me
	End Function
	Function  order(str)
		kl_sqlkey("order")=str
		set order=Me
	End Function
	'==================================
	'sel函数
	'功能：查询数据
	'==================================
	Function sel()
		'on error resume next
		dim top ,fild,where,jin,table,order
		if kl_sqlkey.Exists("top") then top=cstr(kl_sqlkey("top"))
		if kl_sqlkey.Exists("fild") then fild=cstr(kl_sqlkey("fild"))
		if kl_sqlkey.Exists("where") then where="and " &cstr(kl_sqlkey("where"))
		if kl_sqlkey.Exists("jin") then jin=cstr(kl_sqlkey("jin"))
		if kl_sqlkey.Exists("table") then table=cstr(kl_sqlkey("table"))
		if kl_sqlkey.Exists("order") then order=cstr(kl_sqlkey("order"))
		if table="" then call echoerr(0,"数据表不能为空")
		if fild="" then fild="*"
		if jin<>"" then jin="inner join "&jin
		if order<>"" then order=" order by "&order
		kl_sql="select "&top&" "&fild&" from "&table &" "& jin&" "&"where  1=1  "&where&" "&order
		set sel=query("")
		kl_sqlkey.removeAll()
	End Function
	'======================sql连贯操作函数start==============
	'==================================
	'echoErr函数
	'功能：把不同的错误级别输出
	'==================================
	Function echoErr(errnum,errstr)
		select case errnum
			case 0:'致命错误
				kl_err="Error Description:"&err.description
				response.Write errstr&"<br>"
				response.Write kl_err
				response.End()	
			case 1:'一般错误
				kl_err="Error Description:"&err.description
				response.Write errstr&"<br>"
				response.Write kl_err
			case else:'其它
				response.Write errstr&"<br>"
				response.Write("")
		end select
		err.clear
	End Function
	Function getlastsql()
		getlastsql=kl_sql
	end function
End Class
%>