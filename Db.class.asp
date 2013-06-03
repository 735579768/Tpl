<%
'===================================================
'===author 赵克立
'===blog:http://zhaokeli.com/
'===version.v1.0.0
'===access class
'===运行模式类似于smarty不过没有smary强大
'===================================================
const db_path="data/#aspadmindata.mdb"
const db_pwd=""
class AccessDb
	private connkl
	private errstr
	private sqlstr
	private rsset
	public fieldArr
	private debug
	Private Sub Class_Initialize()
		debug = true					'调试模式是否开启
		connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.MapPath(db_path)&";Jet OLEDB:Database Password="&db_pwd&";"
		Set connkl = Server.CreateObject("ADODB.Connection")
		connkl.Open connstr
		'connkl.execute("SET NAMES 'utf8")
	End Sub
	
	Private Sub Class_Terminate()
		Set connkl = Nothing
	End Sub
	
	public Function query(sqlstr)
		set rsset=server.CreateObject("adodb.recordset")
		rsset.open sqlstr,connkl,1,3
		'处理数据
			redim arr(rsset.fields.count)
			for i=0 to rsset.fields.count - 1
				arr(i)=rsset.fields(i).name
			next
			fieldArr=arr
		set query=rsset
	End Function
	
	public Function rsToArr(rss)
		redim rsarr(rss.recordcount)
		if rss.recordcount>0 then 
			for i=0 to rss.recordcount-1
				set keyvalarr=server.CreateObject("Scripting.Dictionary")
				for each  a in rs.fields
					keyvalarr(a.name)=rss(a.name)
				next
				set rsarr(i)=keyvalarr
				set keyvalarr=nothing
			rss.movenext
			next
		end if
		set rss=nothing
		rsToArr=rsarr
		set keyvalarr=nothing
	End Function
end class
%>