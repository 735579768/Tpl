<%
class AspTpl
	private p_err
	private p_var_l
	private p_var_r
	private p_tag_l
	private p_tag_r
	private p_var_list
	private p_tpl_content
	private	p_reg
	private p_fs
	private p_tpl_file
	private p_charset
	public p_tpl_dir
	
	private sub class_Initialize
		p_charset="UTF-8"
		p_var_l = "\{\$"
		p_var_r = "\}"
		p_tag_l = "<\!--\{"
		p_tag_r = "\}-->"
		p_tpl_dir = "tpl"
		set p_fs= server.createobject("Scripting.FileSystemObject")
		set p_var_list = server.CreateObject("Scripting.Dictionary")
		Set p_reg = New RegExp 
		p_reg.IgnoreCase = True
		p_reg.Global = True
	end sub
	
	Private Sub Class_Terminate()
		set p_reg=nothing
		set p_fs=nothing
		set p_var_list=nothing
	end sub
	'==================================================
	'模板目录
	'==================================================	
	Property Let tpldir(a)
		p_tpl_dir=a
	End Property 
	Property Get tpldir
		tpldir=p_tpl_dir
	End Property 
	'==================================================
	'检查模板目录
	'==================================================
	private sub checkTplDirAndFile()
		tdir=server.mappath(p_tpl_dir)
		if not p_fs.FolderExists(tdir) then
			p_fs.CreateFolder(tdir)
		end if
		if p_fs.FileExists(server.mappath(p_tpl_dir &"/"& p_tpl_File)) then
				'set oFile = p_fs.OpenTextFile(server.mappath(p_tpl_dir &"/"& p_tpl_File), 1)
				p_tpl_content = ReadFromTextFile(p_tpl_dir &"/"& p_tpl_File,p_charset)'oFile.ReadAll
				'oFile.Close
			else
				echoErr "<b>ASPTemplate Error: File [" & server.mappath(p_tpl_dir &"/"& p_tpl_File) & "] does not exists!</b><br>"
			end if
	end sub
	'==================================================
	'解析模板变量
	'==================================================
	private Function jiexivar(str)
		a=p_var_list.keys
		p_reg.Pattern =  p_var_l & "(\S*?)" & p_var_r 
		set matches=p_reg.execute(str)
		for each a in matches
			'分析是否有变量过滤器start 
			p_reg.Pattern =  p_var_l &  "(\S*?)\|(\S*?)\=(\S*?)"  & p_var_r 
			set Ms=p_reg.execute(a)
			'处理有过滤器的情况
			if Ms.count>0 then
				keyname=Ms(0).SubMatches(0)'变量键名
				funcname=Ms(0).SubMatches(1)'函数名
				param=Ms(0).SubMatches(2)'调用参数
				str=replace(str,a,filtervar(p_var_list(keyname),funcname,param))
			'没有过滤器
			else
				temval=p_var_list(a.submatches(0))
				if isarray(temval) then '过滤去除数组变量
					str=replace(str,a,"<pre style='color:red;'>'"&a.submatches(0)&" ' is Array </pre>")
				else
					str=replace(str,a,temval)
				end if
			end if
			'分析是否有变量过滤器end 
			
		next
'		'遍历变量键值
'		for each i in p_var_list.keys
'			p_reg.Pattern =  p_var_l & i & p_var_r 
'			if isarray(p_var_list(i)) then '过滤去除数组变量
'				str=p_reg.replace(str,"<pre style='color:red;'>'"&i&" ' is Array </pre>")
'			else
'				str=p_reg.replace(str,p_var_list(i))
'			end if
'		next
		jiexivar=str
	end Function
	'==================================================
	'解析if标签
	'==================================================	
	private Function ifElseTag(str)
		regtag=p_tag_l & "if\s+?(.+?)\s+?"& p_tag_r  &"([\s\S]*?)"& p_tag_l & "else"& p_tag_r & "([\s\S]*?)"& p_tag_l & "/if"& p_tag_r
		p_reg.Pattern=regtag
		'p_tpl_content=p_reg.execute(p_tpl_content,"it is if ")
		Set Matches = p_reg.Execute(str)
		For Each Match in Matches  
			if eval(jiexivar(Match.SubMatches(0))) then
				str=replace(str,Match,jiexivar(Match.SubMatches(1)))
			else
				str=replace(str,Match,jiexivar(Match.SubMatches(2)))
			end if
		next
		ifElseTag=str
	end Function
	'==================================================
	'解析foreach标签
	'功能解析一维数组
	'==================================================	
	private Function foreachTag(str)
		p_reg.Pattern=p_tag_l & "foreach([\s\S]*?)" & p_tag_r&"([\s\S]*?)" & p_tag_l & "/foreach" &p_tag_r
		Set Matches = p_reg.Execute(str)
		For Each Match in Matches  
				temname=getTagParam(Match.SubMatches(0),"name")
				temvar=getTagParam(Match.SubMatches(0),"var")
				temarr=p_var_list(temname)
				str1=""
				if isarray(temarr) then 
					for each a in temarr
					p_reg.Pattern=p_var_l & temvar & p_var_r
					str1=str1&p_reg.replace(Match.SubMatches(1),a)
					next
				end if
				str=replace(str,Match,str1)
		next
		foreachTag=str
	End Function
	'==================================================
	'给模板赋值
	'==================================================
	public sub assign(key,val)
			p_var_list(key)=val
	End sub
	'===============================================================================
	'包含文件
	'在设置模板的时候首先把里面的包含标签替换
	'===============================================================================
	private function includefile(str)
		on error resume next
		p_reg.Pattern =  p_tag_l & "include(\s+?)file=[\""|\'](\S*?)[\""|\'](\s*?)" & p_tag_r 
		set matches=p_reg.Execute(str)
		if matches.count<1 then includefile=str end if
		For Each Match in Matches      ' Iterate Matches collection
			s_str=match '保存在模板文件中查到的包含标签
			m_filestr= match.SubMatches(1) '取出包含的文件名
			'判断要包含的文件是否存在
				set fs = createobject("Scripting.FileSystemObject")
			if fs.FileExists(server.mappath(p_tpl_dir  &"/"& m_filestr)) then
				'读取包含的文件并替换标签
				temstr = ReadFromTextFile(p_tpl_dir  &"/"& m_filestr,p_charset)
				if haveinclude(temstr) then temstr=includefile(temstr)	end if  '递归包含文件
				str=replace(str,s_str,temstr)'替换模板中的包含标签
			else
			response.write "<b>ASPTemplate Error: File [" & server.mappath(p_tpl_dir  &"/"& m_filestr) & "] does not exists!</b><br>"
			end if
		Next
		includefile=str
	end function
	'===============================================================================
	'判断字符串中是否有包含标签,如果有就返回true，否则返回false
	'===============================================================================
	private function haveinclude(str)
		p_reg.IgnoreCase = True
		p_reg.Global = True
		p_reg.Pattern =  p_tag_l &  "include(\s+?)file=\""(\S+?)\""(\s*?)"  & p_tag_r 
		set matches=p_reg.Execute(str)
		if matches.count<1 then
			 haveinclude=false
		else
			 haveinclude=true
		end if
	
	end function
	'==================================================
	'解析模板文件
	'==================================================
	public Function display(tplfile)
		p_tpl_file=tplfile
		checkTplDirAndFile()'载入模板
		p_tpl_content=includefile(p_tpl_content)'包含文件
		p_tpl_content=ifElseTag(p_tpl_content)
		p_tpl_content=foreachTag(p_tpl_content)
		p_tpl_content=jiexivar(p_tpl_content)
		response.Write p_tpl_content
	end Function
	'==================================================
	'输出错误
	'==================================================
	private sub echoErr(str)
		response.Write str
	end sub
	'===============================================================================
	'变量过滤器
	'@param tplvars模板变量将要赋值的字符串
	'@param funcname处理函数
	'@param param函数的参数字符串
	'===============================================================================	
	private Function filterVar(val,funcname,param)
			Select Case funcname
				case "left" filterVar=left(val,Cint(param))
				case else 	filtervar=val
			end select
	End Function
	'===============================================================================
	'取标签中的参数
	'@param str包含标签中键的字符串
	'@param key参数所在的键值
	'===============================================================================	
	Function getTagParam(str,key)
		str1=""
		p_reg.Pattern ="([\s\S]*?)"&key&"=[\""|\']([\s\S]*?)[\""|\']([\s\S]*?)"	
		set ms=p_reg.Execute(str)
		if ms.count>0 then
		str1=ms(0).SubMatches(1)'取sql语句
		end if
		set ms=nothing
		getTagParam=str1
	End Function
	'------------------------------------------------- 
	'函数名称:ReadTextFile 
	'作用:利用AdoDb.Stream对象来读取UTF-8格式的文本文件 
	'---------------------------------------------------- 
	function ReadFromTextFile (FileUrl,CharSet) 
		dim str 
		set stm=server.CreateObject("adodb.stream") 
		stm.Type=2 '以本模式读取 
		stm.mode=3 
		stm.charset=CharSet 
		stm.open 
		stm.loadfromfile server.MapPath(FileUrl) 
		str=stm.readtext 
		stm.Close 
		set stm=nothing 
		ReadFromTextFile=str 
	end function 
	'------------------------------------------------- 
	'函数名称:WriteToTextFile 
	'作用:利用AdoDb.Stream对象来写入UTF-8格式的文本文件 
	'---------------------------------------------------- 
	Sub WriteToTextFile (FileUrl,byval Str,CharSet) 
		set stm=server.CreateObject("adodb.stream") 
		stm.Type=2 '以本模式读取 
		stm.mode=3 
		stm.charset=CharSet 
		stm.open 
		stm.WriteText str 
		stm.SaveToFile server.MapPath(FileUrl),2 
		stm.flush 
		stm.Close 
		set stm=nothing 
	end Sub 
end class
%>