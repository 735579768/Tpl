<%
class AspTpl
	private p_err
	private p_tag_l
	private p_tag_r
	private p_var_list
	private p_tpl_content
	private	p_reg
	private p_fs
	private p_tpl_file
	public p_tpl_dir
	
	private sub class_Initialize
		p_tag_l = "\{\{"
		p_tag_r = "\}\}"
		p_tpl_dir = "tpl"
		set p_fs= server.createobject("Scripting.FileSystemObject")
		set p_var_list = server.CreateObject("Scripting.Dictionary")
		Set p_reg = New RegExp 
		p_reg.IgnoreCase = True
		p_reg.Global = True
	end sub
	
	Private Sub Class_Terminate()
	
	end sub
	'==================================================
	'模板目录属性
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
				set oFile = p_fs.OpenTextFile(server.mappath(p_tpl_dir &"/"& p_tpl_File), 1)
				p_tpl_content = oFile.ReadAll
				oFile.Close
			else
				echoErr "<b>ASPTemplate Error: File [" & server.mappath(p_tpl_dir &"/"& p_tpl_File) & "] does not exists!</b><br>"
			end if
	end sub
	'==================================================
	'解析模板变量
	'==================================================
	private sub jiexivar()
		a=p_var_list.keys
		'遍历变量键值
		for each i in p_var_list.keys
			p_reg.Pattern =  p_tag_l & i & p_tag_r 
			if isarray(p_var_list(i)) then '过滤去除数组变量
				p_tpl_content=p_reg.replace(p_tpl_content,"<pre style='color:red;'>'"&i&" ' is Array </pre>")
			else
				p_tpl_content=p_reg.replace(p_tpl_content,p_var_list(i))
			end if
		next
	end sub
	'==================================================
	'给模板赋值
	'==================================================
	public sub assign(key,val)
			p_var_list(key)=val
	End sub
	'==================================================
	'解析模板文件
	'==================================================
	public Function display(tplfile)
		p_tpl_file=tplfile
		checkTplDirAndFile()
		jiexivar()
		response.Write p_tpl_content
	end Function
	'==================================================
	'输出错误
	'==================================================
	private sub echoErr(str)
		response.Write str
	end sub
end class
%>