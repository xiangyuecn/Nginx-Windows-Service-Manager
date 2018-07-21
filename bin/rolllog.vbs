dim fso
set fso=wscript.CreateObject("Scripting.FileSystemObject")

dim txt,conf
txt=wscript.arguments(0)&""
conf=wscript.arguments(1)&""

dim confStr,y,m,d,h,mi,s
y=year(now)
m=Month(now)
d=day(now)
h=Hour(now)
mi=Minute(now)
s=Second(now)
if m<10 then
	m=0&m
end if
if d<10 then
	d=0&d
end if
if h<10 then
	h=0&h
end if
if mi<10 then
	mi=0&mi
end if
if s<10 then
	s=0&s
end if

on error resume next
	echo "读取文件："&txt
	confStr=readFile(txt,"GBK")
	confStr=replace(confStr,"{y}",y)
	confStr=replace(confStr,"{m}",m)
	confStr=replace(confStr,"{d}",d)
	confStr=replace(confStr,"{h}",h)
	confStr=replace(confStr,"{M}",mi)
	confStr=replace(confStr,"{s}",s)
	
	dim i,exp,matchs,match
	Set exp = New RegExp
	exp.Pattern="^DEF(\w*) ([ \n&<>\/\w\-\u00ff-\uffff]+)=([\s\S]*?) \1END\s*$\n?"
	exp.MultiLine = True
	exp.Global = True
	set matchs=exp.Execute(confStr)
	echo "DEF "&matchs.Count&"个"
	For i=matchs.Count-1 to 0 step -1
		set match=matchs(i)
		confStr=replace(confStr,match.Value,"")
		confStr=replace(confStr,match.SubMatches(1),match.SubMatches(2))
	next
	
	if err then
		echo "打开文件失败:"&err.description
	else
		writeFile conf,confStr,"GBK"
		if err then
			echo "写入文件失败:"&err.description
		else
			echo "操作完成，已写入："&conf
			echo "请稍候..."
			wscript.sleep 1000
			wscript.quit
		end if
	end if
	wscript.quit 1
on error goto 0

'函数***********************
function echo(str)
	wscript.echo str
end function
function newFile(path)
	if not fso.FileExists(path) then
		fso.CreateTextFile path
	end if
end function
function delFile(file)
	fso.DeleteFile file,true
end function
function writeFile(file,str,charset)
	newFile file
	dim stream
	set stream=wscript.CreateObject("ADODB.Stream")
	stream.open
	
	stream.type=2
	stream.charset=charset
	stream.writeText=str
	stream.SaveToFile file,2
end function
function readFile(file,charset)
	dim stream
	set stream=wscript.CreateObject("ADODB.Stream")
	
	stream.open
	stream.type=2
	stream.charset=charset
	stream.loadFromFile file
	readFile=stream.readText(-1)
end function