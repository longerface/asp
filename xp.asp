<%@ language="vbscript" %>
<%
'/**---------------------------------------------------------------
'*  ³ÌÐòÃû³Æ: WanExplorer
'*  ³ÌÐò°æ±¾: Version 1.10 beta 2
'*  ³ÌÐòÉùÃ÷: ±¾³ÌÐòÎª¹²Ïí³ÌÐò£¬½öÇëÊ¹ÓÃÕß×¢Ã÷³ö´¦¡£Çë´ó¼ÒÄÜ°ïÖúÐÞ¸ÄÓëÍêÉÆ¡£
'*  °æÈ¨ËùÓÐ: LONGERFACE.com
'*  ³ÌÐòÖÆ×÷: longerface
'*  µç×ÓÓÊ¼þ: longerface@tom.com
'*  Q Q ºÅÂë: 7281626
'*  ×îºóÐÞ¸Ä£º2007-06-19
'*  Copyright: www.longerface.com - All Rights Reserved.
systempassword = "master"
'*/-----------------------------------------------------------------

'on error resume next
server.scripttimeout = 99999
response.expires=0
response.buffer=false

dim folder,mypage,srceenw
folder = request("folder")
if right(folder,1)="\" and left(right(folder,2),1)<>":" then folder = left(folder,len(folder)-1)
mypage = request.servervariables("script_name")

dim fso,ob_stream,jpeg
set fso = server.CreateObject("scripting.filesystemobject")
set ob_stream = server.CreateObject ("adodb.stream")


set hm = new functionAll
set mm = new menu

if hm.IsObjInstalled("persits.jpeg") then set jpeg = server.CreateObject("persits.jpeg")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  call loginsystem()  'È¥³ý´Ë´¦ rem ¼ÓÈëÃÜÂë½øÈë¹¦ÄÜ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rem ===============================  ÃÜÂëÈÏÖ¤ ===========================================
sub loginsystem()
	dim se,screenWidth
	se = request("se")
	screenWidth = request("screenWidth")
	if se<>"" then
		if se=systempassword then 
			session("se")=se
			screenw=screenWidth
			session("screenWidth") = screenw
		else
			session("se")=""
		end if
	end if
	if session("se")="" then
		response.write "<form method=post action="&mypage&" >ÇëÊäÈëµÇÂ¼ÃÜÂë£º<input type=password name=""se"" id=""sese""><input type=submit value=""½øÈë""><input type=""hidden"" name=""screenWidth"" id=""sw"" value=""""/></form>"&vbcrlf&"<scr"&"ipt defer>document.getElementById(""sese"").focus();document.getElementById(""sw"").value=window.screen.width;</scr"&"ipt>"
		response.end
	end if
end sub
rem =====================================================================================
'------------------ ½ÓÊÕÌõ¼þ½øÈë¹¦ÄÜÇø ------------------------
act = lcase(request("act"))
if act = "open" or act = "edit" then
	call showfile()
elseif act = "menu" then
	call showmenu()
elseif act = "system" then
	call showsystem()
elseif act="task" then
	call showtask()
elseif act="preview" then
	call showimageP()
elseif act="photo" then
	call showphotoHTML()
elseif act="photoarr" then
    call showPhotoArray()
elseif act="search" then
	call showsearch()
elseif act="scan" then
	call showSearchList()
elseif act="searchresult" then
    call showSearchResult(request("searchfolder"),request("keyword"))
elseif act="quit" then
	session.abandon
	response.redirect mypage
end if
'--------------------------------------------------------------

if folder="" then
	Set fo = fso.GetFolder(Server.MapPath("/"))
	'folder = server.MapPath("/")
else
	if mm.righthere("folder",hm.backurl(folder),"") then
	    set fo = fso.getfolder(folder)
	else
	    Set fo = fso.GetFolder(Server.MapPath("/"))
	end if
end if	


class functionAll
	public WebName,WebUrl,SysName,SysVersion
	public ScanResultCount
	private sub Class_Initialize ''''' ÀàµÄ³õÊ¼»¯
		WebName="LONGERFACE.COM"
		WebUrl="http://www.longerface.com/"
		SysName="WanExplorer"
		SysVersion="V1.10 beta"
	end sub
	function superfolder(folder) ''''' Çó³öÉÏ¼¶Ä¿Â¼µØÖ·£¬Èç¸ùÄ¿Â¼Îª¿Õ
		if left(right(folder,2),1)<>":" then
	    	superfolder = left(folder,instrrev(folder,"\"))
		else
			superfolder = ""
		end if
	end function
	public function javapath(tempath) ''''''  ÔÚscriptÀï½«¡°\¡± ±ä³É¡°\\¡± 
		tempath = trim(tempath)
		javapath = replace(tempath,"\","\\")
	end function
	function backpath(tempath)  ''''''  ÔÚscriptÀï½«¡°\\¡± ±ä³É¡°\¡±
		tempath = trim(tempath)
		backpath = replace(tempath,"\\","\")
	end function
	function javaurl(tempath)  ''''''  ÔÚµØÖ·¿òÀï½«ÌØÊâ×Ö·û±ä³É×Ö·ûÂë
		tempath = replace(tempath,"#","%23")
		tempath = replace(tempath," ","%20")
		javaurl = tempath
	end function
	function backurl(tempath)  ''''''  ÔÚµØÖ·¿òÀï½«×Ö·ûÂë±ä³ÉÌØÊâ×Ö·û
		tempath = replace(tempath,"%23","#")
		tempath = replace(tempath,"%20"," ")
		backurl = tempath
	end function
	public function CopyRight '''' °æÈ¨º¯Êý
		response.write "<div align=""center"" style=""cursor:default;"" onclick=""window.open('"&WebUrl&"','_blank');"" title="""&SysName&" "&SysVersion&"   Copyright (C) "&WebName&" All rights reserved.""><font style=""font-size:8pt;color:#666666;font-family:arial;"">CopyRight &copy; <b><font color=336699>"&name8(WebName)&"</font><font color=993300>."&name3(WebName)&"</font></b> <!--All rights reserved.--></font></div>"
	end function
	function location(nowFolder)  ''' ÍøÂçÂ·¾¶
		location = "\\"&request.ServerVariables("LOCAL_ADDR")
		if not isempty(Request.ServerVariables("SERVER_PORT")) and Request.ServerVariables("SERVER_PORT")<>80 then location = location + ":"&Request.ServerVariables("SERVER_PORT")
		location = location + "\"&nowfolder
	end function
	
	function nameAll(namestr)''''' º¯Êý£¬È¡ÎÄ¼þÈ«Ãû
		if instr(namestr,"\")=0 then
			nameAll = namestr
		else
			nameAll = right(namestr,instr(strreverse(namestr),"\")-1)
		end if
	end function
	function name8(namein)'''''''' º¯Êý£¬È¡ÎÄ¼þÖ÷Ãû
		if instr(nameAll(namein),".")=0 then
			name8 = mid(nameAll(namein),1,len(nameAll(namein)))
		else
			name8 = left(nameAll(namein),instr(nameAll(namein),".")-1)
		end if		
	end function
	function name3(namein)'''''''' º¯Êý£¬È¡ÎÄ¼þÀ©Õ¹Ãû
		if instr(nameAll(namein),".")=0 then
			name3=""
		else
			name3 = right(nameAll(namein),instr(strreverse(nameAll(namein)),".")-1)
		end if
	end function
	function getfolderpath(temppath)
		if left(strreverse(temppath),1)="\" then
			getfolderpath = left(temppath,len(temppath)-1)
		else
			getfolderpath = temppath
		end if  
	end function
	function fileimage(namein)''''' º¯Êý,È¡µÃÎÄ¼þÀàÐÍÍ¼Æ¬
		ext = lcase(name3(namein))
		fileimage = "<font style='font-size:30pt;font-family:Wingdings 2;'>/</font>"
		if ext="txt" or ext = "log" then fileimage = "<font style='font-size:30pt;font-family:Wingdings;'>3</font>"
		if ext="htm" or ext="html" then fileimage = "<font style='font-size:30pt;font-family:Wingdings;'>]</font>"
		if ext="asp" or ext="php" or ext="jsp" or ext="java" or ext="vbs" then fileimage = "<font style='font-size:30pt;font-family:Wingdings;'>2</font>"
		if ext="mdb" or ext="dbf" or ext="mdf" then fileimage = "<font style='font-size:25pt;font-family:Wingdings;'>4</font>"
		if ext="exe" or ext="com" then fileimage = "<font style='font-size:25pt;font-family:Wingdings 2;'>4</font>"
		if ext="mp3" or ext="wma" or ext="wav" then fileimage = "<font style='font-size:25pt;font-family:Webdings;'>"&chr(14)&"</font>"
		if ext="doc" then fileimage = "<font style='font-size:30pt;font-family:Wingdings 2;'>2</font>"
		if ext="ini" then fileimage = "<font style='font-size:27pt;font-family:Wingdings;'>></font>"
		if ext="dll" then fileimage = "<font style='font-size:25pt;font-family:Webdings;'>"&chr(96)&"</font>"
		if ext="url" then fileimage = "<font style='font-size:25pt;font-family:Webdings;'>'</font>"
		if ext="msi" then fileimage = "<font style='font-size:25pt;font-family:webdings;'>e</font>"
		if ext="chm" or ext="hlp" then fileimage = "<font style='font-size:20pt;font-family:Webdings;'>s</font>"
		if ext="zip" or ext="rar" then fileimage = "<font style='font-size:28pt;font-family:webdings;'>"&chr(62)&"</font>"
		if ext="jpg" or ext="gif" or ext="jpge" or ext="bmp" then
			if not IsObjInstalled("persits.jpeg") then
				fileimage = "<font style='font-size:25pt;font-family:webdings;'>N</font>"
			else
				fileimage = "<img src="""&mypage&"?act=preview&imagename="&hm.javaurl(namein)&"&wview=50&hview=50"">"
			end if
		end if
	end function
	function fileImageSmall(namein)''''' º¯Êý,È¡µÃÎÄ¼þÀàÐÍÐ¡Í¼Æ¬
		ext = lcase(name3(namein))
		fileImageSmall = "<font style='font-size:20px;font-family:Wingdings 2;'>/</font>"
		if ext="txt" or ext = "log" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings;'>3</font>"
		if ext="htm" or ext="html" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings;'>]</font>"
		if ext="asp" or ext="php" or ext="jsp" or ext="java" or ext="vbs" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings;'>2</font>"
		if ext="mdb" or ext="dbf" or ext="mdf" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings;'>4</font>"
		if ext="exe" or ext="com" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings 2;'>4</font>"
		if ext="mp3" or ext="wma" or ext="wav" then fileImageSmall = "<font style='font-size:20px;font-family:Webdings;'>"&chr(14)&"</font>"
		if ext="doc" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings 2;'>2</font>"
		if ext="ini" then fileImageSmall = "<font style='font-size:20px;font-family:Wingdings;'>></font>"
		if ext="dll" then fileImageSmall = "<font style='font-size:20px;font-family:Webdings;'>"&chr(96)&"</font>"
		if ext="url" then fileImageSmall = "<font style='font-size:20px;font-family:Webdings;'>'</font>"
		if ext="msi" then fileImageSmall = "<font style='font-size:20px;font-family:webdings;'>e</font>"
		if ext="chm" or ext="hlp" then fileImageSmall = "<font style='font-size:20px;font-family:Webdings;'>s</font>"
		if ext="zip" or ext="rar" then fileImageSmall = "<font style='font-size:20px;font-family:webdings;'>"&chr(62)&"</font>"
		if ext="jpg" or ext="gif" or ext="jpge" or ext="bmp" then fileImageSmall = "<font style='font-size:20px;font-family:webdings;'>N</font>"
	end function
	function folderimage(namein)''''' º¯Êý,È¡µÃÎÄ¼þ¼ÐÍ¼Æ¬
		folderimage = "<font style='font-size:25pt;font-family:Wingdings;'>0</font>"
		if lcase(nameAll(namein))="recycled" or lcase(nameAll(namein))="recycler" then folderimage = "<font style='font-size:30pt;font-family:Wingdings 2;'>3</font>"
	end function
	function folderImageSmall(namein)''''' º¯Êý,È¡µÃÎÄ¼þ¼ÐÐ¡Í¼Æ¬
		folderImageSmall = "<font style='font-size:15px;font-family:Wingdings;'>0</font>"
		if lcase(nameAll(namein))="recycled" or lcase(nameAll(namein))="recycler" then folderImageSmall = "<font style='font-size:15px;font-family:Wingdings 2;'>3</font>"
	end function
	function diskimage(num)'''''' º¯Êý,È¡µÃÁËÅÌ·ûÍ¼Æ¬
		select case cint(num)
			Case 1    : diskimage = "<font style='font-size:20pt;font-family:Wingdings;'><</font>"               '1¿ÉÒÆ¶¯´ÅÅÌ
			Case 2    : diskimage = "<font style='font-size:33pt;font-family:Wingdings;'>;</font>"               '2±¾µØÓ²ÅÌ
			Case 3    : diskimage = "<font style='font-size:25pt;font-family:Wingdings 2;'>9</font>"             '3ÍøÂç´ÅÅÌ
			Case 4    : diskimage = "<font style='font-size:25pt;font-family:Wingdings 2;'>8</font>"             '4CD-ROM
			Case 5    : diskimage = "<font style='font-size:25pt;font-family:Wingdings;'>"&chr(128)&"</font>"    '5RAM ´ÅÅÌ
			Case else : diskimage = "<font style='font-size:25pt;font-family:webdings;'>i</font>"                '0Î´Öª
		end Select
	end function
	function diskname(num)'''''' º¯Êý,È¡µÃÁËÅÌ·ûÃû³Æ
		select case cint(num)
			Case 1    : diskname = "¿ÉÒÆ¶¯´ÅÅÌ"               '1¿ÉÒÆ¶¯´ÅÅÌ
			Case 2    : diskname = "±¾µØÓ²ÅÌ"                 '2±¾µØÓ²ÅÌ
			Case 3    : diskname = "ÍøÂç´ÅÅÌ"                 '3ÍøÂç´ÅÅÌ
			Case 4    : diskname = "CD-ROM"                   '4CD-ROM
			Case 5    : diskname = "RAM ´ÅÅÌ"                 '5RAM ´ÅÅÌ
			Case else : diskname = "Î´Öª"                     '0Î´Öª
		end Select
	end function
	public sub savefile(txtpath,txtname,inputtxt,intype) '''''''''''''''  Ð´ÈëÎÄ¼þ¹ý³Ì  ''''''''''''
		if trim(txtname) = empty then txtname = "ÐÂ½¨ ÎÄ±¾ÎÄµµ.txt"
		if inputtxt = empty then inputtxt = ""
		if right(txtpath,1)<>"\" then txtpath = txtpath & "\"
		filepath = txtpath & txtname
		if cint(intype)<>8 and cint(intype)<>2 then  intype = 2
		if not fso.fileexists(filepath) then
	    	set txt = fso.opentextfile(filepath,8,true)
			txt.writeline inputtxt
		 	exit sub
		end if  
		set txt = fso.opentextfile(filepath,cint(intype))'¸ÄÎª8ÊÇ×·¼Ó2ÊÇ¸²¸Ç
		txt.writeline inputtxt
	end sub
	function diskpath(newfolder) '''''''''''''''''' È¡ÅÌ·û
		diskpath = left(trim(newfolder),2)
	end function
	Public Function IsObjInstalled(strClassString) '''''''''''''''¡¡ÑéÖ¤±¾»úÊÇ·ñÒÑ°²×°×é¼þÎ¢(·Ç×÷Õß±àÐ´)
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If Err Then
			IsObjInstalled = False
		else	
			IsObjInstalled = True
		end if
		Set xTestObj = Nothing
	End Function
	public function search(id,searchpath,keyword) '''''''''''''''''' µÝ¹éËÑË÷ÎÄ¼þ¼ÐÏÂµÄÎÄ¼þ
		on error resume next
		if searchpath="" then
			set scanfso = fso.GetFolder(Server.MapPath("./"))
		else
			if right(searchpath,1)<>"\" then searchpath = searchpath & "\"
			set scanfso = fso.GetFolder(searchpath)
		end if
		set lookfolder = scanfso.subfolders
		if lookfolder.count <> 0 then 
			for each foldername in lookfolder    '''''''¶Á³öÎÄ¼þ¼Ð
			  outword = ""
			  id = id + 1
			  set folderdir = fso.getfolder(foldername)
			  if instr(lcase(nameAll(foldername)),lcase(keyword))>0 then
				outword = "<tr><td>&nbsp;<span id='div"&id&"' onclick=""oneClick(this);"" onDblClick=""doubleClick(this,'folder','"&mypage&"?folder="&javapath(javaurl(foldername))&"');"">"&folderImageSmall(foldername)&" "&backurl(nameAll(foldername))&"</span></td>"
				id = id + 1
				outword = outword &	"<td>&nbsp;<span id='div"&id&"' onclick=""oneClick(this);"" onDblClick=""doubleClick(this,'folder','"&mypage&"?folder="&javapath(javaurl(superfolder(foldername)))&"');"">"&backurl(superfolder(foldername))&"</span></td>"
				outword = outword &	"<td>"&"&nbsp;"&" </td>"
				outword = outword &	"<td>&nbsp;"&"ÎÄ¼þ¼Ð"&"</td>"
				outword = outword &	"<td>&nbsp;"&folderdir.datelastmodified&"</td>"
				outword = outword &	"<td style=""display:none;"">&nbsp;</td></tr>"&vbcrlf
				response.write outword
				ScanResultCount = ScanResultCount + 1
			  end if
			  search id,backurl(foldername),keyword  ''µÝ¹éÔÙËÑË÷
			next
		end if
		for each filename in scanfso.files  ''''¶Á³öÎÄ¼þ
		  outword = ""
		  id = id + 1
		  set filedir = fso.getfile(filename)
		  if instr(lcase(nameAll(filename)),lcase(keyword))>0 then
			outword = "<tr><td>&nbsp;<span id='div"&id&"' onclick=""oneClick(this);"" onDblClick=""doubleClick(this,'file','"&javapath(filename)&"');"">"&fileImageSmall(filename)&" "&backurl(nameAll(filename))&"</span></td>"
			id = id + 1
			outword = outword &	"<td>&nbsp;<span id='div"&id&"' onclick=""oneClick(this);"" onDblClick=""doubleClick(this,'folder','"&mypage&"?folder="&javapath(javaurl(superfolder(filename)))&"');"">"&backurl(superfolder(filename))&"</span></td>"
			outword = outword &	"<td align=right>"&mm.formatsize(filedir.size)&"&nbsp;</td>"
			outword = outword &	"<td>&nbsp;"&mm.filename(name3(filename))&"</td>"
			outword = outword &	"<td>&nbsp;"&filedir.datelastmodified&"</td>"
			outword = outword &	"<td style=""display:none;"">&nbsp;</td></tr>"&vbcrlf
			response.write outword
			ScanResultCount = ScanResultCount + 1
		  end if
		next
	end function
	public function rlchar(tempstring,tempchar)   ''' ·µ»ØÈ¥µô´®×óÓÒÖØ¸´µÄ×Ö·û e.g: rlchar(",,,,a,b,,,,c,,d,,,",",") ·µ»Ø a,b,,,,c,,d
		if tempstring="" then
			rlchar=""
			exit function
		end if
		if tempchar="" then
			rlchar = tempstring
			exit function
		end if
		if len(tempchar)>1 then
			tempchar = left(tempchar,1)
		end if
		for i=1 to len(tempstring)
			if mid(tempstring,i,1)<>tempchar then exit for
		next
		tempstring = mid(tempstring,i,len(tempstring))
		if len(tempstring)=0 then
			rlchar = ""
			exit function
		end if
		for i=len(tempstring) to 0 step -1
			if mid(tempstring,i-1,1)<>tempchar then exit for
		next
		tempstring = left(tempstring,i-1)
		rlchar = tempstring
	end function
	public function outrepeat(tempstring,tempdoc)  ''' ·µ»Ø½«×Ö´®ÀïÖØ¸´µÄ×Ö·û±ä³ÉÒ»¸ö×Ö·û e.g: outrepeat("a,b,,,c,,d,,,",",") ·µ»Ø a,b,c,d,
		dim OR_strlen,OR_restring
		if tempstring="" then
			outrepeat = ""
			exit function
		end if
		OR_strlen = len(tempstring)
		if OR_strlen<2 then
			outrepeat = tempstring
			exit function
		end if
		if tempdoc="" then
			outrepeat = tempstring
			exit function
		end if
		if len(tempdoc)>1 then
			tempdoc = left(tempdoc,1)
		end if
		while instr(tempstring,tempdoc&tempdoc)>0
			tempstring = replace(tempstring,tempdoc&tempdoc,tempdoc)
		wend
		OR_restring = tempstring
		outrepeat = OR_restring
	end function
end class


class menu
	public function isdisk(folder) '''''''''''''''''' ÊÇ·ñÊÇ¸ùÅÌ·û
		if len(trim(folder))>3 then
			isdisk=false
		else
			isdisk=true
		end if
	end function
	public function righthere(hereType,herePath,hereName)   '''''''''  ²é¿´ÎÄ¼þ¼Ð»òÎÄ¼þÊÇ·ñ´æÔÚ
		if right(herePath,1)<>"\" then herePath = herePath & "\"
		mypath = herePath & hereName
		if hereType = "file" then
			if fso.fileExists(mypath) then 
				righthere = true
			else
				righthere = false
			end if
		elseif  hereType = "folder" then
			if fso.folderexists(mypath) then
				righthere = true
			else
				righthere = false
			end if
		else
			righthere = false
		end if
	end function
	function isimage(tempfilenameexe)   ''''''''''''''''''' ´«µÄÎÄ¼þÀ©Õ¹ÃûÈçÊÇÍ¼Æ¬¸ñÊ½·´»Ø true
		tempfilenameexe = ucase(tempfilenameexe)
		if tempfilenameexe = "JPG" or tempfilenameexe = "GIF" or tempfilenameexe = "PNG" or tempfilenameexe = "JPGE" or tempfilenameexe = "BMP" then
			isimage = true
		else
			isimage = false
		end if
	end function
	function filename(tempname)   ''''''''''''''''''' ÏÔÊ¾ÎÄ¼þÀàÐÍÃû
		tempname = ucase(tempname)
		filename = tempname
		if tempname = "EXE" then filename = "Ó¦ÓÃ³ÌÐò"
		if tempname = "BAT" then filename = "Windows Åú´¦Àí"
		if tempname = "TXT" then filename = "ÎÄ±¾"
		if tempname = "HTML" or tempname = "HTM" then filename = "HTML Document"
		if tempname = "DOC" then filename = "Microsoft Word ÎÄµµ"
		if tempname = "JPG" or tempname = "GIF" or tempname = "PNG" or tempname = "JPGE" or tempname = "BMP" then filename = tempname & "Í¼Ïñ" 
		if tempname = "MP3" or tempname = "WMA" then filename = tempname & "ÒôÆµ"
		if tempname = "INI" then filename = "ÅäÖÃÉèÖÃ"
		if tempname = "AVI" or tempname = "RMVB" or tempname = "MPEG" then filename = tempname & "ÊÓÆµ"
		if tempname = "DLL" then filename = "Ó¦ÓÃ³ÌÐòÀ©Õ¹"
		if tempname = "MSC" then filename = "Microsoft Í¨ÓÃ¹ÜÀíÎÄµµ"
		if tempname = "INF" then filename = "°²×°ÐÅÏ¢"
		if tempname = "MSI" then filename = "Windows Installer Èí¼þ°ü"
		
		filename = filename & "ÎÄ¼þ"
	end function 
	function formatsize(mysize)  ''''''''''''''''''''  ¼ÆËã³öÎÄ¼þ´óÐ¡
		if mysize>=(1024*1024*1024) then
			formatsize=formatnumber(mysize/(1024*1024*1024),2) & " GB"
		elseif mysize>=(1024*1024) then
			formatsize=formatnumber(mysize/(1024*1024),2) & " MB"
		elseif mysize>=1024 then
			formatsize=formatnumber(mysize/1024,2) & " KB"
		else
			formatsize=mysize & "B"
		end if
	end Function
	public function specialCharacter(tempChar)  ''''''''''''''' ¼ì²éÌØÊâ×Ö·û  /\:*?"<>|      chr(34)ÊÇ"
		if instr(tempChar,"/")<>0 or _
		   instr(tempChar,"\")<>0 or _
		   instr(tempChar,":")<>0 or _
		   instr(tempChar,"*")<>0 or _
		   instr(tempChar,"?")<>0 or _
		   instr(tempChar,chr(34))<>0 or _  
		   instr(tempChar,"<")<>0 or _
		   instr(tempChar,">")<>0 or _
		   instr(tempChar,"|")<>0 or _
		   len(tempChar)<1 _
		then
			response.write "<font size=2>ÃüÃûÊ±²»ÄÜÊ¹ÓÃÌØÊâ×Ö·û:<br>"&vbcrlf&"<font color=red>/ \ : * ? "&chr(34)&" < > | </font>"&vbcrlf&"<br>Çë¸ÄÓÃÆäËüÃû³Æ¡£</font>"
			response.write "<br><br>"&vbcrlf&"<input type=button value=ºóÍË onclick='javascript:history.back();'>"
			response.end
		end if
	end function
	function dtype(num)   ''''''''''''''''''''''''''   ´ÅÅÌÀàÐÍ
		select Case num
			Case 0: dtype = "Î´Öª"
			Case 1: dtype = "¿ÉÒÆ¶¯´ÅÅÌ"
			Case 2: dtype = "±¾µØÓ²ÅÌ"
			Case 3: dtype = "ÍøÂç´ÅÅÌ"
			Case 4: dtype = "CD Çý¶¯Æ÷"
			Case 5: dtype = "RAM ´ÅÅÌ"
		end Select
	end Function
end class

rem ==============================  ÒÔÏÂÉÏ´«ÀàÄÚÈÝ·Ç±¾×÷Õß±àÐ´  ================================
' ÒýÓÃÊ±Î´ÕÒµ½Ô­Öø°æÈ¨
dim oUpFileStream
Class upload_file
dim Form,File,Version
Private Sub Class_Initialize 
   '¶¨Òå±äÁ¿
  dim RequestBinDate,sStart,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
  dim iFileSize,sFilePath,sFileType,sFormvalue,sFileName
  dim iFindStart,iFindEnd
  dim iFormStart,iFormEnd,sFormName
   '´úÂë¿ªÊ¼
  Version="ÎÞ×é¼þÉÏ´«Àà Version 0.96"
  set Form = Server.CreateObject("Scripting.Dictionary")
  set File = Server.CreateObject("Scripting.Dictionary")
  if Request.TotalBytes < 1 then Exit Sub
  set tStream = Server.CreateObject("adodb.stream")
  set oUpFileStream = Server.CreateObject("adodb.stream")
  oUpFileStream.Type = 1
  oUpFileStream.Mode = 3
  oUpFileStream.Open 
  oUpFileStream.Write Request.BinaryRead(Request.TotalBytes)
  oUpFileStream.Position=0
  RequestBinDate = oUpFileStream.Read 
  iFormEnd = oUpFileStream.Size
  bCrLf = chrB(13) & chrB(10)
  'È¡µÃÃ¿¸öÏîÄ¿Ö®¼äµÄ·Ö¸ô·û
  sStart = MidB(RequestBinDate,1, InStrB(1,RequestBinDate,bCrLf)-1)
  iStart = LenB (sStart)
  iFormStart = iStart+2
  '·Ö½âÏîÄ¿
  Do
    iInfoEnd = InStrB(iFormStart,RequestBinDate,bCrLf & bCrLf)+3
    tStream.Type = 1
    tStream.Mode = 3
    tStream.Open
    oUpFileStream.Position = iFormStart
    oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
    tStream.Position = 0
    tStream.Type = 2
    tStream.Charset ="gb2312"
    sInfo = tStream.ReadText      
    'È¡µÃ±íµ¥ÏîÄ¿Ãû³Æ
    iFormStart = InStrB(iInfoEnd,RequestBinDate,sStart)-1
    iFindStart = InStr(22,sInfo,"name=""",1)+6
    iFindEnd = InStr(iFindStart,sInfo,"""",1)
    sFormName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
    'Èç¹ûÊÇÎÄ¼þ
    if InStr (45,sInfo,"filename=""",1) > 0 then
      set oFileInfo= new FileInfo
      'È¡µÃÎÄ¼þÊôÐÔ
      iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
      iFindEnd = InStr(iFindStart,sInfo,"""",1)
      sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
      oFileInfo.FileName = GetFileName(sFileName)
      oFileInfo.FilePath = GetFilePath(sFileName)
      oFileInfo.FileExt = GetFileExt(sFileName)
      iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
      iFindEnd = InStr(iFindStart,sInfo,vbCr)
      oFileInfo.FileType = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
      oFileInfo.FileStart = iInfoEnd
      oFileInfo.FileSize = iFormStart -iInfoEnd -2
      oFileInfo.FormName = sFormName
      file.add sFormName,oFileInfo
    else
    'Èç¹ûÊÇ±íµ¥ÏîÄ¿
      tStream.Close
      tStream.Type = 1
      tStream.Mode = 3
      tStream.Open
      oUpFileStream.Position = iInfoEnd 
      oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
      tStream.Position = 0
      tStream.Type = 2
      tStream.Charset = "gb2312"
      sFormvalue = tStream.ReadText 
      form.Add sFormName,sFormvalue
    end if
    tStream.Close
    iFormStart = iFormStart+iStart+2
    'Èç¹ûµ½ÎÄ¼þÎ²ÁË¾ÍÍË³ö
    loop until (iFormStart+2) = iFormEnd 
  RequestBinDate=""
  set tStream = nothing
End Sub

Private Sub Class_Terminate  
  'Çå³ý±äÁ¿¼°¶ÔÏñ
  if not Request.TotalBytes<1 then
    oUpFileStream.Close
    set oUpFileStream =nothing
    end if
  Form.RemoveAll
  File.RemoveAll
  set Form=nothing
  set File=nothing
End Sub
   
 'È¡µÃÎÄ¼þÂ·¾¶
Private function GetFilePath(FullPath)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
    Else
    GetFilePath = ""
  End If
End function
 
'È¡µÃÎÄ¼þÃû
Private function GetFileName(FullPath)
  If FullPath <> "" Then
    GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
    Else
    GetFileName = ""
  End If
End function

'È¡µÃÀ©Õ¹Ãû
Private function GetFileExt(FullPath)
  If FullPath <> "" Then
    GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
    Else
    GetFileExt = ""
  End If
End function

End Class

'ÎÄ¼þÊôÐÔÀà
Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt = ""
  End Sub
  
'±£´æÎÄ¼þ·½·¨
 Public function SaveToFile(FullPath)
    dim oFileStream,ErrorChar,i
    SaveToFile=1
    if trim(fullpath)="" or right(fullpath,1)="/" then exit function
    set oFileStream=CreateObject("Adodb.Stream")
    oFileStream.Type=1
    oFileStream.Mode=3
    oFileStream.Open
    oUpFileStream.position=FileStart
    oUpFileStream.copyto oFileStream,FileSize
    oFileStream.SaveToFile FullPath,2
    oFileStream.Close
    set oFileStream=nothing 
    SaveToFile=0
  end function
End Class
rem ==============================  ÒÔÉÏÉÏ´«ÀàÄÚÈÝ·Ç±¾×÷Õß±àÐ´  ================================
%>



<%if false then%>
<html><body></body></html>
<%end if%>


<%
sub showcss
%>
<style>
a:link{text-decoration:none;}
a:hover{text-decoration:none;}
a:visited{text-decoration:none;}
.menufont{
	font-size:9pt;color:#333333;
}
.f9p{font-size:12px;}
.menuimage{color:#666666;}
a.menulink:visited,a.menulink:link{TEXT-DECORATION:none;color:333333;}
a.menulink:hover{text-decoration:underline;color:333333;}
.searchwait{width:9px;height:13px;background-color:#9CA2A5;font-size:0px;float:left;margin-left:2px;}
.searchtop{background-color:#D4D0C8;border-top:1px solid #fff;border-left:1px solid #fff;border-right:1px solid #808080;border-bottom:1px solid #808080;height:20px;}
body {
	SCROLLBAR-HIGHLIGHT-COLOR: buttonface; 
	SCROLLBAR-SHADOW-COLOR: buttonface;
	SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;
	SCROLLBAR-TRACK-COLOR: #eeeeee;
	FONT-FAMILY: "ËÎÌå";
	SCROLLBAR-DARKSHADOW-COLOR: buttonshadow;
	margin:0px 0px;
}
</style>
<%
end sub
sub showscript
%>
<script language="javascript">
//ajax
function connxml(){
	var XmlHttp;
	if (window.XMLHttpRequest) {
	¡¡XmlHttp = new XMLHttpRequest()
	¡¡if (XmlHttp.overrideMimeType) {
	¡¡¡¡XmlHttp.overrideMimeType('text/xml');
	¡¡}
	} else 
		if(window.ActiveXObject){  
		  try{   
			  XmlHttp = new ActiveXObject("Msxml2.XMLHTTP");   
		  }catch(e){   
			  try{   
				  XmlHttp = new ActiveXObject("Microsoft.XMLHTTP");   
			  }catch(e){alert("ÄúµÄä¯ÀÀÆ÷²»Ö§³ÖAjax£¬Ä³Ð©¹¦ÄÜ²»ÄÜÊµÏÖ¡£");}   
		  } 
	}
	return XmlHttp;
}
function getdata(path){
	//alert(path);
	var xml = connxml();
	if( navigator.appName =="Microsoft Internet Explorer" ){
		xml.onreadystatechange = function(){
			if(xml.readyState==4){
				if(xml.status==200){
					xml.responseText;
				}
			}else{
				;
			}
		}
		xml.open("GET",path,true);
		xml.send(null);
	}else{ // Èç¹û²»ÊÇIEµÄ´¦Àí·½°¸
		;
		xml.onreadystatechange = function(){
			xml.onload = function(){
				if(xml.status==200){
					xml.responseText;
				}
			}
		}
		xml.open("GET",path,true);
		xml.send(null);
	}
	return ;
}
//function
function getfolder(path){   //µÃµ½Ä¿Â¼
	if(path.length<3) return ;
	var pl = path.lastIndexOf("\\");
	if(pl==2){
		pl = 3;
	}
	var re = path.substr(0,pl);
	return re;
}
function nameAll(path){   //µÃµ½ÎÄ¼þÈ«Ãû
	if(path.indexOf("\\")==-1) return ;
	var pl = path.lastIndexOf("\\");
	var re = path.substr(pl+1);
	return re;
}

</script>
<%
end sub
''''' ----------------------------------  Í¼²ãÏÔÊ¾Í¼±ê  --------------------------------- 
sub showImage(imgId,imgTop,imgLeft,imgName,imgImage,imgLink,imgType,imgDoubleClick)
%>
<div id="image<%=imgId%>" style="position:absolute; width:80px; height:80px; z-index:1; left: <%=imgLeft%>px; top: <%=imgTop%>px;"
 onmousedown="imageNum=<%=imgId%>;down=true;divleft=event.clientX-parseInt(image<%=imgId%>.style.left);divtop=event.clientY-parseInt(image<%=imgId%>.style.top)"onmouseup="down=false">
  <table width="80" height="60" border="0" cellpadding="0" cellspacing="2" style="word-wrap:break-word;word-break:break-all;table-layout:fixed;cursor:default;">
    <tr>
      <td height="40">
	  	<div align="center" 
		  <%if not(imgDoubleClick="") then response.write "onDblClick=""doubleClick("&imgId&",'"&imgDoubleClick&"','"&hm.javapath(hm.javaurl(imgLink))&"');"""%>
		  onclick="oneClick(<%=imgId%>,'<%=imgType%>','<%=hm.javaurl(imgname)%>','<%=hm.javapath(hm.javaurl(folder)&"\"&hm.javaurl(imgname))%>','<%=hm.javapath(hm.javaurl(folder)&"\"&hm.javaurl(imgname))%>');">
			<span id="ico<%=imgId%>" style="color:#333333;"><%=imgImage%></span></div></td>
    </tr>
    <tr>
      <td><div align="center" id="link<%=imgId%>"><a href="<%if imgType="file" then : response.write "javascript:doubleClick("&imgId&",'"&imgDoubleClick&"','"&hm.javapath(hm.javaurl(imgLink))&"');" :else:response.write hm.javaurl(imgLink):end if%>"><font style="font-size:9pt;color:#333333;"><%=hm.backurl(imgName)%></font></a></div></td>
    </tr>
  </table>
</div>
<%
end sub
' --------------------------------------------------------------------------------------
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=hm.location(folder)%></title>
<%showcss%>
<script language="javascript">
var clickid=0;
function doubleClick(theSelfid,theRun,theUrl)
{
	eval("link"+theSelfid).style.backgroundColor = "#cccccc";
	if(clickid!=0 && clickid!=theSelfid){
		eval("link"+clickid).style.backgroundColor = "";	
	}
	clickid = theSelfid;
	if (theRun=="superior")
		location.href = theUrl;//alert("½øÈëÉÏÒ»¼¶Ä¿Â¼ "+theUrl);
	if(theRun=="self")
		location.href = theUrl;//alert("½øÈë±¾Ä¿Â¼ ´«Öµ±¾Ä¿Â¼¼´¿É "+theUrl);
	if(theRun=="disk")
		location.href = theUrl;//alert("´ò¿ª±¾ÅÌ·û "+theUrl);
	if(theRun=="folder")
		location.href = theUrl;//alert("½øÈëÏÂ¼¶Ä¿Â¼ "+theUrl);
	if(theRun=="file")
	{
		//alert("´ò¿ªÎÄ¼þ "+"<%=mypage%>?folder=<%=hm.javapath(folder)%>&file="+theUrl+"&act=open");
		var url = window.open("","popupnav","width=550,height=400,resizable=1,status=yes,scrollbars=no,menubar=no,location=no");
		url.location.href = "<%=mypage%>?folder=<%=hm.javapath(hm.javaurl(folder))%>&file="+escape(theUrl)+"&act=open" ;
	}
}
function oneClick(theSelfid,theType,theDisk,theFolder,theFile)
{
	eval("link"+theSelfid).style.backgroundColor = "#cccccc";
	//eval("ico"+theSelfid).style.filter = "invert()";
	if(clickid!=0 && clickid!=theSelfid){
		eval("link"+clickid).style.backgroundColor = "";	
	}
	clickid = theSelfid;
	//alert(theType +"\n"+ theDisk+"\n" + theFolder+"\n" + theFile+"\n"+"<%=hm.javapath(folder)%>");
	if(theType=="disk")
		menu.location.href = "<%=mypage%>?act=menu&nowfolder=<%=hm.javapath(folder)%>&type=disk&disk="+theDisk+":";
	if(theType=="folder")
		menu.location.href = "<%=mypage%>?act=menu&nowfolder=<%=hm.javapath(hm.javaurl(folder))%>&type=folder&disk="+theDisk+"&folder="+theFolder;
	if(theType=="file")
		menu.location.href = "<%=mypage%>?act=menu&nowfolder=<%=hm.javapath(hm.javaurl(folder))%>&file="+theFile+"&disk="+theDisk+"&type=file";
}
</script>
<script>
var down=false,divleft,divtop;
function move()
{
	if(down)
	{
		if (event.clientX-divleft<200) return false;
		eval("image"+imageNum).style.left=event.clientX-divleft;
		eval("image"+imageNum).style.top=event.clientY-divtop;
	}
}
</script>
</head>
<body onmousemove="move()">
<div id="allpe" style="position:absolute;background-color:#fff;text-align:center;z-index:2;left:250px;top:50px;width:300px;height:30px;display:'none';">
    <div style="border:1px solid #ccc;width:300px;height:5px;text-align:left;padding:1px;margin-bottom:8px;">
	    <div id="pe" style="background-color:#ccc;height:5px;font-size:0px;width:1%;"></div>
    </div>
<div style="font-size:9pt;color:#ccc;width:300px;"><div style="width:180px;float:left;margin-left:15px;">ÕýÔÚ¼ÓÔØ£¬ÇëµÈ´ý(<span id="pernum"></span>/<span id="allnum"></span>)</div><div style="width:70px;float:left;text-align:left;font-weight:bold;"><span id=dot>..........</span></div></div>
<script language="javascript">function showdot(){	var num=0,char='';	num = parseInt(Math.random()*10+1);	for(i=0;i<num;i++)	  char = char +'.';	dot.innerHTML=char;	setTimeout('showdot();',20);};showdot();</script>
</div>
<script>pe.style.width = 0 +"%";</script>
<%
if session("screenWidth")>=1280 then
	session("sw") = 1100
elseif session("screenWidth")>=1024 then
	session("sw") = 914
else
	session("sw") = 690
end if

if not(folder="") then
	itId = 3  '''''''' ÎÄ¼þ»òÎÄ¼þ¼ÐºÅ
	itLeft = 370  '''' ¿ªÊ¼×ó±ßÎ»ÖÃ
	itTop = 0  ''''''' ¿ªÊ¼ÉÏ±ßÎ»ÖÃ
	per = 1
	allnumber = (fo.subfolders.count)+(fo.files.count)
	response.write "<scr"&"ipt>document.getElementById('allpe').style.display='';document.getElementById('allnum').innerHTML='"&allnumber&"';</scr"&"ipt>"
	for each foldername in fo.subfolders  ''''¶Á³öÎÄ¼þ¼Ð
	  'response.write nameAll(foldername)&"<br>"
	  call showImage(itId,itTop,itLeft,hm.nameAll(foldername),hm.folderimage(foldername),"?folder="&foldername,"folder","folder")
	  itId = itId + 1
	  itLeft = itLeft + 80
	  if itLeft > session("sw") then : itLeft=210 : itTop=itTop+80 : end if
	  response.write vbcrlf&"<scr"&"ipt>pe.style.width = "&int((per*100)/allnumber)&" +""%"";pernum.innerHTML='"&per&"';</scr"&"ipt>"&vbcrlf
	  per = per + 1
	next
	for each filename in fo.files  ''''¶Á³öÎÄ¼þ
	  'response.write nameAll(filename)&"<br>"
	  call showImage(itId,itTop,itLeft,hm.nameAll(filename),hm.fileimage(filename),hm.nameAll(filename),"file","file")
	  itId = itId + 1
	  itLeft = itLeft + 80
	  if itLeft > session("sw") then : itLeft=210 : itTop=itTop+80 : end if
	  response.write vbcrlf&"<scr"&"ipt>pe.style.width = "&int((per*100)/allnumber)&" +""%"";pernum.innerHTML='"&per&"';</scr"&"ipt>"&vbcrlf
	  per = per + 1
	next
else
	itId = 1
	itLeft = 210
	itTop = 0
	set dr = fso.drives
	for each disk in dr  '''''''''Ñ­»·³öÅÌ·û
	  'response.write disk.driveletter & disk.drivetype &"<br>"
	  call showImage(itId,ItTop,itLeft,disk.driveletter,hm.diskimage(disk.drivetype),"?folder="&disk&"\","disk","disk")
	  itId = itId + 1
	  itLeft = itLeft + 80
	  if itLeft > 690 then : itLeft=210 : itTop=itTop+80 : end if
	next
end if
%>
<script>document.getElementById('allpe').style.display = 'none';</script>
<table width="100%" height="<%if itTop<613 then :response.write "100%" :else:response.Write itTop+100 : end if%>"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <!---------------  left menu start   ----------------->
	<td width="210">
	<table  style="height:100%;width:204px!important;width:100%;border:1px solid #fff;background-color:#F4F2E7;border-collapse:collapse;" cellpadding="0" cellspacing="0" bgcolor="#7AA1E6">
      <tr>
        <td><table style="height:100%;width:202px!important;width:100%;"  border="0" cellpadding="3" cellspacing="3">
          <tr>
            <td valign="top"><iframe id="menu" src="<%=mypage%>?act=menu&folder=<%=hm.javaurl(folder)%>" height="100%"  frameborder="0" framespacing="0" vspace=0 marginheight=0 marginwidth=0 style="border:0px;width:196px!important;width:100%;"></iframe></td>
          </tr>
          <tr>
            <td height="20"><%=hm.CopyRight%></td>
          </tr>
        </table></td>
      </tr>
    </table>
	</td>
    <!---------------  left menu end   ----------------->
	<!---------------  body start   ----------------->
	<td valign="top">
		<%
		if not(folder="") then
		%>
		<div id="image1" style="position:absolute; width:80px; height:80px; z-index:1; left:210px; top:0px;"
		 onmousedown="imageNum=1;down=true;divleft=event.clientX-parseInt(image1.style.left);divtop=event.clientY-parseInt(image1.style.top)"onmouseup="down=false">
		  <table width="80" height="60" border="0" cellpadding="0" cellspacing="0" style="cursor:default;">
			<tr>
			  <td height="40"><div align="center"><font style="color:#333333;">
				<font style='font-size:25pt;font-family:Wingdings;'
				 onDblClick="doubleClick(1,'self','<%="?folder="&hm.javapath(hm.javaurl(folder))%>');"
				 ><!onclick="oneClick('<%if right(folder,1)=":" then : response.write "disk" : else : response.write "folder" : end if%>','<%=hm.diskpath(folder)%>','<%=hm.javapath(hm.javaurl(folder))%>','<%=hm.javapath(hm.javaurl(folder))%>');">0</font>
			  </font></div></td>
			</tr>
			<tr>
			  <td><div align="center" id="link1"><a href="<%=mypage&"?folder="&hm.javaurl(folder)%>"><font style="font-size:9pt;color:#333333;">.</font></a></div></td>
			</tr>
		  </table>
		</div>
		<div id="image2" style="position:absolute; width:80px; height:80px; z-index:1; left:280px; top:0px;"
		 onmousedown="imageNum=2;down=true;divleft=event.clientX-parseInt(image2.style.left);divtop=event.clientY-parseInt(image2.style.top)"onmouseup="down=false">
		  <table width="80" height="60" border="0" cellpadding="0" cellspacing="0" style="cursor:default;">
			<tr>
			  <td height="40"><div align="center"><font style="color:#333333;">
				<font style="font-size:25pt;font-family:Wingdings;"
				 onDblClick="doubleClick(2,'superior','<%="?folder="&hm.javapath(hm.javaurl(hm.superfolder(folder)))%>');"
				 ><!onclick="oneClick('<%if right(hm.superfolder(folder),1)=":" or hm.superfolder(folder)="" or right(hm.superfolder(folder),1)="\"  then : response.write "disk" : else : response.write "folder" : end if%>','<%=hm.diskpath(folder)%>','<%if right(hm.superfolder(folder),1)="\" then : response.write hm.javapath(hm.javaurl(left(hm.superfolder(folder),len(hm.superfolder(folder))-1))) : else : response.write hm.javapath(hm.javaurl(hm.superfolder(folder))):end if%>','<%=hm.javapath(hm.javaurl(hm.superfolder(folder)))%>');">0</font>
				</font></div></td>
			</tr>
			<tr>
			  <td><div align="center" id="link2"><a href="<%=mypage&"?folder="&hm.superfolder(hm.javaurl(folder))%>"><font style="font-size:9pt;color:#333333;">..</font></a></div></td>
			</tr>
		  </table>
		</div>
		<%
		end if
		%>
	</td>
	<!---------------  body end   ----------------->
  </tr>
</table>
</body>
</html>

<%'---------------------------------------------------   ×Ó¹ý³Ì£¬ÏÔÊ¾»òÐÞ¸ÄÎÄ¼þ   ------------------------------------------------------
sub showfile()
folder = request("folder")
file = request("file")
act = request("act")
fileExt = hm.name3(file)
nameext = lcase(fileExt)
if right(folder,1)<>"\" then folder = folder & "\"
filepath = folder & file
if (nameext="jpeg" or nameext="jpg" or nameext="gif" or nameext="bmp" or nameext="png") and hm.IsObjInstalled("persits.jpeg") then
	showfile_photo(filepath)
	response.end
end if

rem ===== ²»ÖªµÀÊ²Ã´ÎÄ¼þÀàÐÍµÄÒÔÎÄ±¾´ò¿ª ======

if act = "edit" then call hm.savefile(folder,file,request("opentxt"),2)
set gettxt = fso.opentextfile(filepath)
dim con
while not gettxt.atendofstream
	con = con & gettxt.readline & vbcrlf
wend
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=filepath%></title>
</head>
<%showcss%>
<body style="border:0px;">
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
<form name="formtxt" method="post" action="<%=mypage&"?folder="&hm.javaurl(folder)&"&file="&hm.javaurl(file)&"&act=edit"%>">
  <tr bgcolor="#D4D0C8">
    <td height="17">
		<!input type="submit" name="open" value="´ò¿ª" style="border:1 soild;height:100%;"><!input type="button" name="new" value="ÐÂ½¨" style="border:1 soild;height:100%;"><input type="button" name="edit" value="±£´æ" style="border:1 soild;height:100%;" onClick="formtxt.submit();"><input type="reset" name="reset" value="»¹Ô­" style="border:1 soild;height:100%;"><input type="button" name="print" onClick="window.print();" value="´òÓ¡" style="border:1 soild;height:100%;"><input type="button" name="close" value="¹Ø±Õ" style="border:1 soild;height:100%;" onClick="window.close();"></td>
  </tr>
  <tr>
    <td>
      <textarea name="opentxt" style="width:100%;height:100%;border:1px soild #d4d0c8;word-wrap:break-work;word-break:break-all;"><%=server.HTMLEncode(con)%></textarea>
    </td>
  </tr>
</form>
</table>
</body></html>
<%
rem ===========================================
response.end
end sub
'--------------------------------------------------------------------------------------------------------------------------------------------

function showfile_photo(filepath)
jpeg.open filepath
imagewidth = jpeg.originalwidth 
imageheight = jpeg.originalheight
jpeg.close
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
body{margin:0px 0px;}
</style>
<title><%=filepath%> - Í¼Æ¬²é¿´Æ÷</title>
</head>
<script language='javascript' for='window' event='onload()'>
window.resizeTo(600,450);
</script>
<%showscript%>
<script language="javascript">
var mypage = "<%=mypage%>";
var pp = "<%=mypage&"?act=photoarr&path="&hm.javapath(hm.javaurl(filepath))%>";
var pa;
function photobig(){
    window.myimage.showimage.width = (parseInt(window.myimage.showimage.width) + 20);
}
function photosmall(){
    var newwidth = parseInt(window.myimage.showimage.width) - 20;
    if(newwidth>20)
    window.myimage.showimage.width = (parseInt(window.myimage.showimage.width) - 20);
}
function kk(obj) {  /// ´Ë·½·¨ÓÉ±¦±´²»¿Þ°ïÖúÍê³É
	var re = RegExp("Rotation=(.*)\\)", "ig");
	var progid = obj.style.filter.match(/Rotation=(.?)/ig)[0];
	return  parseInt(progid.substring(9));
}
function rotateL(){
	progid = kk(window.myimage.showimage);
	if(progid<3) {
		progid++;
	} else {
		progid=0;
	}
	window.myimage.showimage.style.cssText="cursor:move;FILTER:progid:DXImageTransform.Microsoft.BasicImage(Rotation="+progid+");";
}
function rotateT(){
	progid = kk(window.myimage.showimage);
	if(progid>0) {
		progid--;
	} else {
		progid=3;
	}
	window.myimage.showimage.style.cssText="cursor:move;FILTER:progid:DXImageTransform.Microsoft.BasicImage(Rotation="+progid+");";
}
function becoming(){
	var iw=<%=imagewidth%>,ih=<%=imageheight%>;
	var ipop = iw / ih ;
	var ww = document.body.clientWidth ,wh = document.body.clientHeight; 
	var imgw=0;imgh=0;
	while(true){
		imgw += ipop;
		imgh += 1;
		if(imgw>ww-5)break;
		if(imgh>wh-35)break;
		if(wh<35||ww<5)break;
	}
	window.myimage.imagemove.style.top = "0px";
	window.myimage.imagemove.style.left= "0px";
	window.myimage.showimage.width = imgw;
}
function fact(){
	window.myimage.showimage.width = <%=imagewidth%> ;
}
function photoprev(myself){
    if(pa[0]==myself){
        location.href = mypage+"?folder="+getfolder(pa[pa.length-1])+"&file="+nameAll(pa[pa.length-1])+"&act=open";
        return ;
    }
    for(i=0;i<pa.length;i++){
        if(pa[i]==myself){
            location.href = mypage+"?folder="+getfolder(pa[i-1])+"&file="+nameAll(pa[i-1])+"&act=open";
            return ;
        }
    }
}
function photonext(myself){
    if(pa[pa.length-1]==myself){
        location.href = mypage+"?folder="+getfolder(pa[0])+"&file="+nameAll(pa[0])+"&act=open";
        return ;
    }
    for(i=0;i<pa.length-1;i++){
        if(pa[i]==myself){
            location.href = mypage+"?folder="+getfolder(pa[i+1])+"&file="+nameAll(pa[i+1])+"&act=open";
            return ;
        }
    }    
}
function photoArr(photopath){
    var xml = connxml();
    xml.onreadystatechange = function(){
	    if(xml.readyState==4){
		    if(xml.status==200){
			    var PAstr = xml.responseText;
			    pa = PAstr.split("|");
			    /*for(i=0;i<pa.length;i++){
			        alert(pa[i]);
			    };*/
		    }
	    }
	}
	xml.open("GET",photopath,true);
	xml.send(null);
}

photoArr(pp);
</script>

<body>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" style="word-wrap:break-word;word-break:break-all;table-layout:fixed;">
  <tr>
    <td>
      <iframe id="myimage" name="myimage" src="<%=mypage%>?act=photo&filepath=<%=hm.javaurl(filepath)%>"
	  border=0 frameborder="0" framespacing="0" height="100%" width="100%" marginheight="0" marginwidth="0" noresize scrolling="auto" vspace="0"></iframe>
    </td>
  </tr>
  <tr>
    <td height="30">
		<div align="center">
		<table border="0" cellpadding="0" cellspacing="8" align="center"><tr>
		<td title="ÉÏÒ»¸öÍ¼Ïñ" style="cursor:hand;" onClick="photoprev('<%=hm.javapath(filepath)%>');"><span style="font-family:webdings;color:#666666;"><%=chr(57)%></span></td>
		<td title="ÏÂÒ»¸öÍ¼Ïñ" style="cursor:hand;" onClick="photonext('<%=hm.javapath(filepath)%>');"><span style="font-family:webdings;color:#666666;"><%=chr(58)%></span></td>
		<td style="width:1px;" bgcolor="#333333"></td>
		<td title="×îºÏÊÊ" style="cursor:hand;" onClick="becoming();"><span style="font-family:webdings;color:#666666;"><%=chr(71)%></span></td>
		<td title="Êµ¼Ê´óÐ¡" style="cursor:hand;" onClick="fact()"><span style="font-family:webdings;color:#666666;"><%=chr(67)%></span></td>
		<td title="¿ªÊ¼»ÃµÆÆ¬" style="cursor:hand;display:none;"><span style="font-family:wingdings 3;color:#666666;"><%=chr(93)%></span></td>
		<td style="width:1px;" bgcolor="#333333"></td>
		<td title="·Å´ó" style="cursor:hand;" onClick="photobig();"><span style="font-family:wingdings;color:#666666;"><%=chr(117)%></span></td>
		<td title="ËõÐ¡" style="cursor:hand;" onClick="photosmall();"><span style="font-family:wingdings;color:#666666;"><%=chr(119)%></span></td>
		<td style="width:1px;" bgcolor="#333333"></td>
		<td title="Ë³Ê±ÕëÐý×ª" style="cursor:hand;" onClick="rotateL();"><span style="font-family:wingdings 3;color:#666666;"><%=chr(80)%></span></td>
		<td title="ÄæË³Ê±ÕëÐý×ª" style="cursor:hand;" onClick="rotateT();"><span style="font-family:wingdings 3;color:#666666;"><%=chr(81)%></span></td>
		<td style="width:1px;" bgcolor="#333333"></td>
		<td title="É¾³ý" style="cursor:hand;"
		 onClick="javascript:if(confirm('ÄúÕæµÄÒªÉ¾³ý <%=hm.javapath(hm.backpath(filepath))%> ÎÄ¼þÂð?')){window.open('<%=mypage%>?act=task&run=deleteFile&file=<%=hm.javapath(hm.javaurl(filepath))%>&folder=<%=hm.javapath(hm.javaurl(filepath))%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');}else{return false;}"
		><span style="font-family:wingdings 2;color:#666666;font-size:20px;"><%=chr(51)%></span></td>
		<td title="´òÓ¡" onClick="javascript:window.print();" style="cursor:hand;"><span style="font-family:wingdings 2;color:#666666;font-size:20px;"><%=chr(54)%></span></td>
		<td title="¸´ÖÆµ½" style="cursor:hand;"
		 onClick="javascript:window.open('<%=mypage%>?act=task&run=copyfile&filename=<%=hm.javapath(hm.backpath(hm.javaurl(filepath)))%>&folder=<% =hm.javapath(hm.superfolder(hm.javaurl(filepath))) %>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');"
		><span style="font-family:wingdings;color:#666666;font-size:20px;"><%=chr(60)%></span></td>
		</tr></table>
		</div>
	</td>
  </tr>
</table>
</body></html>
<%	
	response.end
end function
'--------------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÒÔÍøÒ³¼ÓÔØÏÔÊ¾Í¼Æ¬  ----------------------------------------------------------
sub showphotoHTML()
filepath = request("filepath")
jpeg.open filepath
imagewidth = jpeg.originalwidth 
imageheight = jpeg.originalheight
porp = imagewidth / imageheight
while imagewidth>590 or imageheight>360
	imagewidth = imagewidth - porp
	imageheight = imageheight - 1
wend
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style>
body{background-color:#fff;}
</style>
<%showcss%>
<script>
var down=false,divleft,divtop;
function move()
{
	if(down)
	{
		window.imagemove.style.left=event.clientX-divleft;
		window.imagemove.style.top=event.clientY-divtop;
	}
}
function bbimg(o) 
{ 
  var zoom=parseInt(o.style.zoom, 10)||100; 
  zoom -= event.wheelDelta/12; 
  if (zoom>0) o.style.zoom = zoom+'%'; 
} 
</script>
</head>
<body onmousemove="move()">
<div id="imagemove" style="position:absolute; width:100%; height:100%; z-index:1; left:0px; top:0px;"
 onmousedown="down=true;divleft=event.clientX-parseInt(imagemove.style.left);divtop=event.clientY-parseInt(imagemove.style.top)" onMouseUp="down=false">
<div align="center" style="width:100%;height:100%;"><img id="showimage" width="<%=imagewidth%>"  name="showimage" onmousewheel="bbimg(this);" style="cursor:move;FILTER:progid:DXImageTransform.Microsoft.BasicImage(Rotation=0);" src="<%=mypage%>?act=preview&imagename=<%=hm.javaurl(filepath)%>" align="absmiddle" border=1></div>
</div>
</body>
</html>
<%
response.end
end sub
'---------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾Ô¤ÀÀÍ¼Æ¬£¬ÐèÒªaspjpegÖ§³Ö  --------------------------------------------------------------
sub showimageP()
	imgpath = request("imagename")
	getwid = cint(request("wview"))
	gethei = cint(request("hview"))
	jpeg.open imgpath
	if getwid=empty and gethei=empty then
		jpeg.width = jpeg.originalwidth 
		jpeg.height = jpeg.originalheight
	else
		iw = jpeg.originalwidth
		ih = jpeg.originalheight
		nw = iw
		nh = ih
		prop = nw / nh
		do while (nw>getwid or nh>gethei)
			if nw>5 then nw = nw - prop
			if nh>5 then nh = nh - 1
		loop
		if nw < 2 then nw = getwid
		if nh < 2 then nw = gethei
		jpeg.width = nw 
		jpeg.height = nh
	end if
	
	Randomize
	photopath = Server.MapPath(".") & "/"&year(now)&month(now)&day(now)&int(rnd*10000)&".jpg"
	jpeg.Save photopath
	ob_stream.open
	ob_stream.type=1
	ob_stream.loadfromfile photopath
	Response.ContentType = "image/pjpeg"
	response.BinaryWrite ob_stream.read
	
	set del = fso.GetFile(photopath)
	del.delete
	set del = nothing
	response.end
end sub
'---------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬·µ»Ø±¾Ä¿Â¼ÏÂµÄÍ¼Æ¬ÓÃ | ·Ö¸î  --------------------------------------------------------------
sub showPhotoArray()
	imgpath = request("path")
	imgfolderpath = hm.superfolder(imgpath)
	outArray = ""
	if right(imgfolderpath,1)<>"\" then imgfolderpath = imgfolderpath & "\"
	set photofso = fso.GetFolder(imgfolderpath)
	for each filename in photofso.files 
	     fileExt = hm.name3(filename)
         nameext = lcase(fileExt)
         if (nameext="jpeg" or nameext="jpg" or nameext="gif" or nameext="bmp" or nameext="png") then
	        outArray = outArray & hm.javaurl(filename) &"|"
	     end if
	next
	outArray = hm.rlchar(hm.outrepeat(outArray,"|"),"|")
	response.write outArray
	response.end
end sub
'---------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾ÏµÍ³ÊôÐÔ  ----------------------------------------------------------
sub showsystem()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ÏµÍ³ÊôÐÔ</title>
<style>
body{
background-color:#D4D0C8;
margin:0px 0px;
border:0px;
margin:0px;
}
</style>
</head>
<body>
<table width="100%" height="100%"  border="0" cellpadding="5">
  <tr>
    <td>
	  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="20"><input type="button" value="³£¹æ" style="height:19;width:50;font-size:9pt;border-bottom:0 solid ;"></td>
            </tr>
            <tr>
              <td><table width="100%" height="100%"  border="1" style="border-collapse:collapse;">
                <tr>
                  <td><table width="100%" height="100%"  border="0">
                    <tr>
                      <td width="150" valign="top"><table width="100%"  border="0">
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td><div align="center"><font style='font-size:100pt;font-family:Wingdings;'>:</font></div></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                      </table></td>
                      <td valign="top">
					   <table width="100%"  border="0" style="font-size:9pt;">
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td>&nbsp;·þÎñÆ÷Ãû: <%=request.servervariables("server_name")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;·þÎñÆ÷IP: <%=request.ServerVariables("LOCAL_ADDR")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;ÄúµÄIP: <%=request.ServerVariables("remote_addr")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;¶Ë¿ÚºÅ: <%=request.servervariables("server_port")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;±¾µØÂ·¾¶: <%=server.mappath("\")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;www·þÎñ°æ±¾: <%=request.servervariables("server_software")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;ÎÄ¼þÂ·¾¶: <%=request.servervariables("path_translated")%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;³ÌÐòÃû³Æ: <%=hm.SysName & " " & hm.SysVersion%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;×÷ Õß: <%="longerface"%></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                        </tr>
                        <tr>
                          <td><%=hm.CopyRight%></td>
                        </tr>
                      </table>
					 </td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="30"><div align="right">
            <input type="button" value="È·¶¨" style="width:70;height:21;" onClick="window.close();">&nbsp;&nbsp;</div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%response.end
end sub
'-------------------------------------------------------------------------------------------------------------------------------------
%>


<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾²Ëµ¥  --------------------------------------------------------------
sub showmenu()
dim menutype,menufolder,menufile,searchfolder
menutype = request("type")
menudisk = request("disk")
menufolder = request("folder")
menufile = request("file")
nowfolder = request("nowfolder")
if nowfolder="" then
	searchfolder = menufolder
else
	searchfolder = nowfolder
end if

if menutype="" then
	if mm.isdisk(folder) then
		menutype="disk"
	else
	    menutype="folder"
	end if
end if
menufolder = hm.backpath(menufolder)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>menu</title>
<style>
body{
background-color:#f4f2e7;
}
</style>
<%showcss%>
<SCRIPT language=javascript>
function tree(leaf){
if(eval(leaf+"show").style.display == "none"){
   eval(leaf+"show"+".style.display=\"\";");
   eval(leaf+".value=\"<%=chr(126)%>\";");
}else{
   eval(leaf+"show"+".style.display=\"none\";");
   eval(leaf+".value=\"<%=chr(128)%>\";");
};}

</script>
</head>

<body>
<table width="100%"  border="0" cellpadding="0" cellspacing="6">
  <tr>
    <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" style="border-collapse:collapse;">
        <tr onClick="tree('sys');" style="cursor:default;">
          <td height="20" bgcolor="#D4D0C8"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td>&nbsp;&nbsp;<font style="font-weight:bold;font-size:10pt;color:#000000;">ÏµÍ³ÈÎÎñ</font></td>
                <td width="20"><input id="sys" onClick="tree('sys');" type="button" value="<%=chr(126)%>" style="border:0;width:15;height:15;font-family:wingdings 3;font-size:7pt;"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FFFFFF" id="sysshow">
          <td><table width="100%"  border="0">
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:Webdings;">i</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=system','_blank','width=436,height=440,resizable=0,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">²é¿´ÏµÍ³ÐÅÏ¢</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:15pt;font-family:Webdings;">L</font></div></td>
                      <td>&nbsp;<a href="<%=mypage%>?act=search&searchfolder=<%=hm.javaurl(searchfolder)%>" target="_self" class="menulink"><span class="menufont">ËÑË÷ÎÄ¼þ»òÎÄ¼þ¼Ð</span></a></td>
                    </tr>
                </table></td>
              </tr>
			   <%if session("se")<>"" then%>
			   <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:16px;font-family:webdings;">x</font></div></td>
                      <td>&nbsp;<a href="<%=mypage%>?act=quit" target="_parent" class="menulink"><span class="menufont">ÍË³ö¹ÜÀí</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
          </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td>
	<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" style="border-collapse:collapse;<%if menufolder="" and menufile="" then response.write "display:'none';"%>">
        <tr onClick="tree('fil');" style="cursor:default;">
          <td height="20" bgcolor="#D4D0C8"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td>&nbsp;&nbsp;<font style="font-weight:bold;font-size:10pt;color:#000000;">ÎÄ¼þ»òÎÄ¼þ¼ÐÈÎÎñ</font></td>
                <td width="20"><input id="fil" onClick="tree('fil');" type="button" value="<%=chr(126)%>" style="border:0;width:15;height:15;font-family:wingdings 3;font-size:7pt;"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FFFFFF" id="filshow">
          <td><table width="100%"  border="0">
              <%if true then%>
			  <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:14pt;font-family:webdings;">v</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=newfolder&folder=<%if nowfolder="" then : response.write hm.javapath(hm.javaurl(folder)) : else : response.write hm.javapath(hm.javaurl(nowfolder)) : end if%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">´´½¨Ò»¸öÐÂÎÄ¼þ¼Ð</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <%end if%>
			  <%if menutype="folder" and nowfolder<>"" and false then%>
			  <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings;">!</font></div></td>
                      <td>&nbsp;<a onClick="javascript:window.open('<%=mypage%>?act=task&run=FolderRename&menufolder=<%=hm.javapath(hm.javaurl(menufolder))%>&folder=<%=hm.javapath(hm.javaurl(menufolder))%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÖØÃüÃûÕâ¸öÎÄ¼þ¼Ð</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <%end if%>
			  <%if menutype="folder" and nowfolder<>"" and false then%>
			  <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:12pt;font-family:wingdings 2;">%</font></div></td>
                      <td>&nbsp;<span class="menufont">ÒÆ¶¯Õâ¸öÎÄ¼þ¼Ð</span></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="folder" and nowfolder<>"" and false then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings 2;">+</font></div></td>
                      <td>&nbsp;<span class="menufont">¸´ÖÆÕâ¸öÎÄ¼þ¼Ð</span></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="folder" and nowfolder<>"" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:15pt;font-family:wingdings 2;">3</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=FolderDelete&menufolder=<%=hm.javapath(hm.javaurl(menufolder))%>&folder=<%=hm.javapath(hm.javaurl(menufolder))%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">É¾³ýÕâ¸öÎÄ¼þ¼Ð</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if true then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings;">v</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=newfile&folder=<%if nowfolder="" then : response.write hm.javapath(hm.javaurl(folder)) : else : response.write hm.javapath(hm.javaurl(nowfolder)) : end if%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÐÂ½¨Ò»¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="file" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings;">!</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=renamefile&oldfilename=<%=hm.javapath(hm.backpath(hm.javaurl(menufile)))%>&folder=<% =hm.javapath(hm.javaurl(nowfolder)) %>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÖØÃüÃûÕâ¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="file" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:12pt;font-family:wingdings 2;">%</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=movefile&filename=<%=hm.javapath(hm.backpath(hm.javaurl(menufile)))%>&folder=<% =hm.javapath(hm.javaurl(nowfolder)) %>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÒÆ¶¯Õâ¸öÎÄ¼þ</span><a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="file" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings 2;">+</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=copyfile&filename=<%=hm.javapath(hm.backpath(hm.javaurl(menufile)))%>&folder=<% =hm.javapath(hm.javaurl(nowfolder)) %>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">¸´ÖÆÕâ¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="file" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:15pt;font-family:wingdings 2;">3</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:if(confirm('ÄúÕæµÄÒªÉ¾³ý <%=hm.javapath(hm.backpath(menufile))%> ÎÄ¼þÂð?')){window.open('<%=mypage%>?act=task&run=deleteFile&file=<%=hm.javapath(hm.javaurl(menufile))%>&folder=<%=hm.javapath(hm.javaurl(nowfolder))%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');}else{return false;}" class="menulink"><span class="menufont">É¾³ýÕâ¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if menutype="file" then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:15pt;font-family:wingdings 3;">,</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=downFile&file=<%=hm.javapath(hm.javaurl(menufile))%>&folder=<%=hm.javapath(hm.javaurl(nowfolder))%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÏÂÔØÕâ¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
			  <%if true then%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:15pt;font-family:wingdings 3;">X</font></div></td>
                      <td>&nbsp;<a href="javascript:" onClick="javascript:window.open('<%=mypage%>?act=task&run=upFile&folder=<%if nowfolder="" then : response.write hm.javapath(hm.javaurl(folder)) : else : response.write hm.javapath(hm.javaurl(nowfolder)) : end if%>','_blank','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');" class="menulink"><span class="menufont">ÉÏ´«Ò»¸öÎÄ¼þ</span></a></td>
                    </tr>
                </table></td>
              </tr>
			  <%end if%>
          </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" style="border-collapse:collapse;">
        <tr onClick="tree('loc');" style="cursor:default;">
          <td height="20" bgcolor="#D4D0C8"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td>&nbsp;&nbsp;<font style="font-weight:bold;font-size:10pt;color:#000000;">ÆäËüÎ»ÖÃ</font></td>
                <td width="20"><input id="loc"  onclick="tree('loc');" type="button" value="<%=chr(126)%>" style="border:0;width:15;height:15;font-family:wingdings 3;font-size:7pt;"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FFFFFF" id="locshow">
          <td><table width="100%"  border="0">
<%if menutype="folder" then%>
              <tr>
                <td>
	<%
	set drive = fso.GetDrive(fso.GetFolder(hm.backurl(menufolder)).Drive)
	%>
				  <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:19pt;font-family:wingdings;">;</font></div></td>
                      <td>&nbsp;<a href=<%=mypage&"?folder="&drive&"\"%> target=_parent class=menulink><span class="menufont">±¾µØ´ÅÅÌ (<%=drive%>)</span></a></td>
                    </tr>
                  </table>
				</td>
              </tr>
<%end if%>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:17pt;font-family:wingdings;">:</font></div></td>
                      <td>&nbsp;<a href=<%=mypage%> target=_parent class=menulink><span class="menufont">ÎÒµÄµçÄÔ</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:13pt;font-family:wingdings;"><%=chr(38)%></font></div></td>
                      <td>&nbsp;<a href="<%=mypage%>?folder=<%=hm.javaurl(server.mappath("\"))%>" target=_parent class=menulink><span class="menufont">·þÎñÆ÷µØÖ·</span></a></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><div align="right"><font class="menuimage" style="font-size:11pt;font-family:wingdings;">1</font></div></td>
                      <td>&nbsp;<a href=<%=mypage&"?folder=C:\Documents%20and%20Settings\Default%20User\My%20Documents\"%> target=_parent class=menulink><span class="menufont">ÎÒµÄÎÄµµ</span></a></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" style="border-collapse:collapse;">
        <tr onClick="tree('dis');" style="cursor:default;">
          <td height="20" bgcolor="#D4D0C8"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td>&nbsp;&nbsp;<font style="font-weight:bold;font-size:10pt;color:#000000;">ÏêÏ¸ÐÅÏ¢</font></td>
                <td width="20"><input id="dis" onClick="tree('dis');" type="button" value="<%=chr(126)%>" style="border:0;width:15;height:15;font-family:wingdings 3;font-size:7pt;"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FFFFFF" id="disshow">
          <td>
<%if menutype="disk" then%>
		    <table width="100%"  border="0" <%if menutype="disk" then : response.write "style=""display:'';""" : else : response.write "style=""display:'none';""" : end if%>>
<%diskroot = true
  set disks = fso.drives
  for each d in disks
	if lcase(d) = lcase(menudisk) then 
	diskroot = false
	%>
				<tr>
					<td>&nbsp;<span class="menufont"><b><%=mm.dtype(d.drivetype)&" "&"("&d&")"%></b></span><br>
	&nbsp;<span class="menufont"><%=mm.dtype(d.drivetype)%></span></td>
				</tr>
		<%if d.isready then%>
					<tr>
						<td>&nbsp;<span class="menufont">ÎÄ¼þÏµÍ³: <%=d.filesystem%></span></td>
					</tr>
					<tr>
						<td>&nbsp;<span class="menufont">¿ÉÓÃ¿Õ¼ä: <%=mm.formatsize(d.freespace)%></span></td>
					</tr>
					<tr>
						<td>&nbsp;<span class="menufont">×Ü´óÐ¡: <%=mm.formatsize(d.totalsize)%></span></td>
					</tr>
		<%end if%>
	<%
	end if
  next
	if diskroot or menudisk="" then
%>
              <tr>
                <td>&nbsp;<span class="menufont"><b>ÎÒµÄµçÄÔ</b></span><br>
&nbsp;<span class="menufont">ÏµÍ³ÎÄ¼þ¼Ð</span></td>
              </tr>
<%	end if%>
            </table>

<%elseif menutype="folder" then%>
              <table width="100%"  border="0" <%if menutype="folder" then : response.write "style=""display:'';""" : else : response.write "style=""display:'none';""" : end if%> style="word-break:break-all;table-layout:fixed;">
	<%
	set folderdir = fso.getfolder(hm.backurl(menufolder))
	%>
                <tr>
                  <td>&nbsp;<span class="menufont"><b><%=hm.nameAll(menufolder)%></b></span><br>
&nbsp;<span class="menufont">ÎÄ¼þ¼Ð</span></td>
                </tr>
                <tr>
                  <td>&nbsp;<span class="menufont">ÐÞ¸ÄÈÕÆÚ:<%response.write folderdir.datelastmodified%></span></td>
                </tr>
              </table>

<%elseif menutype="file" then%>
              <table width="100%"  border="0" <%if menutype="file" then : response.write "style=""display:'';""" : else : response.write "style=""display:'none';""" : end if%> style="word-break:break-all;table-layout:fixed;">
	<%
	set filedir = fso.getfile(hm.backurl(menufile))
	%>
	<%if mm.isimage(hm.name3(menufile)) and hm.IsObjInstalled("persits.jpeg") then%>
				<tr><td>
				<div id="previewImage" style="margin-top:5px;margin-bottom:5px;margin-left:15px;margin-right:auto;width:120px;height:90px;text-algin:center;padding:2px 2px 2px 2px;background-color:#f5f5f5;border:1px solid #dddddd;" align=center>
				<img src="<%=mypage%>?act=preview&imagename=<%=hm.javaurl(menufile)%>&wview=120&hview=90" align="absmiddle">
				</div>
				</td></tr>
	<%end if%>
				<tr>
                  <td>&nbsp;<span class="menufont"><b><%=hm.nameall(menufile)%></b></span><br>
&nbsp;<span class="menufont"><%=mm.filename(hm.name3(menufile))%></span></td>
                </tr>
                <tr>
                  <td>&nbsp;<span class="menufont">ÐÞ¸ÄÈÕÆÚ:<%=filedir.datelastmodified%></span></td>
                </tr>
                <tr>
                  <td>&nbsp;<span class="menufont">´óÐ¡: <%=mm.formatsize(filedir.size)%></span></td>
                </tr>
            </table>
<%end if%>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%response.end
end sub
'------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾ËÑË÷²Ëµ¥  --------------------------------------------------------------
sub showsearch()
	searchfolder = request("searchfolder")
	keyword      = request("keyword")
	searching    = trim(request("searching"))
	if searching = "yes" then
		searching = "true"
	else
		searching = "false"
	end if
	if searchfolder="" then
		backfolder = mypage
	else
		if instr(searchfolder,";")<>0 then
			foldertemp = split(searchfolder,";")
			backfolder = foldertemp(0)
		else
			backfolder = hm.javapath(searchfolder)
		end if
	end if
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>menu</title>
<%showcss%>
<style>
body {
	background-color:#f4f2e7;
}
</style>
<SCRIPT language=javascript>
function tree(leaf){
if(eval(leaf+"show").style.display == "none"){
   eval(leaf+"show"+".style.display=\"\";");
   eval(leaf+".value=\"<%=chr(126)%>\";");
}else{
   eval(leaf+"show"+".style.display=\"none\";");
   eval(leaf+".value=\"<%=chr(128)%>\";");
};}

</script>
</head>

<body>
<table width="100%"  border="0" cellpadding="0" cellspacing="6">
  <tr>
    <td>
	<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" style="border-collapse:collapse;">
        <tr style="cursor:default;">
          <td height="20" bgcolor="#D4D0C8"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td>&nbsp;&nbsp;<font style="font-weight:bold;font-size:10pt;color:#000000;">ËÑË÷ÖúÀí</font></td>
                <td width="20"><input id="fil" onClick="parent.location.href='<%=mypage&"?folder="&hm.javaurl(backfolder)%>'" type="button" value="r" style="border:0;width:15;height:15;font-family:webdings;font-size:9pt;"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td>
			
			<table width="100%"  border="0" cellpadding="4">
              <form method="post" name="searchform" action="<%=mypage%>?act=scan&nowfolder=<%=hm.javaurl(searchfolder)%>" target="_parent">
              <tr>
                <td><span class="menufont"><b>°´ÏÂÃæÈÎºÎ»òËùÓÐ±ê×¼½øÐÐËÑË÷</b></span></td>
              </tr>
              <tr>
                <td><span class="menufont">È«²¿»ò²¿·ÖÎÄ¼þÃû£º</span></td>
              </tr>
              <tr>
                <td>
					<input type="text" name="keyword" style="width:150px;" value="<%=keyword%>">
			    </td>
              </tr>
              <!--
			  <tr>
                <td><span class="menufont">ÎÄ¼þÖÐµÄÒ»¸ö×Ö»ò´Ê×é£º</span></td>
              </tr>
              <tr>
                <td>
					<input type="text" name="wordinfile" value="" class="textin">
				</td>
              </tr>
              -->
			  <tr>
                <td><span class="menufont">ÔÚÕâÀï²éÕÒ£º</span></td>
              </tr>
              <tr>
                <td>
					<div style="font-family:Verdana;font-size:12px;">
					<input name="searchfolder" id="searchfolder" style="width:135px;height:22px" onFocus="this.select();" value="<%=searchfolder%>">
					<span style="position:absolute;margin:1px 1px 1px -6px">
					<select name="searchfolderselect" id="searchfolderselect" style="font-family:Verdana;margin-left:-132px;width:150px;" onChange="document.getElementById('searchfolder').value=this.value;">
					  	<%if searchfolder<>"" then%>
					  	<option value="<%=searchfolder%>" title="<%=searchfolder%>"><%=searchfolder%></option>
					  	<%end if%>
					  	<optgroup label="ÎÒµÄµçÄÔ">
					  	<option value="<%
					  		set dr = fso.drives
							for each disk in dr
								if disk.drivetype=2 then
									response.Write disk.driveletter&":;"
								end if							
							next
					  	%>" <%if searchfolder="" then response.write " selected"%>>
					  	<%
					  		set dr = fso.drives
							for each disk in dr
								if disk.drivetype=2 then
									response.Write disk.driveletter&":;"
								end if							
							next
					  	%>
					  	</option>
					  	<%
					  		set dr = fso.drives
							for each disk in dr 
							%><option value="<%=disk.driveletter%>:\">
								<%=hm.diskname(disk.drivetype)%>
								(<%=disk.driveletter%>:)
							</option>
							<%
							next
					  	%>
					  	</optgroup>
					  	<!--<option>ä¯ÀÀ...</option>-->
					</select>
					<script language="javascript" defer>document.getElementById('searchfolder').value=document.getElementById('searchfolderselect').value;</script>
					</span>
					</div>

				</td>
              </tr>
              <tr id="searching" style="display:'none';">
                <td align="center">
					<div style="width:120px;height:17px;border:1px solid #666;padding:2px;">
						<marquee direction="right" scrollamount="10" scrolldelay="0" width="116" height="13">
	 						<div class="searchwait"></div><div class="searchwait"></div><div class="searchwait"></div>
						</marquee>
					</div>
                </td>
              </tr>
			  <script language="javascript" defer>if(<%=searching%>){searching.style.display='';}</script>
			  <tr>
                <td><span class="menufont"></span></td>
              </tr>
			  <tr>
                <td><span class="menufont"></span></td>
              </tr>
              <tr>
                <td>
					  <div style="text-align:center;">
					  <input type="button" value="ºóÍË" onClick="window.parent.history.back();"/>
					  <input type="submit" value="ËÑË÷" />
					  </div>
				</td>
              </tr>
			  </form>
			</table>
			
          </td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%response.end
end sub
'------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾ËÑË÷½á¹ûÁÐ±í¿ò¼ÜÒ³  --------------------------------------------------------------
sub showSearchList()
	dim keyword,searchfolder,nowfolder
	keyword = request("keyword")
	searchfolder = request("searchfolder")
	nowfolder = request("nowfolder")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ÏÔÊ¾ËÑË÷½á¹û</title>
<%showcss%>
</head>
<body scroll="no">
<table width="100%" height="<%if itTop<613 then :response.write "100%" :else:response.Write itTop+100 : end if%>"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <!---------------  left menu start   ----------------->
	<td width="210">
		<table  style="height:100%;width:204px!important;width:100%;border:1px solid #fff;background-color:#F4F2E7;border-collapse:collapse;" cellpadding="0" cellspacing="0" bgcolor="#7AA1E6">
		  <tr>
			<td><table style="height:100%;width:202px!important;width:100%;"  border="0" cellpadding="3" cellspacing="3">
			<tr>
				<td valign="top"><iframe id="menu" src="<%=mypage%>?act=search&searching=yes&searchfolder=<%=hm.javaurl(searchfolder)%>&keyword=<%=hm.javaurl(keyword)%>" style="border:0px;width:196px!important;width:100%;" height="100%"  frameborder="0" framespacing="0" vspace=0 marginheight=0 marginwidth=0></iframe></td>
			</tr>
			<tr>
				<td height="20"><%=hm.CopyRight%></td>
			</tr>
			</table></td>
		  </tr>
		</table>
	</td>
    <!---------------  left menu end   ----------------->
	<!---------------  body start   ----------------->
	<td valign="top">
	    <div id="magnifier" style="position:absolute;left:100px;top:100px;width:300px;height:300px;">
        <span style="font-family:webdings;font-size:300px;color:#999;">L</span>
        </div>
        <script>
        var r = 100;
        var pi = Math.PI;
        var sx = 100;
        var sy = 100;
        var bu = 0;
        function scan(bu){
            sx = parseInt(r + r*Math.cos(bu*((2*pi)/360)));
            sy = parseInt(r + r*Math.sin(bu*((2*pi)/360)));
            document.getElementById('magnifier').style.left = sx + 200 + "px";
            document.getElementById('magnifier').style.top  = sy + 20 + "px";
            bu++
            if(bu>360) bu=0;
            setTimeout("scan("+bu+");",1);
        }
        /////////////////////////////////////////
        scan(bu);
        </script>
	    <iframe id="scanresult" src="<%=mypage%>?act=searchResult&nowfolder=<%=hm.javaurl(nowfolder)%>&searchfolder=<%=hm.javaurl(searchfolder)%>&keyword=<%=hm.javaurl(keyword)%>" frameborder="0" framespacing="0" vspace=0 marginheight=0 marginwidth=0 style="border:0px;width:100%;height:100%;"></iframe>
	</td>
	<!---------------  body end   ----------------->
  </tr>
</table>
</body>
</html>
<%
response.end
end sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
%>

<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾ËÑË÷µ½µÄ½á¹û  ----------------------------------------------------------
sub showSearchResult(searchfolder,keyword)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<%showcss%>
<style>
span{cursor:default;}
body{background-color:#fff;}
</style>
<%showscript%>
<script language="javascript">
var clickid=0;
function doubleClick(self,theRun,theUrl)
{
	self.style.backgroundColor = "#cccccc";
	if(clickid!=0 && clickid!=self.id){
		eval(clickid).style.backgroundColor = "";	
	}
	clickid = self.id;
	if(theRun=="disk")
		window.parent.location.href = theUrl;//alert("´ò¿ª±¾ÅÌ·û "+theUrl);
	if(theRun=="folder")
		window.parent.location.href = theUrl;//alert("½øÈëÏÂ¼¶Ä¿Â¼ "+theUrl);
	if(theRun=="file")
	{
		var url = window.open("","popupnav","width=550,height=400,resizable=1,status=yes,scrollbars=no,menubar=no,location=no");
		url.location.href = "<%=mypage%>?folder="+escape(getfolder(theUrl))+"&file="+escape(nameAll(theUrl))+"&act=open" ;
	}
}
function oneClick(self)
{
	self.style.backgroundColor = "#cccccc";
	if(clickid!=0 && clickid!=self.id){
		eval(clickid).style.backgroundColor = "";
	}
	clickid = self.id;
}
</script>
</head>
<body>
        <table style="width:100%;" class="f9p" cellpadding=0 cellspacing=0>
			<tbody>
				<tr style="background-color:#000;">
					<td class="searchtop">&nbsp;&nbsp;Ãû³Æ</td>
					<td class="searchtop">&nbsp;&nbsp;ËùÔÚÎÄ¼þ¼Ð</td>
					<td class="searchtop" style="text-align:right;">´óÐ¡&nbsp;&nbsp;</td>
					<td class="searchtop">&nbsp;&nbsp;ÀàÐÍ</td>
					<td class="searchtop">&nbsp;&nbsp;ÐÞ¸ÄÈÕÆÚ</td>
					<td class="searchtop" style="display:none;">&nbsp;</td>
				</tr>
				<%
				if searchfolder<>"" then
					list = split(searchfolder,";")
					hm.ScanResultCount = 0
					for i=0 to ubound(list)
						if mm.righthere("folder",hm.backurl(list(i)),"") then
							hm.search 0,hm.backurl(list(i)),keyword
						else
							response.Write "<scr"&"ipt>alert('"&hm.javapath(hm.backurl(list(i)))&" ²»ÊÇÒ»¸öÓÐÐ§Â·¾¶¡£');//window.parent.menu.searching.style.display='none';</scr"&"ipt>"
						end if
					next
					if hm.ScanResultCount=0 then response.Write "<tr><td><div style=""margin:5px;"">ËÑË÷Íê±Ï£¬Ã»ÓÐ½á¹û¿ÉÏÔÊ¾¡£</div></td></tr>"
				end if
				%>
			</tbody>
		</table>
<script>
setTimeout("window.parent.menu.searching.style.display='none';",1000);
setTimeout("window.parent.magnifier.style.display='none';",1000);
</script>
</body>
</html>
<%
response.end
end sub
'---------------------------------------------------------------------------------------------------------------------------------------
%>


<%'---------------------------------------------------  ×Ó¹ý³Ì£¬ÏÔÊ¾²Ëµ¥  --------------------------------------------------------------
sub showtask()

run = request("run")
nowfolder = request("folder")


if run = "newfolder" then  ''''''''''''''  ÐÂ½¨ÎÄ¼þ¼ÐÏÔÊ¾Ò³  '''''''''
	newfoldername = request("newfoldername")
	if newfoldername="" then newFolderName = "ÐÂ½¨ÎÄ¼þ¼Ð"
	if right(nowfolder,1)<>"\" then nowfolder = nowfolder & "\"
	if mm.righthere("folder",nowfolder,newfoldername) then
		for foldernumber=2 to 9999
			newFolderName = "ÐÂ½¨ÎÄ¼þ¼Ð" &" ("&foldernumber&")"
			if not(mm.righthere("folder",nowfolder,newfoldername)) then exit for
		next
	end if
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.newFolderName.new_folder.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="true" then response.write "<font size=2>ÎÄ¼þ¼Ð <font  color=blue>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
	response.write "<font size=2>ÄúÒªÔÚ <font color=red>"&nowfolder&"</font> ÎÄ¼þ¼ÐÏÂÐÂ½¨£º</font><br>"&vbcrlf
	response.write "<form name='newFolderName' method='post' action='"&mypage&"?act=task&run=DoNewFolder'><input name='now_folder' type='hidden' value='"&nowfolder&"'><input name='new_folder' value='"&newFolderName&"'>&nbsp;<input type=submit value=ÐÂ½¨></form>"&vbcrlf
	response.end
end if

if run = "DoNewFolder" then  ''''''''''''''  ÐÂ½¨ÎÄ¼þ¼ÐÖ´ÐÐÒ³  '''''''''
	now_folder = request("now_folder")
	new_folder = request("new_folder")
	mm.specialCharacter(new_folder)
	if mm.righthere("folder",now_folder,new_folder) then
		response.redirect mypage&"?act=task&run=newfolder&folder="&now_folder&"&being=true&beingfile="&new_folder
	end if
	fso.createfolder(now_folder&new_folder)
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&now_folder&new_folder&"</font><br> ÎÄ¼þ¼Ð³É¹¦ÐÂ½¨¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	response.end
end if
if run = "FolderDelete" then  ''''''''''''''  É¾³ýÎÄ¼þ¼ÐÏÔÊ¾Ò³  '''''''''
	menufolder = request("menufolder")
	'response.write mm.righthere("folder",nowfolder,"")
	'response.end
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.deleteFolder.delete_folder.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2>ÄúÒªÉ¾³ýÎÄ¼þ¼Ð£º <font color=red><br>"&nowfolder&"</font> </font><br>"&vbcrlf
	response.write "<form name='deleteFolder' method='post' action='"&mypage&"?act=task&run=DoDeleteFolder'><input name='delete_folder' type='hidden' value='"&nowfolder&"'>&nbsp;<input type=submit value=É¾³ý></form>"&vbcrlf
	response.end
end if

if run = "DoDeleteFolder" then  ''''''''''''''  É¾³ýÎÄ¼þ¼ÐÖ´ÐÐÒ³  '''''''''
	delete_folder = request("delete_folder")
	if not mm.righthere("folder",delete_folder,"") then
		response.write "´ËÎÄ¼þ¼Ð²»´æÔÚ¡£"
		response.end
	end if
	fso.deletefolder(hm.getfolderpath(delete_folder))
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&hm.getfolderpath(delete_folder)&"</font><br> ÎÄ¼þ¼ÐÉ¾³ý³É¹¦¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	response.end
end if

if run = "newfile" then  ''''''''''''''  ÐÂ½¨ÎÄ±¾ÎÄ¼þÏÔÊ¾Ò³  '''''''''
	newFileName = "ÐÂ½¨ ÎÄ±¾ÎÄµµ.txt"
	if right(nowfolder,1)<>"\" then nowfolder = nowfolder & "\"
	if mm.righthere("file",nowfolder,newfilename) then
		for filenumber=2 to 9999
			newFileName = "ÐÂ½¨ ÎÄ±¾ÎÄµµ" &" ("&filenumber&").txt"
			if not(mm.righthere("file",nowfolder,newfilename)) then exit for
		next
	end if
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.newFileName.new_file.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="true" then response.write "<font size=2>ÎÄ¼þ <font  color=blue>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
	response.write "<font size=2>ÄúÒªÔÚ <font color=red>"&nowfolder&"</font> ÎÄ¼þ¼ÐÏÂÐÂ½¨£º</font><br>"&vbcrlf
	response.write "<form name='newFileName' method='post' action='"&mypage&"?act=task&run=DoNewFile'><input name='now_folder' type='hidden' value='"&nowfolder&"'><input name='new_file' value='"&newFileName&"'>&nbsp;<input type=submit value=ÐÂ½¨></form>"&vbcrlf
	response.end
end if

if run = "DoNewFile" then  ''''''''''''''  ÐÂ½¨ÎÄ¼þÖ´ÐÐÒ³  '''''''''
	now_folder = request("now_folder")
	new_file = request("new_file")
	mm.specialCharacter(new_file)
	if mm.righthere("file",now_folder,new_file) then
		response.redirect mypage&"?act=task&run=newfile&folder="&now_folder&"&being=true&beingfile="&new_file
	end if
	fso.createtextfile(now_folder&new_file)
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&now_folder&new_file&"</font><br> ÎÄ¼þ³É¹¦ÐÂ½¨¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	response.end
end if

if run="FolderRename" then  ''''''''''''  ÎÄ¼þ¼Ð¸ÄÃûÏÔÊ¾Ò³  '''''''''''''
	'response.write nowfolder
	newFolderName = hm.nameAll(request("menufolder"))
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.modifyFolderName.inputRename_folder.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="true" then response.write "<font size=2>ÎÄ¼þ¼Ð <font  color=blue>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
	response.write "<font size=2>ÄúÒª½« <font color=red>"&nowfolder&"</font> ÎÄ¼þ¼Ð¸ÄÃûÎª£º</font><br>"&vbcrlf
	response.write "<form name='modifyFolderName' method='post' action='"&mypage&"?act=task&run=DoFolderRename'><input name='now_folder' type='hidden' value='"&nowfolder&"'><input name='inputRename_folder' value='"&newFolderName&"'>&nbsp;<input type=submit value=¸ÄÃû></form>"&vbcrlf
	response.end
end if

if run="upFile" then     ''''''''''''  ÎÄ¼þÉÏ´«ÏÔÊ¾Ò³  '''''''''''''
	if request("being")="true" then response.write "<font size=2>ÎÄ¼þ <font  color=blue>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
	response.write "<font size=2>ÄúÒªÉÏ´«ÎÄ¼þµ½ <font color=red>"&nowfolder&"</font> ÎÄ¼þ¼ÐÏÂ£º</font>"&vbcrlf
	response.write "<form name='upfileform' enctype=""multipart/form-data"" method='post' action='"&mypage&"?act=task&run=DoUpFile&folder="&hm.javaurl(nowfolder)&"'>"&vbcrlf&"<input name='upfile' type='file' value=''><br>"&vbcrlf&"&nbsp;<input name='mysubmit' type=button value=ÉÏ´« onclick='uploading();'>&nbsp; <span id='myuploading' style='font-family:arial;font-size:9pt;color:#666666;' style=display:'none';>uploading<span id=dot>..........</span></span></form>"&vbcrlf
	response.write "<script>"&vbcrlf
	response.write "function uploading(){	document.upfileform.mysubmit.disabled = true;	myuploading.style.display='';	showdot(); upfileform.submit();}"&vbcrlf
	response.write "function showdot(){	var num=0,char='';	num = parseInt(Math.random()*10+1);	for(i=0;i<num;i++)	  char = char +'.';	dot.innerHTML=char;	setTimeout('showdot();',20);}"&vbcrlf
	response.write "</script>"
	response.end
end if

if run="DoUpFile" then     ''''''''''''  ÎÄ¼þÉÏ´«Ö´ÐÐÒ³  '''''''''''''
	if right(nowfolder,1)<>"\" then nowfolder = nowfolder & "\"
	upfile = request("upfile")
	upfilename = hm.nameAll(upfile)
	upfilepath = nowfolder & upfilename
	if request("overfile")<>"true" then
		if fso.fileexists(upfilepath) then
			response.write "<form name='upfileform' method='post' enctype=""multipart/form-data"" action='"&mypage&"?act=task&upfile="&hm.javaurl(upfile)&"&run=DoUpFile&folder="&hm.javaurl(nowfolder)&"&overfile=true'>"&vbcrlf&"<!input name='upfile' type='file' value='"&upfile&"'><br>"&vbcrlf&"</form>"&vbcrlf
			response.write "<font size=2><font color=red>"&upfilepath&"</font><br>ÎÄ¼þÒÑ¾­´æÔÚ£¬ÇëÈ·ÈÏ¸²¸Ç´ËÎÄ¼þ£¡</font><br><br>"
			response.write "<input type=button value='¸²¸Ç' onclick='upfileform.submit();'> &nbsp;"
			response.write "<input type=button value='ÍË»Ø' onclick=location.href='"&mypage&"?act=task&run=upFile&beingfile="&hm.javapath(hm.javaurl(upfilepath))&"&folder="&hm.javapath(hm.javaurl(nowfolder))&"&being=true'>"
			response.end
		end if
	end if
	if fso.fileexists(upfilepath) then
		set del = fso.GetFile(upfilepath)
		del.delete
		set del = nothing
	end if
	rem ======================================
	rem ob_stream.type = 1 
	rem ob_stream.open
	rem ob_stream.loadfromfile upfile
	rem ob_stream.savetofile upfilepath
	rem ob_stream.close
	rem ======================================
	rem //////////////////////////////////////
	set upload = new upload_file
	for each formName in upload.File
		set file=upload.File(formName)
		'fileExt=lcase(file.FileExt)
		if file.filesize<5 then
			response.write "<span><font size=5 color=red><b>´íÎó</b></font><br>ÇëÏÈÑ¡ÔñÄãÒªÉÏ´«µÄÎÄ¼þ£¡¡¡<br><a href=""javascript:history.go(-1);"">ÖØÐÂÉÏ´«</a></span>"
			response.end
		end if
		savefilepath = nowfolder&file.FileName
		if file.FileSize>0 then         ''Èç¹û FileSize > 0 ËµÃ÷ÓÐÎÄ¼þÊý¾Ý
			file.SaveToFile savefilepath
		end if
		set file=nothing
	next
	set upload=nothing
	rem //////////////////////////////////////
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2>ÄúÒª³É¹¦ÉÏ´«ÎÄ¼þÔÚ <font color=red>"&upfilepath&"</font> </font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	response.end
end if

if run="downFile" then     ''''''''''''  ÎÄ¼þÏÂÔØÏÔÊ¾Ò³  '''''''''''''
	response.write "<font size=2>ÄúÒªÏÂÔØÎÄ¼þ <font color=red>"&hm.backpath(request("file"))&"</font> </font>"&vbcrlf
	response.write "<br><br><input name='down' type=button value=ÏÂÔØ onclick=""/*downloading();*/;window.open('"&mypage&"?act=task&run=DoDownFile&file="&hm.javapath(hm.backpath(hm.javaurl(request("file"))))&"&folder="&hm.javapath(hm.javaurl(request("folder")))&"','_self','width=300,height=150,resizable=1,status=no,scrollbars=no,menubar=no,location=no');WindowClose.style.display='';"">&nbsp;<input type=button name=WindowClose value='¹Ø±Õ' onclick='window.close();' style='display:none;'>&nbsp; <span id='mydownloading' style='font-family:arial;font-size:9pt;color:#666666;' style=display:'none';>downloading<span id=dot>..........</span></span>"&vbcrlf
	response.write "<script>"&vbcrlf
	response.write "function downloading(){	down.disabled = true;	mydownloading.style.display='';	showdot(); }"&vbcrlf
	response.write "function showdot(){	var num=0,char='';	num = parseInt(Math.random()*10+1);	for(i=0;i<num;i++)	  char = char +'.';	dot.innerHTML=char;	setTimeout('showdot();',20);}"&vbcrlf
	response.write "</script>"
	response.end
end if

if run="DoDownFile" then     ''''''''''''  ÎÄ¼þÏÂÔØÖ´ÐÐÒ³  '''''''''''''
	ob_stream.open
	ob_stream.type=1
	ob_stream.loadfromfile request("file")
	Response.AddHeader "Content-Disposition", "attachment; filename=" & hm.nameAll(request("file"))
	response.ContentType="application/octet-stream"
	
	response.BinaryWrite ob_stream.read
	response.end
end if

if run="deleteFile" then     ''''''''''''  ÎÄ¼þÉ¾³ýÖ´ÐÐÒ³  '''''''''''''
	if fso.fileexists(request("file")) then
	set del = fso.GetFile(request("file"))
	del.delete
	set del = nothing
	response.write "<font size=2>ÎÄ¼þ <font color=red>"&hm.backpath(request("file"))&"</font> É¾³ý³É¹¦¡£</font><br><br><input type=button value=¹Ø±Õ onclick='window.close();'>"&vbcrlf
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	end if
end if

if run="renamefile" then  ''''''''''''  ÎÄ¼þ¸ÄÃûÏÔÊ¾Ò³  '''''''''''''
	newFileName = request("newfilename")
	oldFileName = request("oldfilename")
	if newFileName = "" then newfilename = oldfilename  
	nowfolder = request("folder")
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.modifyFileName.newfilename.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="true" then response.write "<font size=2>ÖØÃüÃûÎÄ¼þ <font  color=blue>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
	response.write "<font size=2>ÄúÒª½« <font color=red>"&oldfilename&"</font> ÎÄ¼þ¸ÄÃûÎª£º</font><br>"&vbcrlf
	response.write "<form name='modifyFileName' method='post' action='"&mypage&"?act=task&run=DoFileRename'><input name='now_folder' type='hidden' value='"&nowfolder&"'><input name='oldfilename' type=hidden value='"&hm.nameAll(oldFileName)&"'><input name='newfilename' value='"&hm.nameAll(newFileName)&"'>&nbsp;<input type=submit value=¸ÄÃû></form>"&vbcrlf
	response.end
end if

if run = "DoFileRename" then  ''''''''''''''  ÎÄ¼þ¸ÄÃûÖ´ÐÐÒ³  '''''''''
	now_folder = request("now_folder")
	new_file = request("newfilename")
	old_file = request("oldfilename")
	mm.specialCharacter(new_file)
	if mm.righthere("file",now_folder,new_file) then
		response.redirect mypage&"?act=task&run=renamefile&newfilename="&hm.javaurl(new_file)&"&oldfilename="&hm.javaurl(old_file)&"&folder="&hm.javaurl(now_folder)&"&being=true&beingfile="&hm.javaurl(new_file)
	end if
	set thefile = fso.GetFile(now_folder&old_file)
	thefile.move(now_folder&new_file)
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&now_folder&new_file&"</font><br> ÎÄ¼þÖØÃüÃû³É¹¦¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	set thefile=nothing
	response.end
end if

if run="movefile" then  ''''''''''''  ÎÄ¼þÒÆ¶¯ÏÔÊ¾Ò³  '''''''''''''
	FileName = request("filename")
	nowfolder = request("folder")
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.moveFileName.inputmove_folder.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="false" then response.write "<font size=2>ÎÄ¼þ¼Ð <font  color=blue>"&request("beingfolder")&"</font> ²»´æÔÚ¡£</font><br>"&vbcrlf
	if request("being")="true" then
		response.write "<font size=2>ÎÄ¼þ <font  color=red>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
		response.write "<form name='moveFileName' method='post' action='"&mypage&"?act=task&run=DoFileMove&overhere=over'><input name='now_file' type='hidden' value='"&FileName&"'><input name='inputmove_folder' value='"&nowfolder&"'>&nbsp;<input type=submit value=¸²¸Ç>&nbsp;<input type=button onclick=window.close() value=È¡Ïû></form>"&vbcrlf
		response.end
	end if
	response.write "<font size=2>ÄúÒª½« <font color=red>"&filename&"</font> ÎÄ¼þ¼ôÇÐµ½£º</font><br>"&vbcrlf
	response.write "<form name='moveFileName' method='post' action='"&mypage&"?act=task&run=DoFileMove'><input name='now_file' type='hidden' value='"&FileName&"'><input name='inputmove_folder' value='"&nowfolder&"'>&nbsp;<input type=submit value=ÒÆ¶¯></form>"&vbcrlf
	response.end
end if

if run = "DoFileMove" then  ''''''''''''''  ÎÄ¼þÒÆ¶¯Ö´ÐÐÒ³  '''''''''
	now_file = request("now_file")
	tofolder = request("inputmove_folder")
	overhere = request("overhere")
	if overhere="over" then
		set del = fso.GetFile(tofolder&hm.nameall(now_file))
		del.delete
		set del = nothing
	end if
	if not(mm.righthere("folder",tofolder,"")) then
		response.redirect mypage&"?act=task&run=movefile&filename="&hm.javaurl(now_file)&"&folder="&hm.javaurl(tofolder)&"&being=false&beingfolder="&hm.javaurl(tofolder)
	end if
	if mm.righthere("file",tofolder,hm.nameAll(now_file)) then
		response.redirect mypage&"?act=task&run=movefile&filename="&hm.javaurl(now_file)&"&folder="&hm.javaurl(tofolder)&"&being=true&beingfile="&hm.javaurl(tofolder&hm.nameAll(now_file))
	end if
	set thefile = fso.GetFile(now_file)
	thefile.move(tofolder&hm.nameAll(now_file))
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&now_folder&new_file&"</font><br> ÎÄ¼þ³É¹¦ÒÆ¶¯µ½ <font color=red>"&tofolder&hm.nameAll(now_file)&"</font>¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	set thefile=nothing
	response.end
end if

if run="copyfile" then  ''''''''''''  ÎÄ¼þ¸´ÖÆÏÔÊ¾Ò³  '''''''''''''
	FileName = request("filename")
	nowfolder = request("folder")
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "document.copyFileName.inputcopy_folder.select();"&vbcrlf
	response.write "</script>"&vbcrlf
	if request("being")="false" then response.write "<font size=2>ÎÄ¼þ¼Ð <font  color=blue>"&request("beingfolder")&"</font> ²»´æÔÚ¡£</font><br>"&vbcrlf
	if request("being")="true" then
		response.write "<font size=2>ÎÄ¼þ <font  color=red>"&request("beingfile")&"</font> ÒÑ¾­´æÔÚ¡£</font><br>"&vbcrlf
		response.write "<form name='copyFileName' method='post' action='"&mypage&"?act=task&run=DoFileCopy&overhere=over'><input name='now_file' type='hidden' value='"&FileName&"'><input name='inputcopy_folder' value='"&nowfolder&"'>&nbsp;<input type=submit value=¸²¸Ç>&nbsp;<input type=button onclick=window.close() value=È¡Ïû></form>"&vbcrlf
		response.end
	end if
	response.write "<font size=2>ÄúÒª½« <font color=red>"&filename&"</font> ÎÄ¼þ¸´ÖÆµ½£º</font><br>"&vbcrlf
	response.write "<form name='copyFileName' method='post' action='"&mypage&"?act=task&run=DoFileCopy'><input name='now_file' type='hidden' value='"&FileName&"'><input name='inputcopy_folder' value='"&nowfolder&"'>&nbsp;<input type=submit value=¸´ÖÆ></form>"&vbcrlf
	response.end
end if

if run = "DoFileCopy" then  ''''''''''''''  ÎÄ¼þ¸´ÖÆÖ´ÐÐÒ³  '''''''''
	now_file = request("now_file")
	tofolder = request("inputcopy_folder")
	overhere = request("overhere")
	if overhere="over" then
		set del = fso.GetFile(tofolder&hm.nameall(now_file))
		del.delete
		set del = nothing
	end if
	if not(mm.righthere("folder",tofolder,"")) then
		response.redirect mypage&"?act=task&run=copyfile&filename="&hm.javaurl(now_file)&"&folder="&hm.javaurl(tofolder)&"&being=false&beingfolder="&hm.javaurl(tofolder)
	end if
	if mm.righthere("file",tofolder,hm.nameAll(now_file)) then
		response.redirect mypage&"?act=task&run=copyfile&filename="&hm.javaurl(now_file)&"&folder="&hm.javaurl(tofolder)&"&being=true&beingfile="&hm.javaurl(tofolder&hm.nameAll(now_file))
	end if
	set thefile = fso.GetFile(now_file)
	thefile.copy(tofolder&hm.nameAll(now_file))
	response.write "<script language='javascript' for='window' event='onload()'>"&vbcrlf
	response.write "window.opener.parent.location.reload();"&vbcrlf
	response.write "</script>"&vbcrlf
	response.write "<font size=2> <font color=red>"&now_folder&new_file&"</font><br> ÎÄ¼þ³É¹¦¸´ÖÆµ½ <font color=red>"&tofolder&hm.nameAll(now_file)&"</font>¡£</font><br><br>"&vbcrlf
	response.write "<input type=button value='¹Ø±Õ' onclick='javascript:window.close();'>"
	set thefile=nothing
	response.end
end if

%>





<%response.end
end sub
'------------------------------------------------------------------------------------------------------------------------------------
%>



<%
set fso = nothing
set ob_stream = nothing
set jpeg = nothing
rem     ================================================================================================================================================================================================================================================================
rem     =   --==folder¶ÔÏó==--                                                                                                                                                                                                                                         =
rem     =   Copy ·½·¨ |  Delete ·½·¨ |  Move ·½·¨ |  OpenAsTextStream ·½·¨                                                                                                                                                                                             =
rem     =   Attributes ÊôÐÔ | DateCreated ÊôÐÔ | DateLastAccessed ÊôÐÔ | DateLastModified ÊôÐÔ | Drive ÊôÐÔ | Files ÊôÐÔ | IsRootFolder ÊôÐÔ | Name ÊôÐÔ | ParentFolder ÊôÐÔ | Path ÊôÐÔ | ShortName ÊôÐÔ | ShortPath ÊôÐÔ | Size ÊôÐÔ | SubFolders ÊôÐÔ | Type ÊôÐÔ   =
rem     =   --==folder¼¯ºÏ==--                                                                                                                                                                                                                                         =
rem     =   Add ·½·¨(Folders)                                                                                                                                                                                                                                          =
rem     =   Count ÊôÐÔ | Item ÊôÐÔ                                                                                                                                                                                                                                     =
rem     =   --==file¶ÔÏó==--                                                                                                                                                                                                                                           =
rem     =   Copy ·½·¨ |  Delete ·½·¨ |  Move ·½·¨ |  OpenAsTextStream ·½·¨                                                                                                                                                                                             =
rem     =    Attributes ÊôÐÔ | DateCreated ÊôÐÔ | DateLastAccessed ÊôÐÔ | DateLastModified ÊôÐÔ | Drive ÊôÐÔ | Name ÊôÐÔ | ParentFolder ÊôÐÔ | Path ÊôÐÔ | ShortName ÊôÐÔ | ShortPath ÊôÐÔ | Size ÊôÐÔ | Type ÊôÐÔ                                                     =
rem     =   --==file¼¯ºÏ==--                                                                                                                                                                                                                                           =
rem     =   Count ÊôÐÔ | Item ÊôÐÔ                                                                                                                                                                                                                                     =
rem     =   --==TextStream¶ÔÏó==--                                                                                                                                                                                                                                     =
rem     =   Close ·½·¨ |  Read ·½·¨ |  ReadAll ·½·¨ |  ReadLine ·½·¨ |  Skip ·½·¨ |  SkipLine ·½·¨ |  Write ·½·¨ |  WriteBlankLines ·½·¨ |  WriteLine ·½·¨                                                                                                             =
rem     =   AtEndOfLine ÊôÐÔ |  AtEndOfStream ÊôÐÔ |  Column ÊôÐÔ |  Line ÊôÐÔ                                                                                                                                                                                         =
rem     ================================================================================================================================================================================================================================================================
%>
