'On error resume next 'comment it for debug

set FSO=CreateObject ("Scripting.FileSystemObject")
erowallfile = fso.GetSpecialFolder(2): if right(erowallfile,1)<>"\" then erowallfile=erowallfile & "\" : erowallfile = erowallfile & "erowall.jpg"

searchstring=""

'=== part 1 ====
sUrlRequest = "https://erowall.com"
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
httpfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing

beg=instr(lcase(httpfile),"/w/")
ef=instr(beg+3,lcase(httpfile),"/")
lnk=mid(httpfile,beg+3,ef-beg-3)

randomize
r = int(rnd*CLng(lnk)) + 1

'=== part 2 ====
url="https://erowall.com/wallpapers/original/" + cstr(r) + ".jpg"
Set oXMLHTTP2 = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP2.Open "GET", url, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile erowallfile, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
'objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", erowallfile, "REG_SZ"
'objWshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", "6", "REG_SZ"
'objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False

'use irfanview if you want
objWshShell.Run "c:\Programs\IrfanView\i_view64.exe """ & erowallfile & """ /wall=3 /killmesoftly", 1, False 


Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing