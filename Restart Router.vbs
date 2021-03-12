'Tested OS: Windows 10 Pro -V: 1909 -B: 18363.1440
'Tested router: ZYXEL VMG3312-B10B
'@gokanil

'Just change the following: "admin" and "ttnet"
'----------------------------------------------
userName =             "admin"
password =             "ttnet"
'----------------------------------------------

'If a different router is used, the CGI addresses can be changed
'---------------------------------------------------------------
loginURL  =   "http://192.168.1.1/login/login-page.cgi"
logoutURL =   "http://192.168.1.1/login/login-logout.cgi"
rebootURL =   "http://192.168.1.1/pages/tabFW/reboot-rebootpost.cgi"
'---------------------------------------------------------------

cookie = ""

dim req
logOut()
loginPost()
cookie = Split(req.getResponseHeader("Set-Cookie"), ";")(0)
setLastCookie()
If NOT Split(cookie, "=")(1) = "" Then
reboot()
wScript.Echo "The restart request to your modem has been SUCCESSFULLY sent."
If NOT req.Status = 200 Then
logOut()
End If
End If

function setLastCookie()
set reg = New RegExp
reg.Pattern = "^cookie\s*=.*"
reg.MULTILINE  = True
reg.IGNORECASE = True

set fileSystem = CreateObject("Scripting.FileSystemObject")
text = fileSystem.OpenTextFile(WScript.ScriptFullName, 1,false).ReadAll
fileSystem.OpenTextFile(WScript.ScriptFullName, 2,false).Write reg.Replace(text, "cookie = """&cookie&"""")
end function

function reboot()
call getRequest(rebootURL)
end function

function logOut()
call getRequest(logoutURL)
end function

function loginPost()
set req = CreateObject("WinHttp.WinHttpRequest.5.1")
req.open "POST", loginURL,false
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.send "AuthName="&userName&"&AuthPassword="&password
end function

function getRequest(url)
Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
req.open "GET", url, False
req.setRequestHeader "Cookie", cookie
req.send
end function