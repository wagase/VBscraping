<%@ CODEPAGE=65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
<% '初期値
	Dim n1,n2,n3,n4,n5,n6,k
	k=0
	n1 = "01"
	n2 = "24"
	n3 = "25"
	n4 = "28"
	n5 = "29"
	n6 = "30"
	if request("1") <>"" Then n1 = request("1") 
	if request("2") <>"" Then n2 = request("2") 
	if request("3") <>"" Then n3 = request("3") 
	if request("4") <>"" Then n4 = request("4") 
	if request("5") <>"" Then n5 = request("5") 
	if request("6") <>"" Then n6 = request("6") 
%>
購入番号</br>
<form action="Default.asp" method="get">
<input type="text" name="1" value="<%=n1%>" size=2>
<input type="text" name="2" value="<%=n2%>" size=2>
<input type="text" name="3" value="<%=n3%>" size=2>
<input type="text" name="4" value="<%=n4%>" size=2>
<input type="text" name="5" value="<%=n5%>" size=2>
<input type="text" name="6" value="<%=n6%>" size=2>
<input type="submit" value="調べる" size=2>
</form>
<%
' *************************** ***************************
' みずほ銀行のHPにアクセスしてロト６の当選番号を取得するサイトスクレイピング
' *************************** ***************************

	Session.CodePage = 65001
	response.charset = "UTF-8"

	' 買った番号の配列
	num = array(n1,n2,n3,n4,n5,n6)
	' 正規表現　/<td class="alnCenter extension"><strong>/ 本数字のhtml要素
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = "<td class=""alnCenter extension""><strong>"
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索

	' 正規表現　/<th colspan="6" class="alnCenter bgf7f7f7">/ 当選回
	Set regEx2 = CreateObject("VBScript.RegExp")
	regEx2.Pattern = "<th colspan=""6"" class=""alnCenter bgf7f7f7"">"
	regEx2.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx2.Global = True ' 文字列全体を検索

	Dim tempDic
	Set tempDic = Server.CreateObject("Scripting.Dictionary")
	Dim array1
	Dim xmlhttp
	Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET","https://www.mizuhobank.co.jp/takarakuji/loto/loto6/index.html", False
	xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	On Error Resume Next
	xmlhttp.Send ""
	m_Status = xmlhttp.status
	m_resText = xmlhttp.responseText
	Set xmlhttp = Nothing
	array1 = split(m_resText,vblf)
	i = 0
	response.write "<table border=1><tr>"
	for each key in array1
		' 6回で区切る
		If i Mod 6 = 0 Then
			Set matches2 = regEx2.Execute(key)
			If matches2.Count <> 0 Then
				key  = replace(key,"<th colspan=""6"" class=""alnCenter bgf7f7f7"">","")
				key  = replace(key,"</th>","")
				Response.write "<td>"&key&"</td>"
			End If
		End If
		Set matches = regEx.Execute(key)
		If matches.Count <> 0 Then
' Debug			MakeLog(key)
			i = i +1
			key  = replace(key,"<td class=""alnCenter extension""><strong>","")
			key  = replace(key,"</strong></td>","")
			Call tempDic.Add(key,"")
			m=0
			For Each n In num
				If tempDic.Exists(n) Then
					k= k+1
					key = "★" & key
				Else
					m=m+1
					If m = 6 Then
						key = "　"&key
					End If
				End If
			Next
			Response.write "<td>"&key&"</td>"
			Set tempDic = Nothing
			Set tempDic = Server.CreateObject("Scripting.Dictionary")
			' 6回で区切る
			If i Mod 6 = 0 Then
				response.write	"<td>"&k&"個一致！</td></tr>" 
				k = 0
			End If
		End If
	Next
	response.write "</table>"
	Set regEx = Nothing
	
	If Err.Number <> 0 Then
		Response.write Err.Number & Err.Description
		Err.Clear()
	End If
	On Error GoTo 0


	' //=================================================================
	' // 関数: MakeLog(str)
	' // 概要: ログ出力用。月単位で出力 debug用
	' //=================================================================
	Private Sub MakeLog(str)
		Dim wYear : wYear = CStr(Year(Date))
		Dim wMonth : wMonth = CStr(Month(Date))
		Dim objLog
		Dim mlog
		Set objLog = Nothing
		Set mlog = Nothing
		If wMonth < 10 Then
			wMonth = "0" & wMonth
		End If
		outputLogFile = wYear & wMonth & "data.log"
		Set objLog = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
			Set mlog = objLog.openTextFile("D:log\" & outputLogFile,8,True)
			If Err.Number = 0 Then
				mlog.WriteLine str
				mlog.Close()
			End If
		On Error Goto 0
		Set objLog = Nothing
		Set mlog = Nothing
	End Sub
%>
</body>
</html>