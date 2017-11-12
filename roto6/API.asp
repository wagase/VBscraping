<%
session.CodePage=65001
Dim KEY
KEY = Request.Form("MYKEY")
If  KEY = "xxxx" Then
		Response.Status = "403 Forbidden"
		response.end
End If

data = Request.Form("DATA") '例　"{""01"",""04"",""05"",""28"",""29"",""30""}"

' 2桁の数値
Set regExsubCd = CreateObject("VBScript.RegExp")
regExsubCd.Pattern = "\d{2}"
regExsubCd.IgnoreCase = False
regExsubCd.Global = True
Set matchSubCd = regExsubCd.Execute(data)
num = array(matchSubCd.Item(0).Value,matchSubCd.Item(1).Value,matchSubCd.Item(2).Value,matchSubCd.Item(3).Value,matchSubCd.Item(4).Value,matchSubCd.Item(5).Value)
Set regExsubCd = Nothing

' *************************** ***************************
' みずほ銀行のHPにアクセスしてロト６の当選番号を取得するサイトスクレイピング
' *************************** ***************************
	response.charset = "UTF-8"
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
	Dim tempDic2,str2
	Set tempDic2 = Server.CreateObject("Scripting.Dictionary")
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
	response.write "{"
	for each key in array1
		' 6回で区切る
		If i Mod 6 = 0 Then
			Set matches2 = regEx2.Execute(key)
			If matches2.Count <> 0 Then
				key  = replace(key,"<th colspan=""6"" class=""alnCenter bgf7f7f7"">","")
				key  = replace(key,"</th>","")
				str2 = str2 & """" &key&""":"""
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
			str2 = str2 &key
			Set tempDic = Nothing
			Set tempDic = Server.CreateObject("Scripting.Dictionary")
			' 6回で区切る
			If i Mod 6 = 0 Then
				str2 = str2 & "　"&k&"個一致!"""
				Call tempDic2.Add(tempDic2.count,str2)
				k = 0
				str2 = ""
			End If
		End If
	Next
	response.write Join(tempDic2.Items(),",")
	response.write "}"
	Set regEx = Nothing
	
	If Err.Number <> 0 Then
		Response.write Err.Number & Err.Description
		Err.Clear()
	End If
	On Error GoTo 0


%>