# scraping
## Windows IIS Classic ASP
クラシックASPでInstagramのウエブスクレイピングを行う  
好きなユーザーの最新画像をマスクされていない扱いやすい形式で表示する  
IDを書き換えると動作  

  
# VBscriptでInstagramのWEBスクレイピングした話
Instagramの画像を簡単に保存する方法  
## 目的
Instagramの画像を扱いやすい形式にしてlocalhostで参照する

## 経緯
Instagramの画像を保存したい  
↓  
Instagramの画像はdivタグでマスクされていて右クリックで簡単に保存したりできないようになっている  
↓  
画像を保存したければF12キーを押して開発者ツールを立ち上げるか外部サービスにURLを貼り付けるのが主流  
↓  
これは面倒くさい  
  
WEBで公開されているものだしウェブスクレイピング（Web scraping）します  
  
## 環境
言語：VBScript  
サーバー：Windows IIS ASP  

## ソース
全体はgithub参照ください  
個別解説  
### WEBサイトへアクセス  

	Private Function getXMLHTTP(byval url)

		Dim xmlhttp,resText
		Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
		xmlhttp.Open "GET",url, False
		xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlhttp.Send ""
		resText = xmlhttp.responseText
		Set xmlhttp = Nothing

		getXMLHTTP = resText
	End Function`
VBSでWEBサイトへアクセスするには  
Server.Createobject("MSXML2.ServerXMLHTTP")  
を使います  
xmlhttp.responseTextで結果がこの場合HTMLが返ってくるのでそれを解析します  

### HTMLの解析
実際にだれかのInstagramにアクセスしてソースをみるとscriptタグでjavascriptの記述があることがわかります。これをテキスト形式で取得することになるので正規表現で解析します。  
`"https://scontent-nrt1-1\.cdninstagram\.com.*?\.jpg"`  
取得したものをHTMLで<img src=～～と書き直してやれば画像を取るだけなら終了です  

### 動画もほしい
動画をとるには工夫が必要でInstagramの記事固有ページに行く必要がありました  

記事固有ページへは上記のjavascriptでcodeと書かれている１１桁の英数字の文字列を  
`https://www.instagram.com/p/(１１桁の英数字)`  
のようにすると行けるようです  
したがってまずは個別ページのcodeを取得します  
正規表現で  
`"""code"":.*?"", ""date"""`  
とかきMid関数で１０文字目から１１文字抜くと良さそうです  

	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = """code"":.*?"", ""date"""
	regEx.IgnoreCase = False ' 大文字と小文字を区別しない
	regEx.Global = True ' 文字列全体を検索

あとは  
`Set matches = regEx.Execute()`  
を使って順番に個別ページにアクセスを行いそれぞれ結果を取得して  
画像なら  
`<img src="(画像URL)">`  
動画なら  
`<video autoplay loop muted controls><source src="(動画URL)" type=""video/mp4"" /></video>`  
と書いてresponse.Writeしてやるとマスクのかかっていない画像の一覧が取得できます  

