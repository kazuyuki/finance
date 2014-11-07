dim URL, TITLE, INTERVAL
'URL = "http://stocks.finance.yahoo.co.jp/stocks/detail/?code=9501.T"
'TITLE = "東京電力(株)【9501】：株式/株価 - Yahoo!ファイナンス"
URL = "http://stocks.finance.yahoo.co.jp/stocks/detail/?code=6701.T"
TITLE = "ＮＥＣ【6701】：株式/株価 - Yahoo!ファイナンス"
INTERVAL = 1000

'===============================================================================
' 監視 & 表示
'===============================================================================
dim ie,  cur, pre
AttachIE ie

cur = 0
pre = 0

While True
	IfMarketOpen
	waitIE ie

	' On Error Resume Next
	cur = ie.Document.getElementsByTagName("TD")(1).outerText
	if cur = "Internet Explorer ではこのページは表示できません" then
		WScript.StdErr.WriteLine Now & vbTab & "[Internet Explorer ではこのページは表示できません] occurred"
		ie.Refresh
		waitIE ie
	end if
	if pre <> cur then
		WScript.StdOut.Write Now & vbTab & cur & vbCrLf
		WScript.StdErr.Write Now & vbTab & cur & vbCrLf
	else 
		WScript.StdErr.Write Now & vbTab & cur & vbCr
	end if
	pre = cur
	WScript.Sleep INTERVAL
Wend
WScript.Quit

'===============================================================================
' IE Window があれば attach なければ open
'===============================================================================
Sub AttachIE (objIE)
	dim flag, sa
	flag = 0
	set sa = CreateObject("Shell.Application")
	for each objIE In sa.Windows
		if TypeName(objIE.Document) = "HTMLDocument" then
			if objIE.Document.Title = TITLE then
				flag = 1
				exit for
			end if
		end if
	next
	if flag = 0 then
		set objIE = CreateObject("InternetExplorer.Application")
		objIE.Visible = True
		objIE.Navigate2 URL
		set objIE = sa.Windows.Item(sa.Windows.Count - 1)
		waitIE objIE
	end if
End Sub

'===============================================================================
' Market Close なら寝る
'===============================================================================
Sub IfMarketOpen()
	dim cur, flag
	flag = 0
	do
		cur = TimeValue(Now)
		if ( cur > #09:00:00# and cur < #11:30:00# ) or ( cur > #12:30:00# and cur < #15:00:00# ) then
			if flag = 1 Then
				flag = 0
				WScript.StdErr.WriteLine Now & vbTab & "Wake up!!"
			end if
			exit do
		else
			if flag = 0 then
				flag = 1
			end if
			WScript.StdErr.Write Now & vbTab & "Sleeping!!" & vbCr
			WScript.Sleep INTERVAL
		end if
	loop
End Sub

'===============================================================================
' IE がビジー状態の間待つ
'===============================================================================
Sub waitIE(objIE)
	While objIE.Busy
		WScript.Sleep 100
	Wend
End Sub


'===============================================================================
' TODO
'===============================================================================
'以下のような出力が来たら、ブラウザをリロードする
'2014/11/06 9:16:18      Internet Explorer ではこのページは表示できません
