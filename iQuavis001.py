import sys
#Excel内のジャンプメソッド
def tojump(data):
	type("g",Key.CTRL)
	wait("1584673703104.png",30)
	wait(0.5)
	paste(data)
	type(Key.TAB*2,Key.SHIFT)
	type(" ")

#iQ内でのタスクフィルター
def filterTask(TaskName):
	#タスクフィルター
	click("1585015900958-1.png")
	wait("1585015955657-1.png",30)
	#フィルター名
	click("1585018332463-1.png")
	#クリア
	type(Key.TAB*2+Key.SPACE)
	#追加
	type(Key.TAB*1+Key.SPACE)
	#タスク名
	wait("1585016551679-1.png",30)
	doubleClick("1585016551679-1.png")
	paste(TaskName)
	#条件(等しいにする)
	type(Key.TAB)
	doubleClick("1585018653584-1.png")
#	type(Key.DOWN*2)
	#フィルター実行
	type(Key.TAB*16,Key.SHIFT)
	type(Key.SPACE)

	
if __name__ == "__main__":
	salesMembers=[u"GuptaPiuesh",u"阿部真人",u"山越隆弘",u"村瀬浩司",u"羽祢田慎一",u"矢野博之",u"浜口満照",u"鄭香丹",u"緑創"]
	for num in range(100):
		#Excelのウィンドウ取得
		if exists("1584949552344.png",0.5):
			click("1584949552344.png")
		else:
			click("1584949166169.png")
			wait("1584949552344.png", 30)
			click("1584949552344.png")
		wait("window pan.png",30)
		type("x")
		wait(0.5)
	    
	# 工番立ち上げ用データの取得
	#	jumpto="INDIRECT(\"R4C\"&TEXT(COLUMN(),\"@\"),FALSE)"
	#	tojump(jumpto)
	#	wait("1584945292672.png",30)
		jumpto="INDIRECT(\"R3C\"&TEXT(COLUMN(),\"@\")&\":R20C\"&TEXT(COLUMN(),\"@\"),FALSE)"
		tojump(jumpto)
		wait("1584945292672.png",30)
		type("c",Key.CTRL)
		wait("1584945183240.png",30)
		fromExcel=App.getClipboard()
		cells=fromExcel.split("\n")
		WorkNumber=cells[0]
		#工番の終わりを検知して終了
		if ""==WorkNumber:
			popup("finish")
			sys.exit(0)
		Customer=cells[1]
		ProductType=cells[2]
		SalesHR=cells[3]
		CarType=cells[4]
		ProjectType=cells[5]
		InspectionMonth=cells[6]
		Tachiai=cells[7]
		Hikitori=cells[8]
		SalesMemo=cells[12]
		Shatachi=cells[17]
		#営業メンバーをフルネームに変更
		foundMember=False
		for member in salesMembers:
			SalesHR=SalesHR.replace("V","")
			SalesHR=SalesHR.replace("*","")
			SalesHR=SalesHR.replace("V","")
			SalesHR=SalesHR.replace("Ⅴ","")
			if member[0]==SalesHR[0] and member[1]==SalesHR[1] :
				SalesHR=member
				foundMember=True
		#未知の名前ならストップ
		if foundMember==False:
			popup("Unknown Sales HR!")
			sys.exit(0)
		#ガント貼り付け用データの取得
		jumpto="INDIRECT(\"R54C\"&TEXT(COLUMN(),\"@\")&\":R101C\"&TEXT(COLUMN(),\"@\"),FALSE)"
		tojump(jumpto)
		wait("1584945292672.png",30)
		type("c",Key.CTRL)
		wait("1584945183240.png",30)
		forGantData=App.getClipboard()
		#次回のためにカーソル移動
		jumpto="INDIRECT(\"R1C\"&TEXT(COLUMN()+1,\"@\"),FALSE)"
		tojump(jumpto)
		wait(0.5)
		
		#iQuavis新規登録用マクロ
		#iQuavisを見つける
		if exists("HomeWindow.png",0.5):
			click("HomeWindow.png")
		else:
			click("1584602881074.png")
			wait("HomeWindow.png", 30)
			click("HomeWindow.png")
		#検索を開始
		type("hsn", Key.ALT)
		wait("1583454054559.png", 30)
		#画面の最大化
		click("1584948558985.png")
		wait("window pan.png",30)
		type("x")
		#検索ワードの入力
		type(Key.TAB*6)
		paste(u"＃＃"+ProjectType[0]+ProjectType[1]+u"）コピー用マスター全行程 V3")
		#曖昧検索の指定
		type(Key.TAB*1+Key.SPACE)
		#テンプレートから検索する指定
		type(Key.TAB*9+Key.SPACE)
		#検索開始
		type(Key.TAB*7+Key.SPACE)
		#コピーを開始
		wait("ProjectFindResult.png", 30)
		rightClick(Pattern("ProjectName.png").targetOffset(-5,53))
		type(Key.DOWN*3+Key.RIGHT+Key.ENTER)
		wait("1585015292635.png", 30)
		#プロジェクト属性の変更
		type(Key.TAB+Key.SPACE)
		wait("1583483369189.png", 30)
		type(Key.TAB*3)
		paste(WorkNumber)
		type(Key.TAB*2) 
		paste(ProductType)
		type(Key.TAB*2) 
		paste(ProjectType[0]+ProjectType[1])
		type(Key.TAB*2) 
		paste(Customer)
		type(Key.TAB*13) 
		paste(SalesHR)
		type(Key.TAB*1) 
		paste(CarType)
		type(Key.TAB*1) 
		paste(InspectionMonth)
		type(Key.TAB*1) 
		paste(SalesMemo)
		#OKボタン
		type(Key.TAB*1+Key.SPACE)
		wait("1584922777615.png", 30)
		wait("1584952154887.png", 30)
		#詳細ボタン
		type(Key.TAB*3,Key.SHIFT)
		type(Key.SPACE)
		wait("1584932433267.png",30)
		type(Key.TAB*4+Key.SPACE)
		type(Key.TAB*1+Key.SPACE)
		#業務タブ
		type(Key.TAB*1,Key.CTRL)
		wait("1584932765847.png",30)
		#アサインを取り込む
		type(Key.TAB*1+Key.SPACE)
		#OKボタン
		type(Key.TAB*1,Key.CTRL)
		type(Key.TAB*4,Key.SHIFT)
		type(" ")
		wait("1584922777615.png", 30)
		#OKボタン
		type(Key.TAB*1+Key.SPACE)
		wait("1584934125726.png",30)
		#OKボタン
		type(Key.TAB*1+Key.SPACE)
		#工番立ち上げ後
		wait("1584944366429.png",30)
		#画面の最大化
		click("1584948138336.png")
		wait("window pan.png",30)
		type("x")
		#編集の開始
		type("p",Key.ALT)
		type("ep")
		wait("1584944395907.png",30)
		wait(0.5)
		#ガントの起動
		type("t",Key.ALT)
		type("g")
		wait("1585036453996.png",30)
		#ガントデータの貼り付け
		type(Key.RIGHT)
		wait(0.5)
		paste(forGantData)
		if Tachiai != "-":
			#フィルター
			filterTask(u"客先立会い")
			#すべて選択
			type("a",Key.CTRL)
			#右クリック
			wait(1)
			type(Key.F10,Key.SHIFT)
			wait("1585018064784.png",30)
			#削除
			type(Key.DOWN*4+Key.ENTER)
			wait("1585030992825.png",30)
			#OK
			type(Key.SPACE)
		if Hikitori != "-":
			pass
		if Shatachi != "-":
			filterTask(u"社内立会い")
			#すべて選択
			type("a",Key.CTRL)
			#右クリック
			wait(1)
			type(Key.F10,Key.SHIFT)
			wait("1585018064784.png",30)
			#削除
			type(Key.DOWN*4+Key.ENTER)
			wait("1585030992825.png",30)
			#OK
			type(Key.SPACE)
			
		#プロジェクト編集の終了
		click("1584948649568.png")
		wait("1584946481215.png",30)
		#履歴の保存
		type(Key.TAB*3+Key.SPACE)
		wait("1584946522888.png",30)
		#履歴内容
		paste(u"初回計画")
		type(Key.TAB*3+Key.SPACE)
		#OKボタン
		type(Key.TAB*2+Key.SPACE)
		#検索画面を閉じる
		wait("1584949331718.png",30)
		wait("1584949351358.png",30)
		wait("1584949363807.png",30)
		click("1584949377298.png")
		type("c")
