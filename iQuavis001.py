from java.awt import Toolkit
from java.awt.datatransfer import Clipboard
from java.awt.datatransfer import DataFlavor
from java.awt.datatransfer import StringSelection
kit = Toolkit.getDefaultToolkit()
clip = kit.getSystemClipboard()

def getclipboard():
    contents = clipboard.getContents(None)
    text = contents.getTransferData(DataFlavor.stringFlavor)
    return text

def setclipboard(text):
    clipboard.setContents(StringSelection(text), None)

if __name__ == "__main__":
	waitTime=30
	keyWaitTime=0.5
	popup((String)waitTime)    
#Excel
    if exists("excel window.png",wait_time):
        click("excel window.png")
    else:
        click("excel icon.png")
        wait("excel window.png", wait_time)
        click("excel window.png")
    wait("window pan.png",wait_time)
    type(Key.ALT) 

#    wait(0.5)
#    type("g",Key.CTRL)
#    wait("1584673703104.png",wait_time)
#    wait(0.5)
#    type("r1c3")
#    type(Key.TAB,Key.SHIFT)
#    type(Key.TAB,Key.SHIFT)
#    type(" ")
#    type(Key.DOWN*2)
#    type(Key.RIGHT,Key.CTRL)
    wait(1)
    type("c",Key.CTRL)
    WorkNumber=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    Customer=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    ProductType=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    SalesHR=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    CarType=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    ProjectType=App.getClipboard()
    wait(0.5)
    type(Key.DOWN) 
    type("c",Key.CTRL)
    InspectionMonth=App.getClipboard()
    wait(0.5)
    type(Key.DOWN*6) 
    type("c",Key.CTRL)
    SalesMemo=App.getClipboard()
    #次回のためにカーソル移動
    type("g",Key.CTRL)
    wait("1584673703104.png",30)
    wait(0.5)
    test003="=OFFSET(INDIRECT(\"R\"&TEXT(ROW(),\"@\")&\"C\"&TEXT(COLUMN(),\"@\"),FALSE),-12,1)"
    paste(test003)
    type(Key.TAB,Key.SHIFT)
    type(Key.TAB,Key.SHIFT)
    type(" ")
 
    
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
    click("FindingWindow.png")
    type(Key.TAB*6)
    paste(u"＃＃新織）コピー用マスター全行程 V3")
    type(Key.TAB)
    type(" ")
    type(Key.TAB*9)
    type(" ")
    
    click(Pattern("1584668756127.png").targetOffset(-40,-2))
    #コピーを開始
    wait("ProjectFindResult.png", 30)
    rightClick(Pattern("ProjectName.png").targetOffset(-5,53))
    type(Key.DOWN*3)
    type(Key.RIGHT+Key.ENTER)
    wait("1583483224021.png", 30)
    type(Key.TAB+" ")
    wait("1583483369189.png", 30)         
    type(Key.TAB*3)
    paste(WorkNumber)
    type(Key.TAB*2) 
    paste(ProductType)
    type(Key.TAB*2) 
    paste(ProjectType)
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
    type(Key.TAB*1)
    type(" ")
    
    wait("1584922777615.png", 30)
    click("1584922816552.png")
    wait("1584932433267.png",30)