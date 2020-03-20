from java.awt import Toolkit
from java.awt.datatransfer import Clipboard
toolkit = Toolkit.getDefaultToolkit()
clipboard = toolkit.getSystemClipboard()
from java.awt.datatransfer import DataFlavor
from java.awt.datatransfer import StringSelection

def getclipboard():
    contents = clipboard.getContents(None)
    text = contents.getTransferData(DataFlavor.stringFlavor)
    return text

def setclipboard(text):
    clipboard.setContents(StringSelection(text), None)


if __name__ == "__main__":
    #Excel
    if exists("excel window.png",0.5):
        click("excel window.png")
    else:
        click("excel icon.png")
        wait("excel window.png", 30)
        click("excel window.png")
    type(Key.ESC) 
    type("c",Key.CTRL)
    WorkNumber=getclipboard()
    type(Key.DOWN) 
    type("c",Key.CTRL)
    Customer=getclipboard()
    type(Key.DOWN) 
    type("c",Key.CTRL)
    SalesHR=getclipboard()
    type(Key.DOWN) 
    type("c",Key.CTRL)
    CarType=getclipboard()
    type(Key.DOWN) 
    type("c",Key.CTRL)
    ProjectType=getclipboard()
    type(Key.DOWN) 
    type("c",Key.CTRL)
    InspectionMonth=getclipboard()

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
    
    click("FindButton.png")
    #コピーを開始
    wait("ProjectFindResult.png", 30)
    rightClick(Pattern("ProjectName.png").targetOffset(-5,53))
    click("rightClickCopy.png")
    click(Pattern("rightClickCopy2.png").targetOffset(0,-36))
    wait("1583483224021.png", 30)
    click(Pattern("1583483224021.png").targetOffset(101,-24)) 
    click(Pattern("1583483316633.png").similar(0.61))
    wait("1583483369189.png", 30)         
    type(Key.TAB*3)
    setclipboard(WorkNumber)
    type("v",Key.CTRL)
    setclipboard(WorkNumber)
    setclipboard(Customer)
    setclipboard(SalesHR)
    setclipboard(CarType)
    setclipboard(ProjectType)
    setclipboard(InspectionMonth)