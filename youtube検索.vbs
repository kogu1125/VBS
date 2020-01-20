Dim ie
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True 
call ie.Navigate("https://www.youtube.com/?gl=JP&hl=ja")

'ページが読み込まれるまで待機
Do While ie.Busy = True Or ie.readyState <> 4
    WScript.Sleep 100        
Loop

Dim doc
Set doc = ie.Document
Dim txt
Set txt = doc.getElementsByName("search_query")
txt.item(0).value = "もこう"

Dim btn
Set btn = doc.getElementById("search-btn")
btn.click