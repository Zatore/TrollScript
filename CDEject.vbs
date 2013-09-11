Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROM's = oWMP.cdromCollection
wscript.sleep  600000
do
if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
NExt
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep 120000
loop
