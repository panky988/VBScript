Option Explicit  
 
Dim objshell,path,DigitalID, Result  
Set objshell = CreateObject("WScript.Shell") 
'Állítsa be a rendszerleíró kulcs útvonalát
Path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" 
'Kulcs értéke
DigitalID = objshell.RegRead(Path & "DigitalProductId") 
Dim ProductName,ProductID,ProductKey,ProductData 
'Név, Azonosító, Kulcs 
ProductName = "Nev: " & objshell.RegRead(Path & "ProductName") 
ProductID = "Azonosito: " & objshell.RegRead(Path & "ProductID") 
ProductKey = "Kulcs: " & ConvertToKey(DigitalID)  
ProductData = ProductName  & vbNewLine & ProductID  & vbNewLine & ProductKey 
'Üzenet fájlba mentéshez 
If vbYes = MsgBox(ProductData  & vblf & vblf & "Menti fajlba?", vbYesNo + vbQuestion, "BackUp Windows Key Information") then 
   Save ProductData  
End If 
 
 
 
'Konvertálja a bináris karaktereket karakterekké
Function ConvertToKey(Key) 
    Const KeyOffset = 52 
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert 
    'Check if OS is Windows 8 
    isWin8 = (Key(66) \ 6) And 1 
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4) 
    i = 24 
    Maps = "BCDFGHJKMPQRTVWXY2346789" 
    Do 
           Current= 0 
        j = 14 
        Do 
           Current = Current* 256 
           Current = Key(j + KeyOffset) + Current 
           Key(j + KeyOffset) = (Current \ 24) 
           Current=Current Mod 24 
            j = j -1 
        Loop While j >= 0 
        i = i -1 
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput 
        Last = Current 
    Loop While i >= 0  
     
    If (isWin8 = 1) Then 
        keypart1 = Mid(KeyOutput, 2, Last) 
        insert = "N" 
        KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0) 
        If Last = 0 Then KeyOutput = insert & KeyOutput 
    End If     
     
 
    ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5) 
    
     
End Function 
'Adatok mentése fájlba
Function Save(Data) 
    Dim fso, fName, txt,objshell,UserName 
    Set objshell = CreateObject("wscript.shell") 
    'Az aktuális felhasználói név beolvasása
    UserName = objshell.ExpandEnvironmentStrings("%UserName%")  
    'Hozzon létre egy szöveges fájlt az asztalon
    fName = "C:\Users\" & UserName & "\Desktop\WindowsKulcsInfo.txt" 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Set txt = fso.CreateTextFile(fName) 
    txt.Writeline Data 
    txt.Close 
End Function
