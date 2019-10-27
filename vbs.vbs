Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("1.txt", 1)
Set fw = fso.CreateTextFile("2.txt", true)
fw.Close
Set fw = fso.OpenTextFile("2.txt", 8, false)
Do While f.AtEndOfStream <> True
    s = f.ReadLine
    If Len(s) > 10 Then
        fw.WriteLine("select * from md_property where id = " & Mid(s, 647, 38) & ";")
    End If
Loop
f.Close
Set f = Nothing
Set fso = Nothing
