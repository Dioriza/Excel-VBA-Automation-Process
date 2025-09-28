Sub ClearCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub




