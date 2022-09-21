Option Explicit
Private WithEvents objinspectors As Outlook.Inspectors

Public Sub Anlagen_speichern()
    Dim myOlApp As Outlook.Application
    Dim myNameSpace As Outlook.NameSpace
    Set myOlApp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Dim olExplorer As Explorer
    Dim olFolder As MAPIFolder
    Dim olSelection As Selection
    Dim olitem As MailItem
    Dim lngAttCount As Long, i As Long
    Dim strAttNames As String
    Dim strSubDir As String
    Dim intAnswer As Integer
    Dim datei, nummer, dateiname, Dateiendung, dateinameneu, Prüfname, geändert
    
      Const strBackupPath As String = "C:\Anlagen"
    
    Set olExplorer = Application.ActiveExplorer
    Set olFolder = Application.ActiveExplorer.CurrentFolder
    Dim Anzahl
    
    If olFolder.DefaultItemType = olMailItem Then
    Set olSelection = olExplorer.Selection
    
    For Each olitem In olSelection
    
      lngAttCount = olitem.Attachments.Count
      If lngAttCount > 0 Then
      strAttNames = ""
      For i = lngAttCount To 1 Step -1
    
       With olitem.Attachments.Item(i)
    
         If Dir(strBackupPath, vbDirectory) = "" Then
         MkDir strBackupPath
         End If
    
        'strSubDir = strBackupPath & Format(olitem.CreationTime, "yyyymmdd")
        strSubDir = Left(strBackupPath, Len(strBackupPath) - 1)
         If Dir(strSubDir, vbDirectory) = "" Then
         MkDir strSubDir
         End If
    
        'Datei schon gespeichert
        datei = strSubDir & "\" & .FileName
        geändert = False
        nummer = 1
        Prüfname = datei
zurück:
    
        nummer = nummer + 1
    
        If CreateObject("Scripting.FileSystemObject").FileExists(Prüfname) = True Then
        dateiname = CreateObject("Scripting.FileSystemObject").GetBaseName(datei)
        Dateiendung = CreateObject("Scripting.FileSystemObject").GetExtensionName(datei)
        dateinameneu = dateiname & " (" & nummer & ")" & "." & Dateiendung
     
        Prüfname = strSubDir & "\" & dateinameneu
        geändert = True
        GoTo zurück:
        Else
        If geändert = False Then dateinameneu = .FileName
        End If
    
        .SaveAsFile strSubDir & "\" & dateinameneu
    
         If olitem.GetInspector.EditorType = olEditorHTML Then
          strAttNames = strAttNames & "file://" & strSubDir & "\" & dateinameneu & "<br>"
          strAttNames = URLDecode(strAttNames)
         Else
          strAttNames = strAttNames & "file://" & strSubDir & "\" & dateinameneu & vbNewLine
          strAttNames = URLDecode(strAttNames)
         End If
        .Delete
       End With
       Anzahl = Anzahl + 1
      Next i
    
    With olitem
    Select Case .GetInspector.EditorType
    
     Case olEditorText
      .Body = vbNewLine & lngAttCount & " Anlage(n) gespeichert unter:" & vbNewLine & strAttNames & String(2, vbNewLine) & .Body
    
     Case olEditorHTML
      .HTMLBody = lngAttCount & " Anlage(n) gespeichert unter:" & "<BR>" & strAttNames & .HTMLBody
    
     Case olEditorRTF, olEditorWord
      intAnswer = MsgBox("Wenn Sie einen Hinweis auf gespeicherte Anlagen einfügen, " & "gehen die Formatierungen der Nachricht verloren!" & vbCr & "Wollen Sie den Hinweis trotzdem einfügen?", vbYesNo)
       If intAnswer = vbYes Then
        .Body = vbNewLine & lngAttCount & " Anlage(n) gespeichert unter:" & vbNewLine & strAttNames & String(2, vbNewLine) & .Body
       End If
    End Select
    
     .Save
    End With
    End If
    
    Next olitem
    Else
    MsgBox "In diesem Ordner befinden sich keine E-Mail-Nachrichten."
    End If
    If Anzahl < 1 Then
    MsgBox "Keine Anlagen vorhanden"
    Else
    If Anzahl < 2 Then
    MsgBox Anzahl & " Anlage gespeichert"
    Else
    MsgBox Anzahl & " Anlagen gespeichert"
    End If
    End If

End Sub
