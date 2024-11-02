Attribute VB_Name = "word_import"
Option Explicit

Sub ImportaTutteLeTabelleWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim Table As Object
    Dim ws As Worksheet
    Dim i As Long, j As Long, tblIndex As Long
    Dim filePath As Variant
    Dim lastRow As Long
    Dim cellText As String
    Dim isNumber As Boolean
    Dim wdCell As Object
    
    ' Mostra una finestra di dialogo per selezionare il file Word
    filePath = Application.GetOpenFilename("File Word (*.docx; *.rtf), *.docx; *.rtf", , "Seleziona il file Word")
    
    ' Se l'utente annulla la selezione o non seleziona un file valido
    If filePath = False Then
        MsgBox "Nessun file selezionato."
        Exit Sub
    End If

    ' Crea un'istanza di Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' Apri il documento Word selezionato dall'utente
    Set wdDoc = wdApp.Documents.Open(filePath, ReadOnly:=True)

    ' Seleziona il primo foglio di lavoro in Excel
    Set ws = ThisWorkbook.Sheets(1)
    
    'Imposto la riga in cui cominciare a scrivere
    lastRow = 1
    
    ' Cicla attraverso tutte le tabelle nel documento Word
    For tblIndex = 1 To wdDoc.Tables.Count
        Set Table = wdDoc.Tables(tblIndex)
        
        ' Trasferisci i dati della tabella corrente di Word in Excel
        For i = 1 To Table.Rows.Count
            For j = 1 To Table.Columns.Count
            
                ' Ottieni la cella di Word
                Set wdCell = Table.cell(i, j).Range
            
                ' Ottieni il testo della cella e rimuovi eventuali caratteri speciali
                cellText = Table.cell(i, j).Range.Text
                
                ' Rimuovi eventuali ritorni a capo o caratteri non visibili
                cellText = Replace(cellText, Chr(13), "")
                cellText = Replace(cellText, Chr(7), "")
                
                ' Controlla se è un numero e rimuovi i punti (separatori delle migliaia)
                isNumber = IsNumeric(Replace(cellText, ".", ""))
                
                If isNumber Then
                    ' Rimuovi i punti dai numeri e converti in numero
                    cellText = Replace(cellText, ".", "")
                End If
                
                ' Se la cella è vuota dopo la pulizia, lasciala vuota in Excel
                If Trim(cellText) = "" Then
                    cellText = ""
                End If
                
                ActiveSheet.Cells(lastRow + i - 1, j).Value = cellText
                
                ' Applica la formattazione in Excel
                With ActiveSheet.Cells(lastRow + i - 1, j)
                    ' Bold
                    If wdCell.Bold = True Then
                        .Font.Bold = True
                    Else
                        .Font.Bold = False
                    End If
                    
                    ' Italic
                    If wdCell.Italic = True Then
                        .Font.Italic = True
                    Else
                        .Font.Italic = False
                    End If
                    
                    ' Colore del testo
                    .Font.Color = wdCell.Font.Color
                End With
                
                ' Debugging: Stampa il testo della cella nella finestra "Immediata" di VBA
                'Debug.Print "Cella (" & i & "," & j & "): " & Table.cell(i, j).Range.Text
                
            Next j
        Next i
        
        ' Aggiungi una riga vuota tra le tabelle (opzionale)
        lastRow = lastRow + Table.Rows.Count + 1
    Next tblIndex
    
    ' Auto-fit delle prime 4 colonne
    ActiveSheet.Columns("A:D").AutoFit
    
    ' Formatta le colonne B, C e D come valuta
    ActiveSheet.Columns("B:D").NumberFormat = "_-* #,##0_-;_-* (#,##0)_-;_-* ""-""_-;_-@_-"


    ActiveSheet.Columns("B:D").Font.Color = RGB(0, 0, 0)

    ' Chiudi Word
    wdDoc.Close False
    wdApp.Quit

    ' Notifica il completamento
    MsgBox "Tutte le tabelle sono state importate correttamente in Excel!"
End Sub
