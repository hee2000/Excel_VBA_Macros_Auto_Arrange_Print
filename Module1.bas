Attribute VB_Name = "Module1"
'Comment #1: This block of code will update the values of each cells with the
'       relevant row number, after the current row is sent to the printer.
Sub iteration()

    Dim NumRows, i&

    NumRows = Application.InputBox("Enter # of rows to print (100 max): ", Type:=1)

    If NumRows = 0 Or NumRows > 100 Then Exit Sub
    
    'Comment #2: For loop
    For i = 2 To NumRows

        Range("C33").Value = i

        ActiveSheet.PrintOut From:=1, To:=1

    Next i

End Sub

'Comment #3: This block of code will automatically print the page without
'                  clicking the print button after each row is inserted
Sub iteration_print()

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, _
    Copies:=1, Collate:=True, IgnorePrintAreas:=False

End Sub


