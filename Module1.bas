Attribute VB_Name = "Module1"
Sub finder()
    Stokgiris.CB_BOLINO.RowSource = "Raw_Data!C2:C" & Sheets("Raw_Data").Range("C100000").End(3).Row
    Stokgiris.CB_SORUMLU.RowSource = "Raw_Data!A2:A" & Sheets("Raw_Data").Range("A100000").End(3).Row
    Stokgiris.CB_GIRISYAPAN.RowSource = "Raw_Data!B2:B" & Sheets("Raw_Data").Range("B100000").End(3).Row
End Sub
