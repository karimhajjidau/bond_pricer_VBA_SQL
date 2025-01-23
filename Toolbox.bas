Attribute VB_Name = "mod_Toolbox"
Option Explicit
Sub sub_PrintData(obj_Rst As Recordset, sht_Data As Worksheet)
'Module récupéré du dernier cours de VBA/SQL, permet d'imprimer une table Access sur une feuille excel

'Declaration of the variables
Dim i As Integer, j As Integer, k As Integer
Dim Field As Variant

'Get last column filled
If sht_Data.Cells(1, 1) = "" Then
    k = 1
Else
    k = sht_Data.Cells(1, Columns.Count).End(xlToLeft).Column + 2
End If

'Prints the header
j = k
If (Not obj_Rst.EOF) Then
    For Each Field In obj_Rst.Fields
        sht_Data.Cells(1, j).Value = Field.Name
        j = j + 1
    Next Field
End If
    
'Loop on the records while obj_Rst.EOF = False
i = 1
While (Not obj_Rst.EOF)
    i = i + 1
    j = k
    
    'Loop on the fields
    For Each Field In obj_Rst.Fields
        sht_Data.Cells(i, j).Value = Field.Value
        j = j + 1
    Next Field
    
    obj_Rst.MoveNext
Wend

Exit Sub
End Sub

