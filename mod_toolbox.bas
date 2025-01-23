{\rtf1\ansi\ansicpg1252\cocoartf2821
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Option Explicit\
Sub sub_PrintData(obj_Rst As Recordset, sht_Data As Worksheet)\
'Module r\'e9cup\'e9r\'e9 du dernier cours de VBA/SQL, permet d'imprimer une table Access sur une feuille excel\
\
'Declaration of the variables\
Dim i As Integer, j As Integer, k As Integer\
Dim Field As Variant\
\
'Get last column filled\
If sht_Data.Cells(1, 1) = "" Then\
    k = 1\
Else\
    k = sht_Data.Cells(1, Columns.Count).End(xlToLeft).Column + 2\
End If\
\
'Prints the header\
j = k\
If (Not obj_Rst.EOF) Then\
    For Each Field In obj_Rst.Fields\
        sht_Data.Cells(1, j).Value = Field.Name\
        j = j + 1\
    Next Field\
End If\
    \
'Loop on the records while obj_Rst.EOF = False\
i = 1\
While (Not obj_Rst.EOF)\
    i = i + 1\
    j = k\
    \
    'Loop on the fields\
    For Each Field In obj_Rst.Fields\
        sht_Data.Cells(i, j).Value = Field.Value\
        j = j + 1\
    Next Field\
    \
    obj_Rst.MoveNext\
Wend\
\
Exit Sub\
End Sub\
\
}