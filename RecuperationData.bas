Attribute VB_Name = "mod_RecuperationData"
Option Explicit
Option Base 1
Public Sub sub_RecuperationData()
'---------------------------------------------------------------------------------------------------
' Summary:
' Ce sous-programme récupre des données spécifiques d'obligations à partir d'une base de données Access..
' Il sélectionne les enregistrements qui ont toutes les informations nécessaires pour un affichage complet
' sur une interface utilisateur dans Excel.
'----------------------------------------------------------------------------------------------------

' Declaration des variables
    Dim obj_Cnn As ADODB.Connection
    Dim obj_Rst As ADODB.Recordset
    Dim str_SQLRequest As String
    
Set obj_Cnn = New ADODB.Connection

' Paramtre de connexionn
obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\hajjka20\Desktop\Data_Projet.accdb"

obj_Cnn.Open

' Efface le contenu de la feuille Excel avant affichage des nouvelles données
sht_Data.Cells.ClearContents

' Exécution de la requte SQL pour récupérer les informations nécessairess
' La requte sélectionne les noms des obligations qui ont toutes les informations requises disponibless
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT CDX_IG_Infos.Name " & _
                 "FROM CDX_IG_Infos, CDX_IG_Prices " & _
                 "WHERE CDX_IG_Prices.Name = CDX_IG_Infos.Name " & _
                 "AND CDX_IG_Infos.[Ref Bond Obligation] IS NOT NULL " & _
                 "AND CDX_IG_Infos.[S&P] IS NOT NULL " & _
                 "AND CDX_IG_Infos.[Moody's] IS NOT NULL " & _
                 "AND CDX_IG_Infos.Fitch IS NOT NULL " & _
                 "AND CDX_IG_Infos.Debt IS NOT NULL " & _
                 "AND CDX_IG_Infos.cur_mkt_cap IS NOT NULL " & _
                 "AND CDX_IG_Infos.GICS_Sector_Name IS NOT NULL " & _
                 "AND CDX_IG_Infos.ICB_Sector_Name IS NOT NULL " & _
                 "AND CDX_IG_Prices.PX_MID IS NOT NULL"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Data) ' Appel à print_data pour afficher les données dans Excel
obj_Rst.Close

' Affiche des informations prédéfinies sur la fréquence des coupons et le type de taux du coupon
sht_Data.Cells(1, 2).Value = "Coupon Frequency"
sht_Data.Cells(2, 2).Value = "Annual"
sht_Data.Cells(3, 2).Value = "Semi-Annual"
sht_Data.Cells(4, 2).Value = "Quarterly"

sht_Data.Cells(1, 3).Value = "Coupon Rate Type"
sht_Data.Cells(2, 3).Value = "Fixed"
sht_Data.Cells(3, 3).Value = "Variable"

End Sub


