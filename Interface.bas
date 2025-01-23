Attribute VB_Name = "mod_Interface"
Option Explicit
Option Base 1

' Ce Sub est destiné au pricing à partir de l'interface remplie par l'utilisateur.
' Il initialise une connexion à une base de données pour récupérer des informations
' sur les taux sans risque, les spreads d'émetteurs et les taux LIBOR.
'ces informations sont imprimées sur des pages excel à l'aide de print_data (mod_Toolbox)
' Ces données sont ensuite utilisées pour configurer des objets spécifiques représentant
' les courbes de taux (objet curve) et un objet bond. Enfin, le code calcule et affiche le planning
' des futurs cash-flows ainsi que le prix et la duration du bond sélectionné par l'utilisateur.
'Ces calculs sont faits à partir de méthodes et fonctions définies dans le modulede classe Bond


Public Sub sub_Interface()

    '-------------------------------------------------------------
    ' DECLARATION DES VARIABLES
    '-------------------------------------------------------------
    
    ' Initialisation et déclaration des objets pour la connexion et le recordset
    Dim obj_Cnn As ADODB.Connection
    Dim obj_Rst As ADODB.Recordset
    
    ' Cha”ne pour la requte SQLL
    Dim str_SQLRequest As String
    
    ' Déclaration des objets personnalisés pour représenter les différentes courbes de taux
    Dim issuer_Bond As New cMod_Bond
    Dim rf_Curve As New CMod_Curve
    Dim Spread_Curve As New CMod_Curve
    Dim Libor_Curve As New CMod_Curve
    
    ' Variables pour stocker les données spécifiques à l'émission
    Dim str_issuer As String
    Dim i As Double
    Dim dbl_rf_rates() As Double
    Dim dbl_rf_maturities() As Double
    Dim dbl_Issuer_Spread() As Double
    Dim dbl_issuer_maturities() As Double
    Dim dbl_Libor_rates() As Double
    Dim dbl_Libor_maturities() As Double
    Dim dbl_date As Double
    Dim dbl_maturity As Double
    Dim lastrow As Double
    
'-------------------------------------------------------------
' INITIALISATION DE LA CONNEXION
'-------------------------------------------------------------
    
' Création et ouverture de la connexion à la base de données
Set obj_Cnn = New ADODB.Connection
obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\hajjka20\Desktop\Data_Projet.accdb"
obj_Cnn.Open
    
'-------------------------------------------------------------
' RƒCUPƒRATION DES DONNƒES DE TAUX SANS RISQUE
'-------------------------------------------------------------
' Effacement des données précédentes sur la feuille de taux
sht_Rates.Cells.ClearContents

' Exécution de la requte SQL pour récupérer les taux sans risque et impression des donnéess
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT * FROM [US Yield Curve]"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Rates)
obj_Rst.Close
    
    ' Détermination de la dernire ligne de données et récupération des taux et des maturitéss
lastrow = sht_Rates.Cells(Rows.Count, 1).End(xlUp).Row
ReDim dbl_rf_rates(0 To lastrow - 2)
ReDim dbl_rf_maturities(0 To lastrow - 2)
For i = 0 To lastrow - 2
    dbl_rf_rates(i) = sht_Rates.Cells(i + 2, 2).Value
    dbl_rf_maturities(i) = sht_Rates.Cells(i + 2, 1).Value
Next i
    
    ' Attribuer les taux sans risque à l'objet rf_Curve
With rf_Curve
    .pName = "Risk Free"
    .pType = "Yield"
    .pMaturities = dbl_rf_maturities
    .pRates = dbl_rf_rates
End With

    '-------------------------------------------------------------
    ' RƒCUPƒRATION DU SPREAD DE L'ENTREPRISE
    '-------------------------------------------------------------
    
' Effacer le contenu précédent de la feuille des spreads
sht_Spread.Cells.ClearContents

' Récupération de la maturité et du nom de l'émetteur à partir de l'interface utilisateur
dbl_maturity = sht_Interface.[rng_maturity].Value
str_issuer = sht_Interface.[rng_company].Value

' Initialisation des maturités pour lesquelles les spreads seront récupérés
ReDim dbl_issuer_maturities(0 To 7)
dbl_issuer_maturities(0) = 0.5
dbl_issuer_maturities(1) = 1
dbl_issuer_maturities(2) = 2
dbl_issuer_maturities(3) = 3
dbl_issuer_maturities(4) = 4
dbl_issuer_maturities(5) = 5
dbl_issuer_maturities(6) = 7
dbl_issuer_maturities(7) = 10

' Récupération des spreads de l'émetteur sélectionné à partir de la base de données
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT CDX_IG_Prices.[6M], CDX_IG_Prices.[1Y], CDX_IG_Prices.[2Y], CDX_IG_Prices.[3Y], CDX_IG_Prices.[4Y], CDX_IG_Prices.[5Y], CDX_IG_Prices.[7Y], CDX_IG_Prices.[10Y] FROM CDX_IG_Prices WHERE CDX_IG_Prices.Name = '" & str_issuer & "'"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Spread)
obj_Rst.Close

' Récupération et conversion des spreads récupérés pour utilisation dans le modlee
ReDim dbl_Issuer_Spread(0 To UBound(dbl_issuer_maturities))
For i = 0 To UBound(dbl_issuer_maturities)
    dbl_Issuer_Spread(i) = sht_Spread.Cells(2, i + 1).Value / 10000 ' Conversion en pourcentage
Next i

' Attribution des spreads à l'objet Curve de l'émetteur
With Spread_Curve
    .pName = str_issuer
    .pType = "Spread"
    .pMaturities = dbl_issuer_maturities
    .pRates = dbl_Issuer_Spread
End With

    '-------------------------------------------------------------
    ' RƒCUPƒRATION DU YIELD (ICI LIBOR)
    '-------------------------------------------------------------
    
' Effacer le contenu précédent de la feuille des taux LIBOR
sht_Libor.Cells.ClearContents

' Récupération des données de taux LIBOR à partir de la base de données
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT * FROM [Libor 3M Curve]"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Libor)
obj_Rst.Close

' Détermination de la dernire ligne de données des taux LIBOR et récupération des tauxx
lastrow = sht_Libor.Cells(Rows.Count, 1).End(xlUp).Row
ReDim dbl_Libor_rates(0 To lastrow - 2)
ReDim dbl_Libor_maturities(0 To lastrow - 2)
For i = 0 To lastrow - 2
    dbl_Libor_rates(i) = sht_Libor.Cells(i + 2, 2).Value
    dbl_Libor_maturities(i) = sht_Libor.Cells(i + 2, 1).Value
Next i

' Attribution des données LIBOR à l'objet Curve correspondant
With Libor_Curve
    .pName = "Libor 3M Curve"
    .pType = "Yield"
    .pMaturities = dbl_Libor_maturities
    .pRates = dbl_Libor_rates
End With

' Remplissage des attributs du bond de l'émetteur à partir de l'interface utilisateur
With issuer_Bond
    .pIssuer = str_issuer
    .pCoupon_Type = sht_Interface.[rng_coupon_rate_type]
    .pFrequency = sht_Interface.[rng_coupon_frequency]
    .pMaturity = sht_Interface.[rng_maturity]
    .pRfRate = rf_Curve
    .pSpread = Spread_Curve
    .pLiborRate = Libor_Curve
End With

' Attribution du taux de coupon ou de la marge, en fonction du type de coupon sélectionné
If issuer_Bond.pCoupon_Type = "Fixed" Then
    issuer_Bond.pCoupon_Rate = sht_Interface.[rng_coupon_rate_or_margin] ' Taux fixe
Else
    issuer_Bond.pMargin = sht_Interface.[rng_coupon_rate_or_margin] ' Marge pour les coupons variables
End If

' Appel de la méthode schedule pour calculer l'échéancier des flux financiers
issuer_Bond.schedule

' Affichage du prix et de la durée du bon sur l'interface utilisateur
sht_Interface.[rng_price].Value = issuer_Bond.fn_price ' Afficher le prix calculé
sht_Interface.[rng_duration].Value = issuer_Bond.fn_duration ' Afficher la durée calculée

End Sub





