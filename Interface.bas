Attribute VB_Name = "mod_Interface"
Option Explicit
Option Base 1

' Ce Sub est destin� au pricing � partir de l'interface remplie par l'utilisateur.
' Il initialise une connexion � une base de donn�es pour r�cup�rer des informations
' sur les taux sans risque, les spreads d'�metteurs et les taux LIBOR.
'ces informations sont imprim�es sur des pages excel � l'aide de print_data (mod_Toolbox)
' Ces donn�es sont ensuite utilis�es pour configurer des objets sp�cifiques repr�sentant
' les courbes de taux (objet curve) et un objet bond. Enfin, le code calcule et affiche le planning
' des futurs cash-flows ainsi que le prix et la duration du bond s�lectionn� par l'utilisateur.
'Ces calculs sont faits � partir de m�thodes et fonctions d�finies dans le modulede classe Bond


Public Sub sub_Interface()

    '-------------------------------------------------------------
    ' DECLARATION DES VARIABLES
    '-------------------------------------------------------------
    
    ' Initialisation et d�claration des objets pour la connexion et le recordset
    Dim obj_Cnn As ADODB.Connection
    Dim obj_Rst As ADODB.Recordset
    
    ' Cha�ne pour la requte SQLL
    Dim str_SQLRequest As String
    
    ' D�claration des objets personnalis�s pour repr�senter les diff�rentes courbes de taux
    Dim issuer_Bond As New cMod_Bond
    Dim rf_Curve As New CMod_Curve
    Dim Spread_Curve As New CMod_Curve
    Dim Libor_Curve As New CMod_Curve
    
    ' Variables pour stocker les donn�es sp�cifiques � l'�mission
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
    
' Cr�ation et ouverture de la connexion � la base de donn�es
Set obj_Cnn = New ADODB.Connection
obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\hajjka20\Desktop\Data_Projet.accdb"
obj_Cnn.Open
    
'-------------------------------------------------------------
' R�CUP�RATION DES DONN�ES DE TAUX SANS RISQUE
'-------------------------------------------------------------
' Effacement des donn�es pr�c�dentes sur la feuille de taux
sht_Rates.Cells.ClearContents

' Ex�cution de la requte SQL pour r�cup�rer les taux sans risque et impression des donn�ess
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT * FROM [US Yield Curve]"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Rates)
obj_Rst.Close
    
    ' D�termination de la dernire ligne de donn�es et r�cup�ration des taux et des maturit�ss
lastrow = sht_Rates.Cells(Rows.Count, 1).End(xlUp).Row
ReDim dbl_rf_rates(0 To lastrow - 2)
ReDim dbl_rf_maturities(0 To lastrow - 2)
For i = 0 To lastrow - 2
    dbl_rf_rates(i) = sht_Rates.Cells(i + 2, 2).Value
    dbl_rf_maturities(i) = sht_Rates.Cells(i + 2, 1).Value
Next i
    
    ' Attribuer les taux sans risque � l'objet rf_Curve
With rf_Curve
    .pName = "Risk Free"
    .pType = "Yield"
    .pMaturities = dbl_rf_maturities
    .pRates = dbl_rf_rates
End With

    '-------------------------------------------------------------
    ' R�CUP�RATION DU SPREAD DE L'ENTREPRISE
    '-------------------------------------------------------------
    
' Effacer le contenu pr�c�dent de la feuille des spreads
sht_Spread.Cells.ClearContents

' R�cup�ration de la maturit� et du nom de l'�metteur � partir de l'interface utilisateur
dbl_maturity = sht_Interface.[rng_maturity].Value
str_issuer = sht_Interface.[rng_company].Value

' Initialisation des maturit�s pour lesquelles les spreads seront r�cup�r�s
ReDim dbl_issuer_maturities(0 To 7)
dbl_issuer_maturities(0) = 0.5
dbl_issuer_maturities(1) = 1
dbl_issuer_maturities(2) = 2
dbl_issuer_maturities(3) = 3
dbl_issuer_maturities(4) = 4
dbl_issuer_maturities(5) = 5
dbl_issuer_maturities(6) = 7
dbl_issuer_maturities(7) = 10

' R�cup�ration des spreads de l'�metteur s�lectionn� � partir de la base de donn�es
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT CDX_IG_Prices.[6M], CDX_IG_Prices.[1Y], CDX_IG_Prices.[2Y], CDX_IG_Prices.[3Y], CDX_IG_Prices.[4Y], CDX_IG_Prices.[5Y], CDX_IG_Prices.[7Y], CDX_IG_Prices.[10Y] FROM CDX_IG_Prices WHERE CDX_IG_Prices.Name = '" & str_issuer & "'"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Spread)
obj_Rst.Close

' R�cup�ration et conversion des spreads r�cup�r�s pour utilisation dans le modlee
ReDim dbl_Issuer_Spread(0 To UBound(dbl_issuer_maturities))
For i = 0 To UBound(dbl_issuer_maturities)
    dbl_Issuer_Spread(i) = sht_Spread.Cells(2, i + 1).Value / 10000 ' Conversion en pourcentage
Next i

' Attribution des spreads � l'objet Curve de l'�metteur
With Spread_Curve
    .pName = str_issuer
    .pType = "Spread"
    .pMaturities = dbl_issuer_maturities
    .pRates = dbl_Issuer_Spread
End With

    '-------------------------------------------------------------
    ' R�CUP�RATION DU YIELD (ICI LIBOR)
    '-------------------------------------------------------------
    
' Effacer le contenu pr�c�dent de la feuille des taux LIBOR
sht_Libor.Cells.ClearContents

' R�cup�ration des donn�es de taux LIBOR � partir de la base de donn�es
Set obj_Rst = New ADODB.Recordset
str_SQLRequest = "SELECT * FROM [Libor 3M Curve]"
Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
Call sub_PrintData(obj_Rst, sht_Libor)
obj_Rst.Close

' D�termination de la dernire ligne de donn�es des taux LIBOR et r�cup�ration des tauxx
lastrow = sht_Libor.Cells(Rows.Count, 1).End(xlUp).Row
ReDim dbl_Libor_rates(0 To lastrow - 2)
ReDim dbl_Libor_maturities(0 To lastrow - 2)
For i = 0 To lastrow - 2
    dbl_Libor_rates(i) = sht_Libor.Cells(i + 2, 2).Value
    dbl_Libor_maturities(i) = sht_Libor.Cells(i + 2, 1).Value
Next i

' Attribution des donn�es LIBOR � l'objet Curve correspondant
With Libor_Curve
    .pName = "Libor 3M Curve"
    .pType = "Yield"
    .pMaturities = dbl_Libor_maturities
    .pRates = dbl_Libor_rates
End With

' Remplissage des attributs du bond de l'�metteur � partir de l'interface utilisateur
With issuer_Bond
    .pIssuer = str_issuer
    .pCoupon_Type = sht_Interface.[rng_coupon_rate_type]
    .pFrequency = sht_Interface.[rng_coupon_frequency]
    .pMaturity = sht_Interface.[rng_maturity]
    .pRfRate = rf_Curve
    .pSpread = Spread_Curve
    .pLiborRate = Libor_Curve
End With

' Attribution du taux de coupon ou de la marge, en fonction du type de coupon s�lectionn�
If issuer_Bond.pCoupon_Type = "Fixed" Then
    issuer_Bond.pCoupon_Rate = sht_Interface.[rng_coupon_rate_or_margin] ' Taux fixe
Else
    issuer_Bond.pMargin = sht_Interface.[rng_coupon_rate_or_margin] ' Marge pour les coupons variables
End If

' Appel de la m�thode schedule pour calculer l'�ch�ancier des flux financiers
issuer_Bond.schedule

' Affichage du prix et de la dur�e du bon sur l'interface utilisateur
sht_Interface.[rng_price].Value = issuer_Bond.fn_price ' Afficher le prix calcul�
sht_Interface.[rng_duration].Value = issuer_Bond.fn_duration ' Afficher la dur�e calcul�e

End Sub





