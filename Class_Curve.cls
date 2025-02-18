VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMod_Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe CMod_Curve : Mod�lise une courbe de taux d'int�rt avec des fonctionnalit�ss
' pour stocker et acc�der � des maturit�s et des taux, ainsi qu'interpoler des taux pour
' des maturit�s non sp�cifiquement list�es dans les donn�es d'entr�e.




Option Explicit

'List of attributes
Private str_Name As String ' Nom de la courbe, par exemple "US Treasury Yield"
Private str_Type_of_bond As String ' Type de la courbe, peut tre "Yield" ou "Spread""
Private dbl_Maturities() As Double ' Tableau stockant les maturit�s de la courbe
Private dbl_rates() As Double ' Tableau stockant les taux correspondants aux maturit�s



'Properties
'Nom de l'actif
Property Get pName() As String
    pName = str_Name
End Property

Property Let pName(Name As String)
    str_Name = Name
End Property

'Type de l'actif
Property Get pType_of_bond() As String
    pType_of_bond = str_Type_of_bond
End Property

Property Let pType(Type_of_bond As String)
    str_Type_of_bond = Type_of_bond
End Property

'Maturit�s
Property Get pMaturities() As Double()
    pMaturities = dbl_Maturities
End Property

Property Let pMaturities(Maturities() As Double)
    dbl_Maturities = Maturities
End Property
'taux ou spread
Property Get pRates() As Double()
    pRates = dbl_rates
End Property

Property Let pRates(rates() As Double)
    dbl_rates = rates
End Property


' Fonction d'interpolation : Calcule un taux pour une maturit� donn�e en interpolant
' les taux disponibles dans la courbe. Cette fonction est utile pour estimer les taux
' pour des maturit�s qui ne sont pas explicitement pr�sentes dans la courbe de taux.
Public Function Interpolate(dbl_maturity As Double) As Double
    Dim i As Double
    Dim dbl_rate As Double
    
    'Gestion des cas o la maturit� demand�e est en dehors des bornes disponibless
    If dbl_maturity < dbl_Maturities(LBound(dbl_Maturities)) Then
        Interpolate = dbl_rates(LBound(dbl_rates))
    End If
    
    If dbl_maturity > dbl_Maturities(UBound(dbl_Maturities)) Then
        Interpolate = dbl_rates(UBound(dbl_rates))
    End If
    
    ' Boucle d'interpolation pour les maturit�s comprises dans les bornes de la courbe
    i = 0
    Do While dbl_maturity > dbl_Maturities(i)
        dbl_rate = dbl_rates(i) + (dbl_rates(i + 1) - dbl_rates(i)) * (dbl_maturity - dbl_Maturities(i)) / (dbl_Maturities(i + 1) - dbl_Maturities(i))
        i = i + 1
    Loop
    Interpolate = dbl_rate
    
End Function











