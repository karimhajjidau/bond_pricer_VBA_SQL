VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMod_Bond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'List of attributes

Private str_issuer As String ' Issuer of the bond
Private dbl_Coupon_rate As Double ' Coupon rate of the bond
Private dbl_Margin As Double ' Margin of the bond
Private str_Coupon_Type As String ' Type of coupon, whether Fixed or Variable
Private Var_Frequency As Variant ' Frequency of coupon payments
Private dbl_maturity As Double ' Maturity of the bond
Private obj_RfRate As CMod_Curve ' Risk-free rate curve
Private obj_Spread As CMod_Curve ' Spread curve
Private obj_LiborRate As CMod_Curve ' Libor rate curve for variable coupon

' Properties
Property Get pIssuer() As String
    pIssuer = str_issuer
End Property
Property Let pIssuer(Issuer As String)
    str_issuer = Issuer
End Property

Property Get pCoupon_Rate() As Double
    pCoupon_Rate = dbl_Coupon_rate
End Property
Property Let pCoupon_Rate(Coupon_Rate As Double)
    dbl_Coupon_rate = Coupon_Rate
End Property

Property Get pMargin() As Double
    pMargin = dbl_Margin
End Property
Property Let pMargin(Margin As Double)
    dbl_Margin = Margin
End Property

Property Get pCoupon_Type() As String
    pCoupon_Type = str_Coupon_Type
End Property
Property Let pCoupon_Type(Coupon_Type As String)
    If Coupon_Type = "Fixed" Or Coupon_Type = "Variable" Then
        str_Coupon_Type = Coupon_Type
    End If
End Property

Property Get pFrequency() As Variant
    pFrequency = Var_Frequency
End Property
Property Let pFrequency(Frequency As Variant)
    Select Case Frequency
        Case "Annual"
            Var_Frequency = 1
        Case "Semi-Annual"
            Var_Frequency = 2
        Case "Quarterly"
            Var_Frequency = 4
        Case Else
            Var_Frequency = 1
    End Select
End Property

Property Get pMaturity() As Double
    pMaturity = dbl_maturity
End Property
Property Let pMaturity(maturity As Double)
    dbl_maturity = maturity
End Property

Property Get pRfRate() As CMod_Curve
    Set pRfRate = obj_RfRate
End Property
Property Let pRfRate(Rfrate As CMod_Curve)
    Set obj_RfRate = Rfrate
End Property

Property Get pSpread() As CMod_Curve
    Set pSpread = obj_Spread
End Property
Property Let pSpread(Spread As CMod_Curve)
    Set obj_Spread = Spread
End Property

Property Get pLiborRate() As CMod_Curve
    Set pLiborRate = obj_LiborRate
End Property
Property Let pLiborRate(LiborRate As CMod_Curve)
    Set obj_LiborRate = LiborRate
End Property

' List of methods

' Function to calculate the price of the bond
Public Function fn_price() As Double
    ' Local variables declaration
    Dim dbl_price As Double
    Dim t As Double
    
    ' Initialize t
    t = dbl_maturity - Int(dbl_maturity) + 1 / Var_Frequency
    ' Initialize price
    dbl_price = 0
    
    ' Loop to calculate bond price
    Do
        If str_Coupon_Type = "Fixed" Then
            dbl_price = dbl_price + dbl_Coupon_rate / Var_Frequency / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        Else
            dbl_price = dbl_price + (obj_LiborRate.Interpolate(t) + dbl_Margin) / Var_Frequency / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        End If
        t = 1 / Var_Frequency + t
    Loop While t < dbl_maturity
    
    ' Final cashflow calculation
    If str_Coupon_Type = "Fixed" Then
        dbl_price = dbl_price + ((dbl_Coupon_rate / Var_Frequency) + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
    Else
        dbl_price = dbl_price + (((obj_LiborRate.Interpolate(t) + dbl_Margin) / Var_Frequency) + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
    End If
    fn_price = dbl_price
End Function

' Function to calculate the duration of the bond
Public Function fn_duration() As Double
    ' Local variables declaration
    Dim dbl_duration As Double
    Dim dbl_price As Double
    Dim t As Double
    
    ' Initialize t
    t = dbl_maturity - Int(dbl_maturity) + 1 / Var_Frequency
    ' Calculate bond price
    dbl_price = fn_price()
    
    ' Loop to calculate bond duration
    Do
        If str_Coupon_Type = "Fixed" Then
            dbl_duration = dbl_duration + (t * dbl_Coupon_rate / Var_Frequency) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        Else
            dbl_duration = dbl_duration + ((t * (obj_LiborRate.Interpolate(t) + dbl_Margin)) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ t))
        End If
        t = 1 / Var_Frequency + t
    Loop While t < dbl_maturity
    
    ' Final duration calculation
    If str_Coupon_Type = "Fixed" Then
        dbl_duration = (dbl_Coupon_rate / Var_Frequency + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t)) * t + dbl_duration
    Else
        dbl_duration = ((obj_LiborRate.Interpolate(t) + dbl_Margin) / Var_Frequency + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t)) * t + dbl_duration
    End If
    fn_duration = dbl_duration / dbl_price
End Function

' Sub to generate the bond cash flow schedule
Public Sub schedule()
    ' Local variables declaration
    Dim i As Double, t As Double
    Dim dbl_Cashflow As Double
    Dim ws As Worksheet
    
    ' Suppress display alerts
    Application.DisplayAlerts = False
    ' Check if "Cash Flow Schedule" sheet exists and delete it
    On Error Resume Next ' Ignore error if sheet doesn't exist
    Set ws = ThisWorkbook.Sheets("Cash Flow Schedule")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    ' Add new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    ' Configure new worksheet
    With ws
        .Name = "Cash Flow Schedule"
        .Cells(1, 1).Value = "Maturity"
        .Cells(1, 2).Value = "Coupon Rate"
        .Cells(1, 3).Value = "Flow"
        .Cells(1, 4).Value = "Discounted Cashflow"
        .Cells(1, 5).Value = "Risk Free Rate"
        .Cells(1, 6).Value = "Spread"
        .Cells(1, 7).Value = "Discount factor"
        .Cells(1, 8).Value = "Discounted Cashflow * i"
    End With
    ' Initialize loop variables
    i = 1
    t = dbl_maturity - Int(dbl_maturity) + 1 / Var_Frequency
    
    ' Loop to generate cash flow schedule
    Do
        If str_Coupon_Type = "Fixed" Then
            dbl_Cashflow = dbl_Coupon_rate / Var_Frequency / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
            ws.Cells(i + 1, 2) = dbl_Coupon_rate / Var_Frequency
            ws.Cells(i + 1, 3) = dbl_Coupon_rate / Var_Frequency
        Else
            dbl_Cashflow = (obj_LiborRate.Interpolate(t) + dbl_Margin) / Var_Frequency / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
            ws.Cells(i + 1, 2) = (obj_LiborRate.Interpolate(t))
            ws.Cells(i + 1, 3) = (obj_LiborRate.Interpolate(t) + dbl_Margin)
        End If
        ' Populate worksheet with cash flow schedule
        With ws
            .Cells(i + 1, 1).Value = t
            .Cells(i + 1, 4).Value = dbl_Cashflow
            .Cells(i + 1, 5).Value = obj_RfRate.Interpolate(t)
            .Cells(i + 1, 6).Value = obj_Spread.Interpolate(t)
            .Cells(i + 1, 7).Value = 1 / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
            .Cells(i + 1, 8).Value = dbl_Cashflow * t
            .Cells(i + 1, 2).Resize(1, 7).NumberFormat = "0.00%"
        End With
        ' Increment loop variables
        i = i + 1
        t = 1 / Var_Frequency + t
    Loop While t < dbl_maturity
    
    ' Final cash flow entry
    If str_Coupon_Type = "Fixed" Then
        dbl_Cashflow = (dbl_Coupon_rate / Var_Frequency + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        ws.Cells(i + 1, 2) = dbl_Coupon_rate / Var_Frequency
        ws.Cells(i + 1, 3) = dbl_Coupon_rate / Var_Frequency + 1
    Else
        dbl_Cashflow = (obj_LiborRate.Interpolate(t) + dbl_Margin + 1) / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        ws.Cells(i + 1, 2) = (obj_LiborRate.Interpolate(t))
        ws.Cells(i + 1, 3) = (obj_LiborRate.Interpolate(t) + dbl_Margin) + 1
    End If
    ' Populate final cash flow entry
    With ws
        .Cells(i + 1, 1).Value = t
        .Cells(i + 1, 4).Value = dbl_Cashflow
        .Cells(i + 1, 5).Value = obj_RfRate.Interpolate(t)
        .Cells(i + 1, 6).Value = obj_Spread.Interpolate(t)
        .Cells(i + 1, 7).Value = 1 / ((1 + obj_RfRate.Interpolate(t) + obj_Spread.Interpolate(t)) ^ (t))
        .Cells(i + 1, 8).Value = dbl_Cashflow * t
        .Cells(i + 1, 2).Resize(1, 7).NumberFormat = "0.00%"
    End With
End Sub


