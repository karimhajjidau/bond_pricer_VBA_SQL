Attribute VB_Name = "mod_Insertion_Access"
Option Explicit
Option Base 1
'Le Sub permet la création d'une nouvelle table Access pour les pricings

Public Sub sub_Creation_Table()
  '-------------------------------------------------------------
    ' DECLARATION DES VARIABLES
    '-------------------------------------------------------------
    
    ' Initialisation et déclaration des objets pour la connexion et le recordset
    Dim obj_Cnn As ADODB.Connection
    Dim obj_Rst As ADODB.Recordset
    
    ' Strings pour l'inserton et création SQL
    Dim str_SQLCreate As String

' Création et ouverture de la connexion à la base de données
Set obj_Cnn = New ADODB.Connection
obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\hajjka20\Desktop\Data_Projet.accdb"
obj_Cnn.Open
    
str_SQLCreate = "CREATE TABLE Resultat_Pricing (" & _
        "ID AUTOINCREMENT, " & _
        "Company_Name TEXT, " & _
        "Date_Pricing DATE, " & _
        "Coupon_rate_type TEXT, " & _
        "Coupon_rate_or_margin DOUBLE, " & _
        "Coupon_frequency TEXT, " & _
        "Maturity DOUBLE, " & _
        "Price DOUBLE, " & _
        "Duration DOUBLE)"
        
obj_Cnn.Execute str_SQLCreate

End Sub

  
Public Sub sub_Insertion_Table()
'le sub permet d'insérer un pricing dans une table
'chaque pricing a un ID selon le moment où il a été ajouté à la base access

  '-------------------------------------------------------------
    ' DECLARATION DES VARIABLES
    '-------------------------------------------------------------
    
    ' Initialisation et déclaration des objets pour la connexion et le recordset
    Dim obj_Cnn As ADODB.Connection
    Dim obj_Rst As ADODB.Recordset
    
    ' Strings pour l'inserton et création SQL
    Dim str_SQLInsert As String
    Dim var_remplissage As String
    Dim str_coupon_rate_or_margin As String
    Dim str_maturity As String
    Dim str_price As String
    Dim str_duration As String

' Création et ouverture de la connexion à la base de données
Set obj_Cnn = New ADODB.Connection
obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\hajjka20\Desktop\Data_Projet.accdb"
obj_Cnn.Open


' Conversion des nombres en chaÃ®nes avec des points comme séparateurs décimaux
str_coupon_rate_or_margin = Replace(CStr(sht_Interface.[rng_coupon_rate_or_margin].Value), ",", ".")
str_maturity = Replace(CStr(sht_Interface.[rng_maturity].Value), ",", ".")
str_price = Replace(CStr(sht_Interface.[rng_price].Value), ",", ".")
str_duration = Replace(CStr(sht_Interface.[rng_duration].Value), ",", ".")

' Construction de la requÃªte SQL avec les valeurs converties
str_SQLInsert = "INSERT INTO Resultat_Pricing " & _
                "([Company_Name], [Date_Pricing], [Coupon_rate_type], [Coupon_rate_or_margin], [Coupon_frequency], [Maturity], [Price], [Duration]) " & _
                "VALUES ('" & sht_Interface.[rng_company].Value & "', #" & Format(Date, "mm/dd/yyyy") & "#, '" & _
                sht_Interface.[rng_coupon_rate_type].Value & "', " & str_coupon_rate_or_margin & ", '" & _
                sht_Interface.[rng_coupon_frequency].Value & "', " & str_maturity & ", " & _
                str_price & ", " & str_duration & ")"
        
obj_Cnn.Execute str_SQLInsert

End Sub

        

        

        

