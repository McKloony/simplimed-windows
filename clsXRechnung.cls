VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXRechnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Benutzerdefinierter Typ für Rechnungspositionen
Private Type InvoiceLineItem
    id As Integer
    Name As String
    Quantity As Currency
    UnitCode As String
    LineExtensionAmount As Currency
    PriceAmount As Currency
    TaxPercent As Currency
End Type

' Invoice properties
Public InvoiceNumber As String
Public InvoiceDate As String
Public InvoiceTax As Currency
Public InvoiceTaxCategory As String
Public InvoicePayment As String
Public InvoiceTypeCode As String
Public DueDate As String
Public CurrencyCode As String
Public BuyerReference As String
Public PaymentMandate As String

' Seller information
Public SellerEndpointID As String
Public SellerStreet As String
Public SellerCity As String
Public SellerPostalZone As String
Public SellerCountryCode As String
Public SellerRegName As String
Public SellerCompanyID As String
Public SellerContactName As String
Public SellerTelephone As String
Public SellerEmail As String
Public SellerIBAN As String
Public SellerBIC As String
Public SellerUID As String
Public SellerAssignedId As String

' Buyer information
Public BuyerEndpointID As String
Public BuyerStreet As String
Public BuyerCity As String
Public BuyerPostalZone As String
Public BuyerCountrySubentity As String
Public BuyerCountryCode As String
Public BuyerRegistrationName As String
Public BuyerCompanyID As String
Public BuyerContactName As String
Public BuyerAccountNumberID As String
Public BuyerAccountShortID As String

' Invoice totals
Public TaxAmount As Currency
Public TaxableAmount As Currency
Public LineExtensionAmount As Currency
Public TaxExclusiveAmount As Currency
Public TaxInclusiveAmount As Currency
Public AllowanceTotalAmount As Currency
Public ChargeTotalAmount As Currency
Public PrepaidAmount As Currency
Public PayableRoundingAmount As Currency
Public PayableAmount As Currency

' Array für die Rechnungspositionen
Private InvoiceLines() As InvoiceLineItem
Private LineCount As Integer
Private Sub Class_Initialize()
    LineCount = 0
End Sub
Public Function AddInvoiceLine(ByVal ItmID As Integer, ByVal Name As String, ByVal Quantity As Currency, ByVal UnitCode As String, ByVal PriceAmount As Currency, ByVal ExtensionAmount As Currency, ByVal TaxPercent As Currency, ByVal NewInvoice As Boolean) As Integer
'Methode zum Hinzufügen einer Rechnungsposition

If NewInvoice = True Then
    LineCount = 0
End If

ReDim Preserve InvoiceLines(LineCount)
InvoiceLines(LineCount).id = ItmID
InvoiceLines(LineCount).Name = Name
InvoiceLines(LineCount).Quantity = Quantity
InvoiceLines(LineCount).UnitCode = UnitCode
InvoiceLines(LineCount).LineExtensionAmount = Format(ExtensionAmount, "0.00")
InvoiceLines(LineCount).PriceAmount = Format(PriceAmount, "0.00")
InvoiceLines(LineCount).TaxPercent = TaxPercent
LineCount = LineCount + 1

AddInvoiceLine = LineCount
    
End Function
Public Sub GenerateXMLFile(FiNam As String)
'Methode zum Erstellen der XML-Datei

Dim AktZa As Integer
Dim xmlDoc As Object
Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

xmlDoc.async = False
xmlDoc.preserveWhiteSpace = True

' Hauptknoten mit Namespaces erstellen
Dim RootNode As Object
Set RootNode = xmlDoc.createElement("ubl:Invoice")
RootNode.setAttribute "xmlns:ubl", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
RootNode.setAttribute "xmlns:cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
RootNode.setAttribute "xmlns:cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
xmlDoc.appendChild RootNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' Hinzufügen von Standard-Identifikationsknoten
Dim CustomizationNode As Object
Set CustomizationNode = CreateNode(xmlDoc, "cbc:CustomizationID", "urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0")
RootNode.appendChild CustomizationNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim ProfileNode As Object
Set ProfileNode = CreateNode(xmlDoc, "cbc:ProfileID", "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0")
RootNode.appendChild ProfileNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim IdNode As Object
Set IdNode = CreateNode(xmlDoc, "cbc:ID", InvoiceNumber)
RootNode.appendChild IdNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim IssueDateNode As Object
Set IssueDateNode = CreateNode(xmlDoc, "cbc:IssueDate", InvoiceDate)
RootNode.appendChild IssueDateNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim DueDateNode As Object
Set DueDateNode = CreateNode(xmlDoc, "cbc:DueDate", DueDate)
RootNode.appendChild DueDateNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim InvoiceTypeNode As Object
Set InvoiceTypeNode = CreateNode(xmlDoc, "cbc:InvoiceTypeCode", InvoiceTypeCode)
RootNode.appendChild InvoiceTypeNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim currencyCodeNode As Object
Set currencyCodeNode = CreateNode(xmlDoc, "cbc:DocumentCurrencyCode", CurrencyCode)
RootNode.appendChild currencyCodeNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

Dim BuyerRefNode As Object
Set BuyerRefNode = CreateNode(xmlDoc, "cbc:BuyerReference", BuyerReference)
RootNode.appendChild BuyerRefNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' Verkäuferinformationen hinzufügen
Dim SupplierNode As Object
Set SupplierNode = xmlDoc.createElement("cac:AccountingSupplierParty")
SupplierNode.appendChild xmlDoc.createTextNode(vbCrLf)
SupplierNode.appendChild xmlDoc.createTextNode(vbTab)
SupplierNode.appendChild xmlDoc.createTextNode(vbTab)

Dim PartyNode As Object
Set PartyNode = xmlDoc.createElement("cac:Party")

PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

' EndpointID
Dim EndpointIDNode As Object
Set EndpointIDNode = CreateNode(xmlDoc, "cbc:EndpointID", BuyerEndpointID)
EndpointIDNode.setAttribute "schemeID", "EM" 'Emailadresse
PartyNode.appendChild EndpointIDNode
PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

If InvoicePayment = "59" Or InvoicePayment = "49" Then
    ' PartyIdentification
    Dim PartyIdentification As Object
    Set PartyIdentification = xmlDoc.createElement("cac:PartyIdentification")

    PartyIdentification.appendChild xmlDoc.createTextNode(vbCrLf)
    PartyIdentification.appendChild xmlDoc.createTextNode(vbTab)
    PartyIdentification.appendChild xmlDoc.createTextNode(vbTab)

    Dim PINIdNode As Object
    Set PINIdNode = CreateNode(xmlDoc, "cbc:ID", PaymentMandate)
    PINIdNode.setAttribute "schemeID", "SEPA"
    
    PartyIdentification.appendChild PINIdNode
    PartyIdentification.appendChild xmlDoc.createTextNode(vbCrLf)
    PartyIdentification.appendChild xmlDoc.createTextNode(vbTab)
    PartyIdentification.appendChild xmlDoc.createTextNode(vbTab)
    PartyNode.appendChild PartyIdentification
    PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PartyNode.appendChild xmlDoc.createTextNode(vbTab)
    PartyNode.appendChild xmlDoc.createTextNode(vbTab)
    'SellerAssignedId
End If

' Verkäuferadresse (PostalAddress)
Dim AddressNode As Object
Set AddressNode = xmlDoc.createElement("cac:PostalAddress")
AddressNode.appendChild CreateNode(xmlDoc, "cbc:StreetName", SellerStreet)
AddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild CreateNode(xmlDoc, "cbc:CityName", SellerCity)
AddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild CreateNode(xmlDoc, "cbc:PostalZone", SellerPostalZone)
AddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)

Dim CountryNode As Object
Set CountryNode = xmlDoc.createElement("cac:Country")
CountryNode.appendChild xmlDoc.createTextNode(vbCrLf)
CountryNode.appendChild xmlDoc.createTextNode(vbTab)
CountryNode.appendChild xmlDoc.createTextNode(vbTab)

CountryNode.appendChild CreateNode(xmlDoc, "cbc:IdentificationCode", SellerCountryCode)
CountryNode.appendChild xmlDoc.createTextNode(vbCrLf)
CountryNode.appendChild xmlDoc.createTextNode(vbTab)
CountryNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild CountryNode
AddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
AddressNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild AddressNode
PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

' party tax scheme
Dim PartyTaxNode As Object
Set PartyTaxNode = xmlDoc.createElement("cac:PartyTaxScheme")
PartyTaxNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)

PartyTaxNode.appendChild CreateNode(xmlDoc, "cbc:CompanyID", SellerUID)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)

Dim SchemeNode As Object
Set SchemeNode = xmlDoc.createElement("cac:TaxScheme")
SchemeNode.appendChild xmlDoc.createTextNode(vbCrLf)
SchemeNode.appendChild xmlDoc.createTextNode(vbTab)
SchemeNode.appendChild xmlDoc.createTextNode(vbTab)

SchemeNode.appendChild CreateNode(xmlDoc, "cbc:ID", "GST")
SchemeNode.appendChild xmlDoc.createTextNode(vbCrLf)
SchemeNode.appendChild xmlDoc.createTextNode(vbTab)
SchemeNode.appendChild xmlDoc.createTextNode(vbTab)

PartyTaxNode.appendChild SchemeNode
PartyTaxNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)
PartyTaxNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild PartyTaxNode
PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

'VAT: Value Added Tax (Mehrwertsteuer)
'GST: Goods and Services Tax (Waren- und Dienstleistungssteuer)
'EXC: Excise Duty(Verbrauchssteuer)
'FC:  Freight Charges(Frachtkosten)
'IGST: Integrated Goods and Services Tax (Integrierte Waren- und Dienstleistungssteuer)
'CGST: Central Goods and Services Tax (Zentrale Waren- und Dienstleistungssteuer)
'SGST: State Goods and Services Tax (Staatliche Waren- und Dienstleistungssteuer)

' Legal Entity für den Verkäufer
Dim LegalEntityNode As Object
Set LegalEntityNode = xmlDoc.createElement("cac:PartyLegalEntity")
LegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)

LegalEntityNode.appendChild CreateNode(xmlDoc, "cbc:RegistrationName", SellerRegName)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
LegalEntityNode.appendChild CreateNode(xmlDoc, "cbc:CompanyID", SellerCompanyID)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
LegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild LegalEntityNode
PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

' Verkäuferkontakt
Dim contactNode As Object
Set contactNode = xmlDoc.createElement("cac:Contact")
contactNode.appendChild CreateNode(xmlDoc, "cbc:Name", SellerRegName)
contactNode.appendChild xmlDoc.createTextNode(vbCrLf)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
contactNode.appendChild CreateNode(xmlDoc, "cbc:Telephone", SellerTelephone)
contactNode.appendChild xmlDoc.createTextNode(vbCrLf)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
contactNode.appendChild CreateNode(xmlDoc, "cbc:ElectronicMail", SellerEmail)
contactNode.appendChild xmlDoc.createTextNode(vbCrLf)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
contactNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild contactNode
PartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)
PartyNode.appendChild xmlDoc.createTextNode(vbTab)

SupplierNode.appendChild PartyNode
SupplierNode.appendChild xmlDoc.createTextNode(vbCrLf)
SupplierNode.appendChild xmlDoc.createTextNode(vbTab)
SupplierNode.appendChild xmlDoc.createTextNode(vbTab)
RootNode.appendChild SupplierNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' Käuferinformationen hinzufügen
Dim CustomerNode As Object
Set CustomerNode = xmlDoc.createElement("cac:AccountingCustomerParty")
CustomerNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerNode.appendChild xmlDoc.createTextNode(vbTab)

Dim CustomerPartyNode As Object
Set CustomerPartyNode = xmlDoc.createElement("cac:Party")
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)

' EndpointID
Dim EndpointNode As Object
Set EndpointNode = CreateNode(xmlDoc, "cbc:EndpointID", SellerUID)
EndpointNode.setAttribute "schemeID", "9930"
CustomerPartyNode.appendChild EndpointNode
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)

' Käuferadresse (PostalAddress)
Dim CustomerAddressNode As Object
Set CustomerAddressNode = xmlDoc.createElement("cac:PostalAddress")
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)

CustomerAddressNode.appendChild CreateNode(xmlDoc, "cbc:StreetName", BuyerStreet)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild CreateNode(xmlDoc, "cbc:CityName", BuyerCity)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild CreateNode(xmlDoc, "cbc:PostalZone", BuyerPostalZone)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)

Dim BuyerCountryNode As Object
Set BuyerCountryNode = xmlDoc.createElement("cac:Country")
BuyerCountryNode.appendChild xmlDoc.createTextNode(vbCrLf)
BuyerCountryNode.appendChild xmlDoc.createTextNode(vbTab)
BuyerCountryNode.appendChild xmlDoc.createTextNode(vbTab)

BuyerCountryNode.appendChild CreateNode(xmlDoc, "cbc:IdentificationCode", BuyerCountryCode)
CustomerAddressNode.appendChild BuyerCountryNode
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerAddressNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild CustomerAddressNode
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)

' Legal Entity für den Käufer
Dim CustomerLegalEntityNode As Object
Set CustomerLegalEntityNode = xmlDoc.createElement("cac:PartyLegalEntity")
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)

CustomerLegalEntityNode.appendChild CreateNode(xmlDoc, "cbc:RegistrationName", BuyerRegistrationName)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerLegalEntityNode.appendChild CreateNode(xmlDoc, "cbc:CompanyID", BuyerCompanyID)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerLegalEntityNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild CustomerLegalEntityNode
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)
CustomerPartyNode.appendChild xmlDoc.createTextNode(vbTab)

CustomerNode.appendChild CustomerPartyNode
CustomerNode.appendChild xmlDoc.createTextNode(vbCrLf)
CustomerNode.appendChild xmlDoc.createTextNode(vbTab)
RootNode.appendChild CustomerNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

'1: Instrument Not defined
'2: Automated clearing house credit
'3: Automated clearing house debit
'10: In cash (Barzahlung)
'20: Cheque (Scheck)
'30: Credit Transfer(Überweisung)
'42: Payment to bank account (Zahlung auf Bankkonto)
'48: Bank card (Kreditkarte, Einkaufskarte und andere Kreditmethoden)
'49: Direct Debit(Lastschrift)
'58: SEPA Credit Transfer (SEPA-Überweisung)
'59: SEPA Direct Debit (SEPA-Lastschrift)
'97: Not defined (Nicht definiert)

' Zahlungsanweisungen hinzufügen (PaymentMeans)
Dim PaymentMeansNode As Object
Set PaymentMeansNode = xmlDoc.createElement("cac:PaymentMeans")
PaymentMeansNode.appendChild CreateNode(xmlDoc, "cbc:PaymentMeansCode", InvoicePayment)
PaymentMeansNode.appendChild xmlDoc.createTextNode(vbCrLf)
PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)
PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)

' Fallunterscheidung für verschiedene Zahlungsarten
If InvoicePayment = "42" Or InvoicePayment = "58" Or InvoicePayment = "10" Or InvoicePayment = "97" Then

    ' Zahlung auf Bankkonto, SEPA-Überweisung, Barzahlung
    Dim PayeeFinAcc As Object
    Set PayeeFinAcc = xmlDoc.createElement("cac:PayeeFinancialAccount")
    PayeeFinAcc.appendChild CreateNode(xmlDoc, "cbc:ID", SellerIBAN)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbCrLf)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)
    PayeeFinAcc.appendChild CreateNode(xmlDoc, "cbc:Name", SellerRegName)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbCrLf)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)

    Dim finnInstBranch As Object
    Set finnInstBranch = xmlDoc.createElement("cac:FinancialInstitutionBranch")
    finnInstBranch.appendChild CreateNode(xmlDoc, "cbc:ID", SellerBIC)
    finnInstBranch.appendChild xmlDoc.createTextNode(vbCrLf)
    finnInstBranch.appendChild xmlDoc.createTextNode(vbTab)
    finnInstBranch.appendChild xmlDoc.createTextNode(vbTab)
    
    PayeeFinAcc.appendChild finnInstBranch
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbCrLf)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)
    PayeeFinAcc.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMeansNode.appendChild PayeeFinAcc
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)

ElseIf InvoicePayment = "48" Then

    ' Bankkartenzahlung
    Dim CartAccount As Object
    Set CartAccount = xmlDoc.createElement("cac:CardAccount")
    CartAccount.appendChild CreateNode(xmlDoc, "cbc:PrimaryAccountNumberID", BuyerAccountShortID) 'IBAN des Kunden
    CartAccount.appendChild xmlDoc.createTextNode(vbCrLf)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)
    CartAccount.appendChild CreateNode(xmlDoc, "cbc:NetworkID", "MC") ' VISA, MC, AMEX, etc.
    CartAccount.appendChild xmlDoc.createTextNode(vbCrLf)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)
    CartAccount.appendChild CreateNode(xmlDoc, "cbc:HolderName", BuyerRegistrationName)
    CartAccount.appendChild xmlDoc.createTextNode(vbCrLf)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)
    CartAccount.appendChild xmlDoc.createTextNode(vbTab)

    PaymentMeansNode.appendChild CartAccount
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Optional: Zeitraum der Gültigkeit (ValidityPeriod)
    Dim PaymentMandateNode1 As Object
    Set PaymentMandateNode1 = xmlDoc.createElement("cac:PaymentMandate")
    
    Dim ValidityPeriodNode As Object
    Set ValidityPeriodNode = xmlDoc.createElement("cac:ValidityPeriod")
    ValidityPeriodNode.appendChild CreateNode(xmlDoc, "cbc:StartDate", "2025-01-01") 'ööö
    ValidityPeriodNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ValidityPeriodNode.appendChild xmlDoc.createTextNode(vbTab)
    ValidityPeriodNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMandateNode1.appendChild ValidityPeriodNode
    PaymentMandateNode1.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMandateNode1.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMandateNode1.appendChild xmlDoc.createTextNode(vbTab)

ElseIf InvoicePayment = "59" Or InvoicePayment = "49" Then
   
   ' Zahlungsmandat erstellen
    Dim PaymentMandateNode As Object
    Set PaymentMandateNode = xmlDoc.createElement("cac:PaymentMandate")
            
    ' Mandatsreferenz-ID hinzufügen
    PaymentMandateNode.appendChild CreateNode(xmlDoc, "cbc:ID", PaymentMandate) 'Mandatsreferenz-ID
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Belastungskonto (PayerFinancialAccount)
    Dim payerFinancialAccountNode As Object
    Set payerFinancialAccountNode = xmlDoc.createElement("cac:PayerFinancialAccount")
    payerFinancialAccountNode.appendChild CreateNode(xmlDoc, "cbc:ID", BuyerAccountNumberID) 'IBAN des Kunden
    payerFinancialAccountNode.appendChild xmlDoc.createTextNode(vbCrLf)
    payerFinancialAccountNode.appendChild xmlDoc.createTextNode(vbTab)
    payerFinancialAccountNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMandateNode.appendChild payerFinancialAccountNode
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbTab)
    PaymentMandateNode.appendChild xmlDoc.createTextNode(vbTab)

    ' Zahlungsmandat dem PaymentMeans-Knoten hinzufügen
    PaymentMeansNode.appendChild PaymentMandateNode
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbCrLf)
    PaymentMeansNode.appendChild xmlDoc.createTextNode(vbTab)
    
End If

' Zahlungsanweisung dem Wurzelknoten hinzufügen
RootNode.appendChild PaymentMeansNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' AllowanceCharge
If AllowanceTotalAmount > "0" Then
    Dim AllowanceChargeNode As Object
    Set AllowanceChargeNode = xmlDoc.createElement("cac:AllowanceCharge")
    AllowanceChargeNode.appendChild CreateNode(xmlDoc, "cbc:ChargeIndicator", "false")
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    
    AllowanceChargeNode.appendChild CreateNode(xmlDoc, "cbc:AllowanceChargeReason", "Discount")
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    
    AllowanceChargeNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:Amount", Replace(CStr(AllowanceTotalAmount), ",", "."), "currencyID", CurrencyCode)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Tax category
    Dim AllowanceChargeTaxSubCategoryNode As Object
    Set AllowanceChargeTaxSubCategoryNode = xmlDoc.createElement("cac:TaxCategory")
    AllowanceChargeTaxSubCategoryNode.appendChild CreateNode(xmlDoc, "cbc:ID", InvoiceTaxCategory)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeTaxSubCategoryNode.appendChild CreateNode(xmlDoc, "cbc:Percent", Replace(CStr(InvoiceTax), ",", "."))
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeTaxSubCategoryNode.appendChild CreateTaxSchemeNode(xmlDoc)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    AllowanceChargeTaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    
    AllowanceChargeNode.appendChild AllowanceChargeTaxSubCategoryNode
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbCrLf)
    AllowanceChargeNode.appendChild xmlDoc.createTextNode(vbTab)
    RootNode.appendChild AllowanceChargeNode
    RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
    RootNode.appendChild xmlDoc.createTextNode(vbTab)
End If

' Steuerinformationen (TaxTotal)
Dim TaxTotalNode As Object
Set TaxTotalNode = xmlDoc.createElement("cac:TaxTotal")
TaxTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:TaxAmount", Replace(CStr(TaxAmount), ",", "."), "currencyID", CurrencyCode)
TaxTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxTotalNode.appendChild xmlDoc.createTextNode(vbTab)
TaxTotalNode.appendChild xmlDoc.createTextNode(vbTab)

Dim TaxSubtotalNode As Object
Set TaxSubtotalNode = xmlDoc.createElement("cac:TaxSubtotal")
TaxSubtotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:TaxableAmount", Replace(CStr(TaxableAmount), ",", "."), "currencyID", CurrencyCode)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubtotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:TaxAmount", Replace(CStr(TaxAmount), ",", "."), "currencyID", CurrencyCode)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)

' Tax category
Dim TaxSubCategoryNode As Object
Set TaxSubCategoryNode = xmlDoc.createElement("cac:TaxCategory")
TaxSubCategoryNode.appendChild CreateNode(xmlDoc, "cbc:ID", InvoiceTaxCategory)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubCategoryNode.appendChild CreateNode(xmlDoc, "cbc:Percent", Replace(CStr(InvoiceTax), ",", "."))
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubCategoryNode.appendChild CreateTaxSchemeNode(xmlDoc)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubtotalNode.appendChild TaxSubCategoryNode
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)
TaxSubtotalNode.appendChild xmlDoc.createTextNode(vbTab)

TaxTotalNode.appendChild TaxSubtotalNode
TaxTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
TaxTotalNode.appendChild xmlDoc.createTextNode(vbTab)
RootNode.appendChild TaxTotalNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' Summeninformationen (LegalMonetaryTotal)
Dim LegalMonetaryTotalNode As Object
Set LegalMonetaryTotalNode = xmlDoc.createElement("cac:LegalMonetaryTotal")
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:LineExtensionAmount", Replace(CStr(LineExtensionAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:TaxExclusiveAmount", Replace(CStr(TaxExclusiveAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:TaxInclusiveAmount", Replace(CStr(TaxInclusiveAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:AllowanceTotalAmount", Replace(CStr(AllowanceTotalAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:ChargeTotalAmount", 0#, "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:PrepaidAmount", Replace(CStr(PrepaidAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:PayableRoundingAmount", 0#, "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)

LegalMonetaryTotalNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:PayableAmount", Replace(CStr(PayableAmount), ",", "."), "currencyID", CurrencyCode)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbCrLf)
LegalMonetaryTotalNode.appendChild xmlDoc.createTextNode(vbTab)
RootNode.appendChild LegalMonetaryTotalNode
RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
RootNode.appendChild xmlDoc.createTextNode(vbTab)

' Rechnungspositionen hinzufügen
For AktZa = 0 To LineCount - 1
    Dim ItemNode As Object
    Set ItemNode = xmlDoc.createElement("cac:InvoiceLine")
    ItemNode.appendChild CreateNode(xmlDoc, "cbc:ID", InvoiceTaxCategory)
    ItemNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:InvoicedQuantity", Replace(CStr(InvoiceLines(AktZa).Quantity), ",", "."), "unitCode", InvoiceLines(AktZa).UnitCode)
    ItemNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:LineExtensionAmount", Replace(CStr(InvoiceLines(AktZa).LineExtensionAmount), ",", "."), "currencyID", CurrencyCode)
    ItemNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Artikeldetails
    Dim ItemDetailNode As Object
    Set ItemDetailNode = xmlDoc.createElement("cac:Item")
    ItemDetailNode.appendChild CreateNode(xmlDoc, "cbc:Name", InvoiceLines(AktZa).Name)
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Steuerkategorie für den Artikel
    Dim TaxCategoryNode As Object
    Set TaxCategoryNode = xmlDoc.createElement("cac:ClassifiedTaxCategory")
    TaxCategoryNode.appendChild CreateNode(xmlDoc, "cbc:ID", InvoiceTaxCategory)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    TaxCategoryNode.appendChild CreateNode(xmlDoc, "cbc:Percent", Replace(CStr(InvoiceLines(AktZa).TaxPercent), ",", "."))
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    TaxCategoryNode.appendChild CreateTaxSchemeNode(xmlDoc)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbCrLf)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    TaxCategoryNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemDetailNode.appendChild TaxCategoryNode
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemDetailNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild ItemDetailNode
    ItemNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    
    ' Preisinformationen für den Artikel
    Dim priceNode As Object
    Set priceNode = xmlDoc.createElement("cac:Price")
    priceNode.appendChild CreateNodeWithAttr(xmlDoc, "cbc:PriceAmount", Replace(CStr(InvoiceLines(AktZa).PriceAmount), ",", "."), "currencyID", CurrencyCode)
    priceNode.appendChild xmlDoc.createTextNode(vbCrLf)
    priceNode.appendChild xmlDoc.createTextNode(vbTab)
    priceNode.appendChild xmlDoc.createTextNode(vbTab)
    ItemNode.appendChild priceNode
    ItemNode.appendChild xmlDoc.createTextNode(vbCrLf)
    ItemNode.appendChild xmlDoc.createTextNode(vbTab)
    
    RootNode.appendChild ItemNode
    RootNode.appendChild xmlDoc.createTextNode(vbCrLf)
    RootNode.appendChild xmlDoc.createTextNode(vbTab)
Next AktZa

' Speichere das XML-Dokument
xmlDoc.Save FiNam

End Sub
Public Function LoadXMLFile(FiNam As String) As String
    ' Methode zum Öffnen der XML-Datei und Formatieren in eine String-Variable
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    xmlDoc.async = False
    xmlDoc.Load FiNam

    If xmlDoc.parseError.errorCode <> 0 Then
        LoadXMLFile = "Fehler beim Laden der XML-Datei: " & xmlDoc.parseError.Reason
        Exit Function
    End If

    Dim RootNode As Object
    Set RootNode = xmlDoc.documentElement

    ' Initialisieren der Ausgabevariable
    Dim result As String
    result = "X-Rechnung Details:" & vbCrLf & vbCrLf

    ' Auslesen und Formatieren der XML-Daten
    result = result & "Rechnungsnummer: " & GetNodeValue(RootNode, "cbc:ID") & vbCrLf
    result = result & "Rechnungsdatum: " & GetNodeValue(RootNode, "cbc:IssueDate") & vbCrLf
    result = result & "Fälligkeitsdatum: " & GetNodeValue(RootNode, "cbc:DueDate") & vbCrLf
    result = result & "Rechnungstyp: " & GetNodeValue(RootNode, "cbc:InvoiceTypeCode") & vbCrLf
    result = result & "Währungscode: " & GetNodeValue(RootNode, "cbc:DocumentCurrencyCode") & vbCrLf

    ' Seller information
    Dim SellerNode As Object
    Set SellerNode = RootNode.selectSingleNode("cac:AccountingSupplierParty/cac:Party")
    If Not SellerNode Is Nothing Then
        result = result & vbCrLf & "Verkäuferinformationen:" & vbCrLf
        result = result & "  Name: " & GetNodeValue(SellerNode, "cac:PartyLegalEntity/cbc:RegistrationName") & vbCrLf
        result = result & "  Straße: " & GetNodeValue(SellerNode, "cac:PostalAddress/cbc:StreetName") & vbCrLf
        result = result & "  Stadt: " & GetNodeValue(SellerNode, "cac:PostalAddress/cbc:CityName") & vbCrLf
        result = result & "  Postleitzahl: " & GetNodeValue(SellerNode, "cac:PostalAddress/cbc:PostalZone") & vbCrLf
        result = result & "  Land: " & GetNodeValue(SellerNode, "cac:PostalAddress/cac:Country/cbc:IdentificationCode") & vbCrLf
    End If

    ' Buyer information
    Dim BuyerNode As Object
    Set BuyerNode = RootNode.selectSingleNode("cac:AccountingCustomerParty/cac:Party")
    If Not BuyerNode Is Nothing Then
        result = result & vbCrLf & "Käuferinformationen:" & vbCrLf
        result = result & "  Name: " & GetNodeValue(BuyerNode, "cac:PartyLegalEntity/cbc:RegistrationName") & vbCrLf
        result = result & "  Straße: " & GetNodeValue(BuyerNode, "cac:PostalAddress/cbc:StreetName") & vbCrLf
        result = result & "  Stadt: " & GetNodeValue(BuyerNode, "cac:PostalAddress/cbc:CityName") & vbCrLf
        result = result & "  Postleitzahl: " & GetNodeValue(BuyerNode, "cac:PostalAddress/cbc:PostalZone") & vbCrLf
        result = result & "  Land: " & GetNodeValue(BuyerNode, "cac:PostalAddress/cac:Country/cbc:IdentificationCode") & vbCrLf
    End If

    ' Gesamtbetrag
    Dim LegalMonetaryTotalNode As Object
    Set LegalMonetaryTotalNode = RootNode.selectSingleNode("cac:LegalMonetaryTotal")
    If Not LegalMonetaryTotalNode Is Nothing Then
        result = result & vbCrLf & "Gesamtbetrag:" & vbCrLf
        result = result & "  Betrag zahlbar: " & GetNodeValue(LegalMonetaryTotalNode, "cbc:PayableAmount") & vbCrLf
    End If

    ' Rechnungspositionen
    Dim LineNode As Object
    For Each LineNode In RootNode.selectNodes("cac:InvoiceLine")
        result = result & vbCrLf & "Rechnungsposition:" & vbCrLf
        result = result & "  ID: " & GetNodeValue(LineNode, "cbc:ID") & vbCrLf
        result = result & "  Menge: " & GetNodeValue(LineNode, "cbc:InvoicedQuantity") & vbCrLf
        result = result & "  Beschreibung: " & GetNodeValue(LineNode, "cac:Item/cbc:Name") & vbCrLf
        result = result & "  Preis: " & GetNodeValue(LineNode, "cac:Price/cbc:PriceAmount") & vbCrLf
    Next LineNode

    LoadXMLFile = result
End Function
Private Function CreateNode(xmlDoc As Object, tagName As String, textValue As String) As Object
' Hilfsfunktion zum Erstellen eines einfachen Knotens

Dim Node As Object

Set Node = xmlDoc.createElement(tagName)
Node.Text = textValue
Set CreateNode = Node

End Function
Private Function CreateNodeWithAttr(xmlDoc As Object, tagName As String, textValue As String, attrName As String, attrValue As String) As Object
' Hilfsfunktion zum Erstellen eines Knotens mit Attribut
    
Dim Node As Object

Set Node = xmlDoc.createElement(tagName)
Node.Text = textValue
Node.setAttribute attrName, attrValue
Set CreateNodeWithAttr = Node

End Function
Private Function CreateTaxSchemeNode(xmlDoc As Object) As Object
' Hilfsfunktion zum Erstellen des Steuer-Schemaknotens

Dim taxSchemeNode As Object

Set taxSchemeNode = xmlDoc.createElement("cac:TaxScheme")
taxSchemeNode.appendChild CreateNode(xmlDoc, "cbc:ID", "VAT")
Set CreateTaxSchemeNode = taxSchemeNode

End Function
Private Function GetNodeValue(parentNode As Object, tagName As String) As String
' Hilfsfunktion zum Abrufen des Werts eines Knotens
Dim Node As Object

Set Node = parentNode.selectSingleNode(tagName)
If Not Node Is Nothing Then
    GetNodeValue = Node.Text
Else
    GetNodeValue = ""
End If

End Function

