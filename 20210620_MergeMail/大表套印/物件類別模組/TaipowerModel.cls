VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaipowerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_CalculationDay As String
Private m_ElectricNumber As String
Private m_ElectricMeterNumber As String
Private m_Multiple As String
Private m_VerificationDeadline As String
Private m_UserName As String
Private m_NextDate As String
Private m_Matter As String
Private m_ElectricAddress As String
Private m_MailAddress As String
Private m_Phone1 As String
Private m_Phone2 As String
Private m_Coordinate As String
Private m_PoleNumber As String
Private m_CurrentTransformer As String

Private m_TableTypes As Variant
Private m_Types As Variant
Private m_CurrentValues As Variant
Private m_DifferentValues As Variant

Public Property Get CalculationDay() As String
    CalculationDay = m_CalculationDay
End Property
Public Property Let CalculationDay(value As String)
    m_CalculationDay = value
End Property

Public Property Get ElectricNumber() As String
    ElectricNumber = m_ElectricNumber
End Property
Public Property Let ElectricNumber(value As String)
    m_ElectricNumber = value
End Property

Public Property Get District() As String
    District = Left(m_ElectricNumber, 2)
End Property

Public Property Get BusinessArea() As String
    BusinessArea = Mid(m_ElectricNumber, 3, 2)
End Property

Public Property Get AccountNumber() As String
    AccountNumber = Mid(m_ElectricNumber, 5, 4)
End Property

Public Property Get CategoryNumber() As String
    CategoryNumber = Mid(m_ElectricNumber, 9, 2)
End Property

Public Property Get CheckNumber() As String
    CheckNumber = Mid(m_ElectricNumber, 11, 1)
End Property

Public Property Get ElectricMeterNumber() As String
    ElectricMeterNumber = m_ElectricMeterNumber
End Property
Public Property Let ElectricMeterNumber(value As String)
    m_ElectricMeterNumber = value
End Property

Public Property Get Ampere() As String
    If m_Multiple = "01" Then
        Ampere = "0030"
    Else
        Ampere = CStr(Val(m_Multiple) * 5)
        If Len(Ampere) < 4 Then
            Ampere = Format(Ampere, "0000")
        End If
    End If
End Property

Public Property Get CurrentTransformer() As String
    If m_Multiple = "01" Then
        CurrentTransformer = "    "
    Else
        CurrentTransformer = CStr(Val(m_Multiple) * 5)
        If Len(CurrentTransformer) < 4 Then
            CurrentTransformer = Format(CurrentTransformer, "0000")
        End If
    End If
End Property

Public Property Get Multiple() As String
    Multiple = m_Multiple
End Property
Public Property Let Multiple(value As String)
    m_Multiple = value
End Property

Public Property Get VerificationDeadline() As String
    VerificationDeadline = m_VerificationDeadline
End Property
Public Property Let VerificationDeadline(value As String)
    m_VerificationDeadline = value
End Property

Public Property Get UserName() As String
    UserName = m_UserName
End Property
Public Property Let UserName(value As String)
    m_UserName = value
End Property

Public Property Get NextDate() As String
    NextDate = m_NextDate
End Property
Public Property Let NextDate(value As String)
    m_NextDate = value
End Property

Public Property Get ElectricAddress() As String
    ElectricAddress = m_ElectricAddress
End Property
Public Property Let ElectricAddress(value As String)
    m_ElectricAddress = value
End Property

Public Property Get MailAddress() As String
    MailAddress = m_MailAddress
End Property
Public Property Let MailAddress(value As String)
    m_MailAddress = value
End Property

Public Property Get Phone1() As String
    Phone1 = m_Phone1
End Property
Public Property Let Phone1(value As String)
    m_Phone1 = value
End Property

Public Property Get Phone2() As String
    Phone2 = m_Phone2
End Property
Public Property Let Phone2(value As String)
    m_Phone2 = value
End Property

Public Property Get Matter() As String
    Matter = m_Matter
End Property
Public Property Let Matter(value As String)
    m_Matter = value
End Property

Public Property Get Coordinate() As String
    Coordinate = m_Coordinate
End Property
Public Property Let Coordinate(value As String)
    m_Coordinate = value
End Property

Public Property Get PoleNumber() As String
    PoleNumber = m_PoleNumber
End Property
Public Property Let PoleNumber(value As String)
    m_PoleNumber = value
End Property

Public Property Get Types() As Variant
    Types = m_Types
End Property
Public Property Let Types(value As Variant)
    m_Types = value
End Property

Public Property Get CurrentValues() As Variant
    CurrentValues = m_CurrentValues
End Property
Public Property Let CurrentValues(value As Variant)
    m_CurrentValues = value
End Property

Public Property Get DifferentValues() As Variant
    DifferentValues = m_DifferentValues
End Property
Public Property Let DifferentValues(value As Variant)
    m_DifferentValues = value
End Property

Public Property Get TableTypes() As Variant
    TableTypes = m_TableTypes
End Property
Public Property Let TableTypes(value As Variant)
    m_TableTypes = value
End Property
