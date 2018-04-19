VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_lIndex            As Long
Private m_sTerm             As String
Private m_sType             As String
Private m_sDefinition       As String
Private m_colAlternative    As Collection
Private m_bSelected         As Boolean


' Property Sets and Gets

' Index
Public Property Let lIndex(ByVal v_lVal As Long)
    m_lIndex = v_lVal
End Property
Public Property Get lIndex() As Long
    lIndex = m_lIndex
End Property

' Term
Public Property Let sTerm(ByVal v_sVal As String)
    m_sTerm = v_sVal
End Property
Public Property Get sTerm() As String
    sTerm = m_sTerm
End Property

' Type
Public Property Let sType(ByVal v_sVal As String)
    m_sType = v_sVal
End Property
Public Property Get sType() As String
    sType = m_sType
End Property

' Definition
Public Property Let sDefinition(ByVal v_sVal As String)
    m_sDefinition = v_sVal
End Property
Public Property Get sDefinition() As String
    sDefinition = m_sDefinition
End Property

' Alternative text
Public Function setcolAlternative(ByVal v_sVal As String)
    
    Dim asString()  As String
    Dim lLowBound   As Long
    Dim lUppBound   As Long
    Dim lCounter    As Long
    
    If v_sVal = "" Then
        'm_colAlternative =
    Else
        asString = Split(v_sVal, ",")
        lLowBound = LBound(asString())
        lUppBound = UBound(asString())
        For lCounter = lLowBound To lUppBound
            Call m_colAlternative.Add(Trim(asString(lCounter)))
        Next lCounter
        
    End If
    
End Function
Public Property Get colAlternative() As Collection
    Set colAlternative = m_colAlternative
End Property

' Selected
Public Property Let bSelected(ByVal v_bVal As Boolean)
    m_bSelected = v_bVal
End Property
Public Property Get bSelected() As Boolean
    bSelected = m_bSelected
End Property


Private Sub Class_Initialize()
    Set m_colAlternative = New Collection
End Sub
