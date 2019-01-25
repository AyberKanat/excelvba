Private pCode As String
Private pTitle As String
Private pSExcName As String
Private pSExcID As String
Private pRgnID As String
Private pStatus As String
Private pRgnType As String
Private pCity As String
Private pSolAttr As Boolean
Private pStartTime As String
Private pTenure As Long
Private pSequence As Long
Private pGMAct as Long
private pGMTrg as Long
private pGIAct as Long
private pGITrg as Long
private pHAct as Long
private pHTrg as Long
private pSAct as Long
private pSTrg as Long
private pDAct as Long
private pDTrg as Long
private pSGAct as Long
private pSGTrg as Long

'Sequence property
Public Property Get Sequence() As Long
    Sequence = pSequence
End Property
Public Property Let Sequence(Value As Long)
    pSequence = Value
End Property
'Code property
Public Property Get Code() As String
    Code = pCode
End Property
Public Property Let Code(Value As String)
    pCode = Value
End Property
'Title property
Public Property Get Title() As String
    Title = pTitle
End Property
Public Property Let Title(Value As String)
    pTitle = Value
End Property
'SExcName property
Public Property Get SExcName() As String
    SExcName = pSExcName
End Property
Public Property Let SExcName(Value As String)
    pSExcName = Value
End Property
'SExcID property
Public Property Get SExcID() As String
    SExcID = pSExcID
End Property
Public Property Let SExcID(Value As String)
    pSExcID = Value
End Property
'RgnID property
Public Property Get RgnID() As String
    RgnID = pRgnID
End Property
Public Property Let RgnID(Value As String)
    pRgnID = Value
End Property
'Status property
Public Property Get Status() As String
    Status = pStatus
End Property
Public Property Let Status(Value As String)
    pStatus = Value
End Property
'RgnType property
Public Property Get RgnType() As String
    RgnType = pRgnType
End Property
Public Property Let RgnType(Value As String)
    pRgnType = Value
End Property
'City property
Public Property Get City() As String
    City = pCity
End Property
Public Property Let City(Value As String)
    pCity = Value
End Property
'SolAttr property
Public Property Get SolAttr() As Boolean
    SolAttr = pSolAttr
End Property
Public Property Let SolAttr(Value As Boolean)
    pSolAttr = Value
End Property
'StartTime property
Public Property Get StartTime() As String
    StartTime = pStartTime
End Property
Public Property Let StartTime(Value As String)
    pStartTime = Value
End Property
'Tenure property
Public Property Get Tenure() As Long
    Tenure = pTenure
End Property
Public Property Let Tenure(Value As Long)
    pTenure = Value
End Property
'GMAct property
Public Property Get GMAct() As Long
    GMAct = pGMAct
End Property
Public Property Let GMAct(Value As Long)
    pGMAct = Value
End Property'GMTrg property
Public Property Get GMTrg() As Long
    GMTrg = pGMTrg
End Property
Public Property Let GMTrg(Value As Long)
    pGMTrg = Value
End Property'GIAct property
Public Property Get GIAct() As Long
    GIAct = pGIAct
End Property
Public Property Let GIAct(Value As Long)
    pGIAct = Value
End Property'GITrg property
Public Property Get GITrg() As Long
    GITrg = pGITrg
End Property
Public Property Let GITrg(Value As Long)
    pGITrg = Value
End Property'HAct property
Public Property Get HAct() As Long
    HAct = pHAct
End Property
Public Property Let HAct(Value As Long)
    pHAct = Value
End Property'HTrg property
Public Property Get HTrg() As Long
    HTrg = pHTrg
End Property
Public Property Let HTrg(Value As Long)
    pHTrg = Value
End Property'SAct property
Public Property Get SAct() As Long
    SAct = pSAct
End Property
Public Property Let SAct(Value As Long)
    pSAct = Value
End Property'STrg property
Public Property Get STrg() As Long
    STrg = pSTrg
End Property
Public Property Let STrg(Value As Long)
    pSTrg = Value
End Property'DAct property
Public Property Get DAct() As Long
    DAct = pDAct
End Property
Public Property Let DAct(Value As Long)
    pDAct = Value
End Property'DTrg property
Public Property Get DTrg() As Long
    DTrg = pDTrg
End Property
Public Property Let DTrg(Value As Long)
    pDTrg = Value
End Property'SGACt property
Public Property Get SGACt() As Long
    SGACt = pSGACt
End Property
Public Property Let SGACt(Value As Long)
    pSGACt = Value
End Property'SGTrg property
Public Property Get SGTrg() As Long
    SGTrg = pSGTrg
End Property
Public Property Let SGTrg(Value As Long)
    pSGTrg = Value
End Property