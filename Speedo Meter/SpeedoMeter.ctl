VERSION 5.00
Begin VB.UserControl SpeedoMeter 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SpeedoMeter.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2160
      Top             =   840
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "SpeedoMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'                              Speedo Meter UserControl
'Class Module made by Alain van Hanegem
'UserControl by Andre Felipe N. D.
'Planet Source Code, 2006

'Description: This UserControl was made to try to made the set up of the Speedo Meter
'             faster and easier. When I first got the Class Module from PSC, I thought
'             it was difficult to configurate, so I made this UserControl to make things
'             easier. If you are an advanced user and didn't have problems with the
'             class module, I suggest you keep using it, because it uses less PC Resouces
'             and has much more potential. But if you are a begginer, I hope this
'             UserControl can be useful for you. I tried to put in it almost all the
'             possible properties of the class module.

Option Explicit
Dim v As clsVUmeter

'Default Property Values:
Const m_def_Size = 1
Const m_def_NeedleMovement = 1
Const m_def_ScaleInterval = 20
Const m_def_Appearance = 0
Const m_def_ControlBorderStyle = 0
Const m_def_TextPlace = 1
Const m_def_ValuePlace = 2
Const m_def_ValueColor = 0
Const m_def_TextColor = 0
Const m_def_TextLabel = "Speed"
Const m_def_ValueLabel = "0 km/hr"
Const m_def_BackColor = &HFFFFFF
Const m_def_BorderColor = &HC00000
Const m_def_BorderThickness = 3
Const m_def_TickSmallColor = 0
Const m_def_TickBigColor = 0
Const m_def_NeedleStyle = 4
Const m_def_NeedleSize = 4
Const m_def_NeedleColor = &HFF&
Const m_def_IntSmall = 5
Const m_def_IntBig = 10
Const m_def_AngleMax = -30
Const m_def_AngleMin = 220
Const m_def_DirectNeedle = False
Const m_def_Value = 0
Const m_def_NeedleQuality = 80
Const m_def_Max = 240
Const m_def_Min = 0

'Property Variables:
Dim m_Appearance As Integer
Dim m_ControlBorderStyle As Integer
Dim m_TextPlace As Integer
Dim m_ValuePlace As Integer
Dim m_DirectNeedle As Boolean
Dim m_NeedleQuality As Integer

'Enums
Public Enum Needle_Style
Very_Thin = 0
Simple = 1
Stretched = 2
Big_Bottom = 3
Normal = 4
Large = 5
End Enum

Public Enum Needle_Movement
Smooth = 0
Slow = 1
End Enum

Public Enum Needle_Size
NoNeedle = 0
SuperShort = 1
Short = 2
Medium = 3
Long_ = 4
SuperLong = 5
End Enum

Public Enum Border_Thickness
No_Border = 0
VeryThin = 1
Thin = 2
Medium = 3
Large = 4
End Enum

Public Enum Text_Place
None = 0
Top = 1
Bottom = 2
End Enum

Public Enum Apperarance_
Flat = 0
Tridimensional = 1
End Enum

Public Enum BorderStyle_
None = 0
Fixed_Single = 1
End Enum

Public Enum Size_
Small = 1
Medium = 2
Big = 3
End Enum

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_NeedleQuality = m_def_NeedleQuality
    m_Value = m_def_Value
    m_DirectNeedle = m_def_DirectNeedle
    m_AngleMax = m_def_AngleMax
    m_AngleMin = m_def_AngleMin
    m_IntSmall = m_def_IntSmall
    m_IntBig = m_def_IntBig
    m_NeedleStyle = m_def_NeedleStyle
    m_NeedleSize = m_def_NeedleSize
    m_NeedleColor = m_def_NeedleColor
    m_TickSmallColor = m_def_TickSmallColor
    m_TickBigColor = m_def_TickBigColor
    m_BorderColor = m_def_BorderColor
    m_BorderThickness = m_def_BorderThickness
    m_BackColor = m_def_BackColor
    m_TextLabel = m_def_TextLabel
    Set m_TextFont = Ambient.Font
    m_ValueLabel = m_def_ValueLabel
    Set m_ValueFont = Ambient.Font
    m_ValueColor = m_def_ValueColor
    m_TextColor = m_def_TextColor
    Set m_ScaleFont = Ambient.Font
    m_TextPlace = m_def_TextPlace
    m_ValuePlace = m_def_ValuePlace
    m_Appearance = m_def_Appearance
    m_ControlBorderStyle = m_def_ControlBorderStyle
    m_ScaleInterval = m_def_ScaleInterval
    m_NeedleMovement = m_def_NeedleMovement
    m_Size = m_def_Size
    'Others
    TextY = (-1 / 3)
    ValueY = (1 / 2)
    StartGauge
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_NeedleQuality = PropBag.ReadProperty("NeedleQuality", m_def_NeedleQuality)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_DirectNeedle = PropBag.ReadProperty("DirectNeedle", m_def_DirectNeedle)
    m_AngleMax = PropBag.ReadProperty("AngleMax", m_def_AngleMax)
    m_AngleMin = PropBag.ReadProperty("AngleMin", m_def_AngleMin)
    m_IntSmall = PropBag.ReadProperty("IntSmall", m_def_IntSmall)
    m_IntBig = PropBag.ReadProperty("IntBig", m_def_IntBig)
    m_NeedleStyle = PropBag.ReadProperty("NeedleStyle", m_def_NeedleStyle)
    m_NeedleSize = PropBag.ReadProperty("NeedleSize", m_def_NeedleSize)
    m_NeedleColor = PropBag.ReadProperty("NeedleColor", m_def_NeedleColor)
    m_TickSmallColor = PropBag.ReadProperty("TickSmallColor", m_def_TickSmallColor)
    m_TickBigColor = PropBag.ReadProperty("TickBigColor", m_def_TickBigColor)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderThickness = PropBag.ReadProperty("BorderThickness", m_def_BorderThickness)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_TextLabel = PropBag.ReadProperty("TextLabel", m_def_TextLabel)
    Set m_TextFont = PropBag.ReadProperty("TextFont", Ambient.Font)
    m_ValueLabel = PropBag.ReadProperty("ValueLabel", m_def_ValueLabel)
    Set m_ValueFont = PropBag.ReadProperty("ValueFont", Ambient.Font)
    m_ValueColor = PropBag.ReadProperty("ValueColor", m_def_ValueColor)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
    Set m_ScaleFont = PropBag.ReadProperty("ScaleFont", Ambient.Font)
    m_TextPlace = PropBag.ReadProperty("TextPlace", m_def_TextPlace)
    m_ValuePlace = PropBag.ReadProperty("ValuePlace", m_def_ValuePlace)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_ControlBorderStyle = PropBag.ReadProperty("ControlBorderStyle", m_def_ControlBorderStyle)
    m_ScaleInterval = PropBag.ReadProperty("ScaleInterval", m_def_ScaleInterval)
    m_NeedleMovement = PropBag.ReadProperty("NeedleMovement", m_def_NeedleMovement)
    m_Size = PropBag.ReadProperty("Size", m_def_Size)
    
    'ValuePlace
    Select Case m_ValuePlace 'To define what to do with the value text
    Case 0 'Remove it. Just put its color equal to the background color
    m_ValueColor = m_BackColor
    Case 1 'In the top - y = (- 1 / 3)
    'Adjust the color
    If m_ValueColor = m_BackColor Then m_ValueColor = m_def_ValueColor
    ValueY = (-1 / 3)
    TextMove
    Case 2 'In the bottom - y = (1 / 2)
    If m_ValueColor = m_BackColor Then m_ValueColor = m_def_ValueColor
    ValueY = (1 / 2)
    TextMove
    End Select
    
    'TextPlace
    Select Case m_TextPlace 'To define what to do with the value text
    Case 0 'Remove it. Just put its color equal to the background color
    m_TextColor = m_BackColor
    Case 1 'In the top - y = (- 1 / 3)
    'Adjust the color
    If m_TextColor = m_BackColor Then m_TextColor = m_def_ValueColor
    TextY = (-1 / 3)
    LabelMove
    Case 2 'In the bottom - y = (1 / 2)
    If m_TextColor = m_BackColor Then m_TextColor = m_def_ValueColor
    TextY = (1 / 2)
    LabelMove
    End Select
    
    StartGauge
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("NeedleQuality", m_NeedleQuality, m_def_NeedleQuality)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("DirectNeedle", m_DirectNeedle, m_def_DirectNeedle)
    Call PropBag.WriteProperty("AngleMax", m_AngleMax, m_def_AngleMax)
    Call PropBag.WriteProperty("AngleMin", m_AngleMin, m_def_AngleMin)
    Call PropBag.WriteProperty("IntSmall", m_IntSmall, m_def_IntSmall)
    Call PropBag.WriteProperty("IntBig", m_IntBig, m_def_IntBig)
    Call PropBag.WriteProperty("NeedleStyle", m_NeedleStyle, m_def_NeedleStyle)
    Call PropBag.WriteProperty("NeedleSize", m_NeedleSize, m_def_NeedleSize)
    Call PropBag.WriteProperty("NeedleColor", m_NeedleColor, m_def_NeedleColor)
    Call PropBag.WriteProperty("TickSmallColor", m_TickSmallColor, m_def_TickSmallColor)
    Call PropBag.WriteProperty("TickBigColor", m_TickBigColor, m_def_TickBigColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderThickness", m_BorderThickness, m_def_BorderThickness)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("TextLabel", m_TextLabel, m_def_TextLabel)
    Call PropBag.WriteProperty("TextFont", m_TextFont, Ambient.Font)
    Call PropBag.WriteProperty("ValueLabel", m_ValueLabel, m_def_ValueLabel)
    Call PropBag.WriteProperty("ValueFont", m_ValueFont, Ambient.Font)
    Call PropBag.WriteProperty("ValueColor", m_ValueColor, m_def_ValueColor)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call PropBag.WriteProperty("ScaleFont", m_ScaleFont, Ambient.Font)
    Call PropBag.WriteProperty("TextPlace", m_TextPlace, m_def_TextPlace)
    Call PropBag.WriteProperty("ValuePlace", m_ValuePlace, m_def_ValuePlace)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("ControlBorderStyle", m_ControlBorderStyle, m_def_ControlBorderStyle)
    Call PropBag.WriteProperty("ScaleInterval", m_ScaleInterval, m_def_ScaleInterval)
    Call PropBag.WriteProperty("NeedleMovement", m_NeedleMovement, m_def_NeedleMovement)
    Call PropBag.WriteProperty("Size", m_Size, m_def_Size)
End Sub

Private Sub StartGauge()
    Set v = New clsVUmeter
    Dim cx As Long
    Dim cy As Long
    cx = P1.ScaleWidth \ 2
    cy = P1.ScaleHeight \ 2
    v.Init_Picture P1.hDC, P1.Image, P1.ScaleWidth, P1.ScaleHeight
    v.SetVUDefaults cx, cy
    v.Draw
    P1.Refresh
    Timer1.enabled = True
End Sub

Private Sub StopGauge()
    Timer1.enabled = False
    Set v = Nothing
End Sub

Private Sub Timer1_Timer()
    v.AnimationLoop
    v.Draw
    P1.Refresh
End Sub

Private Sub UserControl_Resize()
    StopGauge
    Select Case m_Size
    Case 1
    UserControl.Height = 1920
    UserControl.Width = 1920
    NeedleRef = 25
    Case 2
    UserControl.Height = 2760
    UserControl.Width = 2760
    NeedleRef = 20
    Case 3
    UserControl.Height = 3720
    UserControl.Width = 3720
    NeedleRef = 15
    End Select
    P1.Height = UserControl.Height
    P1.Width = UserControl.Width
    StartGauge
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,240
Public Property Get Max() As Integer
Attribute Max.VB_Description = "Sets the Maximum value of the SpeedoMeter."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    If New_Max < m_Min Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than Min.", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_Max = New_Max
    PropertyChanged "Max"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
Attribute Min.VB_Description = "Sets the minimum value ef the SpeedoMeter."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    If New_Min > m_Max Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values smaller than Max.", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_Min = New_Min
    PropertyChanged "Min"
    StartGauge
    End If
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,80
Public Property Get NeedleQuality() As Integer
Attribute NeedleQuality.VB_Description = "Sets (in porcentage) the quality of the Needle's movements. Higher value means better movement, but also higher CPU usage."
Attribute NeedleQuality.VB_ProcData.VB_Invoke_Property = ";Behavior"
    NeedleQuality = m_NeedleQuality
End Property

Public Property Let NeedleQuality(ByVal New_NeedleQuality As Integer)
    If New_NeedleQuality > 99 Or New_NeedleQuality < 0 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values between 0% and 99%.", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    m_NeedleQuality = New_NeedleQuality
    Timer1.interval = 100 - m_NeedleQuality
    PropertyChanged "NeedleQuality"
    End If
End Property
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Value() As Variant
Attribute Value.VB_Description = "Sets the value of the SpeedoMeter."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    m_Value = New_Value
    Select Case m_DirectNeedle
    Case True
    v.SetNeedleValueDirect m_Value
    Case False
    v.SetNeedleValue m_Value
    End Select
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get DirectNeedle() As Boolean
Attribute DirectNeedle.VB_Description = "Sets if the needle goes directly into the new selected value (for low CPU usage)."
Attribute DirectNeedle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DirectNeedle = m_DirectNeedle
End Property

Public Property Let DirectNeedle(ByVal New_DirectNeedle As Boolean)
    m_DirectNeedle = New_DirectNeedle
    PropertyChanged "DirectNeedle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,-30
Public Property Get AngleMax() As Integer
Attribute AngleMax.VB_Description = "Sets the position angle of the Max value."
Attribute AngleMax.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AngleMax = m_AngleMax
End Property

Public Property Let AngleMax(ByVal New_AngleMax As Integer)
    If New_AngleMax > m_AngleMin Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values smaller than AngleMin", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_AngleMax = New_AngleMax
    PropertyChanged "AngleMax"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,220
Public Property Get AngleMin() As Integer
Attribute AngleMin.VB_Description = "Sets the position angle of the Min value."
Attribute AngleMin.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AngleMin = m_AngleMin
End Property

Public Property Let AngleMin(ByVal New_AngleMin As Integer)
    If New_AngleMin < m_AngleMax Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than AngleMax", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_AngleMin = New_AngleMin
    PropertyChanged "AngleMin"
    End If
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,5
Public Property Get IntSmall() As Integer
Attribute IntSmall.VB_Description = "Sets the interval between small ticks."
Attribute IntSmall.VB_ProcData.VB_Invoke_Property = ";Appearance"
    IntSmall = m_IntSmall
End Property

Public Property Let IntSmall(ByVal New_IntSmall As Integer)
    If New_IntSmall <= 0 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than 0.", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_IntSmall = New_IntSmall
    PropertyChanged "IntSmall"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get IntBig() As Integer
Attribute IntBig.VB_Description = "Sets the interval between big ticks."
Attribute IntBig.VB_ProcData.VB_Invoke_Property = ";Appearance"
    IntBig = m_IntBig
End Property

Public Property Let IntBig(ByVal New_IntBig As Integer)
    If New_IntBig <= 0 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than 0.", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_IntBig = New_IntBig
    PropertyChanged "IntBig"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,4
Public Property Get NeedleStyle() As Needle_Style
Attribute NeedleStyle.VB_Description = "Sets the style of the needle."
Attribute NeedleStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NeedleStyle = m_NeedleStyle
End Property

Public Property Let NeedleStyle(ByRef New_NeedleStyle As Needle_Style)
    StopGauge
    m_NeedleStyle = New_NeedleStyle
    PropertyChanged "NeedleStyle"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get NeedleSize() As Needle_Size
Attribute NeedleSize.VB_Description = "Sets the size of the needle."
Attribute NeedleSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NeedleSize = m_NeedleSize
End Property

Public Property Let NeedleSize(ByVal New_NeedleSize As Needle_Size)
    StopGauge
    m_NeedleSize = New_NeedleSize
    PropertyChanged "NeedleSize"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get NeedleColor() As OLE_COLOR
Attribute NeedleColor.VB_Description = "Sets the color of the needle."
    NeedleColor = m_NeedleColor
End Property

Public Property Let NeedleColor(ByVal New_NeedleColor As OLE_COLOR)
    StopGauge
    m_NeedleColor = New_NeedleColor
    PropertyChanged "NeedleColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TickSmallColor() As OLE_COLOR
Attribute TickSmallColor.VB_Description = "Sets the color of small ticks."
Attribute TickSmallColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TickSmallColor = m_TickSmallColor
End Property

Public Property Let TickSmallColor(ByVal New_TickSmallColor As OLE_COLOR)
    StopGauge
    m_TickSmallColor = New_TickSmallColor
    PropertyChanged "TickSmallColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TickBigColor() As OLE_COLOR
Attribute TickBigColor.VB_Description = "Sets the color of big ticks."
Attribute TickBigColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TickBigColor = m_TickBigColor
End Property

Public Property Let TickBigColor(ByVal New_TickBigColor As OLE_COLOR)
    StopGauge
    m_TickBigColor = New_TickBigColor
    PropertyChanged "TickBigColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Sets the color of the SpeedoMeter's circle."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    StopGauge
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,3
Public Property Get BorderThickness() As Border_Thickness
Attribute BorderThickness.VB_Description = "Sets the thickness of the SpeedoMeter's circle."
Attribute BorderThickness.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderThickness = m_BorderThickness
End Property

Public Property Let BorderThickness(ByVal New_BorderThickness As Border_Thickness)
    StopGauge
    m_BorderThickness = New_BorderThickness
    PropertyChanged "BorderThickness"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets the background color."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    StopGauge
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Speed
Public Property Get TextLabel() As String
Attribute TextLabel.VB_Description = "Sets the caption of the text label."
Attribute TextLabel.VB_ProcData.VB_Invoke_Property = ";Text"
    TextLabel = m_TextLabel
End Property

Public Property Let TextLabel(ByVal New_TextLabel As String)
    StopGauge
    m_TextLabel = New_TextLabel
    PropertyChanged "TextLabel"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get TextFont() As Font
Attribute TextFont.VB_Description = "Sets the font of the text label."
Attribute TextFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set TextFont = m_TextFont
End Property

Public Property Set TextFont(ByVal New_TextFont As Font)
    StopGauge
    Set m_TextFont = New_TextFont
    PropertyChanged "TextFont"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0 km/hr
Public Property Get ValueLabel() As String
Attribute ValueLabel.VB_Description = "Sets the caption of the value label."
Attribute ValueLabel.VB_ProcData.VB_Invoke_Property = ";Text"
    ValueLabel = m_ValueLabel
End Property

Public Property Let ValueLabel(ByVal New_ValueLabel As String)
    StopGauge
    If Left(New_ValueLabel, 1) <> 0 Then
    m_ValueLabel = "0 " & New_ValueLabel
    Else
    m_ValueLabel = New_ValueLabel
    PropertyChanged "ValueLabel"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get ValueFont() As Font
Attribute ValueFont.VB_Description = "Sets the font of the value label."
Attribute ValueFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set ValueFont = m_ValueFont
End Property

Public Property Set ValueFont(ByVal New_ValueFont As Font)
    StopGauge
    Set m_ValueFont = New_ValueFont
    PropertyChanged "ValueFont"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ValueColor() As OLE_COLOR
Attribute ValueColor.VB_Description = "Sets the color of the value label."
Attribute ValueColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ValueColor = m_ValueColor
End Property

Public Property Let ValueColor(ByVal New_ValueColor As OLE_COLOR)
    StopGauge
    m_ValueColor = New_ValueColor
    PropertyChanged "ValueColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Sets the color of the text label."
Attribute TextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    StopGauge
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get ScaleFont() As Font
Attribute ScaleFont.VB_Description = "Sets the font of the scale label."
Attribute ScaleFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set ScaleFont = m_ScaleFont
End Property

Public Property Set ScaleFont(ByVal New_ScaleFont As Font)
    StopGauge
    Set m_ScaleFont = New_ScaleFont
    PropertyChanged "ScaleFont"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get TextPlace() As Text_Place
Attribute TextPlace.VB_Description = "Sets the location of the text caption."
Attribute TextPlace.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextPlace = m_TextPlace
End Property

Public Property Let TextPlace(ByVal New_TextPlace As Text_Place)
    StopGauge
    m_TextPlace = New_TextPlace
    Select Case m_TextPlace 'To define what to do with the value text
    Case 0 'Remove it. Just put its color equal to the background color
    TextMove
    m_TextColor = m_BackColor
    Case 1 'In the top - y = (- 1 / 3)
    'Adjust the color
    If m_TextColor = m_BackColor Then m_TextColor = m_def_ValueColor
    TextY = (-1 / 3)
    LabelMove
    Case 2 'In the bottom - y = (1 / 2)
    If m_TextColor = m_BackColor Then m_TextColor = m_def_ValueColor
    TextY = (1 / 2)
    LabelMove
    End Select
    PropertyChanged "TextPlace"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get ValuePlace() As Text_Place
Attribute ValuePlace.VB_Description = "Sets the position of the value label."
Attribute ValuePlace.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ValuePlace = m_ValuePlace
End Property

Public Property Let ValuePlace(ByVal New_ValuePlace As Text_Place)
    StopGauge
    m_ValuePlace = New_ValuePlace
    Select Case m_ValuePlace 'To define what to do with the value text
    Case 0 'Remove it. Just put its color equal to the background color
    LabelMove
    m_ValueColor = m_BackColor
    Case 1 'In the top - y = (- 1 / 3)
    'Adjust the color
    If m_ValueColor = m_BackColor Then m_ValueColor = m_def_ValueColor
    ValueY = (-1 / 3)
    TextMove
    Case 2 'In the bottom - y = (1 / 2)
    If m_ValueColor = m_BackColor Then m_ValueColor = m_def_ValueColor
    ValueY = (1 / 2)
    TextMove
    End Select
    PropertyChanged "ValuePlace"
    StartGauge
End Property

Private Sub LabelMove()
    If ValueY = TextY Then 'avoid distorting the text property
        If ValueY = (-1 / 3) Then
            ValueY = (1 / 2)
            m_ValuePlace = 2
        ElseIf ValueY = (1 / 2) Then
            ValueY = (-1 / 3)
            m_ValuePlace = 1
        End If
    End If
End Sub

Private Sub TextMove()
    If ValueY = TextY Then 'avoid distorting the text property
        If TextY = (-1 / 3) Then
            TextY = (1 / 2)
            m_TextPlace = 2
        ElseIf TextY = (1 / 2) Then
            TextY = (-1 / 3)
            m_TextPlace = 1
        End If
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Apperarance_
Attribute Appearance.VB_Description = "Sets the appearance of the UserControl."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Apperarance_)
    StopGauge
    m_Appearance = New_Appearance
    P1.Appearance = m_Appearance
    PropertyChanged "Appearance"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ControlBorderStyle() As BorderStyle_
Attribute ControlBorderStyle.VB_Description = "Sets the Border Style of the UserControl."
Attribute ControlBorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ControlBorderStyle = m_ControlBorderStyle
End Property

Public Property Let ControlBorderStyle(ByVal New_ControlBorderStyle As BorderStyle_)
    StopGauge
    m_ControlBorderStyle = New_ControlBorderStyle
    P1.borderstyle = m_ControlBorderStyle
    PropertyChanged "ControlBorderStyle"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,20
Public Property Get ScaleInterval() As Integer
Attribute ScaleInterval.VB_Description = "Sets the interval between the scale values in the SpeedoMeter."
Attribute ScaleInterval.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ScaleInterval = m_ScaleInterval
End Property

Public Property Let ScaleInterval(ByVal New_ScaleInterval As Integer)
    If New_ScaleInterval <= 0 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than 0", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_ScaleInterval = New_ScaleInterval
    PropertyChanged "ScaleInterval"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get NeedleMovement() As Needle_Movement
Attribute NeedleMovement.VB_Description = "Sets the style of movement of the needle."
Attribute NeedleMovement.VB_ProcData.VB_Invoke_Property = ";Behavior"
    NeedleMovement = m_NeedleMovement
End Property

Public Property Let NeedleMovement(ByVal New_NeedleMovement As Needle_Movement)
    StopGauge
    m_NeedleMovement = New_NeedleMovement
    PropertyChanged "NeedleMovement"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get SIZE() As Size_
Attribute SIZE.VB_Description = "Sets the size of the SpeedoMeter."
Attribute SIZE.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SIZE = m_Size
End Property

Public Property Let SIZE(ByVal New_Size As Size_)
    StopGauge
    m_Size = New_Size
    PropertyChanged "Size"
    UserControl_Resize
End Property

