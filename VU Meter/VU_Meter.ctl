VERSION 5.00
Begin VB.UserControl VUMeter 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   ScaleHeight     =   3285
   ScaleWidth      =   4425
   ToolboxBitmap   =   "VU_Meter.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   3360
      Top             =   1320
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1
      ScaleLeft       =   24
      ScaleMode       =   0  'User
      ScaleWidth      =   52
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "VUMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'VU Meter ActiveX Control
'Class Module made by Alain van Hanegem
'UserControl made by Andre Felipe N. D
'Planet Source Code, 2006

'Description: This UserControl was made to try to made the set up of the VU Meter
'             faster and easier. When I first got the Class Module from PSC, I thought
'             it was a great control, but difficult to configurate, so I made this
'             UserControl to make things easier. If you are an advanced user and didn't
'             have problems with the class module, I suggest you keep using it, because
'             it has much more potential and use less of your PC resources . But if you
'             are a begginer, I hope this UserControl can be useful for you. I tried to
'             put in it almost all the possible properties of the class module.

Option Explicit
Private v As clsVUmeter

'Property Variables:
Dim m_DirectNeedle As Boolean
Dim m_Value As Double
Dim m_Quality As Integer
Dim m_XPoint As Integer
Dim m_Appearance As Integer
Dim m_BorderStyle As Integer

'Default Property Values:
Const m_def_Quality = 50
Const m_def_Value = 0
Const m_def_DirectNeedle = False
Const m_def_ControlSize = 1
Const m_def_XPoint = 50
Const m_def_MaxPosAngle = 45
Const m_def_MinPosAngle = 135
Const m_def_FadeEndColor = &HFF& 'vm
Const m_def_FadeStartColor = &HFFFFFF 'br
Const m_def_FadeStart = 70
Const m_def_FadeEnd = 100
Const m_def_TicksIntervalBig = 10
Const m_def_TicksIntervalSmall = 5
Const m_def_TickColorBig = &HFF0000 'az
Const m_def_TickColorSmall = &HFF0000 'az
Const m_def_NeedleStyle = 1
Const m_def_Appearance = 0
Const m_def_BorderStyle = 0
Const m_def_MinValue = 0
Const m_def_MaxValue = 100
Const m_def_TextLabel = "Label"
Const m_def_ValueLabel = "0 Value "
Const m_def_BackColor = &HFFFFFF 'br
Const m_def_NeedleColor = 0
Const m_def_ValueColor = 0
Const m_def_TextColor = 0
Const m_def_Screw = True
Const m_def_NeedleSize = 3

'Enums
Public Enum Appearance_
Flat = 0
Tridimensional = 1
End Enum

Public Enum Border_
None = 0
Fixed_Single = 1
End Enum

Public Enum Needle_Style
Very_Thin = 0
Simple = 1
Stretched = 2
Big_Bottom = 3
Normal = 4
Large = 5
End Enum

Public Enum Needle_Size
No_Needle = 0
Short = 1
Medium = 2
Big = 3
Super_Big = 4
End Enum

Public Enum Size_
Small = 1
Medium = 2
Big = 3
Super_Big = 4
End Enum

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TextLabel = m_def_TextLabel
    m_ValueLabel = m_def_ValueLabel
    m_Quality = m_def_Quality
    m_Value = m_def_Value
    m_MinValue = m_def_MinValue
    m_MaxValue = m_def_MaxValue
    m_DirectNeedle = m_def_DirectNeedle
    m_MinValue = m_def_MinValue
    m_MaxValue = m_def_MaxValue
    m_MaxValue = m_def_MaxValue
    m_BackColor = m_def_BackColor
    m_Screw = m_def_Screw
    m_BorderStyle = m_def_BorderStyle
    m_ValueColor = m_def_ValueColor
    m_TextColor = m_def_TextColor
    m_Appearance = m_def_Appearance
    m_NeedleColor = m_def_NeedleColor
    m_NeedleSize = m_def_NeedleSize
    m_NeedleStyle = m_def_NeedleStyle
    m_TickColorSmall = m_def_TickColorSmall
    m_TickColorBig = m_def_TickColorBig
    Set m_TextFont = Ambient.Font
    Set m_ValueFont = Ambient.Font
    m_TicksIntervalBig = m_def_TicksIntervalBig
    m_TicksIntervalSmall = m_def_TicksIntervalSmall
    m_FadeEndColor = m_def_FadeEndColor
    m_FadeStartColor = m_def_FadeStartColor
    m_FadeStart = m_def_FadeStart
    m_FadeEnd = m_def_FadeEnd
    m_MaxPosAngle = m_def_MaxPosAngle
    m_MinPosAngle = m_def_MinPosAngle
    m_XPoint = m_def_XPoint
    m_ControlSize = m_def_ControlSize
    NeedleRef = 0.25
    StartGauge
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_TextLabel = PropBag.ReadProperty("TextLabel", m_def_TextLabel)
    m_ValueLabel = PropBag.ReadProperty("ValueLabel", m_def_ValueLabel)
    m_Quality = PropBag.ReadProperty("Quality", m_def_Quality)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_DirectNeedle = PropBag.ReadProperty("DirectNeedle", m_def_DirectNeedle)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Screw = PropBag.ReadProperty("Screw", m_def_Screw)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ValueColor = PropBag.ReadProperty("ValueColor", m_def_ValueColor)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_NeedleColor = PropBag.ReadProperty("NeedleColor", m_def_NeedleColor)
    m_NeedleSize = PropBag.ReadProperty("NeedleSize", m_def_NeedleSize)
    m_NeedleStyle = PropBag.ReadProperty("NeedleStyle", m_def_NeedleStyle)
    m_TickColorSmall = PropBag.ReadProperty("TickColorSmall", m_def_TickColorSmall)
    m_TickColorBig = PropBag.ReadProperty("TickColorBig", m_def_TickColorBig)
    Set m_TextFont = PropBag.ReadProperty("TextFont", Ambient.Font)
    Set m_ValueFont = PropBag.ReadProperty("ValueFont", Ambient.Font)
    m_TicksIntervalBig = PropBag.ReadProperty("TicksIntervalBig", m_def_TicksIntervalBig)
    m_TicksIntervalSmall = PropBag.ReadProperty("TicksIntervalSmall", m_def_TicksIntervalSmall)
    m_FadeEndColor = PropBag.ReadProperty("FadeEndColor", m_def_FadeEndColor)
    m_FadeStartColor = PropBag.ReadProperty("FadeStartColor", m_def_FadeStartColor)
    m_FadeStart = PropBag.ReadProperty("FadeStart", m_def_FadeStart)
    m_FadeEnd = PropBag.ReadProperty("FadeEnd", m_def_FadeEnd)
    m_MaxPosAngle = PropBag.ReadProperty("MaxPosAngle", m_def_MaxPosAngle)
    m_MinPosAngle = PropBag.ReadProperty("MinPosAngle", m_def_MinPosAngle)
    m_XPoint = PropBag.ReadProperty("XPoint", m_def_XPoint)
    m_ControlSize = PropBag.ReadProperty("ControlSize", m_def_ControlSize)
    NeedleRef = 0.25
    StartGauge
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TextLabel", m_TextLabel, m_def_TextLabel)
    Call PropBag.WriteProperty("ValueLabel", m_ValueLabel, m_def_ValueLabel)
    Call PropBag.WriteProperty("Quality", m_Quality, m_def_Quality)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("DirectNeedle", m_DirectNeedle, m_def_DirectNeedle)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Screw", m_Screw, m_def_Screw)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ValueColor", m_ValueColor, m_def_ValueColor)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("NeedleColor", m_NeedleColor, m_def_NeedleColor)
    Call PropBag.WriteProperty("NeedleSize", m_NeedleSize, m_def_NeedleSize)
    Call PropBag.WriteProperty("NeedleStyle", m_NeedleStyle, m_def_NeedleStyle)
    Call PropBag.WriteProperty("TickColorSmall", m_TickColorSmall, m_def_TickColorSmall)
    Call PropBag.WriteProperty("TickColorBig", m_TickColorBig, m_def_TickColorBig)
    Call PropBag.WriteProperty("TextFont", m_TextFont, Ambient.Font)
    Call PropBag.WriteProperty("ValueFont", m_ValueFont, Ambient.Font)
    Call PropBag.WriteProperty("TicksIntervalBig", m_TicksIntervalBig, m_def_TicksIntervalBig)
    Call PropBag.WriteProperty("TicksIntervalSmall", m_TicksIntervalSmall, m_def_TicksIntervalSmall)
    Call PropBag.WriteProperty("FadeEndColor", m_FadeEndColor, m_def_FadeEndColor)
    Call PropBag.WriteProperty("FadeStartColor", m_FadeStartColor, m_def_FadeStartColor)
    Call PropBag.WriteProperty("FadeStart", m_FadeStart, m_def_FadeStart)
    Call PropBag.WriteProperty("FadeEnd", m_FadeEnd, m_def_FadeEnd)
    Call PropBag.WriteProperty("MaxPosAngle", m_MaxPosAngle, m_def_MaxPosAngle)
    Call PropBag.WriteProperty("MinPosAngle", m_MinPosAngle, m_def_MinPosAngle)
    Call PropBag.WriteProperty("XPoint", m_XPoint, m_def_XPoint)
    Call PropBag.WriteProperty("ControlSize", m_ControlSize, m_def_ControlSize)
End Sub

Private Sub StartGauge()
Set v = New clsVUmeter
'Desenhar o gauge
Dim cx As Long
Dim cy As Long
cx = ((P1.Width * m_XPoint) / 1500)
cy = P1.Height / 18
v.Init_Picture P1.hDC, P1.Image, UserControl.Width, UserControl.Height
v.SetVUDefaults cx, cy
v.Draw
P1.Refresh
Timer1.enabled = True
End Sub

Private Sub StopGauge()
Timer1.enabled = False
Set v = Nothing
End Sub

Private Sub UserControl_Resize()
StopGauge
Select Case m_ControlSize
Case 1
UserControl.Height = 1455
NeedleRef = 0.25
Case 2
UserControl.Height = 1815
NeedleRef = 0.175
Case 3
UserControl.Height = 2055
NeedleRef = 0.13
Case 4
UserControl.Height = 2415
NeedleRef = 0.125
End Select
'Alinhar o tamanho de P1 com o usercontrol
P1.Height = UserControl.Height
P1.Width = UserControl.Width
StartGauge
End Sub

Private Sub Timer1_Timer()
    v.AnimationLoop
    v.Draw
    P1.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
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
'MemberInfo=7,0,0,50
Public Property Get Quality() As Integer
Attribute Quality.VB_Description = "Sets the quality of movement of the needle. Higher values means better quality, but also higher CPU usage."
Attribute Quality.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Quality = m_Quality
End Property

Public Property Let Quality(ByVal New_Quality As Integer)
    If New_Quality > 99 Or New_Quality < 0 Then 'Fora da faixa de valores
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Select a range between 0 (lowest quality) and 99 (highest quality).", vbOKOnly, "Error"
    End Select
    Exit Property
    Else
    m_Quality = New_Quality
    Timer1.interval = 100 - m_Quality
    PropertyChanged "Quality"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ValueLabel() As String
Attribute ValueLabel.VB_Description = "Sets the caption of the value label."
Attribute ValueLabel.VB_ProcData.VB_Invoke_Property = ";Text"
    ValueLabel = m_ValueLabel
End Property

Public Property Let ValueLabel(ByVal New_ValueLabel As String)
    StopGauge
    m_ValueLabel = New_ValueLabel
    If Left(m_ValueLabel, 1) <> "0" Then m_ValueLabel = "0 " + m_ValueLabel
    PropertyChanged "ValueLabel"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Sets the value of the needle"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    Select Case m_DirectNeedle
    Case True 'N ter animação
    DoEvents
    v.SetNeedleValueDirect m_Value
    Case False 'Ter animação
    PropertyChanged "Value"
    DoEvents
    v.SetNeedleValue m_Value
    End Select
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get DirectNeedle() As Boolean
Attribute DirectNeedle.VB_Description = "Sets if the needle goes directly to the selected value (for small CPU usage)."
Attribute DirectNeedle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DirectNeedle = m_DirectNeedle
End Property

Public Property Let DirectNeedle(ByVal New_DirectNeedle As Boolean)
    m_DirectNeedle = New_DirectNeedle
    PropertyChanged "DirectNeedle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get MinValue() As Integer
Attribute MinValue.VB_Description = "Sets the lowest value of the gauge."
Attribute MinValue.VB_ProcData.VB_Invoke_Property = ";Scale"
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Integer)
    StopGauge
    If New_MinValue > m_MaxValue Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Values higher than MaxValue are not allowed"
    End Select
    Exit Property
    End If
    If New_MinValue > m_FadeStart Then
        Dim Difference As Integer
        Difference = m_FadeEnd - m_FadeStart
        m_FadeStart = New_MinValue
        If (m_FadeStart + Difference) > MaxValue Then
            m_FadeEnd = MaxValue
        Else
            m_FadeEnd = m_FadeStart + Difference
        End If
    End If
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get MaxValue() As Integer
Attribute MaxValue.VB_Description = "Sets the highest value of the meter."
Attribute MaxValue.VB_ProcData.VB_Invoke_Property = ";Scale"
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Integer)
    StopGauge
    If New_MaxValue < m_MinValue Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Values smaller than MinValue are not allowed"
    End Select
    Exit Property
    End If
        If New_MaxValue < m_FadeEnd Then
        Dim Difference As Integer
        Difference = m_FadeEnd - m_FadeStart
        m_FadeEnd = New_MaxValue
        If (m_FadeEnd - Difference) < MinValue Then
            m_FadeStart = MinValue
        Else
            m_FadeStart = m_FadeEnd - Difference
        End If
    End If
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
    StartGauge
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
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
'MemberInfo=0,0,0,True
Public Property Get Screw() As Boolean
Attribute Screw.VB_Description = "If sets to true, the meter has the screws. If sets to false, the screws are removed."
Attribute Screw.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Screw = m_Screw
End Property

Public Property Let Screw(ByVal New_Screw As Boolean)
    StopGauge
    m_Screw = New_Screw
    PropertyChanged "Screw"
    StartGauge
End Property
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Border_
Attribute BorderStyle.VB_Description = "Sets the style of the border."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Border_)
    m_BorderStyle = New_BorderStyle
    P1.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
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
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Appearance_
Attribute Appearance.VB_Description = "Sets the appearance of the control."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearance_)
    StopGauge
    m_Appearance = New_Appearance
    P1.Appearance = m_Appearance
    PropertyChanged "Appearance"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get NeedleColor() As OLE_COLOR
Attribute NeedleColor.VB_Description = "Sets the color of the needle."
Attribute NeedleColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NeedleColor = m_NeedleColor
End Property

Public Property Let NeedleColor(ByVal New_NeedleColor As OLE_COLOR)
    StopGauge
    m_NeedleColor = New_NeedleColor
    PropertyChanged "NeedleColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
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

Public Property Get NeedleStyle() As Needle_Style
Attribute NeedleStyle.VB_Description = "Sets the style of the needle."
Attribute NeedleStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NeedleStyle = m_NeedleStyle
End Property

Public Property Let NeedleStyle(ByVal New_NeedleStyle As Needle_Style)
    StopGauge
    m_NeedleStyle = New_NeedleStyle
    PropertyChanged "NeedleStyle"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TickColorSmall() As OLE_COLOR
Attribute TickColorSmall.VB_Description = "Sets the color of the small tick"
Attribute TickColorSmall.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TickColorSmall = m_TickColorSmall
End Property

Public Property Let TickColorSmall(ByVal New_TickColorSmall As OLE_COLOR)
    StopGauge
    m_TickColorSmall = New_TickColorSmall
    PropertyChanged "TickColorSmall"
    StartGauge
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TickColorBig() As OLE_COLOR
Attribute TickColorBig.VB_Description = "Sets the color of the big tick."
Attribute TickColorBig.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TickColorBig = m_TickColorBig
End Property

Public Property Let TickColorBig(ByVal New_TickColorBig As OLE_COLOR)
    StopGauge
    m_TickColorBig = New_TickColorBig
    PropertyChanged "TickColorBig"
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
'MemberInfo=7,0,0,10
Public Property Get TicksIntervalBig() As Integer
Attribute TicksIntervalBig.VB_Description = "Sets the interval of distance between the big ticks."
Attribute TicksIntervalBig.VB_ProcData.VB_Invoke_Property = ";Scale"
    TicksIntervalBig = m_TicksIntervalBig
End Property

Public Property Let TicksIntervalBig(ByVal New_TicksIntervalBig As Integer)
    If New_TicksIntervalBig <= 0 Then 'not permitted value
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "It is not permitted values equal or lower than 0 in this property", vbOKOnly + vbCritical, "Erro"
    End Select
    Exit Property
    Else
    StopGauge
    m_TicksIntervalBig = New_TicksIntervalBig
    PropertyChanged "TicksIntervalBig"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,5
Public Property Get TicksIntervalSmall() As Integer
Attribute TicksIntervalSmall.VB_Description = "Sets the interval of distance between the small ticks."
Attribute TicksIntervalSmall.VB_ProcData.VB_Invoke_Property = ";Scale"
    TicksIntervalSmall = m_TicksIntervalSmall
End Property

Public Property Let TicksIntervalSmall(ByVal New_TicksIntervalSmall As Integer)
    If New_TicksIntervalSmall <= 0 Then 'not permitted value
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "It is not permitted values equal or lower than 0 in this property", vbOKOnly + vbCritical, "Erro"
    End Select
    Exit Property
    Else
    StopGauge
    m_TicksIntervalSmall = New_TicksIntervalSmall
    PropertyChanged "TicksIntervalSmall"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FadeEndColor() As OLE_COLOR
Attribute FadeEndColor.VB_Description = "Sets the second color of the fade."
Attribute FadeEndColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeEndColor = m_FadeEndColor
End Property

Public Property Let FadeEndColor(ByVal New_FadeEndColor As OLE_COLOR)
    StopGauge
    m_FadeEndColor = New_FadeEndColor
    PropertyChanged "FadeEndColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FadeStartColor() As OLE_COLOR
Attribute FadeStartColor.VB_Description = "Sets the first color of the fade."
Attribute FadeStartColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeStartColor = m_FadeStartColor
End Property

Public Property Let FadeStartColor(ByVal New_FadeStartColor As OLE_COLOR)
    StopGauge
    m_FadeStartColor = New_FadeStartColor
    PropertyChanged "FadeStartColor"
    StartGauge
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,70
Public Property Get FadeStart() As Integer
Attribute FadeStart.VB_Description = "Sets the starting position of the fade."
Attribute FadeStart.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeStart = m_FadeStart
End Property

Public Property Let FadeStart(ByVal New_FadeStart As Integer)
    If New_FadeStart >= m_FadeEnd Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Select a value lower than FadeEnd", vbOKOnly + vbCritical, "Error"
    Exit Property
    End Select
    ElseIf New_FadeStart < m_MinValue Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This value can't be smaller than MinValue", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_FadeStart = New_FadeStart
    PropertyChanged "FadeStart"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get FadeEnd() As Integer
Attribute FadeEnd.VB_Description = "Sets the final position of the fade."
Attribute FadeEnd.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeEnd = m_FadeEnd
End Property

Public Property Let FadeEnd(ByVal New_FadeEnd As Integer)
    If New_FadeEnd <= m_FadeStart Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Select a value higher than FadeStart", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    ElseIf New_FadeEnd > m_MaxValue Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This value can't be higher than MaxValue", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_FadeEnd = New_FadeEnd
    PropertyChanged "FadeEnd"
    StartGauge
    End If
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MaxPosAngle() As Integer
Attribute MaxPosAngle.VB_Description = "Sets the angular position of the MaxValue."
Attribute MaxPosAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MaxPosAngle = m_MaxPosAngle
End Property

Public Property Let MaxPosAngle(ByVal New_MaxPosAngle As Integer)
    If New_MaxPosAngle < 15 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values higher than 15", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    ElseIf New_MaxPosAngle > m_MinPosAngle Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Choose a value smaller than MinPosAngle", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_MaxPosAngle = New_MaxPosAngle
    PropertyChanged "MaxPosAngle"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MinPosAngle() As Integer
Attribute MinPosAngle.VB_Description = "Sets the angular position of the MinValue."
Attribute MinPosAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MinPosAngle = m_MinPosAngle
End Property

Public Property Let MinPosAngle(ByVal New_MinPosAngle As Integer)
    If New_MinPosAngle > 165 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "This property only accepts values lower than 165", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    ElseIf New_MinPosAngle < m_MaxPosAngle Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Choose a value higher than MaxPosAngle", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_MinPosAngle = New_MinPosAngle
    PropertyChanged "MinPosAngle"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,50
Public Property Get XPoint() As Integer
Attribute XPoint.VB_Description = "Sets (in porcentage) the horizontal position of the gauge. If the value is 100, the meter will be completely to the right. If the value is 0, The meter will be completely to the left. "
Attribute XPoint.VB_ProcData.VB_Invoke_Property = ";Appearance"
    XPoint = m_XPoint
End Property

Public Property Let XPoint(ByVal New_XPoint As Integer)
    If New_XPoint > 120 Or New_XPoint < -20 Then
    Select Case Ambient.UserMode
    Case True
    Err.Raise 1764 '"The requested operation is not supported."
    Case False
    MsgBox "Select values between -20% And 120%", vbOKOnly + vbCritical, "Error"
    End Select
    Exit Property
    Else
    StopGauge
    m_XPoint = New_XPoint
    PropertyChanged "XPoint"
    StartGauge
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get ControlSize() As Size_
Attribute ControlSize.VB_Description = "Sets the size of the VU Meter."
Attribute ControlSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ControlSize = m_ControlSize
End Property

Public Property Let ControlSize(ByVal New_ControlSize As Size_)
    StopGauge
    m_ControlSize = New_ControlSize
    PropertyChanged "ControlSize"
    UserControl_Resize
End Property

