VERSION 5.00
Begin VB.UserControl dbMagWin 
   AutoRedraw      =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   465
   Begin VB.Label Label1 
      Caption         =   "MW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "dbMagWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum SetToSide
    dbLeft = 0
    dbRight = 1
    dbTop = 2
    dbBottom = 3
End Enum
Public SideSet As Integer
'Default Property Values:
Const m_def_UnloadOnReturn = 1
Const m_def_MoveSpeed = 10
'Const m_def_FormAdhere = ""
'Const m_def_FormAdhere = 0
'Const m_def_FormAdhere = 0
Const m_def_FrontOrBack = 0
Const m_def_SetSide = 0
'Property Variables:
Dim m_UnloadOnReturn As Boolean
Dim m_MoveSpeed As Integer
'Dim m_FormAdhere As Variant
'Dim m_FormAdhere As Variant
'Dim m_FormAdhere As Variant
Dim m_FrontOrBack As Boolean
Dim m_SetSide As Variant


'
'
'
'
'Public Property Let SetSide(ByVal New_SetSide As SetToSide)
'    UserControl.SetSide() = New_SetSide
'    PropertyChanged "SetSide"
'End Property
'Public Property Get SetSide() As SetToSide
'
'End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SetSide() As SetToSide
    SetSide = m_SetSide
End Property

Public Property Let SetSide(ByVal New_SetSide As SetToSide)
    m_SetSide = New_SetSide
    PropertyChanged "SetSide"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SetSide = m_def_SetSide
    m_FrontOrBack = m_def_FrontOrBack
'    m_FormAdhere = m_def_FormAdhere
'    m_FormAdhere = m_def_FormAdhere
'    m_FormAdhere = m_def_FormAdhere
'    UserControl.Timer1.Enabled = False
    m_MoveSpeed = m_def_MoveSpeed
    m_UnloadOnReturn = m_def_UnloadOnReturn
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SetSide = PropBag.ReadProperty("SetSide", m_def_SetSide)
'    Timer1.Enabled = PropBag.ReadProperty("Enabled", True)
'    Timer1.Interval = PropBag.ReadProperty("Interval", 0)
    m_FrontOrBack = PropBag.ReadProperty("FrontOrBack", m_def_FrontOrBack)
'    m_FormAdhere = PropBag.ReadProperty("FormAdhere", m_def_FormAdhere)
'    m_FormAdhere = PropBag.ReadProperty("FormAdhere", m_def_FormAdhere)
'    m_FormAdhere = PropBag.ReadProperty("FormAdhere", m_def_FormAdhere)
    m_MoveSpeed = PropBag.ReadProperty("MoveSpeed", m_def_MoveSpeed)
    m_UnloadOnReturn = PropBag.ReadProperty("UnloadOnReturn", m_def_UnloadOnReturn)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SetSide", m_SetSide, m_def_SetSide)
'    Call PropBag.WriteProperty("Enabled", Timer1.Enabled, True)
'    Call PropBag.WriteProperty("Interval", Timer1.Interval, 0)
    Call PropBag.WriteProperty("FrontOrBack", m_FrontOrBack, m_def_FrontOrBack)
'    Call PropBag.WriteProperty("FormAdhere", m_FormAdhere, m_def_FormAdhere)
'    Call PropBag.WriteProperty("FormAdhere", m_FormAdhere, m_def_FormAdhere)
'    Call PropBag.WriteProperty("FormAdhere", m_FormAdhere, m_def_FormAdhere)
    Call PropBag.WriteProperty("MoveSpeed", m_MoveSpeed, m_def_MoveSpeed)
    Call PropBag.WriteProperty("UnloadOnReturn", m_UnloadOnReturn, m_def_UnloadOnReturn)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function LaunchMe(frmSet As Variant) As Variant
'Dim frmSet As Form
Load UserControl.Parent
UserControl.Parent.Left = frmSet.Left
UserControl.Parent.Top = frmSet.Top
UserControl.Parent.Show
    If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet.ZOrder
            End If
    If SetSide = dbRight Then
            If UserControl.Parent.Width < frmSet.Width Then
                UserControl.Parent.Left = frmSet.Left + (frmSet.Width - UserControl.Parent.Width)
            Else
                UserControl.Parent.Left = frmSet.Left
            End If
            If UserControl.Parent.Height < frmSet.Height Then
                UserControl.Parent.Top = frmSet.Top + ((frmSet.Height - UserControl.Parent.Height) / 2)
            Else
                UserControl.Parent.Top = frmSet.Top
            End If
            
            Static I%
        For I = 0 To UserControl.Parent.Width Step MoveSpeed
            If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet.ZOrder
            End If
            UserControl.Parent.Left = UserControl.Parent.Left + MoveSpeed
        Next I%
            UserControl.Parent.Left = frmSet.Left + frmSet.Width
        End If
        
        If SetSide = dbLeft Then
            
                UserControl.Parent.Left = frmSet.Left
            
            If UserControl.Parent.Height < frmSet.Height Then
                UserControl.Parent.Top = frmSet.Top + ((frmSet.Height - UserControl.Parent.Height) / 2)
            Else
                UserControl.Parent.Top = frmSet.Top
            End If
            
            Static E%
        For E = 0 To UserControl.Parent.Width Step MoveSpeed
            If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet.ZOrder
            End If
            UserControl.Parent.Left = UserControl.Parent.Left - MoveSpeed
        Next E%
            UserControl.Parent.Left = frmSet.Left - UserControl.Parent.Width
        End If
    
    
    If SetSide = dbTop Then
            If UserControl.Parent.Width < frmSet.Width Then
                UserControl.Parent.Left = frmSet.Left + ((frmSet.Width - UserControl.Parent.Width) / 2)
            Else
                UserControl.Parent.Left = frmSet.Left
            End If
            If UserControl.Parent.Height < frmSet.Height Then
                UserControl.Parent.Top = frmSet.Top + ((frmSet.Height - UserControl.Parent.Height) / 2)
            Else
                UserControl.Parent.Top = frmSet.Top
            End If
            
            Static T%
        For T = 0 To UserControl.Parent.Height Step MoveSpeed
            If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet.ZOrder
            End If
            UserControl.Parent.Top = UserControl.Parent.Top - MoveSpeed
            Dim yup As Integer
            yup = UserControl.Parent.Top
        Next T%
            UserControl.Parent.Top = frmSet.Top - UserControl.Parent.Height
        End If
        
        
        If SetSide = dbBottom Then
            If UserControl.Parent.Width < frmSet.Width Then
                UserControl.Parent.Left = frmSet.Left + ((frmSet.Width - UserControl.Parent.Width) / 2)
            Else
                UserControl.Parent.Left = frmSet.Left
            End If
            If UserControl.Parent.Height < frmSet.Height Then
                UserControl.Parent.Top = frmSet.Top + ((frmSet.Height - UserControl.Parent.Height) / 2)
            Else
                UserControl.Parent.Top = frmSet.Top
            End If
            
            Static B%
        For B = 0 To UserControl.Parent.Height Step MoveSpeed
            If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet.ZOrder
            End If
            UserControl.Parent.Top = UserControl.Parent.Top + MoveSpeed
            
        Next B%
            UserControl.Parent.Top = frmSet.Top + frmSet.Height
        End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ReturnMe(frmSet2 As Variant) As Variant
    If FrontOrBack = True Then
                UserControl.Parent.ZOrder
            Else
                frmSet2.ZOrder
            End If
    If SetSide = dbRight Then
            
            
            Static I%
        For I = 0 To UserControl.Parent.Width Step MoveSpeed
            
                frmSet2.ZOrder
            
            UserControl.Parent.Left = UserControl.Parent.Left - MoveSpeed
        Next I%
            If UnloadOnReturn = True Then
                Unload UserControl.Parent
            Else
                UserControl.Parent.Left = frmSet2.Left
            End If
        End If
        
        If SetSide = dbLeft Then
                       
            Static E%
        For E = 0 To UserControl.Parent.Width Step MoveSpeed
            
                frmSet2.ZOrder
            
            UserControl.Parent.Left = UserControl.Parent.Left + MoveSpeed
        Next E%
            If UnloadOnReturn = True Then
                Unload UserControl.Parent
            Else
                UserControl.Parent.Left = frmSet2.Left
            End If
        End If
    
    
    If SetSide = dbTop Then
            
            Static T%
        For T = 0 To UserControl.Parent.Height Step MoveSpeed
            
                frmSet2.ZOrder
            
            UserControl.Parent.Top = UserControl.Parent.Top + MoveSpeed
            
        Next T%
            If UnloadOnReturn = True Then
                Unload UserControl.Parent
            Else
                UserControl.Parent.Top = frmSet2.Top
            End If
        End If
        
        
        If SetSide = dbBottom Then
            
            Static B%
        For B = 0 To UserControl.Parent.Height Step MoveSpeed
            
                frmSet2.ZOrder
            
            UserControl.Parent.Top = UserControl.Parent.Top - MoveSpeed
            
        Next B%
            If UnloadOnReturn = True Then
                Unload UserControl.Parent
            Else
                UserControl.Parent.Top = frmSet2.Top
            End If
        End If
End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Timer1,Timer1,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = Timer1.Enabled
'    If UserControl.Timer1.Enabled = True Then
'        Call SlideForm
'    End If
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    Timer1.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'    If UserControl.Timer1.Enabled = True Then
'        Call SlideForm
'    End If
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Timer1,Timer1,-1,Interval
'Public Property Get Interval() As Long
'    Interval = Timer1.Interval
'End Property
'
'Public Property Let Interval(ByVal New_Interval As Long)
'    Timer1.Interval() = New_Interval
'    PropertyChanged "Interval"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FrontOrBack() As Boolean
Attribute FrontOrBack.VB_Description = "Set to True for form to load in front, or False to load in back."
    FrontOrBack = m_FrontOrBack
End Property

Public Property Let FrontOrBack(ByVal New_FrontOrBack As Boolean)
    m_FrontOrBack = New_FrontOrBack
    PropertyChanged "FrontOrBack"
End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=14,0,0,0
'''Public Property Get FormAdhere() As Variant
'''    FormAdhere = m_FormAdhere
'''End Property
'''
'''Public Property Let FormAdhere(ByVal New_FormAdhere As Variant)
'''    m_FormAdhere = New_FormAdhere
'''    PropertyChanged "FormAdhere"
'''End Property
'''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,1,0,0
''Public Property Get FormAdhere() As Variant
''    FormAdhere = m_FormAdhere
''End Property
''
''Public Property Let FormAdhere(ByVal New_FormAdhere As Variant)
''    If Ambient.UserMode Then Err.Raise 382
''    m_FormAdhere = New_FormAdhere
''    PropertyChanged "FormAdhere"
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,
'Public Property Get FormAdhere() As Variant
'    FormAdhere = m_FormAdhere
'End Property
'
'Public Property Let FormAdhere(ByVal New_FormAdhere As Variant)
'    m_FormAdhere = New_FormAdhere
'    PropertyChanged "FormAdhere"
'End Property
'Public Function SlideForm()
'
'
'
'End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get MoveSpeed() As Integer
Attribute MoveSpeed.VB_Description = "Smaller Numbers For Slower Speed, Larger Numbers For Faster Speed. Recommend between 5 and 100 for best results."
    MoveSpeed = m_MoveSpeed
End Property

Public Property Let MoveSpeed(ByVal New_MoveSpeed As Integer)
    m_MoveSpeed = New_MoveSpeed
    PropertyChanged "MoveSpeed"
    If MoveSpeed < 1 Then
        MoveSpeed = 1
    End If
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get UnloadAtClose() As Boolean
'    UnloadAtClose = m_UnloadAtClose
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get UnloadOnReturn() As Boolean
Attribute UnloadOnReturn.VB_Description = "Unloads Magic Window form when ReturnMe function is called"
    UnloadOnReturn = m_UnloadOnReturn
End Property

Public Property Let UnloadOnReturn(ByVal New_UnloadOnReturn As Boolean)
    m_UnloadOnReturn = New_UnloadOnReturn
    PropertyChanged "UnloadOnReturn"
End Property

