VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnimationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PreviousObject As AnimationObject, m_NextObject As AnimationObject
Private m_Released As Boolean

Private m_ObjectID As Long
Attribute m_ObjectID.VB_VarDescription = "Object's personal ID. Must be a unique value."
Public ObjectTypeID As Long
Attribute ObjectTypeID.VB_VarDescription = "Object type indentifier."

Private m_LayerPriority As Long

Public Visible As Boolean
Attribute Visible.VB_VarDescription = "Indicates whether or not this object should be processed on calls to AnimationWindow.Update."

Private dx_Surface As DirectDrawSurface7
Private dx_SurfaceStaticNum As Long

Private m_SurfaceWidth As Long, m_SurfaceHeight As Long
Private m_ImageWidth As Long, m_ImageHeight As Long
Attribute m_ImageWidth.VB_VarDescription = "The object's image width."
Private m_ImagesPerRow As Long

Private m_ActionSequence As Variant, m_ActionFirst As Long, m_ActionLast As Long
Public ActionFrame As Long, SourceOffsetX As Long, SourceOffsetY As Long
Attribute ActionFrame.VB_VarDescription = "Current frame of action sequence. If this value is -1 then SourceOffsetX and SourceOffsetY are used instead of calculating from the object's action sequence values."
Attribute SourceOffsetX.VB_VarDescription = "X offset in source image surface to find the image. Only used if the ActionFrame value is set to -1, otherwise the normal action sequence is used to calculate where the source image is."
Attribute SourceOffsetY.VB_VarDescription = "Y offset in source image surface to find the image. Only used if the ActionFrame value is set to -1, otherwise the normal action sequence is used to calculate where the source image is."

Public SpecialFX As BlitterFX
Public FXScaleWidth As Double, FXScaleHeight As Double
Public FXTargetColour, FXAlphaBlendLevel As Long, FXRotationAngle As Long

Public PosX_1000ths As Long, PosY_1000ths As Long
Attribute PosX_1000ths.VB_VarDescription = "Object's left point in the animation window (in 1000ths of pixels)."
Attribute PosY_1000ths.VB_VarDescription = "Object's top point in the animation window (in 1000ths of pixels)."
Public VelocityX_1000ths As Long, VelocityY_1000ths As Long
Attribute VelocityX_1000ths.VB_VarDescription = "Object's X velocity."
Public AccelerationX_1000ths As Long, AccelerationY_1000ths As Long
Attribute AccelerationX_1000ths.VB_VarDescription = "Object's X acceleration."

Public CollisionMaskMe As Long, CollisionMaskTarget As Long
Attribute CollisionMaskMe.VB_VarDescription = "Object's collision mask - logically and'ed with other object's CollisionMaskTarget to determine if objects can interact."
Attribute CollisionMaskTarget.VB_VarDescription = "Object's target collision mask - logically and'ed with other object's CollisionMaskMe to determine if objects can interact."
Public CollisionBoxLeft As Long, CollsionBoxRight As Long, CollisionBoxTop As Long, CollisionBoxBottom As Long
Attribute CollisionBoxLeft.VB_VarDescription = "Offset from object center to left of collision box."
Attribute CollsionBoxRight.VB_VarDescription = "Offset from object center to right of collision box."
Attribute CollisionBoxTop.VB_VarDescription = "Offset from object center to top of collision box."
Attribute CollisionBoxBottom.VB_VarDescription = "Offset from object center to bottom of collision box."

Public UserLong1 As Long
Public UserLong2 As Long
Public UserLong3 As Long
Public UserLong4 As Long
Public UserLong5 As Long
Public UserLong6 As Long
Public UserLong7 As Long
Public UserLong8 As Long
Public UserLong9 As Long
Public UserLong0 As Long

Public UserVariant1 As Variant
Public UserVariant2 As Variant
Public UserVariant3 As Variant
Public UserVariant4 As Variant
Public UserVariant5 As Variant

Friend Property Get SurfaceStaticNum() As Long
  SurfaceStaticNum = dx_SurfaceStaticNum
End Property

Friend Property Let SurfaceStaticNum(surfaceNum As Long)
  dx_SurfaceStaticNum = surfaceNum
End Property

Public Property Get Released() As Boolean
Attribute Released.VB_Description = "If true then this object has been removed from animation list and all references to it should be released (set to Nothing)."
  Released = m_Released
End Property

Friend Sub Release()
  m_Released = True
End Sub

Friend Property Set DXSurface(nSurface As DirectDrawSurface7)
  Set dx_Surface = nSurface
End Property

Friend Property Get DXSurface() As DirectDrawSurface7
  Set DXSurface = dx_Surface
End Property

Public Property Get SurfaceWidth() As Long
  SurfaceWidth = m_SurfaceWidth
End Property

Friend Property Let SurfaceWidth(nWidth As Long)
  m_SurfaceWidth = nWidth
End Property

Public Property Get SurfaceHeight() As Long
  SurfaceHeight = m_SurfaceHeight
End Property

Friend Property Let SurfaceHeight(nHeight As Long)
  m_SurfaceHeight = nHeight
End Property

Public Property Get ImageWidth() As Long
  ImageWidth = m_ImageWidth
End Property

Friend Property Let ImageWidth(nWidth As Long)
  m_ImageWidth = nWidth
End Property

Public Property Get ImageHeight() As Long
  ImageHeight = m_ImageHeight
End Property

Friend Property Let ImageHeight(nHeight As Long)
  m_ImageHeight = nHeight
End Property

Public Property Get ImagesPerRow() As Long
  ImagesPerRow = m_ImagesPerRow
End Property

Friend Property Let ImagesPerRow(nIPRow As Long)
  m_ImagesPerRow = nIPRow
End Property

Public Property Get ActionSequenceStart() As Long
Attribute ActionSequenceStart.VB_Description = "Lower bound of the ActionSequence array. (Read Only)"
  ActionSequenceStart = m_ActionFirst
End Property

Public Property Get ActionSequenceStop() As Long
Attribute ActionSequenceStop.VB_Description = "Upperr bound of the ActionSequence array. (Read Only)"
  ActionSequenceStop = m_ActionLast
End Property

Public Property Get ActionSequenceFrame(mFrame As Long) As Long
Attribute ActionSequenceFrame.VB_Description = "Returns the image number for the selected frame. (Read Only)"
  On Error GoTo badvalue
  
  ActionSequenceFrame = m_ActionSequence(mFrame)
  
  Exit Property
  
badvalue:
  ActionSequenceFrame = 0
End Property

Public Property Let ActionSequence(ActionArray As Variant)
Attribute ActionSequence.VB_Description = "A variant containing a one dimensional array of image numbers for each frame in the animation sequence. (Write Only)"
  On Error GoTo badarray
  
  m_ActionSequence = ActionArray
  m_ActionFirst = LBound(ActionArray)
  m_ActionLast = UBound(ActionArray)
  
  Exit Property
  
badarray:
  m_ActionSequence = Empty
  m_ActionFirst = -1
  m_ActionLast = -1
End Property

Friend Property Set PreviousObject(prevObject As AnimationObject)
  Set m_PreviousObject = prevObject
End Property

Public Property Get PreviousObject() As AnimationObject
Attribute PreviousObject.VB_Description = "Previous object in animation window's object list. (Read Only)"
  Set PreviousObject = m_PreviousObject
End Property

Friend Property Set NextObject(nxtObject As AnimationObject)
  Set m_NextObject = nxtObject
End Property

Public Property Get NextObject() As AnimationObject
Attribute NextObject.VB_Description = "Next object in animation window's object list. (Read Only)"
  Set NextObject = m_NextObject
End Property

Public Property Get ObjectID() As Long
Attribute ObjectID.VB_Description = "Object unique identifier. Read Only."
  ObjectID = m_ObjectID
End Property

Friend Property Let ObjectID(objID As Long)
  m_ObjectID = objID
End Property

Public Property Get LayerPriority() As Long
Attribute LayerPriority.VB_Description = "Object's display priority. (Read Only)"
  LayerPriority = m_LayerPriority
End Property

Friend Property Let LayerPriority(layPriority As Long)
  m_LayerPriority = layPriority
End Property

