Attribute VB_Name = "modData"
Option Explicit
Option Base 1

'Enums-------------------------------------

Public Enum eVisualStyle
    Dots = 0
    Wireframe = 1
    Facet = 2
    Smooth = 3
End Enum

Public Enum eDelimiter
    eTab
    eSemicolon
    eComma
    eSpace
    eOther
End Enum
    
Public Enum eEditMF
    NewMesh = 0
    CreateMesh = 1
    ReverseMesh = 2
    DeleteMesh = 3
    ReverseFace = 4
    DeleteFace = 5
End Enum

Public Enum eSelector
    NoSelect = 0
    Rectangular = 1
    Polygonal = 2
End Enum

'Types-------------------------------------

Public Type DELIMITER
    tTab        As Boolean
    tSemicolon  As Boolean
    tComma      As Boolean
    tSpace      As Boolean
    tOther      As Boolean
    tDelimChar  As String
    tFormat     As Byte
End Type

Public Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

Public Type COLORS
    rgbBack1         As RGBQUAD
    rgbBack2         As RGBQUAD
    lObjDots         As Long
    lSelDots         As Long
    lWireframe       As Long
    rgbFaces         As RGBQUAD
    rgbFacesPen      As RGBQUAD
    rgbBackFace      As RGBQUAD
    rgbSelFace       As RGBQUAD
    rgbSelBackFace   As RGBQUAD
    lBox             As Long
    lSelGeo          As Long
End Type

Public Type VECTOR4
    X               As Single
    Y               As Single
    Z               As Single
    W               As Single
End Type

Public Type POINTAPI
    X               As Long
    Y               As Long
End Type

Public Type DOT
    Vector          As VECTOR4  '
    VectorT         As VECTOR4  'Changed position
    Screen          As POINTAPI
    NumUse          As Integer
    Color           As Long
    Selected        As Boolean
    Visible         As Boolean
    Index           As Long
End Type

Public Type EDGE
    Start           As Long
    End             As Long
    Used            As Integer
End Type

Public Type FACE
    A               As Long
    B               As Long
    C               As Long
    Color           As RGBQUAD
End Type

Public Type ORDER
    idxMesh         As Integer
    idxFace         As Long
    ZValue          As Single
End Type

Public Type FACEGROUP
    Name            As String
    NumFaces        As Long
    BorderEdges()   As EDGE
    Faces()         As FACE
    Normals()       As VECTOR4
    NormalsT()      As VECTOR4
End Type

Public Type OBJ_DOTS
    NumDot          As Long
    Dots()          As DOT      'All points
    NumSelDot       As Long
    SelDots()       As DOT      'Selected points
    Box(1 To 8)     As DOT
    Center          As DOT
    CenterZ         As Single
    ClpZ            As Single
End Type

Public Type OBJ_MESH
    NumVertices     As Long
    Vertices()      As VECTOR4       'All points
    NumMeshs        As Integer
    Meshs()         As FACEGROUP
    FaceV()         As ORDER
End Type

Public Type OBJ_CAMERA
    WorldPosition   As VECTOR4
    LookAtPoint     As VECTOR4
    VUP             As VECTOR4
    FOV             As Single
    Zoom            As Single
    ClipFar         As Single
    ClipNear        As Single
End Type

Public Type OBJ_LIGHT
    Position        As VECTOR4
    Normal          As VECTOR4
    NormalT         As VECTOR4
    Length          As Single
    Foton           As Integer
    Dark            As Integer
    Shade           As Single
End Type

Public Type LOCALE
    Rot             As VECTOR4
    Tra             As VECTOR4
    Sca             As Single
End Type

Public Dots1        As OBJ_DOTS
Public Mesh1        As OBJ_MESH
Public Camera1      As OBJ_CAMERA
Public Light1       As OBJ_LIGHT
Public Position     As LOCALE

Public CanBuffer    As New clsDIB
Public BackBuffer   As New clsDIB
Public cWidth       As Long 'canvas width
Public cHeight      As Long 'canvas height
Public HalfWidth    As Long 'canvas center
Public HalfHeight   As Long 'canvas center
Public LoadComplete As Boolean
Public SelectOp     As Boolean

Public Rect1        As POINTAPI
Public Rect2        As POINTAPI

Public Geometry     As clsPolygonal
Public cdiLoad      As clsCommonDialog
Public cfPoints     As clsFilePOINTS
Public cfEPJ        As clsFileEPJ
Public cfOBJ        As clsFileOBJ
Public delim        As DELIMITER
Public Color        As COLORS
Public StartPolygon    As Boolean
Public LastX        As Single
Public LastY        As Single
Public MaxH         As Single

Public SelectedMeshIndex As Integer
Public SelectedFaceIndex As Long

Public ClipFar      As Boolean
Public BigDot       As Boolean
Public ShowBackFace As Boolean
Public ShowBox      As Boolean
Public ShowMeshBorder As Boolean
Public ShowHideDot  As Boolean
Public Unsaved      As Boolean
Public Action       As Boolean

Public VStyle       As eVisualStyle
Public EditMF       As eEditMF
Public SelType      As eSelector

Public InvScl As Single

