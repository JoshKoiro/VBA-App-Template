VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module FormGen
Option Explicit
'Member Variables
Private m_formName As String
Private m_formWidth As Integer
Private m_formHeight As Integer
Private m_controls As Collection
'This is the "constructor class"
Private Sub Class_Initialize()

    m_formName = ""
    m_formWidth = 300
    m_formHeight = 200

End Sub
'Getter for m_formName variable
Property Get formName() As String
    formName = m_formName
End Property
'Letter for m_formName variable
Property Let formName(name As String)
    m_formName = name
End Property
'Getter for the m_formWidth variable
Property Get formWidth() As Integer
    formWidth = m_formWidth
End Property
'Letter for the m_formWidth variable
Property Let formWidth(width As Integer)
    m_formWidth = width
End Property
'Getter for the m_formHeight variable
Property Get formHeight() As Integer
    formHeight = m_formHeight
End Property
'Letter for the m_formHeight variable
Property Let formHeight(height As Integer)
    m_formHeight = height
End Property
'Getter for m_controls collection variable
'Property Get controls() As Collection
    'Set controls = m_controls
'End Property
'Letter for m_controls variable collection
'Property Let controls(control As control)
   ' m_controls.Add control, control.name
'End Property

'Private function to remove white spaces in a string - acquired from https://stackoverflow.com/questions/42306358/removing-all-spaces-in-string
Private Function RemoveWhiteSpace(textString As String) As String
    With New RegExp
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(textString, vbNullString)
    End With
End Function

'Private helper function to change text string to be acceptable for a vba object name (proper case with no spaces) - https://www.automateexcel.com/vba/upper-lower-proper-case-functions/
Private Function formatName(text As String) As String

    formatName = RemoveWhiteSpace(StrConv(text, vbProperCase))

End Function

'This constructs the UserForm object based on the properties defined
Public Sub build()

'verify that the proper information is entered into the class properties

'check m_formName
If m_formName = "" Then
    MsgBox "Please provide a name for the userForm", vbCritical
    Exit Sub
End If

'define the form
Dim frm As VBComponent
Set frm = ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)

'Set the properties of the form
With frm
    .Properties("Caption") = m_formName
    .Properties("Width") = m_formWidth
    .Properties("Height") = m_formHeight
End With

frm.Activate

End Sub
