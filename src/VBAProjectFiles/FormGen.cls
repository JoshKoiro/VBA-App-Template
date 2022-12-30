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
Private pFormName As String
Private pFormWidth As Integer
Private pFormHeight As Integer
Private pControls As Collection
'This is the "constructor class"
Private Sub Class_Initialize()

    pFormName = ""
    pFormWidth = 300
    pFormHeight = 200

End Sub
'Getter for pFormName variable
Property Get formName() As String
    formName = pFormName
End Property
'Letter for pFormName variable
Property Let formName(name As String)
    pFormName = name
End Property
'Getter for the pFormWidth variable
Property Get formWidth() As Integer
    formWidth = pFormWidth
End Property
'Letter for the pFormWidth variable
Property Let formWidth(width As Integer)
    pFormWidth = width
End Property
'Getter for the pFormHeight variable
Property Get formHeight() As Integer
    formHeight = pFormHeight
End Property
'Letter for the pFormHeight variable
Property Let formHeight(height As Integer)
    pFormHeight = height
End Property
'Getter for pControls collection variable
'Property Get controls() As Collection
    'Set controls = pControls
'End Property
'Letter for pControls variable collection
'Property Let controls(control As control)
   ' pControls.Add control, control.name
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

'check if the form name already exists
Dim existingFrm As Object

For Each existingFrm In ThisWorkbook.VBProject.VBComponents

    'item is a userForm
    If existingFrm.Type = vbext_ct_MSForm Then
    
    If formatName(pFormName) = existingFrm.name Then
        MsgBox "Please provide a unique name for the userform. " & existingFrm.name & " has already been taken.", vbCritical
        Exit Sub
    End If
    End If

Next existingFrm

'define the form
Dim frm As VBComponent
Set frm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)

'Set the properties of the form
With frm
    .name = formatName(pFormName)
    .Properties("Caption") = pFormName
    .Properties("Width") = pFormWidth
    .Properties("Height") = pFormHeight
End With

frm.Activate

End Sub
