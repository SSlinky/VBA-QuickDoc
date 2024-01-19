VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickDoc"
Attribute VB_Description = "Generate documents from templates. Quickly!"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

Option Explicit

'-------------------------------------------------------------------------------
'   Class: QuickDoc
'   Generate documents from templates. Quickly!
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mWord As Object
Private mTemplatePath As String
Private mDocPath As String


' Properties
'-------------------------------------------------------------------------------
Public Property Let TemplatePath(var As String)
Attribute TemplatePath.VB_Description = "Sets or gets the template path."
'   Sets or gets the template path.
    mTemplatePath = var
End Property

Public Property Get TemplatePath() As String
    TemplatePath = mTemplatePath
End Property

Public Property Let DocPath(var As String)
Attribute DocPath.VB_Description = "Sets or gets teh document path."
'   Sets or gets teh document path.
    mDocPath = var
End Property

Public Property Get DocPath() As String
    DocPath = mDocPath
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Terminate()
    If Not mWord Is Nothing Then mWord.Quit
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub Generate(options As Dictionary)
Attribute Generate.VB_Description = "Generates a document from the template."
'   Generates a document from the template.
'
'   Args:
'       options: The options dictionary.
'
'   Raises:
'
    Dim doc As Object
    Set doc = WordApp.Documents.Open(mTemplatePath)

    PerformReplacements doc, options.GetValue("replacements", New Dictionary)
    SetCustomProperties doc, options.GetValue("properties", New Dictionary)

    doc.SaveAs2 _
        FileName:=DocPath, _
        ReadOnlyRecommended:= options.GetValue("ReadOnly", False)
    doc.Close
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Function WordApp() As Object
Attribute WordApp.VB_Description = "Lazy instantiates Word."
'   Lazy instantiates Word.
'
    If mWord Is Nothing Then Set mWord = CreateObject("Word.Application")
    Set WordApp = mWord
End Function

Private Sub EnsurePath(directoryPath As String)
Attribute EnsurePath.VB_Description = "Ensures the directory path exists."
'   Ensures the directory path exists.
'
'   Args:
'       directoryPath: The path to create.
'
    If Len(Dir(directoryPath, vbDirectory)) = 0 Then
        MkDir directoryPath
    End If
End Sub

Private Sub PerformReplacements(doc As Object, vars As Dictionary)
Attribute PerformReplacements.VB_Description = "Replaces values in the document."
'   Replaces values in the document.
'
'   Args:
'       doc: The document to act on.
'       vars: The replacements dictionary.
'
    Const wdReplaceAll As Long = 2

    Dim k As Variant
    For Each k In vars.Keys
        doc.Content.Find.Execute _
            FindText:=k, _
            ReplaceWith:=vars(key), _
            Replace:=wdReplaceAll
    Next k
End Sub

Private Sub SetCustomProperties(doc As Object, vars As Dictionary)
Attribute SetCustomProperties.VB_Description = "Sets the custom properties in the document."
'   Sets the custom properties in the document.
'
'   Args:
'       doc: The document to act on.
'       vars: The properties dictionary.
'
    For Each k In vars.Keys
        doc.CustomDocumentProperties.Add _
            Name:=k, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=vars(k)
    Next key
End Sub
