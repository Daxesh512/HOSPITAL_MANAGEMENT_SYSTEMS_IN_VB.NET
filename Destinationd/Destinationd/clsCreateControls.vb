Imports System
Imports System.Windows.Forms

Public Class clsCreateControls

#Region "variables"
    Private Shared height As Integer
    Private Shared width As Integer
    Private Shared left As Integer
    Private Shared top As Integer 'So the controls will be placed in a line, one under the other
    Private Shared text As String
    Private Shared name As String
#End Region

#Region "MakeControl"
    'Add the control to the form
    'AddControlToTheForm(myNewControl)
    Public Shared Function MakeControl(ByVal ControlType As String, ByVal cHeight As Integer, ByVal cWidth As Integer, ByVal cLeft As Integer, ByVal cTop As Integer, ByVal cText As String, ByVal cName As String) As Control
        'As there are four steps to follow
        'in order to create controls at runtime:
        'GET THE ASSEMBLY QUALIFIED NAME,
        'CREATE THE CONTROL,
        'SET THE CONTROL'S PROPERTIES and
        'ADD THE CONTROL TO THE FORM
        'I split method into four
        'methods doing exactly the specified
        'operations
        'Get the assembly qualified name
        Dim assemblyName As String
        assemblyName = GetAssemblyQualifiedName(ControlType)

        'Create the control
        Dim myNewControl As Control
        myNewControl = CreateControl(assemblyName)

        'Set it's properties like Name, 
        'Top, Left etc.
        height = cHeight
        width = cWidth
        left = cLeft
        top = cTop
        text = cText
        name = cName
        SetControlProperties(myNewControl)

        Return myNewControl
    End Function
#End Region

#Region "GetAssemblyQualifiedName"
    Private Shared Function GetAssemblyQualifiedName(ByVal ControlType As String) As String
        'In order to create a control 
        'at runtime you need to know:
        'THE CONTROL TYPE, 
        'THE NAME OF THE FILE that 
        'contains the type, 
        'THE VERION OF THE FILE, 
        'THE CULTURE and 
        'THE PUBLIC TOKEN KEY

        'It would be a better practice to 
        'declare this constants at the 
        'class level.
        'The reason why I placed this here
        'is olny that I wanted to make
        'this tutorial more easy to be 
        'followed and understood
        Dim TYPE_NAME As String = "System.Windows.Forms." + ControlType
        Const FILE_NAME As String = "System.Windows.Forms" 'REMARK! The name of the file without the type (.dll)
        Const VERSION_NUMBER As String = "1.0.3300.0"
        Const CULTURE As String = "neutral"
        Const TOKEN_KEY As String = "b77a5c561934e089"

        'Create the assembly qualified name
        Dim assemblyQualifiedName As String = TYPE_NAME + ", " + FILE_NAME + ", Version=" + VERSION_NUMBER + ", Culture=" + CULTURE + ", PublicKeyToken=" + TOKEN_KEY

        'return the assembly qualified name
        Return assemblyQualifiedName
    End Function
#End Region

#Region "CreateControl"
    Private Shared Function CreateControl(ByVal assemblyName As String) As Control
        'Get the button type
        Dim buttonType As Type = Type.GetType(assemblyName)

        'Get the button contructor
        'We search for the constructor
        'with no parameters
        Dim buttonCtor As System.Reflection.ConstructorInfo
        Dim typesCtorParams() As Type = Type.EmptyTypes
        buttonCtor = buttonType.GetConstructor(typesCtorParams)

        'Create the control by invoking
        'the constructor
        Dim paramsCtor() As Object = Nothing

        Dim myControl As Control = CType(buttonCtor.Invoke(paramsCtor), Control)
        'Return the control
        Return myControl
    End Function
#End Region

#Region "SetControlProperties"
    Private Shared Sub SetControlProperties(ByRef control As Control)
        'First increment the counter so the
        'next control will be visible
        '(otherwise it will be under
        'the previous control)
        'newControlsCounter++;
        'Set some properties
        control.Height = height
        control.Width = width
        control.Left = left
        control.Top = top 'So the controls will be placed in a line, one under the other
        control.Text = text
        control.Name = name 'Note: Multiple controls may have the same name in .Net
    End Sub
#End Region

End Class