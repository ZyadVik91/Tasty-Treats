Public Class Bakery

'Declare variables
Dim name As String
Dim type As String
Dim description As String
Dim occasion As String
Dim price As Integer

'Declare Event
Public Event BakeryCompleted()

'Constructor
Public Sub New(ByVal Name As String, ByVal Type As String, ByVal Description As String, ByVal Occasion As String, ByVal Price As Integer)
    Me.Name = Name
    Me.Type = Type
    Me.Description = Description
    Me.Occasion = Occasion
    Me.Price = Price
End Sub

'Properties
Public Property GetName() As String
    Return Me.name
End Property

Public Property SetName(ByVal value As String)
    Me.name = value
End Property

Public Property GetType() As String
    Return Me.type
End Property

Public Property SetType(ByVal value As String)
    Me.type = value
End Property

Public Property GetDescription() As String
    Return Me.description
End Property

Public Property SetDescription(ByVal value As String)
    Me.description = value
End Property 

Public Property GetOccasion() As String
    Return Me.occasion
End Property

Public Property SetOccasion(ByVal value As String)
    Me.occasion = value
End Property

Public Property GetPrice() As Integer
    Return Me.Price
End Property

Public Property SetPrice(ByVal value As Integer)
    Me.price = value
End Property

'Methods
Public Sub CreateDesserts(ByVal Name As String, ByVal Type As String, ByVal Description As String, ByVal Occasion As String, ByVal Price As Integer)
    Me.Name = Name
    Me.Type = Type
    Me.Description = Description
    Me.Occasion = Occasion
    Me.Price = Price

    RaiseEvent BakeryCompleted()

End Sub

End Class