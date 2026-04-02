Private keyColumnNames() As String
Private attributeColumnNames() As String
Private keyColumnIndices() As Long
Private attributeColumnIndices() As Long

Property Get KeyColumnNames() As String()
    KeyColumnNames = keyColumnNames
End Property

Property Let KeyColumnNames(value As String())
    keyColumnNames = value
End Property

Property Get AttributeColumnNames() As String()
    AttributeColumnNames = attributeColumnNames
End Property

Property Let AttributeColumnNames(value As String())
    attributeColumnNames = value
End Property

Property Get KeyColumnIndices() As Long()
    KeyColumnIndices = keyColumnIndices
End Property

Property Let KeyColumnIndices(value As Long())
    keyColumnIndices = value
End Property

Property Get AttributeColumnIndices() As Long()
    AttributeColumnIndices = attributeColumnIndices
End Property

Property Let AttributeColumnIndices(value As Long())
    attributeColumnIndices = value
End Property

'... (remaining code which uses the new properties to be added) ...