Imports System.Collections
Imports System.Windows.Forms

''' <summary>
''' NumericComparer instances are used as an argument into CollectionBase.Sort()
''' for sorting collections based on a given property of collection members.
''' </summary>
''' <remarks>
''' Based on http://weblogs.asp.net/jan/archive/2003/05/05/6479.aspx
''' 
''' Usage:
''' Dim customers As New ArrayList
''' 'Or you can use the Sort method of the strong typed collection, inheriting from CollectionBase.
''' customers.Sort(New NumericComparer("Name"))
''' 'or
''' customers.Sort(New NumericComparer("Name", SortOrder.Descending))
''' </remarks>
Public Class NumericComparer
    Implements IComparer

    Private propertyToSort As String
    Private mySortOrder As SortOrder

    ''' <summary>
    ''' Create a new instance SimpleComparer. Ascending sort order assumed.
    ''' </summary>
    ''' <param name="aPropertyToSort">Name of the object property to be used for sorting.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal aPropertyToSort As String)
        Me.New(aPropertyToSort, SortOrder.Ascending)
    End Sub

    ''' <summary>
    ''' Create a new instance SimpleComparer.
    ''' </summary>
    ''' <param name="aPropertyToSort">Name of the object property to be used for sorting.</param>
    ''' <param name="aSortOrder">Sort order to use.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal aPropertyToSort As String, ByVal aSortOrder As SortOrder)
        MyBase.new()
        Me.propertyToSort = aPropertyToSort
        Me.mySortOrder = aSortOrder
        Debug.Assert(Me.mySortOrder <> SortOrder.None)
    End Sub

    ''' <summary>
    ''' Comparison function
    ''' </summary>
    ''' <param name="x">First argument for comparison</param>
    ''' <param name="y">Second argument for comparison</param>
    ''' <returns>Return 0 if objects are equal based on their sort property.</returns>
    ''' <remarks>This implementation assumes objects being compared are of the same type.</remarks>
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
        Implements IComparer.Compare
        Dim prop As Reflection.PropertyInfo = x.GetType.GetProperty(Me.propertyToSort)
        Dim result As Integer = Math.Sign(prop.GetValue(x, Nothing) - prop.GetValue(y, Nothing))

        If Me.mySortOrder = SortOrder.Ascending Then
            Return result
        Else
            Return -result
        End If
    End Function

    'Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
    '    Implements IComparer.Compare
    '    Dim prop As Reflection.PropertyInfo = x.GetType.GetProperty(Me.propertyToSort)

    '    If Me.mySortOrder = SortOrder.None OrElse prop.GetValue(x, Nothing) = prop.GetValue(y, Nothing) Then
    '        Return 0
    '    Else
    '        If prop.GetValue(x, Nothing) > prop.GetValue(y, Nothing) Then
    '            If Me.mySortOrder = SortOrder.Ascending Then
    '                Return 1
    '            Else
    '                Return -1
    '            End If
    '        Else
    '            If Me.mySortOrder = SortOrder.Ascending Then
    '                Return -1
    '            Else
    '                Return 1
    '            End If
    '        End If
    '    End If
    'End Function

    'Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
    '    Implements IComparer.Compare
    '    Dim prop As Reflection.PropertyInfo = x.GetType.GetProperty(Me.propertyToSort)

    '    If Me.sortOrder = sortOrder.None OrElse prop.GetValue(x, Nothing) = prop.GetValue(y, Nothing) Then
    '        Return 0
    '    Else
    '        If prop.GetValue(x, Nothing) > prop.GetValue(y, Nothing) Then
    '            If Me.sortOrder = sortOrder.Ascending Then
    '                Return 1
    '            Else
    '                Return -1
    '            End If
    '        Else
    '            If Me.sortOrder = sortOrder.Ascending Then
    '                Return -1
    '            Else
    '                Return 1
    '            End If
    '        End If
    '    End If
    'End Function
End Class
