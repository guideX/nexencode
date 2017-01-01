Imports System.Collections
Imports Freedb

Namespace Freedb
#Region "Interface ISiteCollection"

    ''' <summary>
    ''' Defines size, enumerators, and synchronization methods for strongly
    ''' typed collections of <see cref="Site"/> elements.
    ''' </summary>
    ''' <remarks>
    ''' <b>ISiteCollection</b> provides an <see cref="ICollection"/>
    ''' that is strongly typed for <see cref="Site"/> elements.
    ''' </remarks>

    Public Interface ISiteCollection
#Region "Properties"
#Region "Count"

        ''' <summary>
        ''' Gets the number of elements contained in the
        ''' <see cref="ISiteCollection"/>.
        ''' </summary>
        ''' <value>The number of elements contained in the
        ''' <see cref="ISiteCollection"/>.</value>
        ''' <remarks>Please refer to <see cref="ICollection.Count"/> for details.</remarks>

        ReadOnly Property Count() As Integer

#End Region
#Region "IsSynchronized"

        ''' <summary>
        ''' Gets a value indicating whether access to the
        ''' <see cref="ISiteCollection"/> is synchronized (thread-safe).
        ''' </summary>
        ''' <value><c>true</c> if access to the <see cref="ISiteCollection"/> is
        ''' synchronized (thread-safe); otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="ICollection.IsSynchronized"/> for details.</remarks>

        ReadOnly Property IsSynchronized() As Boolean

#End Region
#Region "SyncRoot"

        ''' <summary>
        ''' Gets an object that can be used to synchronize access
        ''' to the <see cref="ISiteCollection"/>.
        ''' </summary>
        ''' <value>An object that can be used to synchronize access
        ''' to the <see cref="ISiteCollection"/>.</value>
        ''' <remarks>Please refer to <see cref="ICollection.SyncRoot"/> for details.</remarks>

        ReadOnly Property SyncRoot() As Object

#End Region
#End Region
#Region "Methods"
#Region "CopyTo"

        ''' <summary>
        ''' Copies the entire <see cref="ISiteCollection"/> to a one-dimensional <see cref="Array"/>
        ''' of <see cref="Site"/> elements, starting at the specified index of the target array.
        ''' </summary>
        ''' <param name="array">The one-dimensional <see cref="Array"/> that is the destination of the
        ''' <see cref="Site"/> elements copied from the <see cref="ISiteCollection"/>.
        ''' The <b>Array</b> must have zero-based indexing.</param>
        ''' <param name="arrayIndex">The zero-based index in <paramref name="array"/>
        ''' at which copying begins.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <paramref name="arrayIndex"/> is less than zero.</exception>
        ''' <exception cref="ArgumentException"><para>
        ''' <paramref name="arrayIndex"/> is equal to or greater than the length of <paramref name="array"/>.
        ''' </para><para>-or-</para><para>
        ''' The number of elements in the source <see cref="ISiteCollection"/> is greater
        ''' than the available space from <paramref name="arrayIndex"/> to the end of the destination
        ''' <paramref name="array"/>.</para></exception>
        ''' <remarks>Please refer to <see cref="ICollection.CopyTo"/> for details.</remarks>

        Sub CopyTo(array As Site(), arrayIndex As Integer)

#End Region
#Region "GetEnumerator"

        ''' <summary>
        ''' Returns an <see cref="ISiteEnumerator"/> that can
        ''' iterate through the <see cref="ISiteCollection"/>.
        ''' </summary>
        ''' <returns>An <see cref="ISiteEnumerator"/>
        ''' for the entire <see cref="ISiteCollection"/>.</returns>
        ''' <remarks>Please refer to <see cref="IEnumerable.GetEnumerator"/> for details.</remarks>

        Function GetEnumerator() As ISiteEnumerator

#End Region
#End Region
    End Interface

#End Region
#Region "Interface ISiteList"

    ''' <summary>
    ''' Represents a strongly typed collection of <see cref="Site"/>
    ''' objects that can be individually accessed by index.
    ''' </summary>
    ''' <remarks>
    ''' <b>ISiteList</b> provides an <see cref="IList"/>
    ''' that is strongly typed for <see cref="Site"/> elements.
    ''' </remarks>

    Public Interface ISiteList
        Inherits ISiteCollection
#Region "Properties"
#Region "IsFixedSize"

        ''' <summary>
        ''' Gets a value indicating whether the <see cref="ISiteList"/> has a fixed size.
        ''' </summary>
        ''' <value><c>true</c> if the <see cref="ISiteList"/> has a fixed size;
        ''' otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="IList.IsFixedSize"/> for details.</remarks>

        ReadOnly Property IsFixedSize() As Boolean

#End Region
#Region "IsReadOnly"

        ''' <summary>
        ''' Gets a value indicating whether the <see cref="ISiteList"/> is read-only.
        ''' </summary>
        ''' <value><c>true</c> if the <see cref="ISiteList"/> is read-only;
        ''' otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="IList.IsReadOnly"/> for details.</remarks>

        ReadOnly Property IsReadOnly() As Boolean

#End Region
#Region "Item"

        ''' <summary>
        ''' Gets or sets the <see cref="Site"/> element at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index of the
        ''' <see cref="Site"/> element to get or set.</param>
        ''' <value>
        ''' The <see cref="Site"/> element at the specified <paramref name="index"/>.
        ''' </value>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is equal to or greater than
        ''' <see cref="ISiteCollection.Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' The property is set and the <see cref="ISiteList"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="IList.this"/> for details.</remarks>

        Default Property Item(index As Integer) As Site

#End Region
#End Region
#Region "Methods"
#Region "Add"

        ''' <summary>
        ''' Adds a <see cref="Site"/> to the end
        ''' of the <see cref="ISiteList"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to be added to the end of the <see cref="ISiteList"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>The <see cref="ISiteList"/> index at which
        ''' the <paramref name="value"/> has been added.</returns>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="ISiteList"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>ISiteList</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="IList.Add"/> for details.</remarks>

        Function Add(value As Site) As Integer

#End Region
#Region "Clear"

        ''' <summary>
        ''' Removes all elements from the <see cref="ISiteList"/>.
        ''' </summary>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="ISiteList"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>ISiteList</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="IList.Clear"/> for details.</remarks>

        Sub Clear()

#End Region
#Region "Contains"

        ''' <summary>
        ''' Determines whether the <see cref="ISiteList"/>
        ''' contains the specified <see cref="Site"/> element.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to locate in the <see cref="ISiteList"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns><c>true</c> if <paramref name="value"/> is found in the
        ''' <see cref="ISiteList"/>; otherwise, <c>false</c>.</returns>
        ''' <remarks>Please refer to <see cref="IList.Contains"/> for details.</remarks>

        Function Contains(value As Site) As Boolean

#End Region
#Region "IndexOf"

        ''' <summary>
        ''' Returns the zero-based index of the first occurrence of the specified
        ''' <see cref="Site"/> in the <see cref="ISiteList"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to locate in the <see cref="ISiteList"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>
        ''' The zero-based index of the first occurrence of <paramref name="value"/>
        ''' in the <see cref="ISiteList"/>, if found; otherwise, -1.
        ''' </returns>
        ''' <remarks>Please refer to <see cref="IList.IndexOf"/> for details.</remarks>

        Function IndexOf(value As Site) As Integer

#End Region
#Region "Insert"

        ''' <summary>
        ''' Inserts a <see cref="Site"/> element into the
        ''' <see cref="ISiteList"/> at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index at which
        ''' <paramref name="value"/> should be inserted.</param>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to insert into the <see cref="ISiteList"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is greater than
        ''' <see cref="ISiteCollection.Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="ISiteList"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>ISiteList</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="IList.Insert"/> for details.</remarks>

        Sub Insert(index As Integer, value As Site)

#End Region
#Region "Remove"

        ''' <summary>
        ''' Removes the first occurrence of the specified <see cref="Site"/>
        ''' from the <see cref="ISiteList"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to remove from the <see cref="ISiteList"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="ISiteList"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>ISiteList</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="IList.Remove"/> for details.</remarks>

        Sub Remove(value As Site)

#End Region
#Region "RemoveAt"

        ''' <summary>
        ''' Removes the element at the specified index of the
        ''' <see cref="ISiteList"/>.
        ''' </summary>
        ''' <param name="index">The zero-based index of the element to remove.</param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is equal to or greater than
        ''' <see cref="ISiteCollection.Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="ISiteList"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>ISiteList</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="IList.RemoveAt"/> for details.</remarks>

        Sub RemoveAt(index As Integer)

#End Region
#End Region
    End Interface

#End Region
#Region "Interface ISiteEnumerator"

    ''' <summary>
    ''' Supports type-safe iteration over a collection that
    ''' contains <see cref="Site"/> elements.
    ''' </summary>
    ''' <remarks>
    ''' <b>ISiteEnumerator</b> provides an <see cref="IEnumerator"/>
    ''' that is strongly typed for <see cref="Site"/> elements.
    ''' </remarks>

    Public Interface ISiteEnumerator
#Region "Properties"
#Region "Current"

        ''' <summary>
        ''' Gets the current <see cref="Site"/> element in the collection.
        ''' </summary>
        ''' <value>The current <see cref="Site"/> element in the collection.</value>
        ''' <exception cref="InvalidOperationException"><para>The enumerator is positioned
        ''' before the first element of the collection or after the last element.</para>
        ''' <para>-or-</para>
        ''' <para>The collection was modified after the enumerator was created.</para></exception>
        ''' <remarks>Please refer to <see cref="IEnumerator.Current"/> for details, but note
        ''' that <b>Current</b> fails if the collection was modified since the last successful
        ''' call to <see cref="MoveNext"/> or <see cref="Reset"/>.</remarks>

        ReadOnly Property Current() As Site

#End Region
#End Region
#Region "Methods"
#Region "MoveNext"

        ''' <summary>
        ''' Advances the enumerator to the next element of the collection.
        ''' </summary>
        ''' <returns><c>true</c> if the enumerator was successfully advanced to the next element;
        ''' <c>false</c> if the enumerator has passed the end of the collection.</returns>
        ''' <exception cref="InvalidOperationException">
        ''' The collection was modified after the enumerator was created.</exception>
        ''' <remarks>Please refer to <see cref="IEnumerator.MoveNext"/> for details.</remarks>

        Function MoveNext() As Boolean

#End Region
#Region "Reset"

        ''' <summary>
        ''' Sets the enumerator to its initial position,
        ''' which is before the first element in the collection.
        ''' </summary>
        ''' <exception cref="InvalidOperationException">
        ''' The collection was modified after the enumerator was created.</exception>
        ''' <remarks>Please refer to <see cref="IEnumerator.Reset"/> for details.</remarks>

        Sub Reset()

#End Region
#End Region
    End Interface

#End Region
#Region "Class SiteCollection"

    ''' <summary>
    ''' Implements a strongly typed collection of <see cref="Site"/> elements.
    ''' </summary>
    ''' <remarks><para>
    ''' <b>SiteCollection</b> provides an <see cref="ArrayList"/>
    ''' that is strongly typed for <see cref="Site"/> elements.
    ''' </para></remarks>

    <Serializable()> _
    Public Class SiteCollection
        Implements ISiteList
        Implements IList
        Implements ICloneable
#Region "Private Fields"

        Private Const _defaultCapacity As Integer = 16

        Private _array As Site() = Nothing
        Private _count As Integer = 0

        <NonSerialized()> _
        Private _version As Integer = 0

#End Region
#Region "Private Constructors"

        ' helper type to identify private ctor
        Private Enum Tag
            [Default]
        End Enum

        Private Sub New(tag As Tag)
        End Sub

#End Region
#Region "Public Constructors"
#Region "SiteCollection()"

        ''' <overloads>
        ''' Initializes a new instance of the <see cref="SiteCollection"/> class.
        ''' </overloads>
        ''' <summary>
        ''' Initializes a new instance of the <see cref="SiteCollection"/> class
        ''' that is empty and has the default initial capacity.
        ''' </summary>
        ''' <remarks>Please refer to <see cref="ArrayList()"/> for details.</remarks>

        Public Sub New()
            Me._array = New Site(_defaultCapacity - 1) {}
        End Sub

#End Region
#Region "SiteCollection(Int32)"

        ''' <summary>
        ''' Initializes a new instance of the <see cref="SiteCollection"/> class
        ''' that is empty and has the specified initial capacity.
        ''' </summary>
        ''' <param name="capacity">The number of elements that the new
        ''' <see cref="SiteCollection"/> is initially capable of storing.</param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <paramref name="capacity"/> is less than zero.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList(Int32)"/> for details.</remarks>

        Public Sub New(capacity As Integer)
            If capacity < 0 Then
                Throw New ArgumentOutOfRangeException("capacity", capacity, "Argument cannot be negative.")
            End If

            Me._array = New Site(capacity - 1) {}
        End Sub

#End Region
#Region "SiteCollection(SiteCollection)"

        ''' <summary>
        ''' Initializes a new instance of the <see cref="SiteCollection"/> class
        ''' that contains elements copied from the specified collection and
        ''' that has the same initial capacity as the number of elements copied.
        ''' </summary>
        ''' <param name="collection">The <see cref="SiteCollection"/>
        ''' whose elements are copied to the new collection.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="collection"/> is a null reference.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList(ICollection)"/> for details.</remarks>

        Public Sub New(collection As SiteCollection)
            If collection Is Nothing Then
                Throw New ArgumentNullException("collection")
            End If

            Me._array = New Site(collection.Count - 1) {}
            AddRange(collection)
        End Sub

#End Region
#Region "SiteCollection(Site[])"

        ''' <summary>
        ''' Initializes a new instance of the <see cref="SiteCollection"/> class
        ''' that contains elements copied from the specified <see cref="Site"/>
        ''' array and that has the same initial capacity as the number of elements copied.
        ''' </summary>
        ''' <param name="array">An <see cref="Array"/> of <see cref="Site"/>
        ''' elements that are copied to the new collection.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList(ICollection)"/> for details.</remarks>

        Public Sub New(array As Site())
            If array Is Nothing Then
                Throw New ArgumentNullException("array")
            End If

            Me._array = New Site(array.Length - 1) {}
            AddRange(array)
        End Sub

#End Region
#End Region
#Region "Protected Properties"
#Region "InnerArray"

        ''' <summary>
        ''' Gets the list of elements contained in the <see cref="SiteCollection"/> instance.
        ''' </summary>
        ''' <value>
        ''' A one-dimensional <see cref="Array"/> with zero-based indexing that contains all 
        ''' <see cref="Site"/> elements in the <see cref="SiteCollection"/>.
        ''' </value>
        ''' <remarks>
        ''' Use <b>InnerArray</b> to access the element array of a <see cref="SiteCollection"/>
        ''' instance that might be a read-only or synchronized wrapper. This is necessary because
        ''' the element array field of wrapper classes is always a null reference.
        ''' </remarks>

        Protected Overridable ReadOnly Property InnerArray() As Site()
            Get
                Return Me._array
            End Get
        End Property

#End Region
#End Region
#Region "Public Properties"
#Region "Capacity"

        ''' <summary>
        ''' Gets or sets the capacity of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <value>The number of elements that the
        ''' <see cref="SiteCollection"/> can contain.</value>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <b>Capacity</b> is set to a value that is less than <see cref="Count"/>.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Capacity"/> for details.</remarks>

        Public Overridable Property Capacity() As Integer
            Get
                Return Me._array.Length
            End Get
            Set(value As Integer)
                If value = Me._array.Length Then
                    Return
                End If

                If value < Me._count Then
                    Throw New ArgumentOutOfRangeException("Capacity", value, "Value cannot be less than Count.")
                End If

                If value = 0 Then
                    Me._array = New Site(_defaultCapacity - 1) {}
                    Return
                End If

                Dim newArray As Site() = New Site(value - 1) {}
                Array.Copy(Me._array, newArray, Me._count)
                Me._array = newArray
            End Set
        End Property

#End Region
#Region "Count"

        ''' <summary>
        ''' Gets the number of elements contained in the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <value>
        ''' The number of elements contained in the <see cref="SiteCollection"/>.
        ''' </value>
        ''' <remarks>Please refer to <see cref="ArrayList.Count"/> for details.</remarks>

        Public Overridable ReadOnly Property Count() As Integer Implements ISiteCollection.Count
            Get
                Return Me._count
            End Get
        End Property

#End Region
#Region "IsFixedSize"

        ''' <summary>
        ''' Gets a value indicating whether the <see cref="SiteCollection"/> has a fixed size.
        ''' </summary>
        ''' <value><c>true</c> if the <see cref="SiteCollection"/> has a fixed size;
        ''' otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="ArrayList.IsFixedSize"/> for details.</remarks>

        Public Overridable ReadOnly Property IsFixedSize() As Boolean Implements ISiteList.IsFixedSize
            Get
                Return False
            End Get
        End Property

#End Region
#Region "IsReadOnly"

        ''' <summary>
        ''' Gets a value indicating whether the <see cref="SiteCollection"/> is read-only.
        ''' </summary>
        ''' <value><c>true</c> if the <see cref="SiteCollection"/> is read-only;
        ''' otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="ArrayList.IsReadOnly"/> for details.</remarks>

        Public Overridable ReadOnly Property IsReadOnly() As Boolean Implements ISiteList.IsReadOnly
            Get
                Return False
            End Get
        End Property

#End Region
#Region "IsSynchronized"

        ''' <summary>
        ''' Gets a value indicating whether access to the <see cref="SiteCollection"/>
        ''' is synchronized (thread-safe).
        ''' </summary>
        ''' <value><c>true</c> if access to the <see cref="SiteCollection"/> is
        ''' synchronized (thread-safe); otherwise, <c>false</c>. The default is <c>false</c>.</value>
        ''' <remarks>Please refer to <see cref="ArrayList.IsSynchronized"/> for details.</remarks>

        Public Overridable ReadOnly Property IsSynchronized() As Boolean Implements ISiteCollection.IsSynchronized
            Get
                Return False
            End Get
        End Property

#End Region
#Region "IsUnique"

        ''' <summary>
        ''' Gets a value indicating whether the <see cref="SiteCollection"/> 
        ''' ensures that all elements are unique.
        ''' </summary>
        ''' <value>
        ''' <c>true</c> if the <see cref="SiteCollection"/> ensures that all 
        ''' elements are unique; otherwise, <c>false</c>. The default is <c>false</c>.
        ''' </value>
        ''' <remarks>
        ''' <b>IsUnique</b> returns <c>true</c> exactly if the <see cref="SiteCollection"/>
        ''' is exposed through a <see cref="Unique"/> wrapper. 
        ''' Please refer to <see cref="Unique"/> for details.
        ''' </remarks>

        Public Overridable ReadOnly Property IsUnique() As Boolean
            Get
                Return False
            End Get
        End Property

#End Region
#Region "Item: Site"

        ''' <summary>
        ''' Gets or sets the <see cref="Site"/> element at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index of the
        ''' <see cref="Site"/> element to get or set.</param>
        ''' <value>
        ''' The <see cref="Site"/> element at the specified <paramref name="index"/>.
        ''' </value>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is equal to or greater than <see cref="Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException"><para>
        ''' The property is set and the <see cref="SiteCollection"/> is read-only.
        ''' </para><para>-or-</para><para>
        ''' The property is set, the <b>SiteCollection</b> already contains the
        ''' specified element at a different index, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.this"/> for details.</remarks>

        Default Public Overridable Property Item(index As Integer) As Site
            Get
                ValidateIndex(index)
                Return Me._array(index)
            End Get
            Set(value As Site)
                ValidateIndex(index)
                Me._version += 1
                Me._array(index) = value
            End Set
        End Property

#End Region
#Region "IList.Item: Object"

        ''' <summary>
        ''' Gets or sets the element at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index of the element to get or set.</param>
        ''' <value>
        ''' The element at the specified <paramref name="index"/>. When the property
        ''' is set, this value must be compatible with <see cref="Site"/>.
        ''' </value>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is equal to or greater than <see cref="Count"/>.</para>
        ''' </exception>
        ''' <exception cref="InvalidCastException">The property is set to a value
        ''' that is not compatible with <see cref="Site"/>.</exception>
        ''' <exception cref="NotSupportedException"><para>
        ''' The property is set and the <see cref="SiteCollection"/> is read-only.
        ''' </para><para>-or-</para><para>
        ''' The property is set, the <b>SiteCollection</b> already contains the
        ''' specified element at a different index, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.this"/> for details.</remarks>

        Default Private Property IList_Item(index As Integer) As Object Implements IList.this
            Get
                Return Me(index)
            End Get
            Set(value As Object)
                Me(index) = DirectCast(value, Site)
            End Set
        End Property

#End Region
#Region "SyncRoot"

        ''' <summary>
        ''' Gets an object that can be used to synchronize
        ''' access to the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <value>An object that can be used to synchronize
        ''' access to the <see cref="SiteCollection"/>.
        ''' </value>
        ''' <remarks>Please refer to <see cref="ArrayList.SyncRoot"/> for details.</remarks>

        Public Overridable ReadOnly Property SyncRoot() As Object Implements ISiteCollection.SyncRoot
            Get
                Return Me
            End Get
        End Property

#End Region
#End Region
#Region "Public Methods"
#Region "Add(Site)"

        ''' <summary>
        ''' Adds a <see cref="Site"/> to the end of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to be added to the end of the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>The <see cref="SiteCollection"/> index at which the
        ''' <paramref name="value"/> has been added.</returns>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains the specified
        ''' <paramref name="value"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Add"/> for details.</remarks>

        Public Overridable Function Add(value As Site) As Integer
            If Me._count = Me._array.Length Then
                EnsureCapacity(Me._count + 1)
            End If

            Me._version += 1
            Me._array(Me._count) = value
            Return System.Math.Max(System.Threading.Interlocked.Increment(Me._count), Me._count - 1)
        End Function

#End Region
#Region "IList.Add(Object)"

        ''' <summary>
        ''' Adds an <see cref="Object"/> to the end of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">
        ''' The object to be added to the end of the <see cref="SiteCollection"/>.
        ''' This argument must be compatible with <see cref="Site"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>The <see cref="SiteCollection"/> index at which the
        ''' <paramref name="value"/> has been added.</returns>
        ''' <exception cref="InvalidCastException"><paramref name="value"/>
        ''' is not compatible with <see cref="Site"/>.</exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains the specified
        ''' <paramref name="value"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Add"/> for details.</remarks>

        Private Function IList_Add(value As Object) As Integer Implements IList.Add
            Return Add(DirectCast(value, Site))
        End Function

#End Region
#Region "AddRange(SiteCollection)"

        ''' <overloads>
        ''' Adds a range of elements to the end of the <see cref="SiteCollection"/>.
        ''' </overloads>
        ''' <summary>
        ''' Adds the elements of another collection to the end of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="collection">The <see cref="SiteCollection"/> whose elements
        ''' should be added to the end of the current collection.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="collection"/> is a null reference.</exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains one or more elements
        ''' in the specified <paramref name="collection"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.AddRange"/> for details.</remarks>

        Public Overridable Sub AddRange(collection As SiteCollection)
            If collection Is Nothing Then
                Throw New ArgumentNullException("collection")
            End If

            If collection.Count = 0 Then
                Return
            End If
            If Me._count + collection.Count > Me._array.Length Then
                EnsureCapacity(Me._count + collection.Count)
            End If

            Me._version += 1
            Array.Copy(collection.InnerArray, 0, Me._array, Me._count, collection.Count)
            Me._count += collection.Count
        End Sub

#End Region
#Region "AddRange(Site[])"

        ''' <summary>
        ''' Adds the elements of a <see cref="Site"/> array
        ''' to the end of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="array">An <see cref="Array"/> of <see cref="Site"/> elements
        ''' that should be added to the end of the <see cref="SiteCollection"/>.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains one or more elements
        ''' in the specified <paramref name="array"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.AddRange"/> for details.</remarks>

        Public Overridable Sub AddRange(array__1 As Site())
            If array__1 Is Nothing Then
                Throw New ArgumentNullException("array")
            End If

            If array__1.Length = 0 Then
                Return
            End If
            If Me._count + array__1.Length > Me._array.Length Then
                EnsureCapacity(Me._count + array__1.Length)
            End If

            Me._version += 1
            Array.Copy(array__1, 0, Me._array, Me._count, array__1.Length)
            Me._count += array__1.Length
        End Sub

#End Region
#Region "BinarySearch"

        ''' <summary>
        ''' Searches the entire sorted <see cref="SiteCollection"/> for an
        ''' <see cref="Site"/> element using the default comparer
        ''' and returns the zero-based index of the element.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to locate in the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>The zero-based index of <paramref name="value"/> in the sorted
        ''' <see cref="SiteCollection"/>, if <paramref name="value"/> is found;
        ''' otherwise, a negative number, which is the bitwise complement of the index
        ''' of the next element that is larger than <paramref name="value"/> or, if there
        ''' is no larger element, the bitwise complement of <see cref="Count"/>.</returns>
        ''' <exception cref="InvalidOperationException">
        ''' Neither <paramref name="value"/> nor the elements of the <see cref="SiteCollection"/>
        ''' implement the <see cref="IComparable"/> interface.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.BinarySearch"/> for details.</remarks>

        Public Overridable Function BinarySearch(value As Site) As Integer
            Return Array.BinarySearch(Me._array, 0, Me._count, value)
        End Function

#End Region
#Region "Clear"

        ''' <summary>
        ''' Removes all elements from the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Clear"/> for details.</remarks>

        Public Overridable Sub Clear() Implements ISiteList.Clear
            If Me._count = 0 Then
                Return
            End If

            Me._version += 1
            Array.Clear(Me._array, 0, Me._count)
            Me._count = 0
        End Sub

#End Region
#Region "Clone"

        ''' <summary>
        ''' Creates a shallow copy of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <returns>A shallow copy of the <see cref="SiteCollection"/>.</returns>
        ''' <remarks>Please refer to <see cref="ArrayList.Clone"/> for details.</remarks>

        Public Overridable Function Clone() As Object
            Dim collection As New SiteCollection(Me._count)

            Array.Copy(Me._array, 0, collection._array, 0, Me._count)
            collection._count = Me._count
            collection._version = Me._version

            Return collection
        End Function

#End Region
#Region "Contains(Site)"

        ''' <summary>
        ''' Determines whether the <see cref="SiteCollection"/>
        ''' contains the specified <see cref="Site"/> element.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to locate in the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns><c>true</c> if <paramref name="value"/> is found in the
        ''' <see cref="SiteCollection"/>; otherwise, <c>false</c>.</returns>
        ''' <remarks>Please refer to <see cref="ArrayList.Contains"/> for details.</remarks>

        Public Function Contains(value As Site) As Boolean
            Return (IndexOf(value) >= 0)
        End Function

#End Region
#Region "IList.Contains(Object)"

        ''' <summary>
        ''' Determines whether the <see cref="SiteCollection"/> contains the specified element.
        ''' </summary>
        ''' <param name="value">The object to locate in the <see cref="SiteCollection"/>.
        ''' This argument must be compatible with <see cref="Site"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns><c>true</c> if <paramref name="value"/> is found in the
        ''' <see cref="SiteCollection"/>; otherwise, <c>false</c>.</returns>
        ''' <exception cref="InvalidCastException"><paramref name="value"/>
        ''' is not compatible with <see cref="Site"/>.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Contains"/> for details.</remarks>

        Private Function IList_Contains(value As Object) As Boolean Implements IList.Contains
            Return Contains(DirectCast(value, Site))
        End Function

#End Region
#Region "CopyTo(Site[])"

        ''' <overloads>
        ''' Copies the <see cref="SiteCollection"/> or a portion of it to a one-dimensional array.
        ''' </overloads>
        ''' <summary>
        ''' Copies the entire <see cref="SiteCollection"/> to a one-dimensional <see cref="Array"/>
        ''' of <see cref="Site"/> elements, starting at the beginning of the target array.
        ''' </summary>
        ''' <param name="array">The one-dimensional <see cref="Array"/> that is the destination of the
        ''' <see cref="Site"/> elements copied from the <see cref="SiteCollection"/>.
        ''' The <b>Array</b> must have zero-based indexing.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <exception cref="ArgumentException">
        ''' The number of elements in the source <see cref="SiteCollection"/> is greater
        ''' than the available space in the destination <paramref name="array"/>.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.CopyTo"/> for details.</remarks>

        Public Overridable Sub CopyTo(array__1 As Site())
            CheckTargetArray(array__1, 0)
            Array.Copy(Me._array, array__1, Me._count)
        End Sub

#End Region
#Region "CopyTo(Site[], Int32)"

        ''' <summary>
        ''' Copies the entire <see cref="SiteCollection"/> to a one-dimensional <see cref="Array"/>
        ''' of <see cref="Site"/> elements, starting at the specified index of the target array.
        ''' </summary>
        ''' <param name="array">The one-dimensional <see cref="Array"/> that is the destination of the
        ''' <see cref="Site"/> elements copied from the <see cref="SiteCollection"/>.
        ''' The <b>Array</b> must have zero-based indexing.</param>
        ''' <param name="arrayIndex">The zero-based index in <paramref name="array"/>
        ''' at which copying begins.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <paramref name="arrayIndex"/> is less than zero.</exception>
        ''' <exception cref="ArgumentException"><para>
        ''' <paramref name="arrayIndex"/> is equal to or greater than the length of <paramref name="array"/>.
        ''' </para><para>-or-</para><para>
        ''' The number of elements in the source <see cref="SiteCollection"/> is greater than the
        ''' available space from <paramref name="arrayIndex"/> to the end of the destination
        ''' <paramref name="array"/>.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.CopyTo"/> for details.</remarks>

        Public Overridable Sub CopyTo(array__1 As Site(), arrayIndex As Integer)
            CheckTargetArray(array__1, arrayIndex)
            Array.Copy(Me._array, 0, array__1, arrayIndex, Me._count)
        End Sub

#End Region
#Region "ICollection.CopyTo(Array, Int32)"

        ''' <summary>
        ''' Copies the entire <see cref="SiteCollection"/> to a one-dimensional <see cref="Array"/>,
        ''' starting at the specified index of the target array.
        ''' </summary>
        ''' <param name="array">The one-dimensional <see cref="Array"/> that is the destination of the
        ''' <see cref="Site"/> elements copied from the <see cref="SiteCollection"/>.
        ''' The <b>Array</b> must have zero-based indexing.</param>
        ''' <param name="arrayIndex">The zero-based index in <paramref name="array"/>
        ''' at which copying begins.</param>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="array"/> is a null reference.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <paramref name="arrayIndex"/> is less than zero.</exception>
        ''' <exception cref="ArgumentException"><para>
        ''' <paramref name="array"/> is multidimensional.
        ''' </para><para>-or-</para><para>
        ''' <paramref name="arrayIndex"/> is equal to or greater than the length of <paramref name="array"/>.
        ''' </para><para>-or-</para><para>
        ''' The number of elements in the source <see cref="SiteCollection"/> is greater than the
        ''' available space from <paramref name="arrayIndex"/> to the end of the destination
        ''' <paramref name="array"/>.</para></exception>
        ''' <exception cref="InvalidCastException">
        ''' The <see cref="Site"/> type cannot be cast automatically
        ''' to the type of the destination <paramref name="array"/>.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.CopyTo"/> for details.</remarks>

        Private Sub ICollection_CopyTo(array As Array, arrayIndex As Integer) Implements ICollection.CopyTo
            CopyTo(DirectCast(array, Site()), arrayIndex)
        End Sub

#End Region
#Region "GetEnumerator: ISiteEnumerator"

        ''' <summary>
        ''' Returns an <see cref="ISiteEnumerator"/> that can
        ''' iterate through the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <returns>An <see cref="ISiteEnumerator"/>
        ''' for the entire <see cref="SiteCollection"/>.</returns>
        ''' <remarks>Please refer to <see cref="ArrayList.GetEnumerator"/> for details.</remarks>

        Public Overridable Function GetEnumerator() As ISiteEnumerator Implements ISiteCollection.GetEnumerator
            Return New Enumerator(Me)
        End Function

#End Region
#Region "IEnumerable.GetEnumerator: IEnumerator"

        ''' <summary>
        ''' Returns an <see cref="IEnumerator"/> that can
        ''' iterate through the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <returns>An <see cref="IEnumerator"/>
        ''' for the entire <see cref="SiteCollection"/>.</returns>
        ''' <remarks>Please refer to <see cref="ArrayList.GetEnumerator"/> for details.</remarks>

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(GetEnumerator(), IEnumerator)
        End Function

#End Region
#Region "IndexOf(Site)"

        ''' <summary>
        ''' Returns the zero-based index of the first occurrence of the specified
        ''' <see cref="Site"/> in the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to locate in the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>
        ''' The zero-based index of the first occurrence of <paramref name="value"/>
        ''' in the <see cref="SiteCollection"/>, if found; otherwise, -1.
        ''' </returns>
        ''' <remarks>Please refer to <see cref="ArrayList.IndexOf"/> for details.</remarks>

        Public Overridable Function IndexOf(value As Site) As Integer
            Return Array.IndexOf(Me._array, value, 0, Me._count)
        End Function

#End Region
#Region "IList.IndexOf(Object)"

        ''' <summary>
        ''' Returns the zero-based index of the first occurrence of the specified
        ''' <see cref="Object"/> in the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">The object to locate in the <see cref="SiteCollection"/>.
        ''' This argument must be compatible with <see cref="Site"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <returns>
        ''' The zero-based index of the first occurrence of <paramref name="value"/>
        ''' in the <see cref="SiteCollection"/>, if found; otherwise, -1.
        ''' </returns>
        ''' <exception cref="InvalidCastException"><paramref name="value"/>
        ''' is not compatible with <see cref="Site"/>.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.IndexOf"/> for details.</remarks>

        Private Function IList_IndexOf(value As Object) As Integer Implements IList.IndexOf
            Return IndexOf(DirectCast(value, Site))
        End Function

#End Region
#Region "Insert(Int32, Site)"

        ''' <summary>
        ''' Inserts a <see cref="Site"/> element into the
        ''' <see cref="SiteCollection"/> at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index at which <paramref name="value"/>
        ''' should be inserted.</param>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to insert into the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is greater than <see cref="Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains the specified
        ''' <paramref name="value"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Insert"/> for details.</remarks>

        Public Overridable Sub Insert(index As Integer, value As Site)
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot be negative.")
            End If

            If index > Me._count Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot exceed Count.")
            End If

            If Me._count = Me._array.Length Then
                EnsureCapacity(Me._count + 1)
            End If

            Me._version += 1
            If index < Me._count Then
                Array.Copy(Me._array, index, Me._array, index + 1, Me._count - index)
            End If

            Me._array(index) = value
            Me._count += 1
        End Sub

#End Region
#Region "IList.Insert(Int32, Object)"

        ''' <summary>
        ''' Inserts an element into the <see cref="SiteCollection"/> at the specified index.
        ''' </summary>
        ''' <param name="index">The zero-based index at which <paramref name="value"/>
        ''' should be inserted.</param>
        ''' <param name="value">The object to insert into the <see cref="SiteCollection"/>.
        ''' This argument must be compatible with <see cref="Site"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is greater than <see cref="Count"/>.</para>
        ''' </exception>
        ''' <exception cref="InvalidCastException"><paramref name="value"/>
        ''' is not compatible with <see cref="Site"/>.</exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> already contains the specified
        ''' <paramref name="value"/>, and the <b>SiteCollection</b>
        ''' ensures that all elements are unique.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Insert"/> for details.</remarks>

        Private Sub IList_Insert(index As Integer, value As Object) Implements IList.Insert
            Insert(index, DirectCast(value, Site))
        End Sub

#End Region
#Region "ReadOnly"

        ''' <summary>
        ''' Returns a read-only wrapper for the specified <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="collection">The <see cref="SiteCollection"/> to wrap.</param>
        ''' <returns>A read-only wrapper around <paramref name="collection"/>.</returns>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="collection"/> is a null reference.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.ReadOnly"/> for details.</remarks>

        Public Shared Function [ReadOnly](collection As SiteCollection) As SiteCollection
            If collection Is Nothing Then
                Throw New ArgumentNullException("collection")
            End If

            Return New ReadOnlyList(collection)
        End Function

#End Region
#Region "Remove(Site)"

        ''' <summary>
        ''' Removes the first occurrence of the specified <see cref="Site"/>
        ''' from the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">The <see cref="Site"/> object
        ''' to remove from the <see cref="SiteCollection"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Remove"/> for details.</remarks>

        Public Overridable Sub Remove(value As Site)
            Dim index As Integer = IndexOf(value)
            If index >= 0 Then
                RemoveAt(index)
            End If
        End Sub

#End Region
#Region "IList.Remove(Object)"

        ''' <summary>
        ''' Removes the first occurrence of the specified <see cref="Object"/>
        ''' from the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="value">The object to remove from the <see cref="SiteCollection"/>.
        ''' This argument must be compatible with <see cref="Site"/>.
        ''' This argument can be a null reference.
        ''' </param>
        ''' <exception cref="InvalidCastException"><paramref name="value"/>
        ''' is not compatible with <see cref="Site"/>.</exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Remove"/> for details.</remarks>

        Private Sub IList_Remove(value As Object) Implements IList.Remove
            Remove(DirectCast(value, Site))
        End Sub

#End Region
#Region "RemoveAt"

        ''' <summary>
        ''' Removes the element at the specified index of the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="index">The zero-based index of the element to remove.</param>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="index"/> is equal to or greater than <see cref="Count"/>.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.RemoveAt"/> for details.</remarks>

        Public Overridable Sub RemoveAt(index As Integer)
            ValidateIndex(index)

            Me._version += 1
            If index < System.Threading.Interlocked.Decrement(Me._count) Then
                Array.Copy(Me._array, index + 1, Me._array, index, Me._count - index)
            End If

            Me._array(Me._count) = Nothing
        End Sub

#End Region
#Region "RemoveRange"

        ''' <summary>
        ''' Removes the specified range of elements from the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="index">The zero-based starting index of the range
        ''' of elements to remove.</param>
        ''' <param name="count">The number of elements to remove.</param>
        ''' <exception cref="ArgumentException">
        ''' <paramref name="index"/> and <paramref name="count"/> do not denote a
        ''' valid range of elements in the <see cref="SiteCollection"/>.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="count"/> is less than zero.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.RemoveRange"/> for details.</remarks>

        Public Overridable Sub RemoveRange(index As Integer, count As Integer)
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot be negative.")
            End If

            If count < 0 Then
                Throw New ArgumentOutOfRangeException("count", count, "Argument cannot be negative.")
            End If

            If index + count > Me._count Then
                Throw New ArgumentException("Arguments denote invalid range of elements.")
            End If

            If count = 0 Then
                Return
            End If

            Me._version += 1
            Me._count -= count

            If index < Me._count Then
                Array.Copy(Me._array, index + count, Me._array, index, Me._count - index)
            End If

            Array.Clear(Me._array, Me._count, count)
        End Sub

#End Region
#Region "Reverse()"

        ''' <overloads>
        ''' Reverses the order of the elements in the 
        ''' <see cref="SiteCollection"/> or a portion of it.
        ''' </overloads>
        ''' <summary>
        ''' Reverses the order of the elements in the entire <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <exception cref="NotSupportedException">
        ''' The <see cref="SiteCollection"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Reverse"/> for details.</remarks>

        Public Overridable Sub Reverse()
            If Me._count <= 1 Then
                Return
            End If
            Me._version += 1
            Array.Reverse(Me._array, 0, Me._count)
        End Sub

#End Region
#Region "Reverse(Int32, Int32)"

        ''' <summary>
        ''' Reverses the order of the elements in the specified range.
        ''' </summary>
        ''' <param name="index">The zero-based starting index of the range
        ''' of elements to reverse.</param>
        ''' <param name="count">The number of elements to reverse.</param>
        ''' <exception cref="ArgumentException">
        ''' <paramref name="index"/> and <paramref name="count"/> do not denote a
        ''' valid range of elements in the <see cref="SiteCollection"/>.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="count"/> is less than zero.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' The <see cref="SiteCollection"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Reverse"/> for details.</remarks>

        Public Overridable Sub Reverse(index As Integer, count As Integer)
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot be negative.")
            End If

            If count < 0 Then
                Throw New ArgumentOutOfRangeException("count", count, "Argument cannot be negative.")
            End If

            If index + count > Me._count Then
                Throw New ArgumentException("Arguments denote invalid range of elements.")
            End If

            If count <= 1 OrElse Me._count <= 1 Then
                Return
            End If
            Me._version += 1
            Array.Reverse(Me._array, index, count)
        End Sub

#End Region
#Region "Sort()"

        ''' <overloads>
        ''' Sorts the elements in the <see cref="SiteCollection"/> or a portion of it.
        ''' </overloads>
        ''' <summary>
        ''' Sorts the elements in the entire <see cref="SiteCollection"/>
        ''' using the <see cref="IComparable"/> implementation of each element.
        ''' </summary>
        ''' <exception cref="NotSupportedException">
        ''' The <see cref="SiteCollection"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Sort"/> for details.</remarks>

        Public Overridable Sub Sort()
            If Me._count <= 1 Then
                Return
            End If
            Me._version += 1
            Array.Sort(Me._array, 0, Me._count)
        End Sub

#End Region
#Region "Sort(IComparer)"

        ''' <summary>
        ''' Sorts the elements in the entire <see cref="SiteCollection"/>
        ''' using the specified <see cref="IComparer"/> interface.
        ''' </summary>
        ''' <param name="comparer">
        ''' <para>The <see cref="IComparer"/> implementation to use when comparing elements.</para>
        ''' <para>-or-</para>
        ''' <para>A null reference to use the <see cref="IComparable"/> implementation 
        ''' of each element.</para></param>
        ''' <exception cref="NotSupportedException">
        ''' The <see cref="SiteCollection"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Sort"/> for details.</remarks>

        Public Overridable Sub Sort(comparer As IComparer)
            If Me._count <= 1 Then
                Return
            End If
            Me._version += 1
            Array.Sort(Me._array, 0, Me._count, comparer)
        End Sub

#End Region
#Region "Sort(Int32, Int32, IComparer)"

        ''' <summary>
        ''' Sorts the elements in the specified range 
        ''' using the specified <see cref="IComparer"/> interface.
        ''' </summary>
        ''' <param name="index">The zero-based starting index of the range
        ''' of elements to sort.</param>
        ''' <param name="count">The number of elements to sort.</param>
        ''' <param name="comparer">
        ''' <para>The <see cref="IComparer"/> implementation to use when comparing elements.</para>
        ''' <para>-or-</para>
        ''' <para>A null reference to use the <see cref="IComparable"/> implementation 
        ''' of each element.</para></param>
        ''' <exception cref="ArgumentException">
        ''' <paramref name="index"/> and <paramref name="count"/> do not denote a
        ''' valid range of elements in the <see cref="SiteCollection"/>.</exception>
        ''' <exception cref="ArgumentOutOfRangeException">
        ''' <para><paramref name="index"/> is less than zero.</para>
        ''' <para>-or-</para>
        ''' <para><paramref name="count"/> is less than zero.</para>
        ''' </exception>
        ''' <exception cref="NotSupportedException">
        ''' The <see cref="SiteCollection"/> is read-only.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Sort"/> for details.</remarks>

        Public Overridable Sub Sort(index As Integer, count As Integer, comparer As IComparer)
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot be negative.")
            End If

            If count < 0 Then
                Throw New ArgumentOutOfRangeException("count", count, "Argument cannot be negative.")
            End If

            If index + count > Me._count Then
                Throw New ArgumentException("Arguments denote invalid range of elements.")
            End If

            If count <= 1 OrElse Me._count <= 1 Then
                Return
            End If
            Me._version += 1
            Array.Sort(Me._array, index, count, comparer)
        End Sub

#End Region
#Region "Synchronized"

        ''' <summary>
        ''' Returns a synchronized (thread-safe) wrapper
        ''' for the specified <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <param name="collection">The <see cref="SiteCollection"/> to synchronize.</param>
        ''' <returns>
        ''' A synchronized (thread-safe) wrapper around <paramref name="collection"/>.
        ''' </returns>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="collection"/> is a null reference.</exception>
        ''' <remarks>Please refer to <see cref="ArrayList.Synchronized"/> for details.</remarks>

        Public Shared Function Synchronized(collection As SiteCollection) As SiteCollection
            If collection Is Nothing Then
                Throw New ArgumentNullException("collection")
            End If

            Return New SyncList(collection)
        End Function

#End Region
#Region "ToArray"

        ''' <summary>
        ''' Copies the elements of the <see cref="SiteCollection"/> to a new
        ''' <see cref="Array"/> of <see cref="Site"/> elements.
        ''' </summary>
        ''' <returns>A one-dimensional <see cref="Array"/> of <see cref="Site"/>
        ''' elements containing copies of the elements of the <see cref="SiteCollection"/>.</returns>
        ''' <remarks>Please refer to <see cref="ArrayList.ToArray"/> for details.</remarks>

        Public Overridable Function ToArray() As Site()
            Dim array__1 As Site() = New Site(Me._count - 1) {}
            Array.Copy(Me._array, array__1, Me._count)
            Return array__1
        End Function

#End Region
#Region "TrimToSize"

        ''' <summary>
        ''' Sets the capacity to the actual number of elements in the <see cref="SiteCollection"/>.
        ''' </summary>
        ''' <exception cref="NotSupportedException">
        ''' <para>The <see cref="SiteCollection"/> is read-only.</para>
        ''' <para>-or-</para>
        ''' <para>The <b>SiteCollection</b> has a fixed size.</para></exception>
        ''' <remarks>Please refer to <see cref="ArrayList.TrimToSize"/> for details.</remarks>

        Public Overridable Sub TrimToSize()
            Capacity = Me._count
        End Sub

#End Region
#Region "Unique"

        ''' <summary>
        ''' Returns a wrapper for the specified <see cref="SiteCollection"/>
        ''' ensuring that all elements are unique.
        ''' </summary>
        ''' <param name="collection">The <see cref="SiteCollection"/> to wrap.</param>    
        ''' <returns>
        ''' A wrapper around <paramref name="collection"/> ensuring that all elements are unique.
        ''' </returns>
        ''' <exception cref="ArgumentException">
        ''' <paramref name="collection"/> contains duplicate elements.</exception>
        ''' <exception cref="ArgumentNullException">
        ''' <paramref name="collection"/> is a null reference.</exception>
        ''' <remarks><para>
        ''' The <b>Unique</b> wrapper provides a set-like collection by ensuring
        ''' that all elements in the <see cref="SiteCollection"/> are unique.
        ''' </para><para>
        ''' <b>Unique</b> raises an <see cref="ArgumentException"/> if the specified 
        ''' <paramref name="collection"/> contains any duplicate elements. The returned
        ''' wrapper raises a <see cref="NotSupportedException"/> whenever the user attempts 
        ''' to add an element that is already contained in the <b>SiteCollection</b>.
        ''' </para><para>
        ''' <strong>Note:</strong> The <b>Unique</b> wrapper reflects any changes made
        ''' to the underlying <paramref name="collection"/>, including the possible
        ''' creation of duplicate elements. The uniqueness of all elements is therefore
        ''' no longer assured if the underlying collection is manipulated directly.
        ''' </para></remarks>

        Public Shared Function Unique(collection As SiteCollection) As SiteCollection
            If collection Is Nothing Then
                Throw New ArgumentNullException("collection")
            End If

            For i As Integer = collection.Count - 1 To 1 Step -1
                If collection.IndexOf(collection(i)) < i Then
                    Throw New ArgumentException("collection", "Argument cannot contain duplicate elements.")
                End If
            Next

            Return New UniqueList(collection)
        End Function

#End Region
#End Region
#Region "Private Methods"
#Region "CheckEnumIndex"

        Private Sub CheckEnumIndex(index As Integer)
            If index < 0 OrElse index >= Me._count Then
                Throw New InvalidOperationException("Enumerator is not on a collection element.")
            End If
        End Sub

#End Region
#Region "CheckEnumVersion"

        Private Sub CheckEnumVersion(version As Integer)
            If version <> Me._version Then
                Throw New InvalidOperationException("Enumerator invalidated by modification to collection.")
            End If
        End Sub

#End Region
#Region "CheckTargetArray"

        Private Sub CheckTargetArray(array As Array, arrayIndex As Integer)
            If array Is Nothing Then
                Throw New ArgumentNullException("array")
            End If
            If array.Rank > 1 Then
                Throw New ArgumentException("Argument cannot be multidimensional.", "array")
            End If

            If arrayIndex < 0 Then
                Throw New ArgumentOutOfRangeException("arrayIndex", arrayIndex, "Argument cannot be negative.")
            End If
            If arrayIndex >= array.Length Then
                Throw New ArgumentException("Argument must be less than array length.", "arrayIndex")
            End If

            If Me._count > array.Length - arrayIndex Then
                Throw New ArgumentException("Argument section must be large enough for collection.", "array")
            End If
        End Sub

#End Region
#Region "EnsureCapacity"

        Private Sub EnsureCapacity(minimum As Integer)
            Dim newCapacity As Integer = (If(Me._array.Length = 0, _defaultCapacity, Me._array.Length * 2))

            If newCapacity < minimum Then
                newCapacity = minimum
            End If
            Capacity = newCapacity
        End Sub

#End Region
#Region "ValidateIndex"

        Private Sub ValidateIndex(index As Integer)
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument cannot be negative.")
            End If

            If index >= Me._count Then
                Throw New ArgumentOutOfRangeException("index", index, "Argument must be less than Count.")
            End If
        End Sub

#End Region
#End Region
#Region "Class Enumerator"

        <Serializable()> _
        Private NotInheritable Class Enumerator
            Implements ISiteEnumerator
            Implements IEnumerator
#Region "Private Fields"

            Private ReadOnly _collection As SiteCollection
            Private ReadOnly _version As Integer
            Private _index As Integer

#End Region
#Region "Internal Constructors"

            Friend Sub New(collection As SiteCollection)
                Me._collection = collection
                Me._version = collection._version
                Me._index = -1
            End Sub

#End Region
#Region "Public Properties"

            Public ReadOnly Property Current() As Site Implements ISiteEnumerator.Current
                Get
                    Me._collection.CheckEnumIndex(Me._index)
                    Me._collection.CheckEnumVersion(Me._version)
                    Return Me._collection(Me._index)
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current() As Object Implements IEnumerator.Current
                Get
                    Return Current
                End Get
            End Property

#End Region
#Region "Public Methods"

            Public Function MoveNext() As Boolean Implements ISiteEnumerator.MoveNext
                Me._collection.CheckEnumVersion(Me._version)
                Return (System.Threading.Interlocked.Increment(Me._index) < Me._collection.Count)
            End Function

            Public Sub Reset() Implements ISiteEnumerator.Reset
                Me._collection.CheckEnumVersion(Me._version)
                Me._index = -1
            End Sub

#End Region
        End Class

#End Region
#Region "Class ReadOnlyList"

        <Serializable()> _
        Private NotInheritable Class ReadOnlyList
            Inherits SiteCollection
#Region "Private Fields"

            Private _collection As SiteCollection

#End Region
#Region "Internal Constructors"

            Friend Sub New(collection As SiteCollection)
                MyBase.New(Tag.[Default])
                Me._collection = collection
            End Sub

#End Region
#Region "Protected Properties"

            Protected Overrides ReadOnly Property InnerArray() As Site()
                Get
                    Return Me._collection.InnerArray
                End Get
            End Property

#End Region
#Region "Public Properties"

            Public Overrides Property Capacity() As Integer
                Get
                    Return Me._collection.Capacity
                End Get
                Set(value As Integer)
                    Throw New NotSupportedException("Read-only collections cannot be modified.")
                End Set
            End Property

            Public Overrides ReadOnly Property Count() As Integer
                Get
                    Return Me._collection.Count
                End Get
            End Property

            Public Overrides ReadOnly Property IsFixedSize() As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides ReadOnly Property IsSynchronized() As Boolean
                Get
                    Return Me._collection.IsSynchronized
                End Get
            End Property

            Public Overrides ReadOnly Property IsUnique() As Boolean
                Get
                    Return Me._collection.IsUnique
                End Get
            End Property

            Default Public Overrides Property Item(index As Integer) As Site
                Get
                    Return Me._collection(index)
                End Get
                Set(value As Site)
                    Throw New NotSupportedException("Read-only collections cannot be modified.")
                End Set
            End Property

            Public Overrides ReadOnly Property SyncRoot() As Object
                Get
                    Return Me._collection.SyncRoot
                End Get
            End Property

#End Region
#Region "Public Methods"

            Public Overrides Function Add(value As Site) As Integer
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Function

            Public Overrides Sub AddRange(collection As SiteCollection)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub AddRange(array As Site())
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Function BinarySearch(value As Site) As Integer
                Return Me._collection.BinarySearch(value)
            End Function

            Public Overrides Sub Clear()
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Function Clone() As Object
                Return New ReadOnlyList(DirectCast(Me._collection.Clone(), SiteCollection))
            End Function

            Public Overrides Sub CopyTo(array As Site())
                Me._collection.CopyTo(array)
            End Sub

            Public Overrides Sub CopyTo(array As Site(), arrayIndex As Integer)
                Me._collection.CopyTo(array, arrayIndex)
            End Sub

            Public Overrides Function GetEnumerator() As ISiteEnumerator
                Return Me._collection.GetEnumerator()
            End Function

            Public Overrides Function IndexOf(value As Site) As Integer
                Return Me._collection.IndexOf(value)
            End Function

            Public Overrides Sub Insert(index As Integer, value As Site)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Remove(value As Site)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub RemoveAt(index As Integer)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub RemoveRange(index As Integer, count As Integer)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Reverse()
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Reverse(index As Integer, count As Integer)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Sort()
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Sort(comparer As IComparer)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Sub Sort(index As Integer, count As Integer, comparer As IComparer)
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

            Public Overrides Function ToArray() As Site()
                Return Me._collection.ToArray()
            End Function

            Public Overrides Sub TrimToSize()
                Throw New NotSupportedException("Read-only collections cannot be modified.")
            End Sub

#End Region
        End Class

#End Region
#Region "Class SyncList"

        <Serializable()> _
        Private NotInheritable Class SyncList
            Inherits SiteCollection
#Region "Private Fields"

            Private _collection As SiteCollection
            Private _root As Object

#End Region
#Region "Internal Constructors"

            Friend Sub New(collection As SiteCollection)
                MyBase.New(Tag.[Default])

                Me._root = collection.SyncRoot
                Me._collection = collection
            End Sub

#End Region
#Region "Protected Properties"

            Protected Overrides ReadOnly Property InnerArray() As Site()
                Get
                    SyncLock Me._root
                        Return Me._collection.InnerArray
                    End SyncLock
                End Get
            End Property

#End Region
#Region "Public Properties"

            Public Overrides Property Capacity() As Integer
                Get
                    SyncLock Me._root
                        Return Me._collection.Capacity
                    End SyncLock
                End Get
                Set(value As Integer)
                    SyncLock Me._root
                        Me._collection.Capacity = value
                    End SyncLock
                End Set
            End Property

            Public Overrides ReadOnly Property Count() As Integer
                Get
                    SyncLock Me._root
                        Return Me._collection.Count
                    End SyncLock
                End Get
            End Property

            Public Overrides ReadOnly Property IsFixedSize() As Boolean
                Get
                    Return Me._collection.IsFixedSize
                End Get
            End Property

            Public Overrides ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return Me._collection.IsReadOnly
                End Get
            End Property

            Public Overrides ReadOnly Property IsSynchronized() As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides ReadOnly Property IsUnique() As Boolean
                Get
                    Return Me._collection.IsUnique
                End Get
            End Property

            Default Public Overrides Property Item(index As Integer) As Site
                Get
                    SyncLock Me._root
                        Return Me._collection(index)
                    End SyncLock
                End Get
                Set(value As Site)
                    SyncLock Me._root
                        Me._collection(index) = value
                    End SyncLock
                End Set
            End Property

            Public Overrides ReadOnly Property SyncRoot() As Object
                Get
                    Return Me._root
                End Get
            End Property

#End Region
#Region "Public Methods"

            Public Overrides Function Add(value As Site) As Integer
                SyncLock Me._root
                    Return Me._collection.Add(value)
                End SyncLock
            End Function

            Public Overrides Sub AddRange(collection As SiteCollection)
                SyncLock Me._root
                    Me._collection.AddRange(collection)
                End SyncLock
            End Sub

            Public Overrides Sub AddRange(array As Site())
                SyncLock Me._root
                    Me._collection.AddRange(array)
                End SyncLock
            End Sub

            Public Overrides Function BinarySearch(value As Site) As Integer
                SyncLock Me._root
                    Return Me._collection.BinarySearch(value)
                End SyncLock
            End Function

            Public Overrides Sub Clear()
                SyncLock Me._root
                    Me._collection.Clear()
                End SyncLock
            End Sub

            Public Overrides Function Clone() As Object
                SyncLock Me._root
                    Return New SyncList(DirectCast(Me._collection.Clone(), SiteCollection))
                End SyncLock
            End Function

            Public Overrides Sub CopyTo(array As Site())
                SyncLock Me._root
                    Me._collection.CopyTo(array)
                End SyncLock
            End Sub

            Public Overrides Sub CopyTo(array As Site(), arrayIndex As Integer)
                SyncLock Me._root
                    Me._collection.CopyTo(array, arrayIndex)
                End SyncLock
            End Sub

            Public Overrides Function GetEnumerator() As ISiteEnumerator
                SyncLock Me._root
                    Return Me._collection.GetEnumerator()
                End SyncLock
            End Function

            Public Overrides Function IndexOf(value As Site) As Integer
                SyncLock Me._root
                    Return Me._collection.IndexOf(value)
                End SyncLock
            End Function

            Public Overrides Sub Insert(index As Integer, value As Site)
                SyncLock Me._root
                    Me._collection.Insert(index, value)
                End SyncLock
            End Sub

            Public Overrides Sub Remove(value As Site)
                SyncLock Me._root
                    Me._collection.Remove(value)
                End SyncLock
            End Sub

            Public Overrides Sub RemoveAt(index As Integer)
                SyncLock Me._root
                    Me._collection.RemoveAt(index)
                End SyncLock
            End Sub

            Public Overrides Sub RemoveRange(index As Integer, count As Integer)
                SyncLock Me._root
                    Me._collection.RemoveRange(index, count)
                End SyncLock
            End Sub

            Public Overrides Sub Reverse()
                SyncLock Me._root
                    Me._collection.Reverse()
                End SyncLock
            End Sub

            Public Overrides Sub Reverse(index As Integer, count As Integer)
                SyncLock Me._root
                    Me._collection.Reverse(index, count)
                End SyncLock
            End Sub

            Public Overrides Sub Sort()
                SyncLock Me._root
                    Me._collection.Sort()
                End SyncLock
            End Sub

            Public Overrides Sub Sort(comparer As IComparer)
                SyncLock Me._root
                    Me._collection.Sort(comparer)
                End SyncLock
            End Sub

            Public Overrides Sub Sort(index As Integer, count As Integer, comparer As IComparer)
                SyncLock Me._root
                    Me._collection.Sort(index, count, comparer)
                End SyncLock
            End Sub

            Public Overrides Function ToArray() As Site()
                SyncLock Me._root
                    Return Me._collection.ToArray()
                End SyncLock
            End Function

            Public Overrides Sub TrimToSize()
                SyncLock Me._root
                    Me._collection.TrimToSize()
                End SyncLock
            End Sub

#End Region
        End Class

#End Region
#Region "Class UniqueList"

        <Serializable()> _
        Private NotInheritable Class UniqueList
            Inherits SiteCollection
#Region "Private Fields"

            Private _collection As SiteCollection

#End Region
#Region "Internal Constructors"

            Friend Sub New(collection As SiteCollection)
                MyBase.New(Tag.[Default])
                Me._collection = collection
            End Sub

#End Region
#Region "Protected Properties"

            Protected Overrides ReadOnly Property InnerArray() As Site()
                Get
                    Return Me._collection.InnerArray
                End Get
            End Property

#End Region
#Region "Public Properties"

            Public Overrides Property Capacity() As Integer
                Get
                    Return Me._collection.Capacity
                End Get
                Set(value As Integer)
                    Me._collection.Capacity = value
                End Set
            End Property

            Public Overrides ReadOnly Property Count() As Integer
                Get
                    Return Me._collection.Count
                End Get
            End Property

            Public Overrides ReadOnly Property IsFixedSize() As Boolean
                Get
                    Return Me._collection.IsFixedSize
                End Get
            End Property

            Public Overrides ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return Me._collection.IsReadOnly
                End Get
            End Property

            Public Overrides ReadOnly Property IsSynchronized() As Boolean
                Get
                    Return Me._collection.IsSynchronized
                End Get
            End Property

            Public Overrides ReadOnly Property IsUnique() As Boolean
                Get
                    Return True
                End Get
            End Property

            Default Public Overrides Property Item(index As Integer) As Site
                Get
                    Return Me._collection(index)
                End Get
                Set(value As Site)
                    CheckUnique(index, value)
                    Me._collection(index) = value
                End Set
            End Property

            Public Overrides ReadOnly Property SyncRoot() As Object
                Get
                    Return Me._collection.SyncRoot
                End Get
            End Property

#End Region
#Region "Public Methods"

            Public Overrides Function Add(value As Site) As Integer
                CheckUnique(value)
                Return Me._collection.Add(value)
            End Function

            Public Overrides Sub AddRange(collection As SiteCollection)
                For Each value As Site In collection
                    CheckUnique(value)
                Next

                Me._collection.AddRange(collection)
            End Sub

            Public Overrides Sub AddRange(array As Site())
                For Each value As Site In array
                    CheckUnique(value)
                Next

                Me._collection.AddRange(array)
            End Sub

            Public Overrides Function BinarySearch(value As Site) As Integer
                Return Me._collection.BinarySearch(value)
            End Function

            Public Overrides Sub Clear()
                Me._collection.Clear()
            End Sub

            Public Overrides Function Clone() As Object
                Return New UniqueList(DirectCast(Me._collection.Clone(), SiteCollection))
            End Function

            Public Overrides Sub CopyTo(array As Site())
                Me._collection.CopyTo(array)
            End Sub

            Public Overrides Sub CopyTo(array As Site(), arrayIndex As Integer)
                Me._collection.CopyTo(array, arrayIndex)
            End Sub

            Public Overrides Function GetEnumerator() As ISiteEnumerator
                Return Me._collection.GetEnumerator()
            End Function

            Public Overrides Function IndexOf(value As Site) As Integer
                Return Me._collection.IndexOf(value)
            End Function

            Public Overrides Sub Insert(index As Integer, value As Site)
                CheckUnique(value)
                Me._collection.Insert(index, value)
            End Sub

            Public Overrides Sub Remove(value As Site)
                Me._collection.Remove(value)
            End Sub

            Public Overrides Sub RemoveAt(index As Integer)
                Me._collection.RemoveAt(index)
            End Sub

            Public Overrides Sub RemoveRange(index As Integer, count As Integer)
                Me._collection.RemoveRange(index, count)
            End Sub

            Public Overrides Sub Reverse()
                Me._collection.Reverse()
            End Sub

            Public Overrides Sub Reverse(index As Integer, count As Integer)
                Me._collection.Reverse(index, count)
            End Sub

            Public Overrides Sub Sort()
                Me._collection.Sort()
            End Sub

            Public Overrides Sub Sort(comparer As IComparer)
                Me._collection.Sort(comparer)
            End Sub

            Public Overrides Sub Sort(index As Integer, count As Integer, comparer As IComparer)
                Me._collection.Sort(index, count, comparer)
            End Sub

            Public Overrides Function ToArray() As Site()
                Return Me._collection.ToArray()
            End Function

            Public Overrides Sub TrimToSize()
                Me._collection.TrimToSize()
            End Sub

#End Region
#Region "Private Methods"

            Private Sub CheckUnique(value As Site)
                If IndexOf(value) >= 0 Then
                    Throw New NotSupportedException("Unique collections cannot contain duplicate elements.")
                End If
            End Sub

            Private Sub CheckUnique(index As Integer, value As Site)
                Dim existing As Integer = IndexOf(value)
                If existing >= 0 AndAlso existing <> index Then
                    Throw New NotSupportedException("Unique collections cannot contain duplicate elements.")
                End If
            End Sub

#End Region
        End Class

#End Region
    End Class

#End Region
End Namespace