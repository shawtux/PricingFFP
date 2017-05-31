Imports PricingFFP
Imports Microsoft.Office.Interop.Excel

Public Class faredata
    Implements IComparable(Of faredata)
    Private _AP As String
    Private _bookingClass As String
    Private _completedBy As String
    Private _dates As String
    Private _days As String
    Private _F As String
    Private _farebasis As String
    Private _FTC As String
    Private _issueAfter As String
    Private _issueBefore As String
    Private _line As Integer
    Private _max As String
    Private _min As String
    Private _owfare As String
    Private _pen As String
    Private _R As String
    Private _restrictionAfter As String
    Private _restrictionBefore As String
    Private _rtfare As String

    Private _OD As String

    Private _accoundCode As String
    Public Property accountCode() As String
        Get
            Return _accoundCode
        End Get
        Set(ByVal value As String)
            _accoundCode = value
        End Set
    End Property

    Public Property OD() As String
        Get
            Return _OD
        End Get
        Set(ByVal value As String)
            _OD = value
        End Set
    End Property

    Private _seasonsTo As String
    Public Property seasonsTo() As String
        Get
            Return _seasonsTo
        End Get
        Set(ByVal value As String)
            _seasonsTo = value
        End Set
    End Property

    Private _seasonsFrom As String
    Public Property seasonsFrom() As String
        Get
            Return _seasonsFrom
        End Get
        Set(ByVal value As String)
            _seasonsFrom = value
        End Set
    End Property

    Private _blackoutOutbound As String
    Public Property blackoutOutbound() As String
        Get
            Return _blackoutOutbound
        End Get
        Set(ByVal value As String)
            _blackoutOutbound = value
        End Set
    End Property
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="lineNum">numero de linea</param>
    ''' <param name="farebasis">farebasis</param>
    ''' <param name="owFare">oneway fare</param>
    ''' <param name="rtFare">round trip fare</param>
    ''' <param name="bookingClass">booking class</param>
    ''' <param name="penColumn">pen column</param>
    ''' <param name="notUsed">not used yet</param>
    ''' <param name="dateStr">date</param>
    ''' <param name="dayStr">days</param>
    ''' <param name="notUsed1">not used yet</param>
    ''' <param name="ap">ap</param>
    ''' <param name="notUsed2">not used yet</param>
    ''' <param name="minStay">min stay</param>
    ''' <param name="notUsed3">not used yet</param>
    ''' <param name="maxStay">max stay</param>
    ''' <param name="fColumn">F column</param>
    ''' <param name="rColumn">R column</param>
    Public Sub New(lineNum As String, farebasis As String, owFare As String, rtFare As String, bookingClass As String, penColumn As String,
                   notUsed As String, dateStr As String, dayStr As String, notUsed1 As String, ap As String, notUsed2 As String, minStay As String,
                   notUsed3 As String, maxStay As String, fColumn As String, rColumn As String, od As String)
        'Console.WriteLine("value1:" & lineNum)
        'Console.WriteLine("value2:" & farebasis)
        'Console.WriteLine("value3:" & owFare)
        'Console.WriteLine("value4:" & rtFare)
        'Console.WriteLine("value5:" & bookingClass)
        'Console.WriteLine("value6:" & penColumn)
        'Console.WriteLine("value7:" & notUsed)
        'Console.WriteLine("value8:" & dateStr)
        'Console.WriteLine("value9:" & dayStr)
        'Console.WriteLine("value10:" & notUsed1)
        'Console.WriteLine("value11:" & ap)
        'Console.WriteLine("value12:" & notUsed2)
        'Console.WriteLine("value13:" & minStay)
        'Console.WriteLine("value14:" & notUsed3)
        'Console.WriteLine("value15:" & maxStay)
        'Console.WriteLine("value16:" & fColumn)
        'Console.WriteLine("value17:" & rColumn)

        Me.line = lineNum
        Me.farebasis = farebasis
        Me.owfare = owFare
        Me.rtfare = rtFare
        Me.bookingClass = bookingClass
        Me.pen = penColumn

        Me.dates = dateStr
        Me.days = dayStr

        Me.AP = ap

        Me.min = minStay

        Me.max = maxStay
        Me.F = fColumn
        Me.R = rColumn

        Me.OD = od
        Me.seasonsFrom = ""
        Me.seasonsTo = ""

    End Sub

    Public Sub New()

    End Sub

    Public Property AP() As String
        Get
            Return _AP
        End Get
        Set(ByVal value As String)
            _AP = value
        End Set
    End Property

    Public Property bookingClass() As String
        Get
            Return _bookingClass
        End Get
        Set(ByVal value As String)
            _bookingClass = value
        End Set
    End Property

    Public Property completedBy() As String
        Get
            Return _completedBy
        End Get
        Set(ByVal value As String)
            _completedBy = value
        End Set
    End Property

    Public Property dates() As String
        Get
            Return _dates
        End Get
        Set(ByVal value As String)
            _dates = value
        End Set
    End Property

    Public Property days() As String
        Get
            Return _days
        End Get
        Set(ByVal value As String)
            _days = value
        End Set
    End Property

    Public Property F() As String
        Get
            Return _F
        End Get
        Set(ByVal value As String)
            _F = value
        End Set
    End Property

    Public Property farebasis() As String
        Get
            Return _farebasis
        End Get
        Set(ByVal value As String)
            _farebasis = value
        End Set
    End Property

    Public Property FTC() As String
        Get
            Return _FTC
        End Get
        Set(ByVal value As String)
            _FTC = value
        End Set
    End Property

    Public Property issueAfter() As String
        Get
            Return _issueAfter
        End Get
        Set(ByVal value As String)
            _issueAfter = value
        End Set
    End Property

    Public Property issueBefore() As String
        Get
            Return _issueBefore
        End Get
        Set(ByVal value As String)
            _issueBefore = value
        End Set
    End Property

    Public Property line() As Integer
        Get
            Return _line
        End Get
        Set(ByVal value As Integer)
            _line = value
        End Set
    End Property

    Public Property max() As String
        Get
            Return _max
        End Get
        Set(ByVal value As String)
            _max = value
        End Set
    End Property

    Public Property min() As String
        Get
            Return _min
        End Get
        Set(ByVal value As String)
            _min = value
        End Set
    End Property

    Public Property owfare() As String
        Get
            Return _owfare
        End Get
        Set(ByVal value As String)
            _owfare = value
        End Set
    End Property

    Public Property pen() As String
        Get
            Return _pen
        End Get
        Set(ByVal value As String)
            _pen = value
        End Set
    End Property

    Public Property R() As String
        Get
            Return _R
        End Get
        Set(ByVal value As String)
            _R = value
        End Set
    End Property

    Public Property restrictionAfter() As String
        Get
            Return _restrictionAfter
        End Get
        Set(ByVal value As String)
            _restrictionAfter = value
        End Set
    End Property
    Public Property restrictionBefore() As String
        Get
            Return _restrictionBefore
        End Get
        Set(ByVal value As String)
            _restrictionBefore = value
        End Set
    End Property
    Public Property rtfare() As String
        Get
            Return _rtfare
        End Get
        Set(ByVal value As String)
            _rtfare = value
        End Set
    End Property

    Private _blackoutInbound As String
    Public Property blackoutInbound() As String
        Get
            Return _blackoutInbound
        End Get
        Set(ByVal value As String)
            _blackoutInbound = value
        End Set
    End Property

    ''' <summary>
    ''' If two faredatas are equal return 1, else 0
    ''' </summary>
    ''' <param name="other"></param>
    ''' <returns></returns>
    Public Function CompareTo(other As faredata) As Integer Implements IComparable(Of faredata).CompareTo
        If Me.line = other.line And
            Me.farebasis = other.line And
            Me.owfare = other.owfare And
            Me.rtfare = other.rtfare And
            Me.pen = other.pen And
            Me.dates = other.dates And
            Me.days = other.days And
            Me.AP = other.AP And
            Me.min = other.min And
            Me.max = other.max And
            Me.F = other.F And
            Me.R = other.R And
            Me.OD = other.OD And
            Me.accountCode = other.accountCode Then



            Return 1
        Else
            Return 0
        End If

    End Function

    Public Function writeExcelRow() As String(,)
        Dim arrayRow(0, 23) As String
        arrayRow(0, 0) = Me.line
        arrayRow(0, 1) = Me.farebasis
        arrayRow(0, 2) = Me.OD
        arrayRow(0, 3) = Me.owfare
        arrayRow(0, 4) = Me.rtfare
        arrayRow(0, 5) = Me.bookingClass
        arrayRow(0, 6) = Me.pen
        arrayRow(0, 7) = Me.AP
        arrayRow(0, 8) = Me.min
        arrayRow(0, 9) = Me.max
        arrayRow(0, 10) = Me.F
        arrayRow(0, 11) = Me.R
        arrayRow(0, 12) = Me.issueAfter
        arrayRow(0, 13) = Me.issueBefore
        arrayRow(0, 14) = Me.completedBy
        arrayRow(0, 15) = Me.restrictionAfter
        arrayRow(0, 16) = Me.restrictionBefore
        arrayRow(0, 17) = Me.FTC
        arrayRow(0, 18) = Me.blackoutOutbound
        arrayRow(0, 19) = Me.blackoutInbound
        arrayRow(0, 20) = Me.seasonsFrom
        arrayRow(0, 21) = Me.seasonsTo
        arrayRow(0, 22) = Me.accountCode

        Return arrayRow

    End Function

    Public Function writeExcelRowHeader() As String(,)
        Dim arrayRow(0, 23) As String
        arrayRow(0, 0) = "line"
        arrayRow(0, 1) = "farebasis"
        arrayRow(0, 2) = "od"
        arrayRow(0, 3) = "owfare"
        arrayRow(0, 4) = "rtfare"
        arrayRow(0, 5) = "bookingClass"
        arrayRow(0, 6) = "pen"
        arrayRow(0, 7) = "AP"
        arrayRow(0, 8) = "min"
        arrayRow(0, 9) = "max"
        arrayRow(0, 10) = "F"
        arrayRow(0, 11) = "R"
        arrayRow(0, 12) = "ISSUE On/After (A)"
        arrayRow(0, 13) = "ISSUE On/Before (B)"
        arrayRow(0, 14) = "TVL RESTRICTION COMPLETED BY (C)"
        arrayRow(0, 15) = "TVL RESTRICTION On/After (E)"
        arrayRow(0, 16) = "TVL RESTRICTION On/Before (O)"
        arrayRow(0, 17) = "FTC"
        arrayRow(0, 18) = "Blackout Outbound"
        arrayRow(0, 19) = "Blackout Inbound"
        arrayRow(0, 20) = "Seasons From"
        arrayRow(0, 21) = "Seasons To"
        arrayRow(0, 22) = "Account Code"
        Return arrayRow

    End Function

End Class