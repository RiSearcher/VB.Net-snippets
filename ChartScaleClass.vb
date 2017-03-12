Public Class ChartScaleBounds

    ' Calculates nice-looking values for chart scale bounds and number of intervals
    ' Final number of intervals will be closest to the requested (optimal) number

    ' Input parameters
    Private _lower_data_bound As Double
    Private _upper_data_bound As Double
    Private _optimal_num As Integer

    ' Results
    Private _min As Double
    Private _max As Double
    Private _n As Integer
    Private _interval As Double

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="lower_data_bound">Lower data bound</param>
    ''' <param name="upper_data_bound">Upper data bound</param>
    ''' <param name="optimal_num">Optimal number of intervals</param>
    Public Sub New(lower_data_bound As Double, upper_data_bound As Double, optimal_num As Integer)
        _lower_data_bound = lower_data_bound
        _upper_data_bound = upper_data_bound
        _optimal_num = optimal_num

        Call Calc()

    End Sub

    ''' <summary>
    ''' Minimum value of data range (input)
    ''' </summary>
    Public Property LowerDataBound As Double
        Get
            Return _lower_data_bound
        End Get
        Set(value As Double)
            _lower_data_bound = value
            Call Calc()
        End Set
    End Property

    ''' <summary>
    ''' Maximum value of data range (input)
    ''' </summary>
    Public Property UpperDataBound As Double
        Get
            Return _upper_data_bound
        End Get
        Set(value As Double)
            _upper_data_bound = value
            Call Calc()
        End Set
    End Property

    ''' <summary>
    ''' Desired number of intervals (input)
    ''' </summary>
    Public Property OptimalN As Integer
        Get
            Return _optimal_num
        End Get
        Set(value As Integer)
            _optimal_num = value
            Call Calc()
        End Set
    End Property

    ''' <summary>
    ''' Lower scale bound (output)
    ''' </summary>
    Public ReadOnly Property Min As Double
        Get
            Return _min
        End Get
    End Property

    ''' <summary>
    ''' Upper scale bound (output)
    ''' </summary>
    Public ReadOnly Property Max As Double
        Get
            Return _max
        End Get
    End Property

    ''' <summary>
    ''' Number of intervals (output)
    ''' </summary>
    Public ReadOnly Property N As Integer
        Get
            Return _n
        End Get
    End Property

    ''' <summary>
    ''' Size of the interval (output)
    ''' </summary>
    Public ReadOnly Property Interval As Double
        Get
            Return _interval
        End Get
    End Property

    Private Sub Calc()

        Dim MinStep, s(4), tmp_num(4), tmp_min(4), tmp_max(4) As Double
        MinStep = (_upper_data_bound - _lower_data_bound) / _optimal_num

        s(0) = 10 ^ (Math.Ceiling(Math.Log10(MinStep)))
        s(1) = 10 ^ (Math.Ceiling(Math.Log10(MinStep))) / 2
        s(2) = 10 ^ (Math.Ceiling(Math.Log10(MinStep))) / 4
        s(3) = 10 ^ (Math.Ceiling(Math.Log10(MinStep))) / 5
        s(4) = 10 ^ (Math.Ceiling(Math.Log10(MinStep))) / 10

        For i As Integer = 0 To 4
            tmp_min(i) = s(i) * Math.Floor(_lower_data_bound / s(i))
            tmp_max(i) = s(i) * Math.Ceiling(_upper_data_bound / s(i))
            tmp_num(i) = (tmp_max(i) - tmp_min(i)) / s(i)
        Next

        Dim best, diff As Integer
        best = 0
        diff = Math.Abs(_optimal_num - tmp_num(0))
        For i As Integer = 1 To 4
            If Math.Abs(_optimal_num - tmp_num(i)) < diff Then
                diff = Math.Abs(_optimal_num - tmp_num(i))
                best = i
            ElseIf (Math.Abs(_optimal_num - tmp_num(i)) = diff) And (tmp_num(i) < _optimal_num) Then  '// choose smaller number of intevals
                diff = Math.Abs(_optimal_num - tmp_num(i))
                best = i
            End If
        Next

        _n = tmp_num(best)
        _min = tmp_min(best)
        _max = tmp_max(best)
        _interval = (_max - _min) / _n

    End Sub

End Class
