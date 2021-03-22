Attribute VB_Name = "MLGamma"
Option Explicit
''' <summary>Calculates the value of the digamma fuction (logarithmic derivative of the gamma function) for a given positive input value.</summary>
''' <param name="z">Input value >0 for the digamma function.</param>
''' <returns>Value of the digamma function.</returns>
''' <remarks>Based on John Burkardt's implementation of Jose Bernardo's algorithm.</remarks>
''' <see cref="https://people.math.sc.edu/Burkardt/py_src/asa103/asa103.html" />
Private Function Digamma(z As Double) As Double

    Dim result As Double
    Dim x As Double
    Dim r As Double
    
    If (z <= 0#) Then
        Err.Raise Number:=2000, Description:="Numbers <= 0 not supported."
    End If
    
    If (z <= 0.000001) Then
        Digamma = -0.577215664901533 - 1# / z + 1.64493406684823 * z
        Exit Function
    End If
    
    result = 0#
    x = z
    
    While (x < 8.5)
        result = result - 1# / x
        x = x + 1#
    Wend
    
    r = 1# / x
    result = result + Log(x) - 0.5 * r
    r = r * r
    
    result = result - r * (1# / 12# - r * (1# / 120# - r * (1# / 252#)))
    Digamma = result

End Function

''' <summary>Calculates the value of the trigamma fuction (second logarithmic derivative of the gamma function) for a given positive input value.</summary>
''' <param name="z">Input value >0 for the trigamma function.</param>
''' <returns>Value of the trigamma function.</returns>
''' <remarks>Based on John Burkardt's implementation of Jose Bernardo's algorithm.</remarks>
''' <see cref="https://people.math.sc.edu/Burkardt/f_src/asa121/asa121.html" />
Private Function Trigamma(z As Double) As Double

    Dim result As Double
    Dim x As Double
    Dim r As Double
    
    If (z <= 0#) Then
        Err.Raise Number:=2000, Description:="Numbers <= 0 not supported."
    End If
    
    If (z <= 0.0001) Then
        Trigamma = 1# / z / z
        Exit Function
    End If
    
    result = 0#
    x = z
    
    While (x < 5#)
        result = result + 1# / x / x
        x = x + 1#
    Wend
    
    r = 1# / x / x
    
    result = result + 0.5 * r + (1# + r * (1# / 6# + r * (-1# / 30#))) / x
    
    Trigamma = result

End Function

''' <summary>Calculates the maximum likelihood estimator for theshape paramater of a gamma distribution.</summary>
''' <param name="x">One dimensional range in a worksheet containing numbers that are supposed to be gamma distributed.</param>
''' <param name="tolerance">The accuracy of the parameter's iterative calculation.</param>
''' <returns>The maximum likelihood estimator of a gamma distribution's shape parameter. </returns>
''' <reamrks>Implementation of Thomas P. Minka's algorithm. </remarks>
''' <see cref="https://tminka.github.io/papers/minka-gamma.pdf" />
Public Function GammaMLAlpha(x As Range, Optional tolerance As Double = 0.000001) As Double

    Dim i As Long
    Dim avgLogX As Double
    Dim logAvgX As Double
    Dim alpha As Double
    Dim alphaPrev As Double
    Dim xi As Range
    
    If (x.Columns.Count > 1 And x.Rows.Count > 1) Then
        Err.Raise Number:=2001, Description:="Only one dimensional arrays allowed as input."
    End If
    
    i = 0
    avgLogX = 0#
    logAvgX = 0#
    
    For Each xi In x
        If (Not IsNumeric(xi.Value)) Then
            Err.Raise Number:=2002, Description:="Cannot process non-numeric input values for ML estimation."
        End If
        
        If (xi.Value <= 0) Then
            Err.Raise Number:=2003, Description:="Value outside of the gamma distribution's support (0, +inf)"
        End If
        
        i = i + 1
        avgLogX = avgLogX + Log(xi)
        logAvgX = logAvgX + xi
    Next xi
    
    avgLogX = avgLogX / i
    logAvgX = Log(logAvgX / i)
    
    alpha = 0.5 / (logAvgX - avgLogX)
    alphaPrev = 0#
    
    Do While (Abs(alphaPrev - alpha) > tolerance)
        alphaPrev = alpha
        alpha = 1# / (1# / alpha + (avgLogX - logAvgX + Log(alpha) - Digamma(alpha)) / (alpha * alpha * (1# / alpha - Trigamma(alpha))))
    Loop
    
    GammaMLAlpha = alpha
    
End Function

'''<summary>Calculates the maximum likelihood estimator for a gamma distribution's scale parameter.</summary>
'''<param name="mean">Mean of the values which are supposed to be gamma distributed.</param>
'''<param name="alpha">Shape parameter of the assumed gamma distribution.</param>
'''<returns>The maximum likelihood estimator for a gamma distribution's scale parameter.</returns>
Public Function GammaMLBeta(mean As Double, alpha As Double) As Double

    If (mean <= 0) Then
       Err.Raise Number:=2003, Description:="Value outside of the gamma distribution's support (0, +inf)"
    End If
    
    GammaMLBeta = mean / alpha

End Function
