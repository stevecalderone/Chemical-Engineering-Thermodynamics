Attribute VB_Name = "Math"
Option Explicit


' Obtained from => https://mathformeremortals.wordpress.com/
'************************** Module 8 *****************************************
'************* Extra functions *********************************
Function Interpolate(X, x0 As Range, y0 As Range)
Dim n, i, j, k As Integer
  'Check that rows are same size'
  If (x0.Cells.Count <> y0.Cells.Count) Then
    MsgBox ("X and Y vector to interpolate command has to be same size!")
  End If
  
  n = x0.Cells.Count
  
  'Check that x0 are increasing'
  For i = 1 To n - 1
    j = x0(i).Value
    k = x0(i + 1).Value
    If j > k Then
      MsgBox ("X vector to interpolate command has to be increasing!")
      Return
    End If
  Next i
  
  'Check if x<x0(1)'
  If X < x0(1).Value Then
    k = (y0(2).Value - y0(1).Value) / (x0(2).Value - x0(1).Value)
    Interpolate = y0(1).Value + (X - x0(1).Value) * k
  'Check if X0>x0(END)'
  ElseIf X > x0(n).Value Then
    k = (y0(n).Value - y0(n - 1).Value) / (x0(n).Value - x0(n - 1).Value)
    Interpolate = y0(n).Value + (X - x0(n).Value) * k
  Else
   'Loop through values and find where the value are'
   For i = 1 To n
     If X <= x0(i).Value Then
       If (x0(i).Value - x0(i - 1).Value) <> 0 Then
         k = (y0(i).Value - y0(i - 1).Value) / (x0(i).Value - x0(i - 1).Value)
         Interpolate = y0(i).Value + (X - x0(i).Value) * k
       Else
         Interpolate = y0(i).Value + x0(i).Value
       End If
       Exit For
     End If
   Next i
  End If
End Function

'obtained from => https://mathformeremortals.wordpress.com/
' Returns an interpolated point using local bicubic interpolation on table Table.
' The top row and left column in Table must be headers.
Public Function BicubicInterpolation(Table As Range, TopPoint As Double, LeftPoint As Double) As Double
  Dim LeftMinIndex As Long
  Dim TopMinIndex As Long
  Dim i As Long
  Dim j As Long
  Dim Numerator As Double
  Dim Denominator As Double
  Dim Weights(1 To 4) As Double
  Dim Subset(1 To 4) As Double
  ' Choose TopIndex and LeftIndex that yield the 4x4 subset that we will interpolate over...
  ' Which index is the lowest on the top side?
  TopMinIndex = FindIndexBelow(Table.Rows(1), TopPoint)
  ' The leftmost item should be invalid, so the return value should be higher than 1.
  If TopMinIndex <= 2 Then
    ' Slide the range over to the right if it is lower than the source data domain.
    TopMinIndex = 3
  End If
  ' Slide the range over to the left if it is higher than the source data domain.
  If TopMinIndex >= Table.Columns.Count - 1 Then
    TopMinIndex = Table.Columns.Count - 2
  End If
  ' Which index is the lowest on the left side?
  LeftMinIndex = FindIndexBelow(Table.Columns(1), LeftPoint)
  ' The leftmost item should be invalid, so the return value should be higher than 1.
  If LeftMinIndex <= 2 Then
    ' Slide the range over to the right if it is lower than the source data domain.
    LeftMinIndex = 3
  End If
  ' Slide the range over to the left if it is higher than the source data domain.
  If LeftMinIndex >= Table.Rows.Count - 1 Then
    LeftMinIndex = Table.Rows.Count - 2
  End If
  ' Determine weights that will be used for all four rows...
  ' Loop once for each weight.
  For i = LBound(Weights) To UBound(Weights)
    ' Initialize the numerator and denominator.
    Numerator = 1
    Denominator = 1
    ' Loop once for each potential Lagrange coefficient.
    For j = LBound(Weights) To UBound(Weights)
      If i <> j Then
        ' Calculate the numerator for this term.
        Numerator = Numerator * (TopPoint - Table.Cells(1, TopMinIndex - LBound(Weights) - 1 + j))
        ' Calculate the denominator for this term.
        Denominator = Denominator * (Table.Cells(1, TopMinIndex - LBound(Weights) - 1 + i) - Table.Cells(1, TopMinIndex - LBound(Weights) - 1 + j))
      End If
    Next
    ' Populate the Weights array with this weight value.
    Weights(i) = Numerator / Denominator
  Next
  
  ' Generate the 4x1 data subset that will be interpolated over...
  ' Loop once for each interpolated value on the line.
  For i = LBound(Subset) To UBound(Subset)
    ' Initialize this item in the data subset.
    Subset(i) = 0
    ' Loop once for each Lagrange polynomial term.
    For j = LBound(Weights) To UBound(Weights)
      ' Include this Lagrange polynomial term in the data subset.
      Subset(i) = Subset(i) + Table(LeftMinIndex + i - 1 - LBound(Subset), TopMinIndex - LBound(Weights) - 1 + j) * Weights(j)
    Next
  Next
  ' Determine weights for the 4x1 data subset, which is the column interpolated within the dataset...
  ' Loop once for each weight.
  For i = LBound(Weights) To UBound(Weights)
    ' Initialize the numerator and denominator.
    Numerator = 1
    Denominator = 1
    ' Loop once for each potential Lagrange coefficient.
    For j = LBound(Weights) To UBound(Weights)
      If i <> j Then
        ' Calculate the numerator for this term.
        Numerator = Numerator * (LeftPoint - Table.Cells(LeftMinIndex - LBound(Weights) - 1 + j, 1))
        ' Calculate the denominator for this term.
        Denominator = Denominator * (Table.Cells(LeftMinIndex - LBound(Weights) - 1 + i, 1) - Table.Cells(LeftMinIndex - LBound(Weights) - 1 + j, 1))
      End If
    Next
    ' Populate the Weights array with this weight value.
    Weights(i) = Numerator / Denominator
  Next
  ' Assume the result is zero.
  BicubicInterpolation = 0
  ' Finish the interpolation to find the interpolated value...
  ' Loop once for each interpolated value on the subset line.
  For i = LBound(Subset) To UBound(Subset)
    ' The interpolated value is the sum of the product of each Lagrange coefficient and its corresponding function value.
    BicubicInterpolation = BicubicInterpolation + Subset(i) * Weights(i)
  Next
End Function
' Obtained from => https://mathformeremortals.wordpress.com/
' Find the index of a value that is less than or equal to Value.
' If the dataset appears to be in reverse, find the index ABOVE the value.
' If the value A is a Range type, the first cell is ignored.
Function FindIndexBelow(A As Variant, Value As Double) As Long
  Dim i As Long
  ' Assume there is no such value in the array.
  FindIndexBelow = -1
  If VarType(A) = vbArray Then
    ' Are the items in reverse order?
    If A(LBound(A)) > A(LBound(A) + 1) Then
      For i = LBound(A) To UBound(A)
        ' Is this array element less than or equal to Value?
        If A(i) >= Value Then
          ' This is a valid value.
          FindIndexBelow = i
        Else
          ' Stop looking.
          Exit For
        End If
      Next
    Else
      For i = LBound(A) To UBound(A)
        ' Is this array element less than or equal to Value?
        If A(i) <= Value Then
          ' This is a valid value.
          FindIndexBelow = i
        Else
          ' Stop looking.
          Exit For
        End If
      Next
    End If
  ' Is the array a Range type?
  ElseIf VarType(A) = 8204 Then
    ' Are the items in reverse order?
    If A.Cells(2) > A.Cells(3) Then
      For i = 2 To A.Cells.Count
        ' Is this array element less than or equal to Value?
        If A.Cells(i) >= Value Then
          ' This is a valid value.
          FindIndexBelow = i
        Else
          ' Stop looking.
          Exit For
        End If
      Next
    Else
      For i = 2 To A.Cells.Count
        ' Is this array element less than or equal to Value?
        If A.Cells(i) <= Value Then
          ' This is a valid value.
          FindIndexBelow = i
        Else
          ' Stop looking.
          Exit For
        End If
      Next
    End If
  ' We don't know what the argument is!
  Else
    MsgBox "Function FindIndexBelow has an invalid argument ""Value""=" & Value
  End If
End Function

'This is the work of Tomas Co, Michigan Technological Univiersity
Public Function GetLargestRoot(a3 As Double, a2 As Double, a1 As Double, A0 As Double) As Double
    'This work is adapted from work created by Tomas Co, Michigan Technological Univiersity
    ' Computes the maximum real root of the cubic equation
    ' a3 x^3 + a2 x^2 + a1 x + a0 = 0
    
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double
    Dim Z As Double
    Dim Q As Double
    Dim p As Double
    Dim h As Double
    Dim Y As Double
    Dim z1 As Double
    Dim z2 As Double
    Dim z3 As Double
    Dim c1 As Double
    Dim S1 As Double
    Dim m As Double
    Dim R As Double
    Dim Disc As Double
    Dim Theta As Double
    
    Dim myErrorMsg As String

    On Error GoTo myErrorHandler

    If a3 = 0 Then
        myErrorMsg = "GetSmallestRoot error: The first term a3 equals zero!"
        GoTo myErrorHandler
    End If

    A = a2 / a3
    B = a1 / a3
    C = A0 / a3
    p = (-A ^ 2 / 3 + B) / 3
    
    If (9 * A * B - 2 * A ^ 3 - 27 * C) = 0 Then
        myErrorMsg = "GetSmallestRoot error: (9 * A * B - 2 * A ^ 3 - 27 * C) equals zero!"
        GoTo myErrorHandler
    End If
    
    Q = (9 * A * B - 2 * A ^ 3 - 27 * C) / 54
    Disc = Q ^ 2 + p ^ 3
    
    If Disc > 0 Then
        h = Q + Disc ^ (1 / 2)
        
        If (Abs(h)) ^ (1 / 3) = 0 Then
            myErrorMsg = "GetSmallestRoot error: (Abs(h)) ^ (1 / 3) equals zero!"
            GoTo myErrorHandler
        End If
        
        Y = (Abs(h)) ^ (1 / 3)
        If h < 0 Then Y = -Y
        Z = Y - p / Y - A / 3
    Else
        Theta = Atn((-Disc) ^ (1 / 2) / Q)
        c1 = Cos(Theta / 3)
        
        If Q < 0 Then
            S1 = Sin(Theta / 3)
            c1 = (c1 - S1 * 3 ^ (1 / 2)) / 2
        End If
        
        z1 = 2 * (-p) ^ (1 / 2) * c1 - A / 3
        m = A + z1
        R = (m ^ 2 - 4 * (B + m * z1)) ^ (1 / 2)
        z2 = (-m + R) / 2
        z3 = (-m - R) / 2
        Z = z1
        
        If z2 > Z And z2 > 0 Then
            Z = z2
        End If
        If z3 > Z And z3 > 0 Then
            Z = z3
        End If
    End If

GetLargestRoot = Z

Exit Function

myErrorHandler:

    If Err.Number = 0 Then
        Debug.Print myErrorMsg
    Else
        Debug.Print Err.Description
    End If
    
    GetLargestRoot = -500
    
End Function

Public Function GetSmallestRoot(a3 As Double, a2 As Double, a1 As Double, A0 As Double) As Double

    'This work is adapted from work created by Tomas Co, Michigan Technological Univiersity
    ' Computes the minimum real root of the cubic equation
    ' a3 x^3 + a2 x^2 + a1 x + a0 = 0
    
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double
    Dim Z As Double
    Dim Q As Double
    Dim p As Double
    Dim h As Double
    Dim Y As Double
    Dim z1 As Double
    Dim z2 As Double
    Dim z3 As Double
    Dim c1 As Double
    Dim S1 As Double
    Dim m As Double
    Dim R As Double
    Dim Disc As Double
    Dim Theta As Double
    
    Dim myErrorMsg As String
    
    On Error GoTo myErrorHandler

    If a3 = 0 Then
        myErrorMsg = "GetSmallestRoot error: The first term a3 equals zero!"
        GoTo myErrorHandler
    End If
    
    A = a2 / a3
    B = a1 / a3
    C = A0 / a3
    p = (-A ^ 2 / 3 + B) / 3
    
    If (9 * A * B - 2 * A ^ 3 - 27 * C) = 0 Then
        myErrorMsg = "GetSmallestRoot error: (9 * A * B - 2 * A ^ 3 - 27 * C) equals zero!"
        GoTo myErrorHandler
    End If
    
    Q = (9 * A * B - 2 * A ^ 3 - 27 * C) / 54
    Disc = Q ^ 2 + p ^ 3
    
    If Disc > 0 Then
        h = Q + Disc ^ (1 / 2)
        
        If (Abs(h)) ^ (1 / 3) = 0 Then
            myErrorMsg = "GetSmallestRoot error: (Abs(h)) ^ (1 / 3) equals zero!"
            GoTo myErrorHandler
        End If
        
        Y = (Abs(h)) ^ (1 / 3)
        If h < 0 Then Y = -Y
        Z = Y - p / Y - A / 3
    Else
        Theta = Atn((-Disc) ^ (1 / 2) / Q)
        c1 = Cos(Theta / 3)
        
        If Q < 0 Then
            S1 = Sin(Theta / 3)
            c1 = (c1 - S1 * 3 ^ (1 / 2)) / 2
        End If
        z1 = 2 * (-p) ^ (1 / 2) * c1 - A / 3
        m = A + z1
        R = (m ^ 2 - 4 * (B + m * z1)) ^ (1 / 2)
        z2 = (-m + R) / 2
        z3 = (-m - R) / 2
        Z = z1
        
        If z2 < Z And z2 > 0 Then
            Z = z2
        End If
        If z3 < Z And z3 > 0 Then
            Z = z3
        End If
    End If
    
    GetSmallestRoot = Z
    
    Exit Function

myErrorHandler:

    If Err.Number = 0 Then
        Debug.Print myErrorMsg
    Else
        Debug.Print Err.Description
    End If
    
    GetSmallestRoot = -500
    
    End Function
Public Function QUBIC(p As Double, Q As Double, R As Double) As Variant()

', ROOT() As Double)
' Q U B I C - Solves a cubic equation of the form:
' y^3 + Py^2 + Qy + R = 0 for real roots.
' Inputs:
' P,Q,R Coefficients of polynomial.

' Outputs:
' ROOT 3-vector containing only real roots.
' NROOTS The number of roots found. The real roots
' found will be in the first elements of ROOT.

' Method: Closed form employing trigonometric and Cardan
' methods as appropriate.

' Note: To translate and equation of the form:
' O'y^3 + P'y^2 + Q'y + R' = 0 into the form above,
' simply divide thru by O', i.e. P = P'/O', Q = Q'/O',
' etc.

Dim Z(3) As Double
Dim P2 As Double
Dim RMS As Double
Dim A As Double
Dim B As Double
Dim nRoots As Integer
Dim i As Integer
Dim DISCR As Double
Dim T1 As Double
Dim T2 As Double
Dim RATIO As Double
Dim sum As Double
Dim DIF As Double
Dim AD3 As Double
Dim E0 As Double
Dim CPhi As Double
Dim PhiD3 As Double
Dim PD3 As Double
Dim ROOT() As Variant

Const DEG120 As Double = 2.09439510239319
Const Tolerance As Double = 0.00000001
Const Tol2 As Double = 1E-25

' ... Translate equation into the form Z^3 + aZ + b = 0

P2 = p ^ 2
A = Q - P2 / 3
B = p * (2 * P2 - 9 * Q) / 27 + R

RMS = Sqr(A ^ 2 + B ^ 2)
If RMS < Tol2 Then
' ... Three equal roots <= this occurs at the ciritcal point
nRoots = 3
ReDim ROOT(0 To nRoots)
For i = 1 To 3
ROOT(i) = -p / 3
Next i
Exit Function
End If

DISCR = (A / 3) ^ 3 + (B / 2) ^ 2

If DISCR > 0 Then

T1 = -B / 2
T2 = Sqr(DISCR)
If T1 = 0 Then
RATIO = 1
Else
RATIO = T2 / T1
End If

If Abs(RATIO) < Tolerance Then
' ... Three real roots, two (2 and 3) equal. <= 2 phases
nRoots = 3
Z(1) = 2 * QBRT(T1)
Z(2) = QBRT(-T1)
Z(3) = Z(2)
Else
' ... One real root, two complex. Solve using Cardan formula. <= single phase
nRoots = 1
sum = T1 + T2
DIF = T1 - T2
Z(1) = QBRT(sum) + QBRT(DIF)
End If

Else

' ... Three real unequal roots. Solve using trigonometric method. < two phases
nRoots = 3
AD3 = A / 3#
E0 = 2# * Sqr(-AD3)
CPhi = -B / (2# * Sqr(-AD3 ^ 3))
PhiD3 = WorksheetFunction.Acos(CPhi) / 3#
Z(1) = E0 * Cos(PhiD3)
Z(2) = E0 * Cos(PhiD3 + DEG120)
Z(3) = E0 * Cos(PhiD3 - DEG120)

End If

' ... Now translate back to roots of original equation
PD3 = p / 3

ReDim ROOT(1 To nRoots)

For i = 1 To nRoots
ROOT(i) = Z(i) - PD3
Debug.Print ROOT(i)
Next i
QUBIC = ROOT()


End Function

Function QBRT(X As Double) As Double

' Signed cube root function. Used by Qubic procedure.

QBRT = Abs(X) ^ (1 / 3) * Sgn(X)

End Function ' Math for Mere Mortals
' BicubicLagrangeInterpolation_v1
