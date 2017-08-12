Attribute VB_Name = "mcroot"
Option Explicit


Public Function GetLargestRoot(a3 As Double, a2 As Double, a1 As Double, a0 As Double) As Double
'
' Computes the maximum real root of the cubic equation
' a3 x^3 + a2 x^2 + a1 x + a0 = 0
'
Dim A, B, c, D, z, Q, p, h, Y, z1, z2, z3, c1, S1, m, R, Disc, Theta As Double
A = a2 / a3
B = a1 / a3
c = a0 / a3
p = (-A ^ 2 / 3 + B) / 3
Q = (9 * A * B - 2 * A ^ 3 - 27 * c) / 54
Disc = Q ^ 2 + p ^ 3
If Disc > 0 Then
h = Q + Disc ^ (1 / 2)
Y = (Abs(h)) ^ (1 / 3)
If h < 0 Then Y = -Y
z = Y - p / Y - A / 3
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
z = z1
If z2 > z Then z = z2
If z3 > z Then z = z3
End If
GetLargestRoot = z
End Function
Public Function GetSmallestRoot(a3 As Double, a2 As Double, a1 As Double, a0 As Double) As Double
'
' Computes the minimum real root of the cubic equation
' a3 x^3 + a2 x^2 + a1 x + a0 = 0
'
Dim A, B, c, D, z, Q, p, h, Y, z1, z2, z3, c1, S1, m, R, Disc, Theta As Double
A = a2 / a3
B = a1 / a3
c = a0 / a3
p = (-A ^ 2 / 3 + B) / 3
Q = (9 * A * B - 2 * A ^ 3 - 27 * c) / 54
Disc = Q ^ 2 + p ^ 3
If Disc > 0 Then
h = Q + Disc ^ (1 / 2)
Y = (Abs(h)) ^ (1 / 3)
If h < 0 Then Y = -Y
z = Y - p / Y - A / 3
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
z = z1
If z2 < z Then z = z2
If z3 < z Then z = z3
End If
GetSmallestRoot = z
End Function
Public Function QUBIC(p As Double, Q As Double, R As Double) As Double
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

Dim z(3) As Double
Dim P2 As Double
Dim RMS As Double
Dim A As Double
Dim B As Double
Dim nRoots As Integer
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
Dim i As Integer

Const DEG120 = 2.09439510239319
Const Tolerance = 0.00000001
Const Tol2 = 1E-25

' ... Translate equation into the form Z^3 + aZ + b = 0

P2 = p ^ 2
A = Q - P2 / 3
B = p * (2 * P2 - 9 * Q) / 27 + R

RMS = Sqr(A ^ 2 + B ^ 2)
If RMS < Tol2 Then
' ... Three equal roots
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
' ... Three real roots, two (2 and 3) equal.
nRoots = 3
z(1) = 2 * QBRT(T1)
z(2) = QBRT(-T1)
z(3) = z(2)
Else
' ... One real root, two complex. Solve using Cardan formula.
nRoots = 1
sum = T1 + T2
DIF = T1 - T2
z(1) = QBRT(sum) + QBRT(DIF)
End If

Else

' ... Three real unequal roots. Solve using trigonometric method.
nRoots = 3
AD3 = A / 3#
E0 = 2# * Sqr(-AD3)
CPhi = -B / (2# * Sqr(-AD3 ^ 3))
PhiD3 = WorksheetFunction.Acos(CPhi) / 3#
z(1) = E0 * Cos(PhiD3)
z(2) = E0 * Cos(PhiD3 + DEG120)
z(3) = E0 * Cos(PhiD3 - DEG120)

End If

' ... Now translate back to roots of original equation
PD3 = p / 3

ReDim ROOT(0 To nRoots)

For i = 1 To nRoots
ROOT(i) = z(i) - PD3
Debug.Print ROOT(i)
Next i

End Function

Function QBRT(X As Double) As Double

' Signed cube root function. Used by Qubic procedure.

QBRT = Abs(X) ^ (1 / 3) * Sgn(X)

End Function
