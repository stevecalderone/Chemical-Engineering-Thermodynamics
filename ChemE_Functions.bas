Attribute VB_Name = "ChemE_Functions"
    Option Explicit
    Option Base 0
    
'*********************************************************************************************
'License

'Copyright (c) 2016, Steve Calderone, steve@stevecalderone.com
'All rights reserved.
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are
'met:
'
'Redistributions of source code must retain the above copyright
'notice, this list of conditions and the following disclaimer.
'* Redistributions in binary form must reproduce the above copyright
'notice, this list of conditions and the following disclaimer in
'the documentation and/or other materials provided with the distribution
'* Neither the name of the www.stevecalderone.com nor the names
'of its contributors may be used to endorse or promote products derived
'from this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
'ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
'LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
'CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
'SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
'INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
'CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
'ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
'POSSIBILITY OF SUCH DAMAGE.
'*********************************************************************************************

'References
'
'1
'Peng Robinson Equation of State
'A New Two-Constant Equation of State
'Ding-Yu Peng, Donald B. Robinson
'Ind. Eng. Chem. Fundamen., 1976, 15 (1), pp 59–64
'DOI:  10 0.1021 / i160057a011
'Publication Date: February 1976
'
'2
'Spreadsheet for Thermodynamics Instruction
'Phillip Savage - University of Michigan
''ChE classroom'
'Fall 1995
'http://ufdcimages.uflib.ufl.edu/AA/00/00/03/83/00128/AA00000383_00128_00262.pdf
'
'3
'Flash Routine Reference:
'CHEMENG 120 taught by Professor Musgrave during the Spring '04 term at Stanford.
'http://documentslide.com/documents/lecture-5-isothermal-flash-calculations.html
'
'4
'Implementation of Departure Functions on worksheets Departure_Pt1, 2, 3 & 4
'Dr. Phillip Savage - Penn State University
'http://www.che.psu.edu/department/directory-detail.aspx?LandOn=Gen&q=pes15
'
'5
'Bubble and Dew Point Routines adapted from:
'It’s not as easy as it looks - Revisiting Peng–Robinson equation of state convergence issues for dew point, bubble point and flash calculations
'
'Vamshi Krishna Kandula,(a) John C. Telotte (b) and F. Carl Knopf (corresponding author) (a)
'(a) Chemical Engineering Department, Louisiana State University, USA
'e -mail: cknopf@ southalabama.edu
'(b) Chemical Engineering Department, Florida A&M University – Florida State University, USA
'
'International Journal of Mechanical Engineering Education, Volume 41, Number 3 (July 2013), © Manchester University Press
' http://journals.sagepub.com/doi/pdf/10.7227/IJMEE.41.3.2
'
'6
'Cubic Equation VBA Code
'Dr.Tomas b.Co
'Michigan Tech University
'https://www.mtu.edu/chemical/department/faculty/co/
'
'7
'Bicubic Interpolation Code (for LeeKeslerZ() function)
'https://mathformeremortals.wordpress.com/
'
'8
'modArraySupport module and array handling techniques
'Chip Pearson, chip@cpearson.com, www.cpearson.com
'
'9
'Derivation of the enthalpy departure function
'https://shareok.org/bitstream/handle/11244/12606/Thesis-1996-R231e.pdf?sequence=1
'by Abhishek Rastogi
'APPENDIX a
'A DETAILED DERIVATION OF PENG-ROBINSON EQUATION OF STATE ENTHALPY DEPARTURE FUNCTION
'
'10
'PData worksheet physical properties adapted from:
'Properties Databank 1.0
'Pedro Fajardo
'ppfk@ yahoo.com
'http://www.cheresources.com/invision/files/file/125-physical-properties-ms-excel-add-in/
'
'11
'Thermodynamic Properties Involving Derivatives
'Using the Peng-Robinson Equation of State
'R.M. Pratt Ph. D.
'The National University of Malaysia
'(now Associate Professor of Mathematics and Sciences, Fresno Pacific University)
'http://ufdcimages.uflib.ufl.edu/AA/00/00/03/83/00150/AA00000383_00150_00112.pdf
'
'12
'Chemical Equilibrium by Gibbs Energy Minimization on Spreadsheets
'Int. J. Engng Ed. Vol. 16, No. 4, pp. 335±339, 2000
'Y.LWIN
'Department of Chemical Engineering, Rangoon Institute of Technology, Insein P. O., Rangoon, Burma.
'e -mail: ylwin@ yahoo.com
'http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.476.5931&rep=rep1&type=pdf
'
'13
'Getting a Handle on Advanced Cubic Equations of State
'www.cepmagazine.org, November 2002, CEP
'Chorng H. Twu, Wayne D. Sim and Vince Tassone, Aspen Technology, Inc.
'http://people.clarkson.edu/~wwilcox/Design/adv-ceos.pdf

'14
'VLE predictions with the Peng.Robinson equation of state and
'temperature dependent kij calculated through a group contribution method
'Jean-No¡§el Jaubert., Fabrice Mutelet
'Laboratoire de Thermodynamique des Milieux Polyphas¢¥es, Institut National Polytechnique de Lorraine, Ecole Nationale Sup¢¥erieure des Industries
'Chimiques, 1rue Grandville, 54000 Nancy, France
'Received 24 January 2004; accepted 25 June 2004

'15
'Prediction of Thermodynamic Properties of Alkyne-Containing
'Mixtures with the E.PPR78 Model
'Xiaochun Xu,õ Jean-Noe.l Jaubert,*,õ Romain Privat,õ and Philippe Arpentinierö
'õEcole Nationale Supe.rieure des Industries Chimiques, Laboratoire Re.actions et Ge.nie des Proce.de.s (UMR CNRS 7274),
'Universite. de Lorraine, 1 rue Grandville, 54000 Nancy, France
'öCentre de Recherche Paris Saclay, Air Liquide, 1 chemin de la porte des loges, BP 126, 78354 Jouy-en-Josas, France

'16
'Efficient flash calculations for chemical process
'design — extension of the Boston–Britt
'‘‘Inside–out’’ flash algorithm to extreme
'conditions and new flash types
'Vipul S. Parekh and Paul M. Mathias
'Computers Chem. Engng Vol. 22, No. 10, pp. 1371—1380, 1998
'( 1998 Published by Elsevier Science Ltd.
'All rights reserved. Printed in Great Britain



'*********************************************************************************************



    Const GasLawR As Double = 0.000083144621 'Units = m3-bar/gmol/K,  GasLawR * 100,000 = 8.3144621 kJ/kg-mole/K = 8.3144621 J/g-mole/K

    Enum iColumns                   '<= This enumeration is used to manage the range and array indexes used in many of these user defined functions
        MW = 0                      '<= References are found of this enumeration in the code
        tc = 1                      'for example: 'iColumns.MW' refers to the molecular weight column (zero base index) of the dataset
        pc = 2
        omega = 3
        zc = 4
        Ki = 5
        bi = 6
        tb = 7
        hvap = 8
        iHf298 = 9
        iS298 = 10
        iGf298 = 11
        CpDataType = 12             '<= These indexes are for the Cp range index selection
        NIST_Mn1 = 13               '<= These indexes are for the Cp range index selection
        NIST_Mx1 = 14               '<= These are for information processed in the validateDataset function and required by other functions
        NIST_MN2 = NIST_Mn1 + 10
        NIST_MX2 = NIST_Mx1 + 10
        NIST_MN3 = NIST_MN2 + 10
        NIST_MX3 = NIST_MX2 + 10
        NIST_MN4 = NIST_MN3 + 10
        NIST_MX4 = NIST_MX3 + 10
        NIST_MN5 = NIST_MN4 + 10
        NIST_MX5 = NIST_MX4 + 10
        NIST_MN6 = NIST_MN5 + 10
        NIST_Mx6 = NIST_MX5 + 10
        lastCpIndex = 72
        alphaType = 73              '<= This is a flag for processing create_alphaiArray normally or using Twu volume tranlation method
        errMsgsOn = 74              '<= This is a flag indicating if 'error messages on' was found in the top left most cell of the dataset.
        LiquidsFound = 75           '<=
        LiquidIndex = 76            '<= The base zero index of the the liquid first liquid species
        iSpecies = 77               '<= Storage for the upper bound for the rows of the dataset. This is equal to the number of species minus one.
        globalErrmsg = 78           '<= This is to store private function error message to pass along to the public function.
        predictive = 79
        FinalIndex = 80
        MoleFraction = 0            '<= Used in the validateMoles function
        moles = 1                   '<= Used in the validateMoles function
        TempK = 0                   '<= These indexes are for the seleceCpDataRange, Enthalpy and Entropy functions. this stores NIST-MNT index for TempK
        Vap298 = 1                  '<= These indexes are for the seleceCpDataRange, Enthalpy and Entropy functions. This stores NIST-MNT index for Vapor at 298K
        NBPVap = 2                  '<= These indexes are for the seleceCpDataRange, Enthalpy and Entropy functions. This stores NIST-MNT for index vapor at normal boiling point
        NBPLiq = 3                  '<= These indexes are for the seleceCpDataRange, Enthalpy and Entropy functions. This stores NIST-MNT for index liquid at normal boiling point
        dadT_constV = 0
        dPdv_constT = 1
        dPdT_constV = 2
        dadT_constP = 3
        dBdT_constP = 4
        dZdT_constP = 5
        dVdT_constP = 6
        sumb = 7
        suma = 8
        a = 9
        b = 10
        Z = 11
        vol = 12
    End Enum

Private Function calculate_T_BubDew_Est(dataset As Variant, moleComp() As Double, pbara As Double) As Double()
    
    '***************************************************************************
    'This function is called by BubbleT and DewT functions to calculate the initial guess of the dew or bubble temperature
    'More on this function can be found in Reference 5 and Reference 15
    '***************************************************************************

    
    On Error GoTo myErrorHandler:
    
    Dim T_Lo As Double
    Dim T_Hi As Double
    Dim T_New As Double
    Dim T_NBP As Double
    Dim y_Sum_Lo As Double
    Dim y_Sum_Hi As Double
    Dim y_Sum As Double
    Dim x_Sum_Lo As Double
    Dim x_Sum_Hi As Double
    Dim x_Sum As Double
    Dim Ki() As Double
    Dim outputArray() As Double
    Dim i As Integer
    Dim T_Dew_Est As Double
    Dim T_Bub_Est As Double
    Dim Counter As Integer
    Dim T_Bub_RoughEst_Temp As Double
    Dim T_Dew_RoughEst_Temp As Double
    Dim BubTempFound As Boolean
    Dim DewTempFound As Boolean
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "calculate_T_BubDew_Est"

    ReDim Ki(dataset(0, iColumns.iSpecies))
    
    T_Bub_Est = 0
    T_Dew_Est = 0
    T_Lo = 0
    T_Hi = 0
    y_Sum = 0
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        T_Bub_Est = T_Bub_Est + moleComp(i) * dataset(i, iColumns.tc)
    Next i
    
    T_Bub_Est = T_Bub_Est * 0.7
    T_Bub_RoughEst_Temp = T_Bub_Est
    
    Counter = 0
        
    Do While BubTempFound = False
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            If T_Bub_Est = 0 And dataset(i, iColumns.tb) >= 0 And dataset(i, iColumns.tc) = 0 Or pbara = 0 Or (1 / dataset(i, iColumns.tc) - 1 / dataset(i, iColumns.tb)) = 0 Then
                myErrorMsg = "Species " & i & " error: The supplied pressure or TC equals zero or there is a problem with the boiling point."
                GoTo myErrorHandler
            End If
            
            On Error Resume Next
            'Not sure how to derive the Ki(i) estimating equation below. Need to contact the author as it's derivation is not presented in the referenced paper. Maybe it is only valid at the Dew and Bubble points. Check if it can be adopted for FlashTP.
            Ki(i) = dataset(i, iColumns.pc) ^ (((1 / T_Bub_Est - 1 / dataset(i, iColumns.tb)) / _
            ((1 / dataset(i, iColumns.tc)) - (1 / dataset(i, iColumns.tb)))))                               'Ki etimate & Dew/Bubble point adapted from => It’s not as easy as it looks: revisiting Peng–Robinson equation of state convergence issues for dew point, bubble point and flash calculations
            Ki(i) = Ki(i) / pbara                                                                           'Vamshi Krishna Kandula (a), John C. Telotteb (b) and F. Carl Knopf - E-mail: knopf@lsu.edu (a) (corresponding author)
                                                                                                            'a - Chemical Engineering Department, Louisiana State University, USA
            If Err.Number = 6 Then           'This is an overflow error often caused by                     'b - Chemical Engineering Department, Florida A&M University – Florida State University, USA
                ReDim outputArray(1)         'dividing a number by a very small number
                outputArray(0) = T_Bub_RoughEst_Temp
                outputArray(1) = 1.1 * T_Bub_RoughEst_Temp
                calculate_T_BubDew_Est = outputArray()
                Err.Clear
                myErrorMsg = "Overflow error. Tried to divided a number by a very small number in Ki estimate calc found in LSU/FAM Dew & Bubble T paper."
                GoTo myErrorHandler
            End If
        Next i
            
        y_Sum = 0
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            y_Sum = y_Sum + moleComp(i) * Ki(i)
        Next i
                
        If y_Sum < 1 Then
            T_Lo = T_Bub_Est
            y_Sum_Lo = y_Sum - 1
            T_New = T_Bub_Est * 1.1
        End If
        
        If y_Sum > 1 Then
            T_Hi = T_Bub_Est
            y_Sum_Hi = y_Sum - 1
            T_New = T_Bub_Est / 1.1
        End If
        
        If y_Sum = 1 Then
            BubTempFound = True
        End If
        
        If T_Lo * T_Hi > 0 Then
            T_New = (y_Sum_Hi * T_Lo - y_Sum_Lo * T_Hi) / (y_Sum_Hi - y_Sum_Lo)
        End If
                
        If Abs(T_Bub_Est - T_New) < 0.001 Then
            BubTempFound = True
        End If
        
        If Abs(y_Sum - 1) < 0.00001 Then
            BubTempFound = True
        End If
        
        Counter = Counter + 1
        
        T_Bub_Est = T_New
        
        If Counter = 1000 Then
            myErrorMsg = "Counter is 100 iterations."
            GoTo myErrorHandler
        End If
    Loop
    
        Counter = 0
        T_Dew_Est = 1.1 * T_Bub_Est
        T_Dew_RoughEst_Temp = T_Dew_Est
        T_Lo = 0
        T_Hi = 0
                
    Do While DewTempFound = False
        For i = 0 To dataset(0, iColumns.iSpecies)
            Ki(i) = dataset(i, iColumns.pc) ^ ((((1 / T_Dew_Est) - (1 / dataset(i, iColumns.tb))) / _
            (1 / dataset(i, iColumns.tc) - 1 / dataset(i, iColumns.tb))))
            Ki(i) = Ki(i) / pbara
        Next i
                    
        x_Sum = 0
                    
        For i = 0 To dataset(0, iColumns.iSpecies)
            x_Sum = x_Sum + moleComp(i) / Ki(i)
        Next i
                        
        If x_Sum < 1 Then
            T_Lo = T_Dew_Est
            x_Sum_Lo = x_Sum - 1
            T_New = T_Dew_Est / 1.1
        End If
                        
        If x_Sum > 1 Then
            T_Hi = T_Dew_Est
            x_Sum_Hi = x_Sum - 1
            T_New = T_Dew_Est * 1.1
        End If
                        
        If x_Sum = 1 Then
            DewTempFound = True
        End If
        
        If T_Lo * T_Hi > 0 Then
            T_New = (T_Lo * x_Sum_Hi - T_Hi * x_Sum_Lo) / (x_Sum_Hi - x_Sum_Lo)
        End If
        
        If Abs(T_Dew_Est - T_New) < 0.001 Then
            DewTempFound = True
        End If
        
        If Abs(x_Sum - 1) < 0.00001 Then
            DewTempFound = True
        End If
                        
        Counter = Counter + 1
        
        T_Dew_Est = T_New
        
        If Counter = 1000 Then
            myErrorMsg = "More than 100 iterations!"
            GoTo myErrorHandler
        End If
                        
    Loop

        ReDim outputArray(1)
        outputArray(0) = T_Bub_Est
        outputArray(1) = T_Dew_Est
        
        calculate_T_BubDew_Est = outputArray()
        
        Exit Function
                        
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

        
            ReDim outputArray(1)
            outputArray(0) = T_Bub_RoughEst_Temp                '<= Provide very rough estimate of bubble point temp if this function fails
            outputArray(1) = T_Bub_RoughEst_Temp * 1.1          '<= Provide very rough estimate of dew point temp if this function fails
            
            calculate_T_BubDew_Est = outputArray()
            
        End Function

    Public Function AntoineVP(Species As Variant, temperature As Variant, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'This function does not utilize a dataset.
    'This function calculates the pure component vapor pressure given temperature
    'based upon the Antoin coefficients found in the PData worksheet.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim TempK As Double
    Dim minT As Double
    Dim maxT As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim UDF_Range As Range
    Dim i As Integer
    Dim Name As String
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim TempC As Double
    Dim PData As Worksheet
    Dim SpeciesName As String
   
    fcnName = "AntoineVP"
                           
    myErrorMsg = ""
    
    If TypeName(Species) = "Range" Then
        If Species.Rows.Count <> 1 Or Species.Columns.Count <> 1 Then
            myErrorMsg = "The supplied species name should be a string or a single cell reference to a string equal to 'Vapor' or 'Liquid'."
            GoTo myErrorHandler
        End If
    End If
    
    SpeciesName = CStr(Species)
    
    i = getPdataWorksheetIndex
    
    Set PData = ThisWorkbook.Worksheets(i)
    
    Set UDF_Range = Application.Caller
    

    
    TempC = checkInputTemperature(temperature)
        
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is a problem with the supplied temperature."
        GoTo myErrorHandler
    End If

    On Error Resume Next
    
    i = Application.WorksheetFunction.Match("(Species)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "The '(Species)' column could not be found."
        GoTo myErrorHandler
    Else
        If SpeciesName <> Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0) Then
            myErrorMsg = "The species could not be found."
            GoTo myErrorHandler
        End If
    End If
    
    i = Application.WorksheetFunction.Match("(ANT-TMN)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Could not find (ANT-TMN) column in PData worksheet."
        GoTo myErrorHandler
    Else
        minT = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0)
        If IsNumeric(minT) = False Then
            myErrorMsg = "(ANT-tmn) in PData worksheet is not numeric." & myErrorMsg
            GoTo myErrorHandler
        End If
    End If
    
    i = Application.WorksheetFunction.Match("(ANT-TMX)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Could not find (ANT-TMX) column in PData worksheet." & myErrorMsg
        GoTo myErrorHandler
    Else
        maxT = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0)
        If IsNumeric(maxT) = False Then
            myErrorMsg = "(ANT-tmx) in PData worksheet is not numeric." & myErrorMsg
            GoTo myErrorHandler
        End If
    End If
    
    i = Application.WorksheetFunction.Match("(ANT-A)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Could not find (ANT-A) column in PData worksheet." & myErrorMsg
        GoTo myErrorHandler
    Else
        a = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0)
        If IsNumeric(a) = False Then
            myErrorMsg = "(ANT-A) in PData worksheet is not numeric." & myErrorMsg
            GoTo myErrorHandler
        End If
    End If
    
    i = Application.WorksheetFunction.Match("(ANT-B)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
            myErrorMsg = "Could not find (ANT-B) column in PData worksheet." & myErrorMsg
            GoTo myErrorHandler
    Else
        b = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0)
        If IsNumeric(b) = False Then
            myErrorMsg = "(ANT-B) in PData worksheet is not numeric." & myErrorMsg
            GoTo myErrorHandler
        End If
    End If
    
    i = Application.WorksheetFunction.Match("(ANT-C)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Could not find (ANT-C) column in PData worksheet." & myErrorMsg
        GoTo myErrorHandler
    Else
        c = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), i, 0)
        If IsNumeric(c) = False Then
            myErrorMsg = "(ANT-C) in PData worksheet is not numeric." & myErrorMsg
            GoTo myErrorHandler
        End If
    End If
    
    On Error GoTo myErrorHandler
    
    If TempK > minT And TempK < maxT And (TempK) + c > 10 ^ -12 Then
        AntoineVP = 14.5038 * Exp(a - b / ((TempK) + c)) / 760# / 14.5038               '<convert from mmHG to atm to bara
    Else
        myErrorMsg = "The supplied temperature (" & TempC & " C) is outside of the valid temperature range. Tmin = " & minT - 273.15 & " C and Tmax = " & maxT - 273.15 & " C"
        GoTo myErrorHandler
    End If
    
    On Error GoTo 0
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    
    Application.ScreenUpdating = True
    Exit Function
    
myErrorHandler:

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    AntoineVP = 0
    Application.ScreenUpdating = True
    End Function

    Public Function LiquidWaterViscCP(temperature As Variant, Optional errMsgsOn As Boolean = False) As Double '<= dynamic viscosity of water
    Dim myErrorMsg As String
    
    '***************************************************************************
    'The constants below are created from data found here:
    'http://www.viscopedia.com/viscosity-tables/substances/water/
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    Application.ScreenUpdating = False
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim TempC As Double
    Dim TempK As Double
    
    fcnName = "LiquidWaterViscCP"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    TempC = checkInputTemperature(temperature)
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is something wrong with the supplied temperature."
        GoTo myErrorHandler
    End If
    
    If TempC < 1 Or TempC > 80 Then
        myErrorMsg = "The supplied temperature is outside the valid range. Tmin = 1 C and Tmax = 80 C."
        GoTo myErrorHandler
    End If
         
    LiquidWaterViscCP = -7.00115362651037 * 10 ^ -10 * TempC ^ 5 + 1.90681026542577 * 10 ^ -7 * TempC ^ 4 - 2.16713835341039E-05 * TempC ^ 3 + 1.39734725302587E-03 * TempC ^ 2 _
    - 5.99578184253486E-02 * TempC + 1.78622754441384
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    Exit Function
    
myErrorHandler:

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    LiquidWaterViscCP = 0
    
    Application.ScreenUpdating = True
    
    End Function
    Public Function LiquidWaterViscM2BySec(temperature As Variant, Optional errMsgsOn As Boolean = False) As Double  '<= kinematic viscosity of water
    
    '***************************************************************************
    'The constants below are created from data found here:
    'http://www.viscopedia.com/viscosity-tables/substances/water/
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim myErrorMsg As String
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim TempC As Double
    Dim TempK As Double
    
    fcnName = "LiquidWaterViscM2BySec"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    TempC = checkInputTemperature(temperature)
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is something wrong with the supplied temperature."
        GoTo myErrorHandler
    End If
    
    If TempC < 1 Or TempC > 80 Then
        myErrorMsg = "The supplied temperature is outside the valid range. Tmin = 1 C and Tmax = 80 C."
        GoTo myErrorHandler
    End If
     
    LiquidWaterViscM2BySec = (-7.12697702590059 * 10 ^ -10 * TempC ^ 5 + 1.93812004157535 * 10 ^ -7 * TempC ^ 4 - 2.19695476816367E-05 * TempC ^ 3 + _
        1.41020562534985E-03 * TempC ^ 2 - 6.00397258025756E-02 * TempC + 1.78642374134197) * 10 ^ -6
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    LiquidWaterViscM2BySec = 0
    
    Application.ScreenUpdating = True
    
    End Function
    Public Function LiquidWaterDensity(temperature As Variant, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'The constants below are created from data found here:
    'http://www.viscopedia.com/viscosity-tables/substances/water/
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim myErrorMsg As String
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim TempC As Double
    Dim TempK As Double
    
    fcnName = "LiquidWaterDensity"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    TempC = checkInputTemperature(temperature)
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is something wrong with the supplied temperature."
        GoTo myErrorHandler
    End If
    
    If TempC < 1 Or TempC > 80 Then
        myErrorMsg = "The supplied temperature is outside the valid range. Tmin = 1 C and Tmax = 80 C."
        GoTo myErrorHandler
    End If
    
    LiquidWaterDensity = 2.48894114505504E-12 * TempC ^ 5 - 7.01566908289643E-10 * TempC ^ 4 + 8.73643506922288E-08 * TempC ^ 3 - _
    9.00827374796632E-06 * TempC ^ 2 + 6.8062375849856E-05 * TempC + 0.999845097660858
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    LiquidWaterDensity = 0
    
    Application.ScreenUpdating = True
    
    End Function
    
    Public Function LeeKeslerZ(Species As Variant, temperature As Variant, pressure As Double, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'This function utilizes the PData_LK1NonPolar_Z0 & PData_LK1NonPolar_Z1 named ranges found in the PData worksheet
    'and the BicubicInterpolation function in the Math module
    'This function calculates the pure component compressibility given T and P.
    'This function does not utilize a dataset.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    'valid for non-polar to slightly polar gases
    Dim TempK As Double
    Dim z0 As Double
    Dim z1 As Double
    Dim omega As Double
    Dim PR As Double
    Dim pc As Double
    Dim Tr As Double
    Dim tc As Double
    Dim i As Integer
    Dim sourceBook As Workbook
    Dim PDataExists As Boolean
    Dim PropertyNamesExist As Boolean
    Dim int_TC As Integer
    Dim int_PC As Integer
    Dim int_Species As Integer
    Dim int_Omega As Integer
    Dim sourceSheet As Worksheet
    Dim PDataName As Name
    Dim PropertiesExist As Boolean
    Dim PData_LK1NonPolar_Z0 As Boolean
    Dim PData_LK1NonPolar_Z1 As Boolean
    Dim myErrorMsg As String
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim TempC As Double
    Dim pbara As Double
    Dim PData As Worksheet
    Dim SpeciesName As String
    
    
    i = getPdataWorksheetIndex
    
    Set PData = ThisWorkbook.Worksheets(i)
    
    If TypeName(Species) = "Range" Then
        If Species.Rows.Count <> 1 Or Species.Columns.Count <> 1 Then
            myErrorMsg = "The supplied species name should be a string or a single cell reference to a string equal to 'Vapor' or 'Liquid'."
            GoTo myErrorHandler
        End If
    End If
    
    SpeciesName = CStr(Species)
    
    fcnName = "LeeKeslerZ"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    TempC = checkInputTemperature(temperature)
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is something wrong with the supplied temperature."
        GoTo myErrorHandler
    End If
    
    pbara = checkInputPressure(pressure)
    If pbara = -1 Then
        myErrorMsg = "There is something wrong with the supplied pressure."
        GoTo myErrorHandler
    End If
    
    If ValidateWorkbook = False Then
        myErrorMsg = "Workbook failed validation."
        GoTo myErrorHandler
    End If
    
    Set sourceSheet = ThisWorkbook.Worksheets(getPdataWorksheetIndex)
    
    i = 1
    For Each PDataName In ThisWorkbook.Names                        '<=ValidateWorkbook function does not check for existance of 'PData_LK1NonPolar_Z0' or 'PData_LK1NonPolar_Z1' named ranges.
                                                                    '<=so check for them here'
        If ThisWorkbook.Names(i).Name = "PData_LK1NonPolar_Z0" Then
            PData_LK1NonPolar_Z0 = True
        End If
        
        If ThisWorkbook.Names(i).Name = "PData_LK1NonPolar_Z1" Then
            PData_LK1NonPolar_Z1 = True
        End If
        
        i = i + 1
    Next PDataName

    If PData_LK1NonPolar_Z0 = True And PData_LK1NonPolar_Z1 = False Then
        myErrorMsg = "the 'PData_LK1NonPolar_Z0' and 'PData_LK1NonPolar_Z1' named ranges do not exist."
        GoTo myErrorHandler
    End If
    
    On Error Resume Next
    
    int_Species = Application.WorksheetFunction.Match("(Species)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "The '(species)' column could not be found."
        GoTo myErrorHandler
    Else
        If SpeciesName <> Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), int_Species, 0) Then
            myErrorMsg = "The species could not be found."
            GoTo myErrorHandler
        End If
    End If
    
    int_Species = Application.WorksheetFunction.Match("(Species)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "The '(species)' column could not be found."
        GoTo myErrorHandler
    End If
          
    int_TC = Application.WorksheetFunction.Match("(TC, K)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(TC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_PC = Application.WorksheetFunction.Match("(PC, bara)", sourceSheet.Range("PData_PropertyNames"), 0)
        If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(PC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_Omega = Application.WorksheetFunction.Match("(OMEGA)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(OMEGA)' not found in PData worksheet"
        GoTo myErrorHandler
    End If
           
    On Error GoTo myErrorHandler
    
    TempK = TempC + 273.15
    
    pc = Application.WorksheetFunction.VLookup(SpeciesName, sourceSheet.Range("PData_Properties"), int_PC, 0)
    If IsNumeric(pc) = False Then
        myErrorMsg = "The critical pressure is not numeric."
        GoTo myErrorHandler
    End If

    tc = Application.WorksheetFunction.VLookup(SpeciesName, sourceSheet.Range("PData_Properties"), int_TC, 0)
    If IsNumeric(tc) = False Then
        myErrorMsg = "The critical pressure is not numeric."
        GoTo myErrorHandler
    End If
    

    omega = Application.WorksheetFunction.VLookup(SpeciesName, sourceSheet.Range("PData_Properties"), int_Omega, 0)
    If IsNumeric(omega) = False Then
        myErrorMsg = "The critical pressure is not numeric."
        GoTo myErrorHandler
    End If
    
    If pc = 0 Or tc = 0 Then
        myErrorMsg = "Some critical properties are zero."
        GoTo myErrorHandler
    End If
    
    PR = pbara / pc
    Tr = TempK / tc
    
    z0 = BicubicInterpolation(sourceSheet.Range("PData_LK1NonPolar_Z0"), PR, Tr)
    
    
    z1 = BicubicInterpolation(sourceSheet.Range("PData_LK1NonPolar_Z1"), PR, Tr)
    
    LeeKeslerZ = z0 + omega * z1
        
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    LeeKeslerZ = 0
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
        
    Application.ScreenUpdating = True
        
    End Function
    Public Function LeeKeslerVP(Species As Variant, temperature As Variant, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'This function returns pure component vapor pressure in bar given temp in C
    'valid for non-polar species
    'For more informaton - https://en.wikipedia.org/wiki/Lee%E2%80%93Kesler_method
    'The prediction error can be up to 10% for polar components and small pressures and the calculated pressure is typically too low.
    'For pressures above 1 bar, that means, above the normal boiling point, the typical errors are below 2%.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim TempK As Double
    Dim z0 As Double
    Dim z1 As Double
    Dim omega As Variant
    Dim PR As Double
    Dim pc As Variant
    Dim Tr As Double
    Dim tc As Variant
    Dim N, i As Integer
    Dim LKConstants() As Double
    Dim f1 As Double
    Dim f2 As Double
    Dim sourceBook As Workbook
    Dim PDataExists As Boolean
    Dim PropertyNamesExist As Boolean
    Dim int_Species As Integer
    Dim int_TC As Integer
    Dim int_PC As Integer
    Dim int_Omega As Integer
    Dim int_TB As Integer
    Dim myErrorMsg As String
    Dim sourceSheet As Worksheet
    Dim tb As Double
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim TempC As Double
    Dim PData As Worksheet
    Dim SpeciesName As String
    
    i = getPdataWorksheetIndex
    
    Set PData = ThisWorkbook.Worksheets(i)
    
    fcnName = "LeeKeslerVP"
                          
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    If TypeName(Species) = "Range" Then
        If Species.Rows.Count <> 1 Or Species.Columns.Count <> 1 Then
            myErrorMsg = "The supplied species name should be a string or a single cell reference to a string equal to 'Vapor' or 'Liquid'."
            GoTo myErrorHandler
        End If
    End If
    
    SpeciesName = CStr(Species)
    
    TempC = checkInputTemperature(temperature)
    If TempC <> -273.15 Then
        TempK = TempC + 273.15
    Else
        myErrorMsg = "There is something wrong with the supplied temperature."
        GoTo myErrorHandler
    End If
    
    If ValidateWorkbook = False Then
        myErrorMsg = "Workbook validation failed"
        GoTo myErrorHandler
    End If
    
    Set sourceSheet = ThisWorkbook.Worksheets(getPdataWorksheetIndex)

    On Error Resume Next
    
    int_Species = Application.WorksheetFunction.Match("(Species)", PData.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "The '(species)' column could not be found."
        GoTo myErrorHandler
    Else
        If SpeciesName <> Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("PData_Properties"), int_Species, 0) Then
            If Err.Number = 1004 Then
                myErrorMsg = "The species could not be found."
                GoTo myErrorHandler
            End If
        End If
    End If
     
    int_TC = Application.WorksheetFunction.Match("(TC, K)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(TC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_PC = Application.WorksheetFunction.Match("(PC, bara)", sourceSheet.Range("PData_PropertyNames"), 0)
        If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(PC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_Omega = Application.WorksheetFunction.Match("(OMEGA)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(OMEGA)' not found in PData worksheet"
        GoTo myErrorHandler
    End If
    
    pc = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("A7:CZ1000"), int_PC, 0)
    
    If IsNumeric(pc) = False Then
        myErrorMsg = "The critical pressure found in the PData worksheet is not numeric."
        GoTo myErrorHandler
    End If
    
    tc = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("A7:CZ1000"), int_TC, 0)
    If IsNumeric(tc) = True Then
        Tr = TempK / tc
    Else
        myErrorMsg = "The critical temperature found in the PData worksheet is not numeric."
        GoTo myErrorHandler
    End If
    
    omega = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("A7:CZ1000"), int_Omega, 0)
    If IsNumeric(omega) <> True Then
        myErrorMsg = "The acentricity factor found in the PData worksheet is not numeric."
        GoTo myErrorHandler
    End If
    
    int_TB = Application.WorksheetFunction.Match("(TB, K)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(TB, K)' not found in PData worksheet. Cannot check if supplied temperature is above or below the normal boiling point of " & Species & "."
    Else
        tb = Application.WorksheetFunction.VLookup(SpeciesName, PData.Range("A7:CZ1000"), int_TB, 0)
            If TempK < tb And IsNumeric(tb) = True Then
                myErrorMsg = "The supplied temperature is below the normal boiling point (" & tb - 273.15 & " C). Accuracy may be degraded."
            End If
    End If
    
    On Error GoTo myErrorHandler
    
    Tr = TempK / tc
    
    If TempK > tc Then
        myErrorMsg = "The supplied temperature is above the critical temperature."
        GoTo myErrorHandler
    End If
        
    f1 = 5.92714 - 6.09648 / Tr - 1.28862 * Log(Tr) + 0.169347 * (Tr) ^ 6
    f2 = 15.2518 - 15.6875 / Tr - 13.4721 * Log(Tr) + 0.43577 * (Tr) ^ 6
    
    LeeKeslerVP = (Exp(f1 + omega * f2) * pc) * 14.6959487776468 / 14.5038
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    LeeKeslerVP = 0
    
    Application.ScreenUpdating = True
        
    End Function
    
    Public Function Get_MW(DataRange As Range, Optional errMsgsOn As Boolean = False) As Variant()

    '***************************************************************************
    'Returns the molecular weights of the vapor phase species found
    'in the passed 'DataRange' range of cells
    '***************************************************************************

    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim dataset() As Variant
    Dim outputArray() As Variant
    Dim speciesNames() As String
    Dim myErrorMsg As String
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim Phase As String
    Dim datasetErrMsgsOn As Boolean
    
    fcnName = "Get_MW"
    myErrorMsg = ""
    
    
    
    Set UDF_Range = Application.Caller
    
    Phase = "Vapor"

    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = dataset(i, iColumns.MW)
    Next i
    
    Get_MW = outputArray()
    Get_MW = Application.Transpose(Get_MW)
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(3)
        For i = 0 To 3
            outputArray(i) = 0
        Next i
    Else
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i) = 0
        Next i
    End If
    
    Get_MW = outputArray()
    Get_MW = Application.Transpose(Get_MW)
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
        
    Application.ScreenUpdating = True
        
    End Function
    Public Function GetCriticalConstants(DataRange As Range, Optional errMsgsOn As Boolean = False) As Variant()
    
    '***************************************************************************
    'Returns the critial constants of the vapor species found in the passed DataRange of cells
    'for vapor and liquid (if present)
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Index As Integer
    Dim dataset() As Variant
    
    Dim outputArray() As Variant
    Dim speciesNames() As String
    Dim myErrorMsg As String
    
    Dim UDF_Range As Range
    
    Dim fcnName As String
    Dim datasetErrMsgsOn As Boolean
    
    myErrorMsg = ""
    
    fcnName = "GetCriticalConstants"

    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    On Error GoTo myErrorHandler

    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    ReDim outputArray(0 To dataset(0, iColumns.iSpecies), 0 To 3)

    For Index = 0 To dataset(0, iColumns.iSpecies)
        outputArray(Index, 0) = dataset(Index, iColumns.tc)
        outputArray(Index, 1) = dataset(Index, iColumns.pc)  '<= send this back out as bara
        outputArray(Index, 2) = dataset(Index, iColumns.omega)
        outputArray(Index, 3) = dataset(Index, iColumns.zc)
    Next Index
    
    GetCriticalConstants = outputArray()
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:
    
    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(3, 3)
        For i = 0 To 3
            outputArray(i, 0) = 0
            outputArray(i, 1) = 0
            outputArray(i, 2) = 0
            outputArray(i, 3) = 0
        Next i
    Else
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i, 0) = 0
            outputArray(i, 1) = 0
            outputArray(i, 2) = 0
            outputArray(i, 3) = 0
        Next i
    End If

    GetCriticalConstants = outputArray()
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function
    
    Private Function calculate_B(dataset As Variant, sum_b As Double, TempK As Double, pbara As Double) As Double
    
    '***************************************************************************
    'The function is called by the 'CreateDataset' function
    'Calculates B (see reference 1) from PR1978 EOS for all vapor species and stores
    'the result in the dataset
    '***************************************************************************
    
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "calculate_B"
    
    On Error GoTo myErrorHandler
    
    calculate_B = sum_b * pbara / (GasLawR * TempK)
    
    Exit Function
    
myErrorHandler:

    calculate_B = 0
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function

    Private Function calculate_sum_b(dataset As Variant, molarComp() As Double) As Double
    
    '***************************************************************************
    'The function is called by all of the PR1978 functions
    'Calculates sum(bi) (see reference 1) from PR1978 EOS for all vapor
    '***************************************************************************
    
    Dim i As Integer
    Dim sum_b As Double
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "calculate_sum_b"
    
    On Error GoTo myErrorHandler
    
    sum_b = 0
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        sum_b = sum_b + dataset(i, iColumns.bi) * molarComp(i)
    Next i
    
    calculate_sum_b = sum_b
    
    Exit Function
    
myErrorHandler:
    
    sum_b = 0
    calculate_sum_b = sum_b
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    
    Private Function calculate_A(dataset As Variant, sum_a As Double, TempK As Double, pbara As Double) As Double
    
    '***************************************************************************
    'The function is called by all of the PR1978 functions
    'Calculates A (see reference 1) from PR1978 EOS for all vapor
    '***************************************************************************
    
    Dim myErrorMsg As String
    
    On Error GoTo myErrorHandler
    
    Dim fcnName As String
    fcnName = "calculate_A"
    
        calculate_A = sum_a * pbara / (GasLawR * TempK) ^ 2
    
    Exit Function
    
myErrorHandler:

    calculate_A = 0
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    Private Function calculate_sum_a(dataset As Variant, aij_Array() As Double, molarComp() As Double) As Double
    
    '***************************************************************************
    'The function is called by all of the PR1978 functions
    'Calculates sum(ai) (see reference 1) from PR1978 EOS for all vapor
    '***************************************************************************
    
    Dim i As Integer
    Dim j As Integer
    Dim sum_a As Double
    
    Dim fcnName As String
    Dim myErrorMsg As String
    
    fcnName = "calculate_sum_a"
    
    On Error GoTo myErrorHandler
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            sum_a = sum_a + aij_Array(i, j) * molarComp(i) * molarComp(j)
        Next j
    Next i

    calculate_sum_a = sum_a
    
    Exit Function

myErrorHandler:

    sum_a = 0
    calculate_sum_a = sum_a
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    Private Function create_aijArray(dataset As Variant, BinariesUsed As Boolean, ai_Array() As Double, kij0() As Double, kijT() As Double, TempK As Double) As Double()
    
    '***************************************************************************
    'The function is called by all of the PR1978 functions
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim aijArray() As Double
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "create_aijArray"

    ReDim aijArray(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            If BinariesUsed = True Or dataset(0, iColumns.predictive) = 1 Then
                If kij0(0, 0) <> -1 And kijT(0, 0) <> -1 Then
                    aijArray(i, j) = ((ai_Array(i) * ai_Array(j)) ^ 0.5) * (1 - (kij0(i, j) + kijT(i, j) * TempK))
                End If
            Else
                aijArray(i, j) = (ai_Array(i) * ai_Array(j)) ^ 0.5
            End If
        Next j
    Next i
            
    create_aijArray = aijArray

    Exit Function

myErrorHandler:

    aijArray(0, 0) = -1
    create_aijArray = aijArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
 
    End Function
    Private Function calculate_Derivatives(Phase As String, dataset() As Variant, moleComp() As Double, TempK As Double, pbara As Double, BinariesUsed As Boolean, _
                                         kij0() As Double, kijT() As Double, Optional passed_aiArray As Variant) As Double()
    
    '***************************************************************************
    'The function calculates various derivatives of the PR1978 EOS
    'Calculates/returns an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Dim a As Double
    Dim b As Double
    Dim Z As Double
    Dim i As Integer
    
    Dim alpha_aiArray() As Double
    Dim ai_Array() As Double
    Dim aij_Array() As Double
    
    Dim sum_a As Double
    Dim sum_b As Double
    Dim vol As Double
    Dim dPdT_V As Double
    Dim dadT_V As Double
    Dim dadT_P As Double
    Dim dBdT_P As Double
    Dim dZdT_P As Double
    Dim dPdv_T As Double
    Dim dVdT_P As Double
    Dim outputArray() As Double
    
    Dim myErrorMsg As String
    
    Dim fcnName As String
    
    ReDim outputArray(12)
    
    fcnName = "calculate_Derivatives"

    On Error GoTo myErrorHandler

    myErrorMsg = ""
    
    If IsMissing(passed_aiArray) = True Then
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        ai_Array = create_aiArray(dataset, alpha_aiArray)
    Else
        ReDim aiArray(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            aiArray(i) = passed_aiArray(i)
        Next i
    End If
    
    If alpha_aiArray(0) = -1 Then
        myErrorMsg = "create_alphaiArray returned an error."
    End If
    
    If ai_Array(0) = -1 Then
        myErrorMsg = "create_aiArray returned an error."
    End If

    aij_Array = create_aijArray(dataset, BinariesUsed, ai_Array, kij0, kijT, TempK)
    
    If ai_Array(0) = -1 Then
        myErrorMsg = "create_aijArray returned an error."
    End If
    
    sum_a = calculate_sum_a(dataset, aij_Array(), moleComp())
    
    If sum_a = 0 Then
        myErrorMsg = "calculate_sum_a returned an error."
    End If
    
    sum_b = calculate_sum_b(dataset, moleComp())
    
    If sum_b = 0 Then
        myErrorMsg = "calculate_sum_b returned an error."
    End If
    
    a = calculate_A(dataset, sum_a, TempK, pbara)
    
    If a = 0 Then
        myErrorMsg = "calculate_A returned an error."
    End If
    
    b = calculate_B(dataset, sum_b, TempK, pbara)
    
    If b = 0 Then
        myErrorMsg = "calculate_B returned an error."
    End If
    
    Z = calculate_EOS_Root(dataset, a, b, Phase)
    
    If Z = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error."
    End If
    
    If vol - b < 10 ^ -35 Then
        myErrorMsg = "The term 'volume - b' is too close to zero. Check phase."
    End If
    
    vol = Z * GasLawR * TempK / pbara
    
    dadT_V = calculate_dadt(dataset, TempK, moleComp, aij_Array, alpha_aiArray)
    
    If (vol * (vol + sum_b) + sum_b * (vol - sum_b)) <> 0 Then
        dPdv_T = -((GasLawR * TempK) / (vol - sum_b) ^ 2) + ((2 * sum_a * (vol + sum_b)) / (vol * (vol + sum_b) + sum_b * (vol - sum_b)) ^ 2)
    Else
        dPdv_T = 0
        myErrorMsg = "The term '(vol * (vol + sum_b) + sum_b * (vol - sum_b))' equals zero. dPdv_T cannot be calculated. The phase must be vapor."
    End If
    
    If (vol * (vol + sum_b) + sum_b * (vol - sum_b)) <> 0 And (vol - sum_b) <> 0 Then
        dPdT_V = ((GasLawR / (vol - sum_b)) - (dadT_V / (vol * (vol + sum_b) + sum_b * (vol - sum_b))))
    Else
        dPdT_V = 0
        myErrorMsg = "The term '(vol * (vol + sum_b) + sum_b * (vol - sum_b))' or '(vol - sum_b)' equal zero. dPdT_V cannot be calculated. The phase must be vapor."
    End If
    

    
    dadT_P = pbara * (dadT_V - ((2 * sum_a) / TempK)) / ((GasLawR) ^ 2 * TempK ^ 2)

    dBdT_P = -sum_b * pbara / (GasLawR * TempK ^ 2)
    
    If (3 * Z ^ 2 + 2 * (b - 1) * Z + (a - 2 * b - 3 * b ^ 2)) <> 0 Then
        dZdT_P = (dadT_P * (b - Z) + dBdT_P * (6 * b * Z + 2 * Z - 3 * b ^ 2 - 2 * b + a - Z ^ 2)) / (3 * Z ^ 2 + 2 * (b - 1) * Z + (a - 2 * b - 3 * b ^ 2))
    Else
        dZdT_P = 0
        myErrorMsg = "The term '(3 * z ^ 2 + 2 * (B - 1) * z + (A - 2 * B - 3 * B ^ 2))' equals zero. dZdT_P cannot be calculated. Check phase."
    End If
    
    dVdT_P = (GasLawR / pbara) * (TempK * (dZdT_P) + Z)
    
    outputArray(iColumns.dadT_constV) = dadT_V
    outputArray(iColumns.dPdv_constT) = dPdv_T
    outputArray(iColumns.dPdT_constV) = dPdT_V
    outputArray(iColumns.dadT_constP) = dadT_P
    outputArray(iColumns.dBdT_constP) = dBdT_P
    outputArray(iColumns.dZdT_constP) = dZdT_P
    outputArray(iColumns.dVdT_constP) = dVdT_P
    outputArray(iColumns.sumb) = sum_b
    outputArray(iColumns.suma) = sum_a
    outputArray(iColumns.a) = a
    outputArray(iColumns.b) = b
    outputArray(iColumns.Z) = Z
    outputArray(iColumns.vol) = vol
    
    calculate_Derivatives = outputArray
    
    Exit Function

myErrorHandler:
    
    ReDim ouputArray(1)
    ouputArray(0) = 987654321.12345
    calculate_Derivatives = outputArray                                      '<= Error flag
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    
    Private Function calculate_PhaseZ(Phase As String, dataset() As Variant, molesArray() As Double, TempK As Double, pbara As Double, BinariesUsed As Boolean, _
                                         kij0() As Double, kijT() As Double, Optional passed_aiArray As Variant) As Double
    
    '***************************************************************************
    'This function is called by all of the PR1978 functions
    'This function calculates the PR1978 EOS compressibility factor
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Dim a As Double
    Dim b As Double
    Dim Z As Double
    Dim i As Integer
    
    Dim alpha_aiArray() As Double
    Dim ai_Array() As Double
    Dim bi_Array() As Double
    Dim aij_Array() As Double
    
    Dim sum_a As Double
    Dim sum_b As Double
    
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "calculate_PhaseZ"

    On Error GoTo myErrorHandler

    myErrorMsg = ""
    
    If IsMissing(passed_aiArray) = True Then
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        ai_Array = create_aiArray(dataset, alpha_aiArray)
    Else
        ReDim aiArray(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            aiArray(i) = passed_aiArray(i)
        Next i
    End If
    
    If alpha_aiArray(0) = -1 Then
        myErrorMsg = "create_alphaiArray returned an error."
    End If
    
    If ai_Array(0) = -1 Then
        myErrorMsg = "create_aiArray returned an error."
    End If

    aij_Array = create_aijArray(dataset, BinariesUsed, ai_Array, kij0, kijT, TempK)
    
    If ai_Array(0) = -1 Then
        myErrorMsg = "create_aijArray returned an error."
    End If
    
    sum_a = calculate_sum_a(dataset, aij_Array(), molesArray())
    
    If sum_a = 0 Then
        myErrorMsg = "calculate_sum_a returned an error."
    End If
    
    sum_b = calculate_sum_b(dataset, molesArray())
    
    If sum_b = 0 Then
        myErrorMsg = "calculate_sum_b returned an error."
    End If
    
    a = calculate_A(dataset, sum_a, TempK, pbara)
    
    If a = 0 Then
        myErrorMsg = "calculate_A returned an error."
    End If
    
    b = calculate_B(dataset, sum_b, TempK, pbara)
    
    If b = 0 Then
        myErrorMsg = "calculate_B returned an error."
    End If
    
    Z = calculate_EOS_Root(dataset, a, b, Phase)
    
    If Z = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error."
    End If
    
    calculate_PhaseZ = Z
    
    Exit Function

myErrorHandler:
    
    calculate_PhaseZ = -500                                       '<= Error flag
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    Private Function calculate_Phi(Phase As String, dataset() As Variant, molesArray() As Double, TempK As Double, pbara As Double, BinariesUsed As Boolean, _
                                         kij0() As Double, kijT() As Double, Optional passed_aiArray As Variant) As Variant()
    
    '***************************************************************************
    'This function is called by all of the PR1978 functions
    'This function calculates the PR1978 EOS fugacity coefficients
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Dim i As Integer
    Dim j As Integer
    
    Dim a As Double
    Dim b As Double
    Dim Z As Double
    Dim sum_a As Double
    Dim sum_b As Double
    Dim errorTest As Double
    
    Dim alpha_aiArray() As Double
    Dim ai_Array() As Double
    Dim bi_Array() As Double
    Dim aij_Array() As Double
    Dim xi_aijArray() As Double

    Dim Phi() As Variant
    
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "calculate_Phi"
    
    ReDim Phi(dataset(0, iColumns.iSpecies) + 2)
    
    myErrorMsg = ""
    
    If IsArrayAllocated(passed_aiArray) = False Then
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        ai_Array = create_aiArray(dataset, alpha_aiArray)
    Else
        ReDim ai_Array(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            ai_Array(i) = passed_aiArray(i)
        Next i
    End If
    

    aij_Array = create_aijArray(dataset, BinariesUsed, ai_Array, kij0, kijT, TempK)

    
    sum_a = calculate_sum_a(dataset, aij_Array(), molesArray())
    
    
    sum_b = calculate_sum_b(dataset, molesArray())
    
    If sum_b = 0 Then
        myErrorMsg = "sum_b is zero and it will cause a divide by zero in the Phi() function."
        GoTo myErrorHandler
    End If
    
    a = calculate_A(dataset, sum_a, TempK, pbara)
    
    b = calculate_B(dataset, sum_b, TempK, pbara)
    
    Z = calculate_EOS_Root(dataset, a, b, Phase)
    
    If Z - b < 0 Then
        myErrorMsg = "the term z - B is less than zero. Check supplied Phase."
        GoTo myErrorHandler
    End If
    
    xi_aijArray = create_xi_aijArray(dataset, aij_Array, molesArray)
        
    If ((Z + (2 ^ 0.5 + 1) * b) / (Z - (2 ^ 0.5 - 1) * b)) < 0 Then
        myErrorMsg = "Check supplied Phase. The term ((z + (2 ^ 0.5 + 1) * B) / (z - (2 ^ 0.5 - 1) * B)) is less than zero. This will cause and error in the ln() function of the Phi() expression."
        GoTo myErrorHandler
    End If
    
    On Error Resume Next
    
    For i = 0 To dataset(0, iColumns.iSpecies)
    
        Phi(i) = (sum_a / (GasLawR * TempK * sum_b * 2 * 2 ^ 0.5)) * ((2 * xi_aijArray(i) / sum_a) - (dataset(i, iColumns.bi) / sum_b)) * _
                                                        Log((Z + (2 ^ 0.5 + 1) * b) / (Z - (2 ^ 0.5 - 1) * b))
        If Err.Number = 5 Then
            myErrorMsg = "The Phi() function attemped to take the ln() of a negative number. Check supplied Phase."
            Err.Clear
            GoTo myErrorHandler
        ElseIf Err.Number = 6 Then
            myErrorMsg = "sum_b or (z - (2 ^ 0.5 - 1) * B) equal zero. Divide by zero error."
            Err.Clear
            GoTo myErrorHandler
        Else
            If Err.Number <> 0 Then
                myErrorMsg = "Check phase."
                Err.Clear
            GoTo myErrorHandler
            End If
        End If
    
        Phi(i) = (dataset(i, iColumns.bi) / sum_b) * (Z - 1) - Log(Z - b) - Phi(i)

        If Err.Number = 6 Or Err.Number = 5 Then
            myErrorMsg = "Either divide by by zero or log() of a negative number error. Typically the specified phase is incorrect when this happens."
            Err.Clear
            GoTo myErrorHandler
        Else
            If Err.Number <> 0 Then
                myErrorMsg = "Check phase."
                Err.Clear
            GoTo myErrorHandler
            End If
        End If
        
        On Error GoTo myErrorHandler
            
        Phi(i) = Exp(Phi(i))
        
    Next i
    
    Phi(dataset(0, iColumns.iSpecies) + 1) = Z                      '<= Add z to bottom of array because in can! : )
    Phi(dataset(0, iColumns.iSpecies) + 2) = 0                      '<= Error flag - 0 = no error, -500 = error
    
    calculate_Phi = Phi
    
    Exit Function

myErrorHandler:
    
    Phi(dataset(0, iColumns.iSpecies) + 2) = -500                   '<= Error flag - 0 = no error, -500 = error
    calculate_Phi = Phi()
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
  
    End Function
    
    Private Function calculate_S_Departure(Phase As String, dataset() As Variant, moleComp() As Double, TempK As Double, pbara As Double, BinariesUsed As Boolean, _
                                         kij0() As Double, kijT() As Double, Optional passed_aiArray As Variant) As Double
    
    '***************************************************************************
    'This function is called by the Entropy function
    'This function calculates the PR1978 EOS entropy departure function
    '***************************************************************************
    
    On Error GoTo myErrorHandler:

    Dim a As Double
    Dim A0 As Double
    Dim b As Double
    Dim B0 As Double
    Dim z0 As Double
    Dim Z As Double
    Dim i As Integer
    Dim j As Integer
    Dim alpha_aiArray() As Double
    Dim ai_Array() As Double
    Dim bi_Array() As Double
    Dim aij_Array() As Double
    Dim xi_aijArray() As Double
    Dim sum_a As Double
    Dim sum_b As Double
    Dim Phi() As Double
    Dim dadT As Double
    Dim departS As Double
    Dim departS0 As Double
    Dim dadt_Array() As Double
    Dim myErrorMsg As String
    
    Dim fcnName As String
    
    fcnName = "calculate_S_Departure"
    
    myErrorMsg = ""
    
    If IsMissing(passed_aiArray) = True Then
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        ai_Array = create_aiArray(dataset, alpha_aiArray)
    Else
        ReDim aiArray(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            aiArray(i) = passed_aiArray(i)
        Next i
    End If

    aij_Array = create_aijArray(dataset, BinariesUsed, ai_Array, kij0, kijT, TempK)
    
    sum_a = calculate_sum_a(dataset, aij_Array(), moleComp())
    
    sum_b = calculate_sum_b(dataset, moleComp())
    
    If sum_b = 0 Then
        myErrorMsg = "calculate_sum_b returned an error."
    End If
    
    a = calculate_A(dataset, sum_a, TempK, pbara)
    
    A0 = calculate_A(dataset, sum_a, TempK, 1)
    
    b = calculate_B(dataset, sum_b, TempK, pbara)
    
    B0 = calculate_B(dataset, sum_b, TempK, 1)
    
    Z = calculate_EOS_Root(dataset, a, b, Phase)
    
    If Z = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error."
    End If
    
    z0 = calculate_EOS_Root(dataset, A0, B0, Phase)
    
    If z0 = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error."
    End If
    
    If Z - b < 0 Or z0 - B0 < 0 Then
        myErrorMsg = "the term z - B is less than zero. Check supplied Phase."
    End If
        
    dadT = calculate_dadt(dataset, TempK, moleComp, aij_Array, alpha_aiArray)
    
    On Error Resume Next
    
    departS = 100000 * ((GasLawR) * Log(Z - b) + (dadT / ((2 * 2 ^ 0.5) * sum_b)) * Log((Z + (2 ^ 0.5 + 1) * b) / (Z - (2 ^ 0.5 - 1) * b)))
                        'the 100,000 factor above converts from m3-bar/K-g-mole to kJ/kg-mole
    If Err.Number = 5 Then
        myErrorMsg = "Check supplied Phase. The Phi() function attemped to take the ln() of a negative number."
        Err.Clear
        GoTo myErrorHandler
    ElseIf Err.Number = 6 Then
        myErrorMsg = "Divide by zero error. The terms sum_b or (z - (2 ^ 0.5 - 1) * B) equal zero."
        Err.Clear
        GoTo myErrorHandler
    Else
        If Err.Number <> 0 Then
            myErrorMsg = Err.Description
            Err.Clear
        GoTo myErrorHandler
        End If
    End If
                        
    departS0 = 100000 * ((GasLawR) * Log(z0 - B0) + (dadT / ((2 * 2 ^ 0.5) * sum_b)) * Log((z0 + (2 ^ 0.5 + 1) * B0) / (z0 - (2 ^ 0.5 - 1) * B0)))
                        'the 100,000 factor above converts from m3-bar/K-g-mole to kJ/kg-mole
    If Err.Number = 5 Then
        myErrorMsg = "Check supplied Phase. The Phi() function attemped to take the ln() of a negative number."
        Err.Clear
        GoTo myErrorHandler
    ElseIf Err.Number = 6 Then
        myErrorMsg = "Divide by zero error. The terms sum_b or (z0 - (2 ^ 0.5 - 1) * B0) equal zero."
        Err.Clear
        GoTo myErrorHandler
    Else
        If Err.Number <> 0 Then
            myErrorMsg = Err.Description
            Err.Clear
        GoTo myErrorHandler
        End If
    End If
    
    
                        
    calculate_S_Departure = departS - departS0
    
    Exit Function
    
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    calculate_S_Departure = 987654321.123457
    
    End Function

    Private Function calculate_H_Departure(Phase As String, dataset() As Variant, moleComp() As Double, TempK As Double, pbara As Double, BinariesUsed As Boolean, _
                                         kij0() As Double, kijT() As Double, Optional passed_aiArray As Variant) As Double
    
    '***************************************************************************
    'This function is called by the Enthalpy function
    'This function calculates the PR1978 EOS enthalpy departure function
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Dim a As Double
    Dim A0 As Double
    Dim b As Double
    Dim B0 As Double
    Dim z0 As Double
    Dim Z As Double
    Dim i As Integer
    Dim j As Integer
    Dim alpha_aiArray() As Double
    Dim ai_Array() As Double
    Dim bi_Array() As Double
    Dim aij_Array() As Double
    Dim xi_aijArray() As Double
    Dim sum_a As Double
    Dim sum_b As Double
    Dim Phi() As Double
    Dim dadT As Double
    Dim departH As Double
    Dim departH0 As Double
    Dim errorTest As Double
    Dim myErrorMsg As String
    Dim dadt_Array() As Double
    Dim fcnName As String

    fcnName = "calculate_H_Departure"

    myErrorMsg = ""
    
    If IsMissing(passed_aiArray) = True Then
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        ai_Array = create_aiArray(dataset, alpha_aiArray)
    Else
        ReDim aiArray(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            aiArray(i) = passed_aiArray(i)
        Next i
    End If

    aij_Array = create_aijArray(dataset, BinariesUsed, ai_Array, kij0, kijT, TempK)
    
    sum_a = calculate_sum_a(dataset, aij_Array(), moleComp())
    
    sum_b = calculate_sum_b(dataset, moleComp())
    
    If sum_b = 0 Then
        myErrorMsg = "sum_b is zero and will cause a divide by zero error."
        GoTo myErrorHandler
    End If
    
    a = calculate_A(dataset, sum_a, TempK, pbara)
    
    A0 = calculate_A(dataset, sum_a, TempK, 1)
    
    b = calculate_B(dataset, sum_b, TempK, pbara)
    
    B0 = calculate_B(dataset, sum_b, TempK, 1)
    
    Z = calculate_EOS_Root(dataset, a, b, Phase)
    
    If Z = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error when calculating z."
        GoTo myErrorHandler
    End If
    
    z0 = calculate_EOS_Root(dataset, A0, B0, Phase)
    
    If z0 = -500 Then
        myErrorMsg = "calculate_EOS_Root returned an error when calculating z0."
        GoTo myErrorHandler
    End If
       
    dadT = calculate_dadt(dataset, TempK, moleComp, aij_Array, alpha_aiArray)
    
    If (Z + (2 ^ 0.5 + 1) * b) / (Z - (2 ^ 0.5 - 1) * b) <= 0 Then
        myErrorMsg = "Check supplied Phase. The term (z + (2 ^ 0.5 + 1) * B) / (z - (2 ^ 0.5 - 1) * B) is less than or equal to zero and will cause an error in the ln() function of the Phi() expression."
    End If
    
    On Error Resume Next
    
    departH = 100000 * (GasLawR * TempK * (Z - 1) + (TempK * dadT - sum_a) / (2 * (2 ^ 0.5) * sum_b) * Log((Z + (2 ^ 0.5 + 1) * b) / (Z - (2 ^ 0.5 - 1) * b)))
                        'the 100,000 factor above converts from m3-bar/K-g-mole to kJ/kg-mole
                        
        If Err.Number = 5 Then
            myErrorMsg = "Check supplied Phase. The Phi() function attemped to take the ln() of a negative number."
            Err.Clear
            GoTo myErrorHandler
        ElseIf Err.Number = 6 Then
            myErrorMsg = "Divide by zero error. The terms sum_b or (z - (2 ^ 0.5 - 1) * B) equal zero."
            Err.Clear
            GoTo myErrorHandler
        Else
            If Err.Number <> 0 Then
                myErrorMsg = Err.Description
                Err.Clear
            GoTo myErrorHandler
            End If
        End If
    
    departH0 = 100000 * (GasLawR * TempK * (z0 - 1) + (TempK * dadT - sum_a) / (2 * (2 ^ 0.5) * sum_b) * Log((z0 + (2 ^ 0.5 + 1) * B0) / (z0 - (2 ^ 0.5 - 1) * B0)))
                        'the 100,000 factor above converts from m3-bar/K-g-mole to kJ/kg-mole
        If Err.Number = 5 Then
            myErrorMsg = "Check supplied Phase. The Phi() function attemped to take the ln() of a negative number."
            Err.Clear
            GoTo myErrorHandler
        ElseIf Err.Number = 6 Then
            myErrorMsg = "Divide by zero error. The terms sum_b or (z0 - (2 ^ 0.5 - 1) * B0) equal zero."
            Err.Clear
            GoTo myErrorHandler
        Else
            If Err.Number <> 0 Then
                myErrorMsg = Err.Description
                Err.Clear
            GoTo myErrorHandler
            End If
        End If
    
    calculate_H_Departure = departH - departH0
    
    Exit Function
    
myErrorHandler:
    
    calculate_H_Departure = 987654321.123457
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
   
    End Function

    Private Function calculate_dadt(dataset As Variant, TempK As Double, moleComp() As Double, aij_Array() As Double, alpha_aiArray() As Double) As Double
    
    '***************************************************************************
    'This function is called by the Enthalpy, Entropy and Derivatives function
    'This function calculates the PR1978 EOS da/dT function
    '***************************************************************************
    
    Dim i As Integer
    Dim j As Integer
    Dim dadt_Array() As Double
    Dim myErrorMsg As String
    Dim dadT As Double
    
    Dim fcnName As String
    
    dadT = 0

    fcnName = "calculate_dadt"
    
    On Error GoTo myErrorHandler
    
    ReDim dadt_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    
    On Error Resume Next
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
                    dadt_Array(i, j) = -(moleComp(i) * moleComp(j) * aij_Array(i, j) / (2 * TempK ^ 0.5)) * _
                    ((dataset(j, iColumns.Ki) / (alpha_aiArray(j) * dataset(j, iColumns.tc)) ^ 0.5) + _
                    dataset(i, iColumns.Ki) / (alpha_aiArray(i) * dataset(i, iColumns.tc)) ^ 0.5)
                    
                    dadT = dadT + dadt_Array(i, j)
                    
            If Err.Number = 6 Or Err.Number = 5 Then
                myErrorMsg = "(alpha_aiArray(i or j) * Dataset(i or j, iColumns.tc)) ^ 0.5 = 0. Divide by zero error."
                Err.Clear
                GoTo myErrorHandler
            Else
                If Err.Number <> 0 Then
                    myErrorMsg = Err.Description
                    Err.Clear
                    GoTo myErrorHandler
                End If
            End If
        Next j
    Next i
    
    calculate_dadt = dadT
    
    Exit Function
    
myErrorHandler:

    calculate_dadt = 0

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    Private Function calculate_EOS_Root(dataset As Variant, a As Double, b As Double, Phase As String)
    
    '***************************************************************************
    'This function is called by all of the PR1978 EOS functions
    'This function calculates the the root to the PR1978 EOS (reference 1)
    'This function uses the GetLargestRoot function in the Math module (reference 6)
    '***************************************************************************
    
    Dim Z As Double
    
    Dim myErrorMsg As String
    Dim fcnName As String

    fcnName = "calculate_EOS_Root"
    
    On Error GoTo myErrorHandler
    
    If LCase(Phase) = "vapor" Then
        Z = GetLargestRoot(1, (b - 1), a - 3 * b ^ 2 - 2 * b, (-a * b + b ^ 2 + b ^ 3))
    End If
    
    If LCase(Phase) = "liquid" Then
        Z = GetSmallestRoot(1, (b - 1), a - 3 * b ^ 2 - 2 * b, (-a * b + b ^ 2 + b ^ 3))
    End If
    
    calculate_EOS_Root = Z
    
    If Z = -500 Then
        myErrorMsg = "GetSmallestRoot/GetLargestRoot function returned an errror."
        GoTo myErrorHandler
    End If
    
    Exit Function
    
myErrorHandler:

    Z = -500
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    Public Function createPredictivekijTArray(dataset As Variant, deComp As Variant, ByRef TempK As Double, aiArray() As Double)
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim array_1to17() As Variant
    Dim array_18to31() As Variant
    Dim array_28to31() As Variant
    Dim grpInteractionParamA() As Double
    Dim grpInteractionParamB() As Double
    Dim tempArray() As Variant
    Dim decompArray() As Double
    Dim alphaiArray() As Double
    Dim DoubleSum() As Double
    Dim predict_kijT() As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    
    ReDim grpInteractionParamA(30, 30)
    ReDim grpInteractionParamB(30, 30)
    ReDim predict_kijT(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    ReDim DoubleSum(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    
    fcnName = createPredictivekijTArray
    
    If deComp.Rows.Count - 1 <> dataset(0, iColumns.iSpecies) + 1 Or deComp.Columns.Count <> 1 + 31 Then
        myErrorMsg = "The supplied decomposition range row count must equal the number of species and the column count must equal 31 groups plus 1 for the species column."
        GoTo myErrorHandler
    End If
        
    If dataset(0, iColumns.predictive) = 1 Then
    
        ReDim decompArray(dataset(0, iColumns.iSpecies), 30)
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            For j = 0 To 30
                decompArray(i, j) = deComp(i + 2, j + 2)
            Next j
        Next i
        
        ReDim array_1to17(16, 30)
        array_1to17 = Array( _
        Array(0, 65.54, 214.9, 431.6, 28.48, 3.775, 98.83, 103.6, 624.9, 43.58, 293.4, 144.8, 38.09, 159.6, 789.6, 3557, 7.892, 48.73, 102.6, 47.01, 174, 91.24, 416.3, 11.27, 322.2, 86.1, 0, 0, 0, 0, 0), _
        Array(65.54, 0, 39.05, 134.5, 37.75, 29.85, 25.05, 5.147, -17.84, 8.579, 63.48, 141.4, 83.73, 136.6, 439.9, 4324, 59.71, 9.608, 64.85, 34.31, 155.4, 44, 520.52, 113.6, 55.9, 107.4, 0, 0, 0, 0, 0), _
        Array(214.9, 39.05, 0, -86.13, 131.4, 156.1, 56.62, 48.73, 0, 73.09, -120.8, 191.8, 383.6, 192.5, 374, 971.4, 147.9, 84.76, 91.62, 0, 326, 0, 728.1, 185.8, -70, 0, 0, 0, 0, 0, 0), _
        Array(431.6, 134.5, -86.13, 0, 309.5, 388.1, 170.5, 128.3, 0, 208.6, 25.05, 377.5, 341.8, 330.8, 685.9, 0, 366.8, 181.2, 0, 0, 548.3, 0, 0, 899, 0, 0, 0, 0, 0, 0, 0), _
        Array(28.48, 37.75, 131.4, 309.5, 0, 0, 9.951, 67.26, 106.7, 249.1, 33.97, 188, 136.57, 30.88, 190.1, 701.7, 2277.12, 19.22, 48.73, 0, 0, 156.1, 14.43, 394.5, 15.97, 205.89, 0, 0, 44.61, 436.14, 0), _
        Array(3.775, 29.85, 156.1, 388.1, 9.951, 0, 41.18, 67.94, 0, 12.7, 118, 136.2, 61.59, 157.2, 0, 2333, 7.549, 26.77, 0, 0, 137.6, 15.42, 581.3, 43.81, 0, 0, 0, 0, 0, 0, 0), _
        Array(98.83, 25.05, 56.62, 170.5, 67.26, 41.18, 0, -16.47, 52.5, 28.82, 129, 98.48, 185.3, 21.28, 277.6, 2268, 25.74, 9.951, -16.47, 3.775, 288.9, 153.4, 753.6, 195.6, 37.1, 233.4, 0, 0, 0, 0, 0), _
        Array(103.6, 5.147, 48.73, 128.3, 106.7, 67.94, -16.47, 0, -328, 37.4, -99.17, 154.4, 343.8, 9.608, 1002, 543.5, 97.8, -48.38, 343.1, 242.9, 400.1, 125.77, 753.6, 0, -196.6, 177.1, 0, 0, 0, 0, 0), _
        Array(624.9, -17.84, 0, 0, 249.1, 0, 52.5, -328, 0, 140.7, -99.17, 331.1, 702.4, 9.608, 1002, 1340, 209.7, 669.8, 0, 0, 602.9, 197, 753.6, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(43.58, 8.579, 73.09, 208.6, 33.97, 12.7, 28.82, 37.4, 140.7, 0, 139, 144.1, 179.5, 117.4, 493.1, 4211, 35.34, -15.44, 159.6, 31.91, 236.1, 113.1, 0, 1269, 0, 181.2, -27.5, 0, 0, 0, 0), _
        Array(293.4, 63.48, -120.8, 25.05, 188, 118, 129, -99.17, -99.17, 139, 0, 216.2, 331.5, 71.37, 463.2, 244, 297.2, 260.1, 0, 151.3, -51.82, 0, 0, 0, 0, 102.3, 0, 0, 0, 0, 0), _
        Array(144.8, 141.4, 191.8, 377.5, 136.57, 136.2, 98.48, 154.4, 331.1, 144.1, 216.2, 0, 113.92, 135.2, 0, 559.3, 73.09, 60.74, 74.81, 87.85, 261.13, 87.85, 685.9, 177.75, 54.9, 154.42, 5.1, 83.04, 0, 124.91, 3.77), _
        Array(38.09, 83.73, 383.6, 341.8, 30.88, 61.59, 185.3, 343.8, 702.4, 179.5, 331.5, 113.92, 0, 319.5, 0, 2574, 45.3, 59.71, 541.5, 0, 65.2, 23.33, 204.7, 6.488, 282.4, 2.4, 258.73, 0, 585.75, 263.54, 101.57), _
        Array(159.6, 136.6, 192.5, 330.8, 190.1, 157.2, 21.28, 9.608, 9.608, 117.4, 71.37, 135.2, 319.5, 0, 0, -157, 603.94, 0, 0, 0, 0, 145.84, 278.63, 0, 0, 0, 0, 0, 101.91, 0, 0), _
        Array(789.6, 439.9, 374, 685.9, 701.7, 0, 277.6, 1002, 1002, 493.1, 463.2, 0, 0, -157, 0, 30.88, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(3557, 4324, 971.4, 0, 2277.12, 2333, 2268, 543.5, 1340, 4211, 244, 559.3, 2574, 603.94, 30.88, 0, 1650, 2243, 0, 0, 830.76, 278.63, 0, 0, 374.4, 1376, 0, 0, -550.06, 0, 568.94), _
        Array(7.892, 59.71, 147.9, 366.8, 19.22, 7.549, 25.74, 97.8, 209.7, 35.34, 297.2, 73.09, 45.3, 0, 0, 1650, 0, 14.76, -518, -98.8, 151.3, 84.55, 569.6, 0, 0, 0, 0, 0, 0, 0, 0))

        ReDim array_18to31(13, 30)
        array_18to31 = Array( _
        Array(48.73, 9.608, 84.76, 181.2, 48.73, 26.77, 9.951, -48.38, 669.8, -15.44, 260.1, 60.74, 59.71, 0, 0, 2243, 14.76, 0, 24.71, 14.07, 175.7, 0, 644.3, 203, 26.8, 0, 0, 0, 0, 0, 0), _
        Array(102.6, 64.85, 91.62, 0, 0, 0, -16.47, 343.1, 0, 159.6, 0, 74.81, 541.5, 0, 0, 0, -518, 24.71, 0, 23.68, 621.4, 0, 0, 0, -141, 0, 0, 0, 0, 0, 0), _
        Array(47.01, 34.31, 0, 0, 0, 0, 3.775, 242.9, 0, 31.91, 151.3, 87.85, 0, 0, 0, 0, -98.8, 14.07, 23.68, 0, 460.8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(174, 155.4, 326, 548.3, 156.1, 137.6, 288.9, 400.1, 602.9, 236.1, -51.82, 261.13, 65.2, 145.84, 0, 830.76, 151.3, 175.7, 621.4, 460.8, 0, 75.84, 138.7, 128.2, 0, 0, 0, 0, 701.73, 0, 0), _
        Array(91.24, 44, 0, 0, 14.43, 15.42, 153.4, 125.77, 197, 113.1, 0, 87.85, 23.33, 278.63, 0, 278.63, 84.55, 0, 0, 0, 75.84, 0, 260.1, 4.042, 0, 0, 309.17, 0, 0, 0, 0), _
        Array(416.3, 520.52, 728.1, 0, 394.5, 581.3, 753.6, 753.6, 753.6, 0, 0, 685.9, 204.7, 0, 0, 0, 569.6, 644.3, 0, 0, 138.7, 260.1, 0, 243.1, 0, 0, 0, 0, 0, 0, 0), _
        Array(11.27, 113.6, 185.8, 899, 15.97, 43.81, 195.6, 0, 0, 1269, 0, 177.75, 6.488, 0, 0, 0, 0, 203, 0, 0, 128.2, 4.042, 243.1, 0, 299.91, 4.8, 110.84, 0, 630.02, 278.63, 0), _
        Array(322.2, 55.9, -70, 0, 205.89, 0, 37.1, -196.6, 0, 0, 0, 54.9, 282.4, 0, 0, 374.4, 0, 26.8, -141, 0, 0, 0, 0, 299.91, 0, 0, 339.94, 172.26, 0, 0, 0), _
        Array(86.1, 107.4, 0, 0, 0, 0, 233.4, 177.1, 0, 181.2, 102.3, 154.42, 2.4, 0, 0, 1376, 0, 0, 0, 0, 0, 0, 0, 4.8, 339.94, 0, 0, 0, 0, 271.09, 120.1), _
        Array(0, 0, 0, 0, 0, 0, 0, 0, 0, -27.5, 0, 5.1, 258.73, 0, 0, 0, 0, 0, 0, 0, 0, 309.17, 0, 110.84, 172.26, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 44.61, 0, 0, 0, 0, 0, 0, 83.04, 0, 101.91, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 436.14, 0, 0, 0, 0, 0, 0, 0, 585.75, 0, 0, -550.06, 0, 0, 0, 0, 701.73, 0, 0, 630.02, 0, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 124.91, 263.54, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 278.63, 0, 271.09, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 74.81, 0, 0, 0, 0, 0, 0, 3.77, 101.57, 0, 0, 568.94, 0, 0, 0, 0, 0, 0, 0, 0, 0, 120.1, 0, 0, 0, 0, 0))

        For i = 0 To 16
            For j = 0 To 30
                grpInteractionParamA(i, j) = array_1to17(i)(j) * 10 '<= Convert from MPa to bara
            Next j
        Next i
        
        For i = 0 To 13
            For j = 0 To 30
                grpInteractionParamA(i + 16, j) = array_18to31(i)(j) * 10 '<= Convert from MPa to bara
            Next j
        Next i
        
        ReDim array_1to17(16, 30)
        array_1to17 = Array( _
        Array(0, 105.7, 249.9, 575, 20.25, 8.922, 136.2, 103.6, 774.1, 60.05, 170.9, 401.5, 88.19, 227.8, 1829, 11195, 35, 44.27, 260.1, 169.5, 239.5, 94.24, 513.4, 55.48, 201.4, 87.5, 0, 0, 0, 0, 0), _
        Array(105.7, 0, 41.59, 183.9, 74.81, 65.88, 64.51, -7.549, -4.118, 27.79, -74.46, 237.1, 188.7, 124.6, 504.8, 12126, 82.35, 50.79, 51.82, 51.13, 240.9, 45.55, 673.22, 231.6, -28.5, 200.8, 0, 0, 0, 0, 0), _
        Array(249.9, 41.59, 0, 85.1, 157.5, 96.77, 129.7, -89.22, 0, 71.37, 18.53, 380.9, 375.4, 562.8, 520.9, 567.6, -55.59, 193.2, 54.9, 0, 287.9, 0, 750.9, 634.2, 233.7, 0, 0, 0, 0, 0, 0), _
        Array(575, 183.9, 85.1, 0, 35.69, -224.8, 284.1, 189.1, 0, 294.4, 81.33, 162.7, 635.2, -297.2, 1547, 0, -219.3, 419, 0, 0, 2343, 0, 0, 4655, 0, 0, 0, 0, 0, 0, 0), _
        Array(20.25, 74.81, 157.5, 35.69, 0, 13.37, 167.5, 190.8, 408.3, 5.49, 473.9, 214.81, 37.06, 307.46, 1318, 4719.63, 33.29, 68.29, 0, 0, 92.99, 20.92, 378.1, 24.48, 323.59, 0, 0, -95.05, 958.75, 0, 107.06), _
        Array(8.922, 65.88, 96.77, -224.8, 13.37, 0, 50.79, 210.7, 0, 73.43, -212.8, 235.7, 84.92, 217.1, 0, 5147, 20.93, -5.147, 0, 0, 150, 33.3, 517.1, 53.1, 0, 0, 0, 0, 0, 0, 0), _
        Array(136.2, 64.51, 129.7, 284.1, 167.5, 50.79, 0, 16.47, 251.2, 65.54, 36.72, 253.6, 490.7, 6.177, 449.5, 62.18, 78.92, 19.9, 61.42, 1.716, 189.1, 153.4, 590.5, 361.3, -23.7, 404.9, 0, 0, 0, 0, 0), _
        Array(103.6, -7.549, -89.22, 189.1, 190.8, 210.7, 16.47, 0, -569.3, 53.53, -193.5, 374.4, 1712, -36.72, -736.4, 411.8, 67.94, 27.79, 880.2, -7.206, 1201, -231.1, 590.5, 0, -397.4, 2559.4, 0, 0, 0, 0, 0), _
        Array(774.1, -4.118, 0, 0, 408.3, 0, 251.2, -569.3, 0, 277.6, -193.5, 276.6, 1889, -36.72, -736.4, -65.88, 3819, 589.5, 0, 0, 1463, -238.8, 590.5, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(60.05, 27.79, 71.37, 294.4, 5.49, 73.43, 65.54, 53.53, 277.6, 0, 35.69, 354.1, 546.6, 166.4, 832.1, 13031, 52.5, 24.36, 140.7, 69.32, 192.5, 143.6, 0, 18666, 0, 281.4, 50.1, 0, 0, 0, 0), _
        Array(170.9, -74.46, 18.53, 81.33, 473.9, -212.8, 36.72, -193.5, -193.5, 35.69, 0, -132.8, 389.8, -127.7, -337.7, -60.39, -647.2, 134.9, 0, 2.745, 34.31, 0, 0, 0, 0, 988, 0, 0, 0, 0, 0), _
        Array(401.5, 237.1, 380.9, 162.7, 214.81, 235.7, 253.6, 374.4, 276.6, 354.1, -132.8, 0, 212.41, 199.02, 0, 277.95, 106.7, 183.9, -266.6, 66.91, 300.94, 190.79, 559.3, 86.82, 59.02, 109.81, 48.38, 165.74, 0, 241.57, 14.07), _
        Array(88.19, 188.7, 375.4, 635.2, 37.06, 84.92, 490.7, 1712, 1889, 546.6, 389.8, 212.41, 0, 550.06, 0, 5490.33, 92.65, 227.2, 94.71, 0, 70.1, -25.4, 222.8, 8.77, 362.71, 4.8, 100.54, 0, 1011.25, 255.99, 230.94), _
        Array(227.8, 124.6, 562.8, -297.2, 307.46, 217.1, 6.177, -36.72, -36.72, 166.4, -127.7, 199.02, 550.06, 0, 153.7, 599.13, 0, 0, 0, 0, 823.55, 404.23, 0, 0, 0, 0, 0, 98.14, 0, 0, 0), _
        Array(1829, 504.8, 520.9, 1547, 1318, 0, 449.5, -736.4, -736.4, 832.1, -337.7, 0, 0, 153.7, 0, -113, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(11195, 12126, 567.6, 0, 4719.63, 5147, 62.18, 411.8, -65.88, 13031, -60.39, 277.95, 5490.33, 599.13, -113, 0, 1661, 5199, 0, 0, -137.94, -89.9, 0, 1211.3, 148.58, 1609.35, 0, 0, -1404.15, 0, -144.81), _
        Array(35, 82.35, -55.59, -219.3, 33.29, 20.93, 78.92, 67.94, 3819, 52.5, -647.2, 106.7, 92.65, 0, 0, 1661, 0, 11.32, 6815, 1809, 165.1, -7.51, 536.7, 0, 0, 0, 0, 0, 0, 0, 0))
        
        ReDim array_18to31(13, 30)
        array_18to31 = Array( _
        Array(44.27, 50.79, 193.2, 419, 68.29, -5.147, 19.9, 27.79, 589.5, 24.36, 134.9, 183.9, 227.2, 0, 0, 5199, 11.32, 0, 121.8, -12.3, 373, 0, 687.7, -11.7, 26.8, 0, 0, 0, 0, 0, 0), _
        Array(260.1, 51.82, 54.9, 0, 0, 0, 61.42, 880.2, 0, 140.7, 0, -266.6, 94.71, 0, 0, 0, 6815, 121.8, 0, 87.5, 873.6, 0, 0, 0, -151, 0, 0, 0, 0, 0, 0), _
        Array(169.5, 51.13, 0, 0, 0, 0, 1.716, -7.206, 0, 69.32, 2.745, 66.91, 0, 0, 0, 0, 1809, -12.3, 87.5, 0, 2167, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(239.5, 240.9, 287.9, 2343, 92.99, 150, 189.1, 1201, 1463, 192.5, 34.31, 300.94, 70.1, 823.55, 0, -137.94, 165.1, 373, 873.6, 2167, 0, 74.81, 95.49, 102.9, 0, 0, 0, 0, 931.3, 0, 0), _
        Array(94.24, 45.55, 0, 0, 20.92, 33.3, 153.4, -231.1, -238.8, 143.6, 0, 190.79, -25.4, 404.23, 0, -89.9, -7.51, 0, 0, 0, 74.81, 0, 259.9, 8.18, 0, 0, 28.82, 0, 0, 0, 0), _
        Array(513.4, 673.22, 750.9, 0, 378.1, 517.1, 590.5, 590.5, 590.5, 0, 0, 559.3, 222.8, 0, 0, 0, 536.7, 687.7, 0, 0, 95.49, 259.9, 0, 305.6, 0, 0, 0, 0, 0, 0, 0), _
        Array(55.48, 231.6, 634.2, 4655, 24.48, 53.1, 361.3, 0, 0, 18666, 0, 86.82, 8.77, 0, 0, 1211.3, 0, -11.7, 0, 0, 102.9, 8.18, 305.6, 0, 354.13, 7.89, 155.45, 0, 1793.97, 274.52, 0), _
        Array(201.4, -28.5, 233.7, 0, 323.59, 0, -23.7, -397.4, 0, 0, 0, 59.02, 362.71, 0, 0, 148.58, 0, 26.8, -151, 0, 0, 0, 0, 354.13, 0, 665.7, 1343, 0, 0, 0, 0), _
        Array(87.5, 200.8, 0, 0, 0, 0, 404.9, 2559.4, 0, 281.4, 988, 109.81, 4.8, 0, 0, 1609.35, 0, 0, 0, 0, 0, 0, 0, 7.89, 665.7, 0, 0, 0, 0, 362.36, 105.69), _
        Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 50.1, 0, 48.38, 100.54, 0, 0, 0, 0, 0, 0, 0, 0, 28.82, 0, 155.45, 1343, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, -95.05, 0, 0, 0, 0, 0, 0, 165.74, 0, 98.14, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 958.75, 0, 0, 0, 0, 0, 0, 0, 1011.25, 0, 0, -1404.15, 0, 0, 0, 0, 931.3, 0, 0, 1793.97, 0, 0, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 241.57, 255.99, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 274.52, 0, 362.36, 0, 0, 0, 0, 0), _
        Array(0, 0, 0, 0, 107.06, 0, 0, 0, 0, 0, 0, 14.07, 230.94, 0, 0, -144.81, 0, 0, 0, 0, 0, 0, 0, 0, 0, 105.69, 0, 0, 0, 0, 0))

        For i = 0 To 16
            For j = 0 To 30
                grpInteractionParamB(i, j) = array_1to17(i)(j) * 10 '<= Convert from MPa to bara
            Next j
        Next i
        
        For i = 0 To 13
            For j = 0 To 30
                grpInteractionParamB(i + 16, j) = array_18to31(i)(j) * 10 '<= Convert from MPa to bara
            Next j
        Next i
        
        ReDim predict_kijT(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))

        For i = 0 To dataset(0, iColumns.iSpecies)
            For j = 0 To dataset(0, iColumns.iSpecies)
                If i = j Then
                    predict_kijT(i, j) = 0
                    predict_kijT(j, i) = 0
                Else
                    If j > i Then
                        For k = 0 To 30
                            For l = 0 To 30
                                If l <> k Then
                                    If grpInteractionParamA(k, l) <> 0 And grpInteractionParamB(k, l) <> 0 Then
                                            DoubleSum(i, j) = DoubleSum(i, j) + _
                                            (decompArray(i, k) - decompArray(j, k)) * (decompArray(i, l) - decompArray(j, l)) * grpInteractionParamA(k, l) * _
                                             (298.15 / TempK) ^ ((grpInteractionParamB(k, l) / grpInteractionParamA(k, l)) - 1)
                                    End If
                                End If
                            Next l
                        Next k
                        If j > i And DoubleSum(i, j) <> 0 Then
                            predict_kijT(i, j) = -(1 / 2) * DoubleSum(i, j)
                            predict_kijT(i, j) = predict_kijT(i, j) - (aiArray(i) ^ (1 / 2) / dataset(i, iColumns.bi) - aiArray(j) ^ (1 / 2) / dataset(j, iColumns.bi)) ^ 2
                            predict_kijT(i, j) = predict_kijT(i, j) / (2 * (aiArray(i) * aiArray(j)) ^ (1 / 2) / (dataset(i, iColumns.bi) * dataset(j, iColumns.bi)))
                            predict_kijT(j, i) = predict_kijT(i, j)
                        End If
                    End If
                    
                End If
            Next j
        Next i
    End If
    
    createPredictivekijTArray = predict_kijT
    
    Exit Function
    
myErrorHandler:

        dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
        
    End Function
    
    Public Function createKij0Array(dataset As Variant, BinariesUsed As Boolean, Optional kij0 As Variant) As Double()
    
    '***************************************************************************
    'This function is called by all of the PR1978 EOS functions that require binary interaction coefficients.
    'This function creates an array of values.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim errorSum As Double
    Dim kij0_Array() As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim array_1to14() As Variant
    Dim array_15to27() As Variant
    Dim grpInteractionParamA() As Double
    Dim grpInteractionParamB() As Double
    Dim tempArray() As Variant
    Dim decompArray() As Variant
    
    ReDim grpInteractionParamA(26, 26)
    ReDim grpInteractionParamB(26, 26)
    
    fcnName = "createKij0Array"
    
    ReDim kij0_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies)) '<= Populates kij0_Array with all zeros. If BinariesUsed = False then send back all zeros.
    
    If BinariesUsed = True And IsMissing(kij0) = False And dataset(0, iColumns.predictive) = 0 Then
        If kij0.Columns.Count - 1 = dataset(0, iColumns.iSpecies) And kij0.Rows.Count - 1 = dataset(0, iColumns.iSpecies) Then
            ReDim kij0_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
            For i = 0 To dataset(0, iColumns.iSpecies)
                For j = 0 To dataset(0, iColumns.iSpecies)
                    kij0_Array(i, j) = kij0(i + 1, j + 1)
                Next j
            Next i
        Else
            myErrorMsg = "Binary array should be two dimensional array with each dimension equal to the number of species."
            GoTo myErrorHandler
        End If
    End If



    errorSum = 0

    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            errorSum = errorSum + kij0_Array(i, j) - kij0_Array(j, i)
        Next j
    Next i
    
    If errorSum <> 0 Then
        myErrorMsg = "Binaries are imbalanced. The condition Kij0(i,j) = Kij0(j,i) must be true."
        GoTo myErrorHandler
    End If
    
    createKij0Array = kij0_Array

    Exit Function

myErrorHandler:
  
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    Public Function createKijTArray(dataset As Variant, Optional kijT As Variant) As Double()
    
    '***************************************************************************
    'This function is called by all of the PR1978 EOS functions that require binary interaction coefficients.
    'This function creates an array of values.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim errorSum As Double
    Dim kijT_Array() As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    
    fcnName = "createKijTArray"
    
    ReDim kijT_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    
    If IsMissing(kijT) = False Then
        If kijT.Columns.Count - 1 = dataset(0, iColumns.iSpecies) And kijT.Rows.Count - 1 = dataset(0, iColumns.iSpecies) Then
            ReDim kijT_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
            For i = 0 To dataset(0, iColumns.iSpecies)
                For j = 0 To dataset(0, iColumns.iSpecies)
                    kijT_Array(i, j) = kijT(i + 1, j + 1)
                Next j
            Next i
        Else
            myErrorMsg = "Binary array should be two dimensional array with each dimension equal to the number of species."
            GoTo myErrorHandler
        End If
    End If


    errorSum = 0

    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            errorSum = errorSum + kijT_Array(i, j) - kijT_Array(j, i)
        Next j
    Next i
    
    If errorSum <> 0 Then
        myErrorMsg = "Binaries are imbalanced. The condition Kijt(i,j) = Kijt(j,i) must be true."
        GoTo myErrorHandler
    End If
    
    createKijTArray = kijT_Array

    Exit Function
    
myErrorHandler:
    
    ReDim kijT_Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))                  '<= Returning zeros with a warning message
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            kijT_Array(i, j) = 0
        Next j
    Next i
    
    createKijTArray = kijT_Array
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    
    Private Function create_xi_aijArray(dataset As Variant, aij_Array() As Double, molarComp() As Double) As Double()
    
    '***************************************************************************
    'The function is called by calculate_Phi function
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim yj_aijArray() As Double
    Dim i As Integer
    Dim j As Integer
    Dim Aij() As Double
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "create_xi_aijArray"
       
    ReDim Aij(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    ReDim yj_aijArray(dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        yj_aijArray(i) = 0
        For j = 0 To dataset(0, iColumns.iSpecies)
                Aij(i, j) = molarComp(j) * aij_Array(i, j)
            yj_aijArray(i) = Aij(i, j) + yj_aijArray(i)
        Next j
    Next i
        
    create_xi_aijArray = yj_aijArray

    Exit Function

myErrorHandler:

    yj_aijArray(0, 0) = -1
    create_xi_aijArray = yj_aijArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
        
    End Function

    Public Function Get_Cp_TempRange(DataRange As Range, Optional errMsgsOn As Boolean = False) As Variant
    
    
    '***************************************************************************
    'Returns the min and max heat capacity data temperatures
    'for vapor and liquid (if present)
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim myErrorMsg As String
    Dim outputArray() As Variant
    Dim Phase As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    
    Dim datasetErrMsgsOn As Boolean
    
    

    fcnName = "Get_CP_TempRange"

    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    Phase = "Vapor"
    
    For i = 1 To DataRange.Rows.Count
        If LCase(DataRange(i, 1)) = "liquid" Then
            Phase = "Liquid"
        End If
    Next i

    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, True)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    If LCase(Phase) = "vapor" Then
        ReDim outputArray(dataset(0, iColumns.iSpecies), 1)
    Else
        ReDim outputArray(2 * dataset(0, iColumns.iSpecies) + 2, 1)
    End If
    
    For i = 0 To UBound(dataset, 1)
        If dataset(i, iColumns.CpDataType) = "NIST" Or dataset(i, iColumns.CpDataType) = "HSC" Then
            outputArray(i, 0) = dataset(i, iColumns.NIST_Mn1) - 273.15
            For j = 0 To 5
                If dataset(i, iColumns.NIST_Mx6 - j * 10) <> 0 Then
                    outputArray(i, 1) = dataset(i, iColumns.NIST_Mx6 - j * 10) - 273.15
                    j = 5
                End If
            Next j
        End If
    Next i
    
    Get_Cp_TempRange = outputArray()
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(3, 1)
        For i = 0 To 3
            outputArray(i, 0) = 0
            outputArray(i, 1) = 0
        Next i
    Else
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i, 0) = 0
            outputArray(i, 1) = 0
        Next i
    End If
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function
    Private Function create_aiArray(dataset, alpha_aiArray() As Double) As Double()
    
    '***************************************************************************
    'The function is called directly of indirectly by all PR1978 functions
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim aiArray() As Double
    Dim i As Integer
    Dim outputArray() As Double
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "create_aiArray"
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = 0.457235529 * alpha_aiArray(i) * ((GasLawR * dataset(i, iColumns.tc)) ^ 2 / (dataset(i, iColumns.pc)))
    Next i
             
    create_aiArray = outputArray

    Exit Function

myErrorHandler:

    outputArray(0) = -1
    create_aiArray = outputArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
       
    End Function
    Private Function create_d2aidT2Array(dataset, TempK) As Double()
    
    '***************************************************************************
    'The function is called directly of indirectly by vaporCv, PhaseCp, Derivatives, SpeedOfSound and JTCoef
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim aiArray() As Double
    Dim i As Integer
    Dim outputArray() As Double
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "create_d2aidT2Array"
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = (0.45724 * ((GasLawR * dataset(i, iColumns.tc)) ^ 2 / (dataset(i, iColumns.pc)))) * dataset(i, iColumns.Ki)
        outputArray(i) = outputArray(i) * (dataset(i, iColumns.tc) / TempK) ^ 0.5 * (1 + dataset(i, iColumns.Ki))
        outputArray(i) = outputArray(i) / (2 * TempK * dataset(i, iColumns.tc))
    Next i
             
    create_d2aidT2Array = outputArray

    Exit Function

myErrorHandler:

    outputArray(0) = -1
    create_d2aidT2Array = outputArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
       
    End Function
    
    Private Function create_daidTArray(dataset, aiArray() As Double, TempK) As Double()
    
    '***************************************************************************
    'The function is called directly of indirectly by vaporCv, PhaseCp, Derivatives, SpeedOfSound and JTCoef
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim outputArray() As Double
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "create_daidTArray"
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = -dataset(i, iColumns.Ki) * aiArray(i) / (((1 + dataset(i, iColumns.Ki) * (1 - (TempK / dataset(i, iColumns.tc)) ^ 0.5)) * (TempK * dataset(i, iColumns.tc)) ^ 0.5))
    Next i
             
    create_daidTArray = outputArray

    Exit Function

myErrorHandler:

    outputArray(0) = -1
    create_daidTArray = outputArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
       
    End Function
    Private Function calculate_d2adT2(dataset, moleComp() As Double, TempK As Double, aiArray() As Double, daidTArray() As Double, d2aidT2Array() As Double, _
                                                    BinariesUsed As Boolean, Optional kij0Array As Variant, Optional kijTArray As Variant) As Double
    
    '***************************************************************************
    'The function is called directly of indirectly by vaporCv, PhaseCp, Derivatives, SpeedOfSound and JTCoef
    'Calculates an array of values
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim d2adT2Array() As Double
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "create_d2aidT2Array"
    
    calculate_d2adT2 = 0
    
    ReDim d2adT2Array(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
        
            d2adT2Array(i, j) = (1 / 2) * ((daidTArray(i) ^ 2 * aiArray(j) ^ 0.5 / ((aiArray(i) ^ 3) ^ 0.5)) + (daidTArray(j) ^ 2 * aiArray(i) ^ 0.5 / (aiArray(j) ^ 3) ^ 0.5))
            d2adT2Array(i, j) = d2aidT2Array(i) * aiArray(j) ^ 0.5 / aiArray(i) ^ 0.5 + d2aidT2Array(j) * aiArray(i) ^ 0.5 / aiArray(j) ^ 0.5 - d2adT2Array(i, j)
            d2adT2Array(i, j) = daidTArray(i) * daidTArray(j) / (aiArray(i) * aiArray(j)) ^ 0.5 + d2adT2Array(i, j)
            
            If BinariesUsed = True Then
                d2adT2Array(i, j) = moleComp(i) * moleComp(j) * (1 - kij0Array(i, j) + kijTArray(i, j) * TempK) * d2adT2Array(i, j)
            Else
                d2adT2Array(i, j) = moleComp(i) * moleComp(j) * d2adT2Array(i, j)
            End If
            
            calculate_d2adT2 = calculate_d2adT2 + d2adT2Array(i, j)
            
            
            
        Next j
    Next i
    
    calculate_d2adT2 = calculate_d2adT2 / 2
                        
    Exit Function

myErrorHandler:

    calculate_d2adT2 = -1
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
       
    End Function
    Private Function create_alphaiArray(dataset, TempK As Double) As Double()
    
    '***************************************************************************
    'The function is called directly of indirectly by all PR1978 functions
    'Calculates an array of values. If the dataset range contains 'Twu alpha' in cell (1,1) then
    'the Twu volume translation method is used otherwise it calculates normally
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    
    Dim alpha0 As Double
    Dim alpha1 As Double
    Dim Tr As Double
    
    Dim outputArray() As Double
    Dim LMN_Array() As Double
    
    Dim myErrorMsg As String
    
    Dim fcnName As String
    fcnName = "create_alphaiArray"
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
       
    On Error Resume Next
    
    If dataset(0, iColumns.alphaType) = 0 Then                          '<= alphaType is a variable created in the validateDataset function. Set in the (0,0) index of the dataset
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i) = (1 + dataset(i, iColumns.Ki) * (1 - (TempK / dataset(i, iColumns.tc)) ^ 0.5)) ^ 2
            If Err.Number = 6 Then
                myErrorMsg = "One of the critical temperatures equals zero."
                GoTo myErrorHandler
            Else
                If Err.Number <> 0 Then
                    myErrorMsg = Err.Description
                End If
            End If
        Next i
    Else                                                                 '<= Twu type alpha is being used - cell (1,1) of dataset contains the value 'Twu_alpha'
    
        ReDim LMN_Array(2, 3)
        LMN_Array(0, 0) = 0.272838
        LMN_Array(0, 1) = 0.625701
        LMN_Array(0, 2) = 0.373949
        LMN_Array(0, 3) = 0.0239035
        LMN_Array(1, 0) = 0.924779
        LMN_Array(1, 1) = 0.792014
        LMN_Array(1, 2) = 4.7302
        LMN_Array(1, 3) = 1.24615
        LMN_Array(2, 0) = 1.19764
        LMN_Array(2, 1) = 2.46022
        LMN_Array(2, 2) = -0.2
        LMN_Array(2, 3) = -8
                                                      
        For i = 0 To dataset(0, iColumns.iSpecies)
            Tr = TempK / dataset(i, iColumns.tc)

            If Tr <= 1 Then
                alpha0 = Tr ^ (1.19764 * (0.924779 - 1)) * Exp(0.272838 * (1 - Tr ^ (1.19764 * 0.924779)))
                alpha1 = Tr ^ (2.46022 * (0.792014 - 1)) * Exp(0.625701 * (1 - Tr ^ (2.46022 * 0.792014)))
            Else
                alpha0 = Tr ^ (1.19764 * (0.924779 - 1)) * Exp(0.272838 * (1 - Tr ^ (1.19764 * 0.924779)))
                alpha1 = Tr ^ (2.46022 * (0.792014 - 1)) * Exp(0.625701 * (1 - Tr ^ (2.46022 * 0.792014)))
            End If
            
            outputArray(i) = alpha0 + dataset(i, iColumns.omega) * (alpha1 - alpha0)
        Next i
    End If
     
    create_alphaiArray = outputArray

    Exit Function

myErrorHandler:

    outputArray(0, 0) = -1
    create_alphaiArray = outputArray

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    
    Private Function selectCpDataRanges(dataset As Variant, TempK As Double, Phase As String, moleComp() As Double) As Integer()
    
    '***************************************************************************
    'The function is called by Enthalpy and Entropy
    'This function compares tempK, 298 K and if necessary T Boil K against the 6 available heat capacity temperature ranges for each species
    'This function returns the required index values for TMN. For vapor calculation only two comparisons are required - tempK and 298 K.
    'For liquids four comparisons are required - tempK, 298 K, T boil liquid, and T boil for vapor
    'If heat capacity data is not found the a -500 flag is raised and zero indexes and passed to the calling funciton
    'If max temp heat capacity data found is below tempk the a -400 flag is raised and the highest available heat capacity data is passed to the calling function
    '***************************************************************************
    
    Dim T298_Found As Boolean
    Dim TK_Found As Boolean
    Dim TVNBP_Found As Boolean
    Dim TLNBP_Found As Boolean
    
    Dim i As Integer
    Dim j As Integer
    Dim i_Vap As Integer
    Dim g_Liq As Integer
    
    Dim CpRangeIndexes() As Integer
    Dim LastTMX_Indexes() As Integer
    Dim fcnName As String
    Dim myErrorMsg As String

    fcnName = "selectCpDataRanges"
    
    T298_Found = False
    TK_Found = False
    TVNBP_Found = False
    TLNBP_Found = False
    
    On Error GoTo myErrorHandler
    
    If LCase(Phase) = "liquid" Then
        ReDim CpRangeIndexes(dataset(0, iColumns.iSpecies), 4)                       '<= This is the result array used to store the indexes of the Cp equation with valid NIST-TMN and NIST-TMX values for input parameter TempK
    Else                                                                             '<= Index 4 stores an error flag. If a required temperature does not fall within the available Cp data temperture range then the species will be ingored.
        ReDim CpRangeIndexes(dataset(0, iColumns.iSpecies), 2)                       '<= Index 2 stores an error flag. If a required temperature does not fall within the available Cp data temperture range then the species will be ingored.
    End If
    
    ReDim LastTMX_Indexes(UBound(dataset, 1))
    
    For i = 0 To UBound(dataset, 1)
        For j = iColumns.NIST_Mx6 To iColumns.NIST_Mx1 Step -10                       '<= Create and initialize an array to hold the index of the highest valid NIST-TMXn for each species
            LastTMX_Indexes(i) = 0
        Next j
    Next i
    
    For i = 0 To UBound(dataset, 1)
        For j = iColumns.NIST_Mx6 To iColumns.NIST_Mx1 Step -10
            If LastTMX_Indexes(i) = 0 Then                                            '<= Find highest valid NIST-TMXn for each species and store it in LastTMX_Indexes()
                If IsNumeric(dataset(i, j)) = False Then
                    LastTMX_Indexes(i) = 0
                ElseIf dataset(i, j) > 0 Then
                    LastTMX_Indexes(i) = j
                End If
            End If
        Next j
    Next i
      
    If LCase(Phase) = "vapor" Then
        For i = 0 To UBound(dataset, 1)
            T298_Found = False
            TK_Found = False
            For j = iColumns.NIST_Mx1 To LastTMX_Indexes(i) Step 10
                If TK_Found = False Then
                    If TempK < dataset(i, iColumns.NIST_Mn1) Or TempK > dataset(i, LastTMX_Indexes(i)) Then       '<= Test if parameter TempK is between NIST-MN1 and the highest available max NIST temperature
                        If LastTMX_Indexes(i) <> 0 And TempK > dataset(i, iColumns.NIST_Mn1) Then
                                        CpRangeIndexes(i, 2) = -400                                               '<= Error flag: Species does not have valide Cp data - this species will be ignored
                                        CpRangeIndexes(i, iColumns.TempK) = LastTMX_Indexes(i) + 1
                                        TK_Found = True
                                    Else
                            CpRangeIndexes(i, 2) = -500                                                           '<= Error flag: Species does not have valide Cp data - this species will be ignored
                            TK_Found = True
                        End If
                    ElseIf dataset(i, iColumns.NIST_Mn1) <= TempK And TempK <= dataset(i, j) Then                 '<= Test if parameter TempK is between NIST-MN1 and the next lower available max NIST temperature
                        CpRangeIndexes(i, iColumns.TempK) = j + 1
                        TK_Found = True
                    End If
                End If
                
                If T298_Found = False Then
                    If 298.15 < dataset(i, iColumns.NIST_Mn1) Or 298.15 > dataset(i, LastTMX_Indexes(i)) Then     '<= Test if 298 k is between NIST-MN1 and the first highest max NIST temperature
                        CpRangeIndexes(i, 2) = -500                                                               '<= Error flag: Species does not have valide Cp data - this species will be ignored
                        T298_Found = True
                    ElseIf dataset(i, iColumns.NIST_Mn1) <= 298.15 And 298.15 <= dataset(i, j) Then               '<= Test if parameter TempK is between NIST-MN1 and the next lower available max NIST temperature
                        CpRangeIndexes(i, iColumns.Vap298) = j + 1
                        T298_Found = True
                    End If
                End If
                If TK_Found And T298_Found Then
                    j = LastTMX_Indexes(i)
                End If
            Next j
        Next i
    End If
    
    If LCase(Phase) = "liquid" Then
        If dataset(g_Liq, iColumns.LiquidIndex) <> (UBound(dataset, 1) + 1) / 2 Then
            myErrorMsg = "First liquid species in wrong position of dataset."
            GoTo myErrorHandler
        End If
         g_Liq = ((UBound(dataset, 1) + 1) / 2) - 1                                                                 '<= The g_Liq iterator is for the liquid species
        For i_Vap = 0 To ((UBound(dataset, 1) + 1) / 2) - 1                                                         '<= The i_Vap iterator is for the vapor species
            T298_Found = False
            TK_Found = False
            TVNBP_Found = False
            TLNBP_Found = False
            g_Liq = g_Liq + 1
            If dataset(g_Liq, iColumns.CpDataType) <> "No Data" And dataset(g_Liq, iColumns.CpDataType) <> "Not Found!" Then                                                            '<= Ingoring liquid species present at low concentrations
                If LastTMX_Indexes(g_Liq) <> 0 Then
                    For j = iColumns.NIST_Mx1 To LastTMX_Indexes(g_Liq) Step 10
                        If TK_Found = False Then
                            If TempK < dataset(g_Liq, iColumns.NIST_Mn1) Or TempK > dataset(g_Liq, LastTMX_Indexes(g_Liq)) _
                                                                            Or LastTMX_Indexes(g_Liq) = 0 Then                      '<= Test if parameter TempK is between NIST-MN1 and the highest available max NIST temperature for liquid species (g_Liq - iterator)
                                If LastTMX_Indexes(g_Liq) <> 0 And TempK > dataset(g_Liq, iColumns.NIST_Mn1) Then
                                    CpRangeIndexes(i_Vap, 4) = -400                                                                 '<= Error flag: Species does not have valide Cp data - this species will be ignored
                                    CpRangeIndexes(i_Vap, iColumns.TempK) = LastTMX_Indexes(g_Liq) + 1
                                    TK_Found = True
                                Else
                                    CpRangeIndexes(i_Vap, 4) = -500
                                    TK_Found = True
                                End If
                            ElseIf dataset(g_Liq, iColumns.NIST_Mn1) <= TempK And TempK <= dataset(g_Liq, j) Then                '<= Test if parameter TempK is between NIST-MN1 and the next available max NIST temperature for liquid species (g_Liq - iterator)
                                CpRangeIndexes(i_Vap, iColumns.TempK) = j + 1
                                TK_Found = True
                            End If
                        End If
                        
                        If TLNBP_Found = False Then
                            If dataset(i_Vap, iColumns.tb) < dataset(g_Liq, iColumns.NIST_Mn1) Or _
                                                dataset(i_Vap, iColumns.tb) > dataset(g_Liq, LastTMX_Indexes(g_Liq)) Then               '<= Test if parameter specie normal boiling point is between NIST-MN1 and the highest available max NIST temperature for liquid species (g_Liq - iterator)
                                If LastTMX_Indexes(g_Liq) <> 0 And dataset(i_Vap, iColumns.tb) > dataset(g_Liq, iColumns.NIST_Mn1) Then
                                    CpRangeIndexes(i_Vap, 4) = -400                                                                     '<= Error flag: Species does not have valide Cp data - this species will be ignored
                                    CpRangeIndexes(i_Vap, iColumns.NBPLiq) = LastTMX_Indexes(g_Liq) + 1
                                    TK_Found = True
                                Else
                                    CpRangeIndexes(i_Vap, 4) = -500
                                    TK_Found = True
                                End If                                                                                          '<= Using vapor species normal boiling point data
                            ElseIf dataset(g_Liq, iColumns.NIST_Mn1) <= dataset(i_Vap, iColumns.tb) And _
                                                                dataset(i_Vap, iColumns.tb) <= dataset(g_Liq, j) Then           '<= Test if parameter specie normal boiling point is between NIST-MN1 and the next available max NIST temperature for liquid species (g_Liq - iterator)
                                CpRangeIndexes(i_Vap, iColumns.NBPLiq) = j + 1
                                TLNBP_Found = True
                            End If
                        End If
                            
                        If T298_Found = False Then
                            If 298.15 < dataset(i_Vap, iColumns.NIST_Mn1) Or 298.15 > dataset(i_Vap, LastTMX_Indexes(i_Vap)) Then       '<= Test if 298 K is between NIST-MN1 and the highest available max NIST temperature for vapor species (i_Vap - iterator)
                                CpRangeIndexes(i_Vap, 4) = -500                                                                         '<= Error flag: Species does not have valide Cp data - this species will be ignored
                                T298_Found = True
                            ElseIf dataset(i_Vap, iColumns.NIST_Mn1) <= 298.15 And 298.15 <= dataset(i_Vap, j) Then                     '<= Test if parameter 298 K is between NIST-MN1 and the next available max NIST temperature for vapor species (i_Vap - iterator)
                                CpRangeIndexes(i_Vap, iColumns.Vap298) = j + 1
                                T298_Found = True
                            End If
                        End If
                        
                        If TVNBP_Found = False Then
                            If dataset(i_Vap, iColumns.tb) < dataset(i_Vap, iColumns.NIST_Mn1) Or _
                                                dataset(i_Vap, iColumns.tb) > dataset(i_Vap, LastTMX_Indexes(i_Vap)) Then           '<= Test if parameter species normal boiling point is between NIST-MN1 and the highest available max NIST temperature for vapor species (i_Vap - iterator)
                                CpRangeIndexes(i_Vap, 4) = -500                                                                     '<= Error flag: Species does not have valide Cp data - this species will be ignored
                                TVNBP_Found = True                                                                                  '<= Using vapor species normal boiling point data
                            ElseIf dataset(i_Vap, iColumns.NIST_Mn1) <= dataset(i_Vap, iColumns.tb) _
                                                                And dataset(i_Vap, iColumns.tb) <= dataset(i_Vap, j) Then           '<= Test if parameter species normal boiling point is between NIST-MN1 and the next available max NIST temperature for vapor species (i_Vap - iterator)
                                CpRangeIndexes(i_Vap, iColumns.NBPVap) = j + 1
                                TVNBP_Found = True
                            End If
                        End If
                        If TK_Found And T298_Found And TLNBP_Found And TVNBP_Found Then
                        j = LastTMX_Indexes(g_Liq)
                    End If
                    Next j
                End If
            Else
            CpRangeIndexes(i_Vap, 4) = -500
            End If
        Next i_Vap
    End If

    'range of TMN, TMX and Polynomual Coefficients index i ranges from 1 to 8 to match the Shomate equation coefficients A through H
    'Data for species can be added to the PData worksheet. NIST data is organized as follows:
    'Cp° = A + B * T + C * T2 + D * t3 + E / T2
    'H° - H°298.15= A*t + B*t2/2 + C*t3/3 + D*t4/4 - E/t + F - H
    'S° = A * Ln(T) + B * T + C * T2 / 2 + D * t3 / 3 - E / (2 * T2) + g
    'Cp = heat capacity (J/mol*K)
    'H° = standard enthalpy (kJ/mol)
    'S° = standard entropy (J/mol*K)
    'T = temperature(K) / 1000
    'uses TemK/1000 and calculates .
    'if NIST data is entered into the PData worksheet the column labeled '(CPDATA)' must contain 'NIST' to divide TempK by 1000
        
    selectCpDataRanges = CpRangeIndexes
    
    Exit Function

myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    
Private Sub errorSub(UDF_Range As Range, fncName As String, myErrorMsg As String, Optional errMsgsOn As Boolean = False, _
                                Optional datasetErrMsgsOn As Boolean = False)

    '***************************************************************************
    'This function manages errors written to and cleared from the cell commets fields
    '***************************************************************************

    Dim cell As Range
    Dim CommentSize As Long
    Dim commentWidth As Integer
    Dim commentLength As Integer

    If errMsgsOn Or datasetErrMsgsOn Then
        UDF_Range.ClearComments
        If Err.Number <> 0 Or myErrorMsg <> "" Then
        
            If myErrorMsg <> "" Then
                    UDF_Range.Cells(1, 1).AddComment.Text fncName & myErrorMsg      '<=Handled error
            Else
                If Err.Description <> "" Then
                    UDF_Range.Cells(1, 1).AddComment.Text fncName & Err.Description         '<=Unhandled error
                Else
                    UDF_Range.Cells(1, 1).AddComment.Text fncName & "Unhandled error."
                End If
            End If
            
            Application.ScreenUpdating = False
            For Each cell In UDF_Range
                If Len(cell.NoteText) <> 0 Then
                    commentLength = Len(cell.NoteText)
                    With cell.Comment.Shape
                        .TextFrame.AutoSize = True
                        commentWidth = .Width
                        .Width = 400                                                            '<=400 is roughly 73 characters wide
                        .Height = (Int(commentLength / 73) + 1) * 15                            '<=15 is roughly a row height
                    End With
                End If
            Next cell
            Application.ScreenUpdating = True
        Else
            UDF_Range.ClearComments
        End If
    Else
        UDF_Range.ClearComments
    End If
    Application.ScreenUpdating = True
End Sub

Private Function checkInputTemperature(temperature As Variant, Optional dataset As Variant) As Double

    '***************************************************************************
    'This function checks for user input errors
    '***************************************************************************

    Dim myErrorMsg As String
    Dim fcnName As String
    
    fcnName = "checkInputTemperature"
    
    If TypeName(temperature) = "Range" Then
        If temperature.Rows.Count <> 1 Or temperature.Columns.Count <> 1 Then
            myErrorMsg = "The Supplied temperature should be a number or a reference to a single cell."
            GoTo myErrorHandler
        End If
    End If
        
    If IsNumeric(temperature) = False Then
        myErrorMsg = "The supplied temperature is not numeric"
        GoTo myErrorHandler
    End If
    
    If temperature <= -273.15 Then
        myErrorMsg = "The supplied temperature cannot equal -273.15."
        GoTo myErrorHandler
    End If
    
    checkInputTemperature = CDbl(temperature)
    
    Exit Function

myErrorHandler:

    If IsMissing(dataset) = False Then
        dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    End If
    
    checkInputTemperature = -273.15
    
End Function

    Private Function validateDataset(DataRange As Range, Phase As String, Optional CpDataRequired As Boolean = False) As Variant()
                                                                            
    '***************************************************************************
    'This function reads the input DataRange cells and creates/returns and array of vapor or vapor and liquid species data
    'of a known format for all functions that utilise a dataset in this module. If the CpDataRequired is True
    'and if the Phase passed is liquid then both vapor and liquid species
    'will be assembled into the return array otherwise the return array will only contain
    'vapor species. Old versions of this program supported HHV and LHV calculations. That support has been dropped.
    'Backward compatibilty has been added for workbooks utilizing older versions of the PData worksheet.
    '***************************************************************************
                                                                            
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim m As Integer
    Dim strCpData As String
    Dim CPData() As Double
    Dim InputArrayBaseOne() As Variant
    Dim InputArray() As Variant
    Dim outputArray() As Variant
    Dim UpToDateColumnHeaderArray() As Variant
    Dim LegacyColumnHeaderArray() As Variant
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim LiquidSpeciesFound As Boolean
    Dim LiquidIndex As Integer
    Dim boolLegacyDataset As Boolean
    Dim UDF_Range As Range
    Dim cellRef As String
    Dim boolNewColumnHeaders As Boolean
    
    cellRef = DataRange.Address

    fcnName = "validateDataset"
    
    If TypeName(DataRange) <> "Range" Then
        myErrorMsg = "Dataset is not a range."
        GoTo myErrorHandler
    End If
    
    InputArrayBaseOne() = DataRange.Value2
    
    If DataRange.Columns.Count - 1 <> 74 Then
        myErrorMsg = "The dataset has the wrong number of columns."
    End If
    
    ReDim InputArray(DataRange.Rows.Count - 1, DataRange.Columns.Count - 1)

    For i = 1 To UBound(InputArrayBaseOne, 1)
        For j = 1 To UBound(InputArrayBaseOne, 2)
            InputArray(i - 1, j - 1) = InputArrayBaseOne(i, j)
        Next j
    Next i
        
    'Initialize
    LiquidSpeciesFound = False
    LiquidIndex = 0
    
    boolLegacyDataset = True                '<=Assue that we have old column headers then prove that they are not old
    boolNewColumnHeaders = False
    
    If DataRange.Columns.Count = 74 Then
        ReDim LegacyColumnHeaderArray(0 To DataRange.Columns.Count - 1)
        'First row of dataset should have these labels for original dataset without S298 and Gf298. Using them for error checking
        LegacyColumnHeaderArray() = Array("(MW)", "(TC, K)", "(PC, bara)", "(OMEGA)", "(ZC)", "k(i)", "b(i)", "(TB, K)", "(Hvap, kJ/kg-mole)", "(Hf, kJ/g-mole)", "(HHV, kJ/g-mole)", _
        "(LHV, kJ/g-mole)", "(CPData)", "(NIST-TMN1)", "(NIST-TMX1)", "(NIST-A1)", "(NIST-B1)", "(NIST-C1)", "(NIST-D1)", _
        "(NIST-E1)", "(NIST-F1)", "(NIST-G1)", "(NIST-H1)", "(NIST-TMN2)", "(NIST-TMX2)", "(NIST-A2)", "(NIST-B2)", "(NIST-C2)", "(NIST-D2)", _
        "(NIST-E2)", "(NIST-F2)", "(NIST-G2)", "(NIST-H2)", "(NIST-TMN3)", "(NIST-TMX3)", "(NIST-A3)", "(NIST-B3)", "(NIST-C3)", "(NIST-D3)", "(NIST-E3)", _
        "(NIST-F3)", "(NIST-G3)", "(NIST-H3)", "(NIST-TMN4)", "(NIST-TMX4)", "(NIST-A4)", "(NIST-B4)", "(NIST-C4)", "(NIST-D4)", "(NIST-E4)", "(NIST-F4)", _
        "(NIST-G4)", "(NIST-H4)", "(NIST-TMN5)", "(NIST-TMX5)", "(NIST-A5)", "(NIST-B5)", "(NIST-C5)", "(NIST-D5)", "(NIST-E5)", "(NIST-F5)", "(NIST-G5)", _
        "(NIST-H5)", "(NIST-TMN6)", "(NIST-TMX6)", "(NIST-A6)", "(NIST-B6)", "(NIST-C6)", "(NIST-D6)", "(NIST-E6)", "(NIST-F6)", "(NIST-G6)", "(NIST-H6)")
    
        
        For i = 0 To UBound(LegacyColumnHeaderArray, 1)
            If boolLegacyDataset = True Then
                If LegacyColumnHeaderArray(i) = InputArray(0, i + 1) Then
                    boolLegacyDataset = True
                Else
                    boolLegacyDataset = False
                End If
            End If
        Next i
    End If
    
                 '<= Assume that the headers are correct and prove that they are not
    
    If boolLegacyDataset = False Then
        boolNewColumnHeaders = True
        If DataRange.Columns.Count = 74 Then
            ReDim UpToDateColumnHeaderArray(0 To DataRange.Columns.Count - 1)
            'First row of dataset should have these labels after addition of S298 and Gf298. Using them for error checking.
            UpToDateColumnHeaderArray() = Array("(MW)", "(TC, K)", "(PC, bara)", "(OMEGA)", "(ZC)", "k(i)", "b(i)", "(TB, K)", "(Hvap, kJ/kg-mole)", "(Hf298, kJ/g-mole)", "(S298, J/g-mole/K)", _
            "(Gf298, kJ/g-mole)", "(CPData)", "(NIST-TMN1)", "(NIST-TMX1)", "(NIST-A1)", "(NIST-B1)", "(NIST-C1)", "(NIST-D1)", _
            "(NIST-E1)", "(NIST-F1)", "(NIST-G1)", "(NIST-H1)", "(NIST-TMN2)", "(NIST-TMX2)", "(NIST-A2)", "(NIST-B2)", "(NIST-C2)", "(NIST-D2)", _
            "(NIST-E2)", "(NIST-F2)", "(NIST-G2)", "(NIST-H2)", "(NIST-TMN3)", "(NIST-TMX3)", "(NIST-A3)", "(NIST-B3)", "(NIST-C3)", "(NIST-D3)", "(NIST-E3)", _
            "(NIST-F3)", "(NIST-G3)", "(NIST-H3)", "(NIST-TMN4)", "(NIST-TMX4)", "(NIST-A4)", "(NIST-B4)", "(NIST-C4)", "(NIST-D4)", "(NIST-E4)", "(NIST-F4)", _
            "(NIST-G4)", "(NIST-H4)", "(NIST-TMN5)", "(NIST-TMX5)", "(NIST-A5)", "(NIST-B5)", "(NIST-C5)", "(NIST-D5)", "(NIST-E5)", "(NIST-F5)", "(NIST-G5)", _
            "(NIST-H5)", "(NIST-TMN6)", "(NIST-TMX6)", "(NIST-A6)", "(NIST-B6)", "(NIST-C6)", "(NIST-D6)", "(NIST-E6)", "(NIST-F6)", "(NIST-G6)", "(NIST-H6)")
            
            For i = 0 To UBound(UpToDateColumnHeaderArray, 1)
                If UpToDateColumnHeaderArray(i) <> InputArray(0, i + 1) Then
                    boolNewColumnHeaders = False
                End If
            Next i
        End If
    End If
    
    If boolNewColumnHeaders = False And boolLegacyDataset = False Then
        myErrorMsg = "The dataset has the wrong number of columns."
        GoTo myErrorHandler
    End If
    
    If boolLegacyDataset = True Then
        For i = 0 To UBound(LegacyColumnHeaderArray, 1)
            If LegacyColumnHeaderArray(i) <> InputArray(0, i + 1) Then
                ReDim outputArray(0, 0)
                outputArray(0, 0) = -1
                myErrorMsg = "Something is wrong with the dataset. Column lables are incorrect."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    If boolNewColumnHeaders = True Then
        For i = 0 To UBound(UpToDateColumnHeaderArray, 1)
            If UpToDateColumnHeaderArray(i) <> InputArray(0, i + 1) Then
                ReDim outputArray(0, 0)
                outputArray(0, 0) = -1
                myErrorMsg = "Something is wrong with the dataset. Column lables are incorrect."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    For i = 0 To UBound(InputArray, 1)
        
        If LCase(InputArray(i, 0)) = "liquid" Then                          '<= look for the key work "Liquid" or "liquid" to detect liquid species in dataset
            If (UBound(InputArray, 1) - 1) / (i - 1) = 2 Then               '<= make sure liquid dey word is in correct loaction to confirm correct structure of dataset
                LiquidSpeciesFound = True
                LiquidIndex = i - 1
            Else
                myErrorMsg = "Liquid species found in wrong position of dataset."
                GoTo myErrorHandler
            End If
        End If
    Next i
    

    
    If LCase(Phase) = "vapor" Or CpDataRequired = False Then
        If LiquidSpeciesFound = False Then
            ReDim outputArray(UBound(InputArray, 1) - 1, iColumns.FinalIndex)               '<= Need all species present in dataset in case that no liquid species exist
        Else
            ReDim outputArray(((UBound(InputArray, 1) - 1) / 2) - 1, iColumns.FinalIndex)   '<= Need only vapor species in dataset. Do not need the liquid species.
        End If
    End If
    
    If LCase(Phase) = "liquid" And CpDataRequired = True Then
        If LiquidSpeciesFound = True Then
            ReDim outputArray(UBound(InputArray, 1) - 2, iColumns.FinalIndex)               '<=Need liquid and vapor species.
        Else
            myErrorMsg = "Phase is liquid but liquid keyword in dataset not found."         '<=Need liquid and vapor species. If not present need to bail.
            GoTo myErrorHandler
        End If
    End If
    
    m = 0
    For i = 0 To UBound(InputArray, 1)
        If LCase(InputArray(i, 0)) <> "liquid" And i <> 0 Then
            For j = 0 To iColumns.iS298 - 1
                outputArray(m, j) = InputArray(i, j + 1)
            Next j                                                                           '<= Strip header row and species names by advancing inputarray indexes
            
            If boolLegacyDataset = True Then
                outputArray(m, iColumns.iS298) = 0
                outputArray(m, iColumns.iGf298) = 0
            Else
                outputArray(m, iColumns.iS298) = InputArray(i, iColumns.iS298)
                outputArray(m, iColumns.iGf298) = InputArray(i, iColumns.iGf298)
            End If
            
            For j = iColumns.iGf298 + 1 To iColumns.lastCpIndex
                If boolLegacyDataset = True Then
                    outputArray(m, j) = InputArray(i, j + 1)
                Else
                    outputArray(m, j) = InputArray(i, j + 1)
                End If
            Next j
            
            If m = UBound(outputArray, 1) Then
                i = UBound(InputArray, 1)
            Else
                m = m + 1
            End If
            
        End If
    Next i
    
    If InStr(1, CStr(InputArray(0, 0)), "Twu alpha", vbTextCompare) <> 0 Then               '<=vbTextCompare eliminates the need to check the string for capital letters or needing to use lcase fcn
        outputArray(0, iColumns.alphaType) = 1                                              '<= 1 forces create_alphaiArray() function to use the Twu Alpha method
    Else
        outputArray(0, iColumns.alphaType) = 0                                              '<= 0 the allows create_alphaiArray() function to calculate normally
    End If
    
    If InStr(1, CStr(InputArray(0, 0)), "predictive", vbTextCompare) <> 0 Then               '<=vbTextCompare eliminates the need to check the string for capital letters or needing to use lcase fcn
        outputArray(0, iColumns.predictive) = 1                                              '<= 1 forces create_alphaiArray() function to use the Twu Alpha method
    Else
        outputArray(0, iColumns.predictive) = 0                                              '<= 0 the allows create_alphaiArray() function to calculate normally
    End If
    
    If InStr(1, CStr(InputArray(0, 0)), "Error Messages On", vbTextCompare) <> 0 Then       '<=global dataset level error messages flag on
        outputArray(0, iColumns.errMsgsOn) = 1
    Else
        outputArray(0, iColumns.errMsgsOn) = 0                                              '<=global dataset level error messages flag off
    End If
    
    If LiquidSpeciesFound = True And LiquidIndex <> 0 Then                                  '<=store this data for use outside of this procedure
        outputArray(0, iColumns.iSpecies) = LiquidIndex - 1
        outputArray(0, iColumns.LiquidIndex) = LiquidIndex
        outputArray(0, iColumns.LiquidsFound) = True
    Else
        outputArray(0, iColumns.iSpecies) = UBound(outputArray, 1)
        outputArray(0, iColumns.LiquidIndex) = 0
        outputArray(0, iColumns.LiquidsFound) = False
    End If
    
    outputArray(0, iColumns.globalErrmsg) = ""
    
    
        
    validateDataset = outputArray
   
Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = fcnName & ": " & Err.Description
    Else
        
    End If
    
    ReDim outputArray(0, 0)
    outputArray(0, 0) = myErrorMsg
    validateDataset = outputArray
    
End Function
    Private Function validateMoles(dataset As Variant, moles As Variant) As Double()
    
    '***************************************************************************
    'This function checks for user input errors and returns
    'a two dimersional array of mole amounts and mole composition
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim TotalMoles As Double
    Dim InputArrayBaseOne() As Variant
    Dim InputArray() As Variant
    Dim outputArray() As Double
    Dim myErrorMsg As String
    Dim MoleFractionSum As Double
    Dim inputRange As Range
    Dim fcnName As String

    fcnName = "validateMoles"
    
    If TypeName(moles) = "Double" Then
        If moles <> 1 Then
            myErrorMsg = "The moles parameter must be either a range of values or the number 1 indicating a pure component dataset containing a single component with or without a liquid phase."
            GoTo myErrorHandler
        Else
            If dataset(0, iColumns.iSpecies) = 0 Then   '<= For the case a single species dataset allowing for mole composition = 1
                ReDim outputArray(0)
                outputArray(0) = 1#
                validateMoles = outputArray
                Exit Function
            End If
        End If
    End If
    
    
    '=================== Validate Moles and calculate mole fractions
    
        If moles.Count - 1 <> dataset(0, iColumns.iSpecies) Then                                      '< cannot continue if selected number of mole amount do not equal the number of species
            myErrorMsg = "Number of species does not equal the number of mole amounts!"
            GoTo myErrorHandler
        End If
            
        If moles Is Nothing = True Then
            myErrorMsg = "No range of cells selected for moles"
            GoTo myErrorHandler
        Else
            If moles.Columns.Count > 1 Then
                myErrorMsg = "Moles range must only contain one column."
                GoTo myErrorHandler
            Else
                Set inputRange = moles                                                     '<= Minimize references to shpreadsheet by working with VBA range object
                If inputRange(1).Value2 < 0 Then
                    myErrorMsg = "Mole amount equals zero."
                    GoTo myErrorHandler
                Else
                    If inputRange.Rows.Count = 1 And inputRange(1).Value2 = 1 Then
                        ReDim outputArray(0)
                        outputArray(0) = 1#
                        validateMoles = outputArray
                        Exit Function
                    Else
                        ReDim InputArrayBaseOne(1 To moles.Rows.Count)
                        For i = 1 To inputRange.Rows.Count
                            InputArrayBaseOne(i) = moles(i).Value2                                       '<=value2 preserves the most precision
                        Next i
                    End If
                End If
                
                ReDim InputArray(moles.Rows.Count - 1, iColumns.moles)
                For i = 0 To UBound(InputArray, 1)
                    InputArray(i, iColumns.moles) = InputArrayBaseOne(i + 1)          '<= switch from base one array to base zero array
                Next i
                
                TotalMoles = 0
                
                For i = 0 To UBound(InputArray, 1)
                    If InputArray(i, iColumns.moles) >= 0 And IsNumeric(InputArray(i, iColumns.moles)) = True Then
                        TotalMoles = TotalMoles + InputArray(i, iColumns.moles)
                    Else
                        myErrorMsg = "Some moles amounts are below zero or not numeric!"
                        GoTo myErrorHandler
                    End If
                Next i
                
                MoleFractionSum = 0
                If TotalMoles > 0 Then
                    For i = 0 To UBound(InputArray, 1)
                        InputArray(i, iColumns.MoleFraction) = InputArray(i, iColumns.moles) / TotalMoles
                        MoleFractionSum = InputArray(i, iColumns.MoleFraction) + MoleFractionSum
                    Next i
                Else
                    myErrorMsg = "No moles amounts or all amounts are zero."
                    GoTo myErrorHandler
                End If
                
                If MoleFractionSum > 1.00000000001 Or 0.99999999999 > MoleFractionSum Then
                    myErrorMsg = "Mole fractions do not add up to zero."
                    GoTo myErrorHandler
                End If
                    
                ReDim outputArray(dataset(0, iColumns.iSpecies))
            
                For i = 0 To dataset(0, iColumns.iSpecies)
                    If TotalMoles >= 0 Then
                        outputArray(i) = InputArray(i, iColumns.MoleFraction)                           '<index 2 stores the mole fractions, index 1 stores the moles
                    Else
                        myErrorMsg = "validateDatasetMole amount data is bad!"
                        GoTo myErrorHandler
                    End If
                Next i
            End If
        End If

    validateMoles = outputArray
        
        Exit Function
        
myErrorHandler:
    
    ReDim outputArray(1)
    outputArray(0) = -1
    
    If myErrorMsg = "" Then
        myErrorMsg = "validateMoles - " & Err.Description
    End If
    
    validateMoles = outputArray
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    End Function
    
    Private Function getPdataWorksheetIndex() As Integer
    
    '***************************************************************************
    'This function finds the index to the PData worksheet
    '***************************************************************************
    
    Dim i As Integer
    Dim j As Integer
    Dim testString As String
    Dim PDataExists As Boolean
    Dim sourceSheet As Worksheet
    Dim myErrorMsg As String
    Dim fcnName As String

    fcnName = "getPdataWorksheetIndex"
    
        On Error Resume Next

            For i = 1 To ThisWorkbook.Worksheets.Count                  'Physical property worksheet must have 'PDATA' in cell(a1,a1)
            Err.Clear
            testString = CStr(ThisWorkbook.Sheets(i).Cells(1, 1))
                If LCase(ThisWorkbook.Sheets(i).Cells(1, 1)) = "pdata" Then
                    If Err.Number = 0 Then
                        Set sourceSheet = ThisWorkbook.Sheets(i)
                        PDataExists = True
                        j = i
                        i = ThisWorkbook.Worksheets.Count       '<= Escape loop as soon as PData is found
                    End If
                End If
            Next i
            
            If PDataExists = True Then
                getPdataWorksheetIndex = j
            End If
            
        Exit Function
        
myErrorHandler:
    Debug.Print "getPdataWorksheetIndex error"


    End Function
    Public Function ValidateWorkbook(Optional errMsgsOn As Boolean = False) As Boolean
    
    '***************************************************************************
    'This function tests the current workbook for the correct configuration. It
    'looks for the text 'pdata' in cell(1,1) of each worksheet and looks for the
    'PData_Properties and PData_PropertyNames named ranges.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim PropertyNamesExist As Boolean
    Dim PropertiesExist As Boolean
    Dim PDataExists As Boolean

    Dim i As Integer
    Dim sourceSheet As Worksheet
    Dim myErrorMsg As String
    Dim testString As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim PDataIndex As Integer

    fcnName = "ValidateWorkbook"
    
    Set UDF_Range = Application.Caller
    
    On Error Resume Next

            For i = 1 To ThisWorkbook.Worksheets.Count                  'Physical property worksheet must have 'PDATA' in cell(a1,a1)
            Err.Clear
            testString = CStr(ThisWorkbook.Sheets(i).Cells(1, 1))
                If LCase(ThisWorkbook.Sheets(i).Cells(1, 1)) = "pdata" Then
                    If Err.Number = 0 Then
                        Set sourceSheet = ThisWorkbook.Sheets(i)
                        PDataExists = True
                        PDataIndex = i
                        i = ThisWorkbook.Worksheets.Count       '<= Escape loop as soon as PData is found
                    End If
                End If
            Next i
            
        For i = 1 To ThisWorkbook.Names.Count
        Err.Clear
            If ThisWorkbook.Names(i).Name = "PData_PropertyNames" Then
                PropertyNamesExist = True
            End If
            
            If ThisWorkbook.Names(i).Name = "PData_Properties" Then
                PropertiesExist = True
            End If
        Next i
            
        If PDataExists And PropertyNamesExist And PropertiesExist = False Then

            If PDataExists = False Then
                myErrorMsg = "None of the worksheets in this workbook contain 'PData' in Cell A1." & myErrorMsg
            End If
            
            If PropertyNamesExist = False Then
                myErrorMsg = "The named range 'PData_PropertyNames' does not exist" & myErrorMsg
            End If
            
            If PropertiesExist = False Then
                myErrorMsg = "The named range 'PData_Properties' does not exist" & myErrorMsg
            End If
            
        End If
        
        If myErrorMsg <> "" Then
            GoTo myErrorHandler
        End If
        
        ValidateWorkbook = True
        
        Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
        
        Application.ScreenUpdating = True
        
        Exit Function
        
        
myErrorHandler:
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn)
    
    ValidateWorkbook = False
    
    Application.ScreenUpdating = True

    End Function


    Public Function MessageCount(Optional cmmType As String = "") As Integer
    
    '***************************************************************************
    'This looks for all comments on the worksheet that contain the test value of 'warning' or 'error'
    'and returns totals for either or both types of comments.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim cmm As Comment
    Dim Count As Integer
    Dim UDF_Range As Range
    
    Count = 0
    
    Set UDF_Range = Application.Caller
    
    Set ws = UDF_Range.Parent
    
    For Each cmm In ws.Comments
    
        Select Case cmmType
           Case Is = ""
           If InStr(1, cmm.Text, "error:", vbTextCompare) <> 0 Or InStr(1, cmm.Text, "warning:", vbTextCompare) <> 0 Then
               Count = Count + 1
           End If
           
           Case Is = "Warnings"
           If InStr(1, cmm.Text, "warning:", vbTextCompare) <> 0 Then
               Count = Count + 1
           End If
           
           Case Is = "Errors"
           If InStr(1, cmm.Text, "error:", vbTextCompare) <> 0 Then
               Count = Count + 1
           End If
           
           Case Else
           If InStr(1, cmm.Text, "error:", vbTextCompare) <> 0 Or InStr(1, cmm.Text, "warning:", vbTextCompare) <> 0 Then
               Count = Count + 1
           End If
        End Select
        
    Next cmm

    MessageCount = Count
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    Application.ScreenUpdating = True

End Function
    Public Function CreateDecomposition(Species As Range, Optional errMsgsOn As Boolean = False) As Variant()
    
    '***************************************************************************
    'This function lookups up values in the PData worksheet and returns an array of species data in a defined formate
    'expected by the functions contained in this module.
    'Once the range of values are returned the range should be copied and pasted in place as values only.
    'This prevents the function from constantly re-calculating.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = True

    Dim LiquidPhaseExists As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim m As Integer
    Dim int_Species As Integer
    Dim int_Group1 As Integer
    Dim int_TC As Integer
    Dim int_PC As Integer
    Dim int_TMN1 As Integer
    Dim int_Omega As Integer
    Dim int_ZC As Integer
    Dim int_TB As Integer
    Dim int_HVap As Integer
    Dim int_Hf298 As Integer
    Dim int_Gf298 As Integer
    Dim int_S298 As Integer
    Dim int_HHV As Integer
    Dim int_LHV As Integer
    Dim int_CpDataType As Integer
    Dim intSum As Integer
    Dim NumberOfDatasetRows As Integer
    Dim outputArray() As Variant
    Dim sourceSheet As Worksheet
    Dim myErrorMsg As String
    Dim cellRef As String
    Dim LiquidSpeciesFound As Boolean
    Dim LiquidIndex As Integer
    Dim selectedRange As Range
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim Phase As String
    Dim SfNotFound As Boolean
    Dim GfNotFound As Boolean
    
    fcnName = "CreateDecomposition"

    myErrorMsg = ""
    LiquidIndex = 0
    
    Set UDF_Range = Application.Caller
    
    Set sourceSheet = ThisWorkbook.Sheets(getPdataWorksheetIndex)
    
    cellRef = Application.Caller.Address
    
    Set selectedRange = Range(cellRef)
    
    If selectedRange.Cells.Count / selectedRange.Rows.Count <> 31 Then
        myErrorMsg = "The selected range required to create a dataset is (number of species + 1 rows) by (1 + 31 + 1 columns)."
        GoTo myErrorHandler
    End If
    
    If ValidateWorkbook <> True Then
        myErrorMsg = "Pdata worksheet, or Cell A1:A1 in PData worksheet does not contain 'PData' or Pdata named ranges 'PData_Properties' or 'PData_PropertyNames' do not exist"
        GoTo myErrorHandler
    End If

    On Error Resume Next
    
    int_Species = Application.WorksheetFunction.Match("(SPECIES)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(SPECIES)' not found in PData worksheet"
        GoTo myErrorHandler
        Exit Function
    End If
    
    For i = 1 To 31
        If i = 1 Then
            int_Group1 = Application.WorksheetFunction.Match("(GROUP " & i & ")", sourceSheet.Range("PData_PropertyNames"), 0)
            If Err.Number = 1004 Then
                myErrorMsg = "Column with label " & "(GROUP " & i & ") " & "not found in PData worksheet"
                GoTo myErrorHandler
                Exit Function
            End If
        End If
        
        If "(Group " & i & ")" = Application.WorksheetFunction.Match("(GROUP " & i & ")", sourceSheet.Range("PData_PropertyNames"), 0) Then
        End If

        If Err.Number = 1004 Then
            myErrorMsg = "Column with label '(MW)' not found in PData worksheet"
            GoTo myErrorHandler
            Exit Function
        End If
    Next i
    

    
    On Error GoTo myErrorHandler
    
    If Species Is Nothing = True Or Species.Rows.Count = 1 Then                                     '<= If Range parameter is missing or only contains one row then no species are present so exit
        myErrorMsg = "No range of species selected."
        GoTo myErrorHandler
    Else
        ReDim speciesNames(0 To Species.Rows.Count - 1)                                             '<= If range contains 3 rows or more then at least one species is present so continue
        For i = 0 To UBound(speciesNames)
            speciesNames(i) = Species(i + 1)
        Next i
    End If
    
    Err.Clear
    On Error Resume Next
    
    For i = 0 To UBound(speciesNames)
         If IsNumeric(speciesNames(i)) = True Then                                                 '<= Species names cannot start with a number so flag species but continue eventhough this is probably an error
             speciesNames(i) = "Not Found!"
         End If
         
         If IsNull(speciesNames(i)) = True Then                                                    '<= Species names cannot be null
             speciesNames(i) = "Not Found!"
         End If
         
        If LCase(speciesNames(i)) <> "liquid" And i <> 0 Then
             If speciesNames(i) <> "Not Found!" Then
                 If speciesNames(i) <> Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Species, 0) Then
                     If LCase(speciesNames(i)) = "liquid" Or LCase(speciesNames(i)) = "vapor" Then
                        If LCase(speciesNames(i)) = "liquid" Then
                            LiquidIndex = i
                        End If
                     Else
                        If i < LiquidIndex Then
                            myErrorMsg = "Some vapor species are not found in the PData worksheet '(SPECIES)' column"
                            GoTo myErrorHandler
                        End If
                            speciesNames(i) = "Not Found!"
                     End If
                 End If
             End If
         Else
            If LCase(speciesNames(i)) = "liquid" Then
                LiquidIndex = i
                LiquidPhaseExists = True
            End If
         End If
    Next i
    
   On Error GoTo myErrorHandler
   
    If LiquidIndex = 0 Then
        For i = 1 To UBound(speciesNames, 1)
            If speciesNames(i) = "Not Found!" Then
                myErrorMsg = "Not all vapor phase species are found in PData worksheet."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    If LiquidPhaseExists = True Then
        For i = 1 To LiquidIndex
            If speciesNames(i) = "Not Found!" Then
                myErrorMsg = "Not all vapor phase species are found in PData worksheet."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    
    NumberOfDatasetRows = UBound(speciesNames)

    ReDim outputArray(NumberOfDatasetRows, 31)
    
    For i = 0 To 30
        outputArray(0, i) = Application.WorksheetFunction.VLookup("(SPECIES)", sourceSheet.Range("PData_Properties"), int_Group1 + i, 0)
        outputArray(0, 31) = "(Total Groups)"
    Next i
    
    For i = 1 To UBound(speciesNames)
        intSum = 0
        For j = 0 To 30
            outputArray(i, j) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Group1 + j, 0)
            If IsEmpty(outputArray(i, j)) = False Then
                intSum = intSum + outputArray(i, j)
            Else
                outputArray(i, j) = 0
            End If

        Next j
        
        If intSum <> 0 Then
            For m = 0 To 26
                outputArray(i, m) = outputArray(i, m) / intSum
            Next m
        Else
            outputArray(i, m) = 0
        End If

    Next i
    
    CreateDecomposition = outputArray
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)
    
    Application.ScreenUpdating = True
    
    Exit Function
  
myErrorHandler:

    ReDim outputArray(NumberOfDatasetRows, iColumns.lastCpIndex)
    For i = 0 To NumberOfDatasetRows
        For j = 0 To iColumns.lastCpIndex
            outputArray(i, j) = 0
        Next j
    Next i
    
    CreateDecomposition = outputArray
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, True)
    
    Application.ScreenUpdating = True
    
    End Function
    
    Public Function CreateDataset(Species As Range, Optional errMsgsOn As Boolean = False) As Variant()
    
    '***************************************************************************
    'This function lookups up values in the PData worksheet and returns an array of species data in a defined formate
    'expected by the functions contained in this module.
    'Once the range of values are returned the range should be copied and pasted in place as values only.
    'This prevents the function from constantly re-calculating.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = True

    Dim LiquidPhaseExists As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim m As Integer
    Dim int_Species As Integer
    Dim int_MW As Integer
    Dim int_TC As Integer
    Dim int_PC As Integer
    Dim int_TMN1 As Integer
    Dim int_Omega As Integer
    Dim int_ZC As Integer
    Dim int_TB As Integer
    Dim int_HVap As Integer
    Dim int_Hf298 As Integer
    Dim int_Gf298 As Integer
    Dim int_S298 As Integer
    Dim int_HHV As Integer
    Dim int_LHV As Integer
    Dim int_CpDataType As Integer
    Dim NumberOfDatasetRows As Integer
    Dim outputArray() As Variant
    Dim sourceSheet As Worksheet
    Dim myErrorMsg As String
    Dim cellRef As String
    Dim LiquidSpeciesFound As Boolean
    Dim LiquidIndex As Integer
    Dim selectedRange As Range
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim Phase As String
    Dim SfNotFound As Boolean
    Dim GfNotFound As Boolean
    
    fcnName = "CreateDataset"

    myErrorMsg = ""
    LiquidIndex = 0
    
    Set UDF_Range = Application.Caller
    
    Set sourceSheet = ThisWorkbook.Sheets(getPdataWorksheetIndex)
    
    cellRef = Application.Caller.Address
    
    Set selectedRange = Range(cellRef)
    
    If selectedRange.Cells.Count / selectedRange.Rows.Count <> 73 Then
        myErrorMsg = "The selected range required to create a dataset is (number of species + 1 rows) by (73 columns)."
        GoTo myErrorHandler
    End If
    
    If ValidateWorkbook <> True Then
        myErrorMsg = "Pdata worksheet, or Cell A1:A1 in PData worksheet does not contain 'PData' or Pdata named ranges 'PData_Properties' or 'PData_PropertyNames' do not exist"
        GoTo myErrorHandler
    End If

    On Error Resume Next
    
    int_Species = Application.WorksheetFunction.Match("(SPECIES)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(SPECIES)' not found in PData worksheet"
        GoTo myErrorHandler
        Exit Function
    End If
    
    int_MW = Application.WorksheetFunction.Match("(MW)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(MW)' not found in PData worksheet"
        GoTo myErrorHandler
        Exit Function
    End If
    
    int_TC = Application.WorksheetFunction.Match("(TC, K)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(TC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_PC = Application.WorksheetFunction.Match("(PC, bara)", sourceSheet.Range("PData_PropertyNames"), 0)
        If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(PC)' not found in PData worksheet"
        Exit Function
    End If

    int_Omega = Application.WorksheetFunction.Match("(OMEGA)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(OMEGA)' not found in PData worksheet"
        GoTo myErrorHandler
    End If

    int_ZC = Application.WorksheetFunction.Match("(ZC)", sourceSheet.Range("PData_PropertyNames"), 0)
        If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(ZC)' not found in PData worksheet"
        GoTo myErrorHandler
    End If
    
    int_TB = Application.WorksheetFunction.Match("(TB, K)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(TB)' not found in PData worksheet"
        GoTo myErrorHandler
    End If
    
    int_HVap = Application.WorksheetFunction.Match("(Hvap, kJ/kg-mole)", sourceSheet.Range("PData_PropertyNames"), 0)
    If Err.Number = 1004 Then
        myErrorMsg = "Column with label '(Hvap, kJ/kg-mole)' not found in PData worksheet"
        GoTo myErrorHandler
    End If
    
    int_Hf298 = Application.WorksheetFunction.Match("(Hf298, kJ/g-mole)", sourceSheet.Range("PData_PropertyNames"), 0)

    If Err.Number = 1004 Then
        Err.Clear
        int_Hf298 = Application.WorksheetFunction.Match("(Hf, kJ/g-mole)", sourceSheet.Range("PData_PropertyNames"), 0)
    End If
    
    If Err.Number = 1004 Then
            myErrorMsg = "validateDataset: (Hf298, kJ/g-mole) column is not found in PData worksheet"
            GoTo myErrorHandler
    End If
    
    int_S298 = Application.WorksheetFunction.Match("(S298, J/g-mole/K)", sourceSheet.Range("PData_PropertyNames"), 0)

    If Err.Number = 1004 Then
        myErrorMsg = "validateDataset: Old PData worksheet being used."
        SfNotFound = True
        Err.Clear
    End If
    
    int_Gf298 = Application.WorksheetFunction.Match("(Gf298, kJ/g-mole)", sourceSheet.Range("PData_PropertyNames"), 0)

    If Err.Number = 1004 Then
        myErrorMsg = "validateDataset: Old PData worksheet being used."
        GfNotFound = True
        Err.Clear
    End If
    
    int_TMN1 = Application.WorksheetFunction.Match("(NIST-TMN1)", sourceSheet.Range("PData_PropertyNames"), 0)
    
    If Err.Number = 1004 Then
        Debug.Print "(NIST-TMN1) column is not found in PData worksheet"
        GoTo myErrorHandler
    End If
        
    int_CpDataType = Application.WorksheetFunction.Match("(CPData)", sourceSheet.Range("PData_PropertyNames"), 0)
    
    If Err.Number = 1004 Then
        Debug.Print "(CPData) column is not found in PData worksheet"
        GoTo myErrorHandler
    End If
    
    On Error GoTo myErrorHandler
    
    If Species Is Nothing = True Or Species.Rows.Count = 1 Then                                     '<= If Range parameter is missing or only contains one row then no species are present so exit
        myErrorMsg = "No range of species selected."
        GoTo myErrorHandler
    Else
        ReDim speciesNames(0 To Species.Rows.Count - 1)                                             '<= If range contains 3 rows or more then at least one species is present so continue
        For i = 0 To UBound(speciesNames)
            speciesNames(i) = Species(i + 1)
        Next i
    End If
    
    Err.Clear
    On Error Resume Next
    
    For i = 0 To UBound(speciesNames)
         If IsNumeric(speciesNames(i)) = True Then                                                  '<= Species names cannot start with a number so flag species but continue eventhough this is probably an error
             speciesNames(i) = "Not Found!"
         End If
         
         If IsNull(speciesNames(i)) = True Then                                                     '<= Species names cannot be null
             speciesNames(i) = "Not Found!"
         End If
         
        If LCase(speciesNames(i)) <> "liquid" And i <> 0 Then
             If speciesNames(i) <> "Not Found!" Then
                 If speciesNames(i) <> Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Species, 0) Then
                     If LCase(speciesNames(i)) = "liquid" Or LCase(speciesNames(i)) = "vapor" Then
                        If LCase(speciesNames(i)) = "liquid" Then
                            LiquidIndex = i
                        End If
                     Else
                        If i < LiquidIndex Then
                            myErrorMsg = "Some vapor species are not found in the PData worksheet '(SPECIES)' column"
                            GoTo myErrorHandler
                        End If
                            speciesNames(i) = "Not Found!"
                     End If
                 End If
             End If
         Else
            If LCase(speciesNames(i)) = "liquid" Then
                LiquidIndex = i
                LiquidPhaseExists = True
            End If
         End If
    Next i
    
   On Error GoTo myErrorHandler
   
    If LiquidIndex = 0 Then
        For i = 1 To UBound(speciesNames, 1)
            If speciesNames(i) = "Not Found!" Then
                myErrorMsg = "Not all vapor phase species are found in PData worksheet."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    If LiquidPhaseExists = True Then
        For i = 1 To LiquidIndex
            If speciesNames(i) = "Not Found!" Then
                myErrorMsg = "Not all vapor phase species are found in PData worksheet."
                GoTo myErrorHandler
            End If
        Next i
    End If
    
    
    NumberOfDatasetRows = UBound(speciesNames)

    ReDim outputArray(NumberOfDatasetRows, iColumns.lastCpIndex)
    
    outputArray(0, iColumns.MW) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_MW, 0)
    outputArray(0, iColumns.tc) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_TC, 0)
    outputArray(0, iColumns.pc) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_PC, 0)
    outputArray(0, iColumns.omega) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_Omega, 0)
    outputArray(0, iColumns.zc) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_ZC, 0)
    outputArray(0, iColumns.Ki) = "k(i)"
    outputArray(0, iColumns.bi) = "b(i)"
    outputArray(0, iColumns.tb) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_TB, 0)
    outputArray(0, iColumns.hvap) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_HVap, 0)
    outputArray(0, iColumns.iHf298) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_Hf298, 0)
    
    If SfNotFound = False Then
        outputArray(0, iColumns.iS298) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_S298, 0)
    Else
        outputArray(0, iColumns.iS298) = "(HHV, kJ/g-mole)"
    End If
    
    If GfNotFound = False Then
        outputArray(0, iColumns.iGf298) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_Gf298, 0)
    Else
        outputArray(0, iColumns.iGf298) = "(LHV, kJ/g-mole)"
    End If
        
    outputArray(0, iColumns.CpDataType) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_CpDataType, 0)
    outputArray(0, iColumns.NIST_Mn1) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_TMN1, 0)
    
    m = 0
    For j = iColumns.NIST_Mx1 To iColumns.lastCpIndex
        outputArray(0, j) = Application.WorksheetFunction.VLookup("(Species)", sourceSheet.Range("PData_Properties"), int_TMN1 + m + 1, 0)
        m = m + 1
    Next j

    For i = 1 To NumberOfDatasetRows
        If speciesNames(i) = "Liquid" Or speciesNames(i) = "liquid" Then
            For j = 0 To iColumns.lastCpIndex
                outputArray(i, j) = ""
            Next j
            outputArray(i, iColumns.CpDataType) = "Liquid"
        End If
          
        If speciesNames(i) = "Not Found!" Then
            For j = 0 To iColumns.lastCpIndex
                outputArray(i, j) = 0
            Next j
            outputArray(i, iColumns.CpDataType) = "No Data"
        End If
        
        If speciesNames(i) = "Not Found!" Or LCase(speciesNames(i)) = "liquid" Or LCase(speciesNames(i)) = "vapor" Then
            If LCase(speciesNames(i)) = "liquid" Then
                outputArray(i, iColumns.CpDataType) = "Liquid"
            Else
                outputArray(i, iColumns.CpDataType) = "Not Found!"
            End If
        Else
                outputArray(i, iColumns.MW) = CDbl(Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_MW, 0))
            If i < LiquidIndex Or LiquidIndex = 0 Then
                outputArray(i, iColumns.tc) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_TC, 0)
                
                If myErrorMsg = "validateDataset: Old PData worksheet being used." = True Then
                    outputArray(i, iColumns.pc) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_PC, 0) * 1.01324999970844
                Else
                    outputArray(i, iColumns.pc) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_PC, 0)
                End If
                
                outputArray(i, iColumns.omega) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Omega, 0)
                outputArray(i, iColumns.zc) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_ZC, 0)
                outputArray(i, iColumns.tb) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_TB, 0)
                outputArray(i, iColumns.hvap) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_HVap, 0)
                outputArray(i, iColumns.iHf298) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Hf298, 0)
                If SfNotFound = False Then
                    outputArray(i, iColumns.iS298) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_S298, 0)
                Else
                    outputArray(i, iColumns.iS298) = 0
                End If
                If GfNotFound = False Then
                    outputArray(i, iColumns.iGf298) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_Gf298, 0)
                Else
                    outputArray(i, iColumns.iGf298) = 0
                End If

            End If
                outputArray(i, iColumns.CpDataType) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_CpDataType, 0)
                outputArray(i, iColumns.NIST_Mn1) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_TMN1, 0)
    
                m = 0
                For j = iColumns.NIST_Mx1 To iColumns.lastCpIndex
                    outputArray(i, j) = Application.WorksheetFunction.VLookup(speciesNames(i), sourceSheet.Range("PData_Properties"), int_TMN1 + m + 1, 0)
                    m = m + 1
                Next j
        End If

    Next i
    
    If LiquidPhaseExists = True Then
        m = LiquidIndex - 1
    Else
        m = NumberOfDatasetRows
    End If
    
    For i = 1 To m                                                                                      '<=Only checking the vapor species. These are the only ones used for Peng-Robinson calculations
        If IsNumeric(outputArray(i, iColumns.pc)) = False Or IsNumeric(outputArray(i, iColumns.tc)) = False Or _
                                        IsNumeric(outputArray(i, iColumns.omega)) = False Then       '<= make sure TC and PC values are non-zero
            myErrorMsg = "Some TC, PC or acentric factor values are not numeric!"
            GoTo myErrorHandler
        End If
        If outputArray(i, iColumns.omega) < 0.491 Then                                               '<= m or k for PR1978
            outputArray(i, iColumns.Ki) = 0.37464 + 1.54226 * outputArray(i, iColumns.omega) - 0.26992 * outputArray(i, iColumns.omega) ^ 2
        Else
            outputArray(i, iColumns.Ki) = 0.379642 + 1.487503 * outputArray(i, iColumns.omega) - _
                                                                        0.164423 * outputArray(i, iColumns.omega) ^ 2 + 0.016666 * outputArray(i, iColumns.omega) ^ 3 '<= m or k for PR1978
        End If

        If outputArray(i, iColumns.pc) <= 0 Or outputArray(i, iColumns.tc) <= 0 Then              '<= make sure TC and PC values are non-zero
            myErrorMsg = "Some TC or PC values are zero!"
            GoTo myErrorHandler
        End If
        outputArray(i, iColumns.bi) = 0.0778 * GasLawR * outputArray(i, iColumns.tc) / (outputArray(i, iColumns.pc))                 '<=b
    Next i
    
    For i = 1 To NumberOfDatasetRows                                                                    '<= Check both liquid and vapor species for valid Cp data
        If LCase(speciesNames(i)) <> "liquid" And speciesNames(i) <> "Not Found!" Then
            If IsNumeric(outputArray(i, iColumns.hvap)) = True Then
                If outputArray(i, iColumns.hvap) < 0 Then
                    outputArray(i, iColumns.hvap) = 0
                End If                                                                                  '<=Here only throw a warning
            End If
            If IsNumeric(outputArray(i, iColumns.hvap)) = False Then
                outputArray(i, iColumns.hvap) = 298.15
            End If
            If IsNumeric(outputArray(i, iColumns.tb)) = False Then
                outputArray(i, iColumns.tb) = 0
            End If
        End If
    Next i
    
    If LiquidSpeciesFound = True Then
        For i = 1 To (NumberOfDatasetRows - 1) / 2
            If LiquidIndex <> i And speciesNames(i + 1 + (NumberOfDatasetRows - 1) / 2) <> "Not Found!" Then
                If outputArray(i, iColumns.hvap) <> outputArray(i + 1 + (NumberOfDatasetRows - 1) / 2, iColumns.hvap) Then
                End If
            End If
        Next i
    End If
    
    CreateDataset = outputArray
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn)
    
    Application.ScreenUpdating = True
    
    Exit Function
  
myErrorHandler:


    
    ReDim outputArray(NumberOfDatasetRows, iColumns.lastCpIndex)
    For i = 0 To NumberOfDatasetRows
        For j = 0 To iColumns.lastCpIndex
            outputArray(i, j) = 0
        Next j
    Next i
    
    CreateDataset = outputArray
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, True)
    
    Application.ScreenUpdating = True
    
    End Function
    Private Function calculate_LiquidEntropy(dataset() As Variant, moleComp() As Double, TempK As Double, CpRanges() As Integer) As Double
    
    '***************************************************************************
    'This function is called by the Entropy function.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim LiquidCpDataRequired As Boolean
    Dim g As Integer
    Dim Denominator As Double
    Dim CpEquation As Double
    Dim local_Tempk As Double
    Dim NBP_Temp As Double
    Dim myErrorMsg As String
    Dim fcnName As String

    fcnName = "calculate_LiquidEntropy"

    calculate_LiquidEntropy = 0
            
            For g = 0 To dataset(0, iColumns.iSpecies)
                        If CpRanges(g, UBound(CpRanges, 2)) <> -500 And (dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) <> "No Data" _
                    And dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) <> "Not Found!") Then
                               
            If CpRanges(g, UBound(CpRanges, 2)) = -400 Then
                local_Tempk = dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) - 1)
            Else
                local_Tempk = TempK
            End If
                               
                If dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) = "NIST" Then
                    Denominator = 1000
                Else
                    Denominator = 1
                End If
                
                    CpEquation = 0

                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK)) * Log(local_Tempk / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 1) * (local_Tempk / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 2) * (local_Tempk / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 3) * (local_Tempk / Denominator) ^ 3 / 3
                    CpEquation = CpEquation - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 4) / (2 * (local_Tempk / Denominator) ^ 2)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 6)
                    calculate_LiquidEntropy = calculate_LiquidEntropy + moleComp(g) * CpEquation
                    
                    CpEquation = 0
                    
                    If CpRanges(g, UBound(CpRanges, 2)) = -400 Then
                        NBP_Temp = dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) - 1)
                    Else
                        NBP_Temp = dataset(g, iColumns.tb)
                    End If
                    
                    
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq)) * Log(NBP_Temp / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 1) * (NBP_Temp / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 2) * (NBP_Temp / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 3) * (NBP_Temp / Denominator) ^ 3 / 3
                    CpEquation = CpEquation - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 4) / (2 * (NBP_Temp / Denominator) ^ 2)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 6)
                    calculate_LiquidEntropy = calculate_LiquidEntropy - moleComp(g) * CpEquation
        
                    calculate_LiquidEntropy = calculate_LiquidEntropy - moleComp(g) * dataset(g, iColumns.hvap) / (dataset(g, iColumns.tb))
                
                    If CpRanges(g, UBound(CpRanges, 2)) = -400 Then
                        NBP_Temp = dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) - 1)
                    Else
                        NBP_Temp = dataset(g, iColumns.tb)
                    End If
                
                    CpEquation = 0
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap)) * Log(dataset(g, iColumns.tb) / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 1) * (dataset(g, iColumns.tb) / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 2) * (dataset(g, iColumns.tb) / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 3) * (dataset(g, iColumns.tb) / Denominator) ^ 3 / 3
                    CpEquation = CpEquation - dataset(g, CpRanges(g, iColumns.NBPVap) + 4) / (2 * (dataset(g, iColumns.tb) / Denominator) ^ 2)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 6)
                    calculate_LiquidEntropy = calculate_LiquidEntropy + moleComp(g) * CpEquation
                    
                    CpEquation = 0
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298)) * Log(298.15 / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 1) * (298.15 / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 2) * (298.15 / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 3) * (298.15 / Denominator) ^ 3 / 3
                    CpEquation = CpEquation - dataset(g, CpRanges(g, iColumns.Vap298) + 4) / (2 * (298.15 / Denominator) ^ 2)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 6)
                    calculate_LiquidEntropy = calculate_LiquidEntropy - moleComp(g) * CpEquation

SkipSpecies:

        End If
    Next g
    
'    NIST Data (units for H & S are different)
'    Cp = heat capacity (J/mol*K)
'    H° = standard enthalpy (kJ/mol)
'    S° = standard entropy (J/mol*K)
'    t = temperature(k) / 1000

'   HSC Data
'   Cp j/mol/K (units are the same for H & S)
'   T = temperaqture(k)
         
    calculate_LiquidEntropy = calculate_LiquidEntropy * (1000 / 1000)       '<= convert Cp data from j/g-mole/K to kJ/kg-mole/K
    
    Exit Function
    
myErrorHandler:

    calculate_LiquidEntropy = 987654321.123457 '<=error flag

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    End Function
    
    Private Function calculate_LiquidEnthalpy(dataset() As Variant, moleComp() As Double, TempK As Double, CpRanges() As Integer) As Double
    
    '***************************************************************************
    'This function is called by the Enthalpy function.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Dim LiquidCpDataRequired As Boolean
    Dim g As Integer
    Dim Denominator As Double
    Dim J_to_kJ As Double
    Dim CpEquation As Double
    Dim NBP_Temp As Double
    Dim local_Tempk As Double
    Dim myErrorMsg As String
    Dim fcnName As String

    fcnName = "calculate_LiquidEnthalpy"

    calculate_LiquidEnthalpy = 0
                
    For g = 0 To dataset(0, iColumns.iSpecies)
        If CpRanges(g, UBound(CpRanges, 2)) <> -500 And dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) <> "No Data" _
                    And dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) <> "Not Found!" Then
                    
            If CpRanges(g, UBound(CpRanges, 2)) = -400 Then
                local_Tempk = dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) - 1)
            Else
                local_Tempk = TempK
            End If
                                                                                                                      
                If dataset(g + dataset(0, iColumns.iSpecies) + 1, iColumns.CpDataType) = "NIST" Then
                    Denominator = 1000
                    J_to_kJ = 1         'NIST enthalpy data is already in Kj/g-mole
                Else
                    Denominator = 1
                    J_to_kJ = 1000      'HSC data enthalpy is in j/g-mole
                End If
                
                    CpEquation = 0

                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK)) * local_Tempk / Denominator
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 1) * (local_Tempk / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 2) * (local_Tempk / Denominator) ^ 3 / 3
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 3) * (local_Tempk / Denominator) ^ 4 / 4
                    CpEquation = CpEquation - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 4) / (local_Tempk / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 5) - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.TempK) + 7)
                    
                    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy + moleComp(g) * CpEquation / J_to_kJ
                    
                    If CpRanges(g, UBound(CpRanges, 2)) = -400 Then
                        NBP_Temp = dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) - 1)
                    Else
                        NBP_Temp = dataset(g, iColumns.tb)
                    End If
                    
                    CpEquation = 0
                                                                                                         
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq)) * NBP_Temp / Denominator
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 1) * (NBP_Temp / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 2) * (NBP_Temp / Denominator) ^ 3 / 3
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 3) * (NBP_Temp / Denominator) ^ 4 / 4
                    CpEquation = CpEquation - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 4) / (NBP_Temp / Denominator)
                    CpEquation = CpEquation + dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 5) - dataset(g + dataset(0, iColumns.iSpecies) + 1, CpRanges(g, iColumns.NBPLiq) + 7)
                    
                    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy - moleComp(g) * CpEquation / J_to_kJ
                    
                    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy - moleComp(g) * dataset(g, iColumns.hvap) / 1000
  
                If dataset(g, iColumns.CpDataType) = "NIST" Then
                    Denominator = 1000
                    J_to_kJ = 1
                Else
                    Denominator = 1
                    J_to_kJ = 1000
                End If
  
                    CpEquation = 0
                    
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap)) * dataset(g, iColumns.tb) / Denominator
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 1) * (dataset(g, iColumns.tb) / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 2) * (dataset(g, iColumns.tb) / Denominator) ^ 3 / 3
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 3) * (dataset(g, iColumns.tb) / Denominator) ^ 4 / 4
                    CpEquation = CpEquation - dataset(g, CpRanges(g, iColumns.NBPVap) + 4) / (dataset(g, iColumns.tb) / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.NBPVap) + 5) - dataset(g, CpRanges(g, iColumns.NBPVap) + 7)
                    
                    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy + moleComp(g) * CpEquation / J_to_kJ
                    
                    CpEquation = 0
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298)) * 298.15 / Denominator
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 1) * (298.15 / Denominator) ^ 2 / 2
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 2) * (298.15 / Denominator) ^ 3 / 3
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 3) * (298.15 / Denominator) ^ 4 / 4
                    CpEquation = CpEquation - dataset(g, CpRanges(g, iColumns.Vap298) + 4) / (298.15 / Denominator)
                    CpEquation = CpEquation + dataset(g, CpRanges(g, iColumns.Vap298) + 5) - dataset(g, CpRanges(g, iColumns.Vap298) + 7)
                    
                    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy - moleComp(g) * CpEquation / J_to_kJ
                                                                   
                     
SkipSpecies:

        End If
    Next g
    
'    NIST Data (units for H & S are different)
'    Cp = heat capacity (J/mol*K)
'    H° = standard enthalpy (kJ/mol)
'    S° = standard entropy (J/mol*K)
'    t = temperature(k) / 1000

'   HSC Data
'   Cp j/mol/K (units are the same for H & S)
'   T = temperaqture(k)
          
    calculate_LiquidEnthalpy = calculate_LiquidEnthalpy * 1000 ' <= Convert from kJ/g-mole to kJ/kg-mole
    
    Exit Function
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    calculate_LiquidEnthalpy = 987654321.123457 '<=error flag
    
    End Function

    
    Public Function Entropy(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                vaporOrLiquid As Variant, BinariesUsed As Boolean, Optional kij0 As Variant, _
                                Optional kijT As Variant, Optional deComp As Variant, Optional calcAsIdealGas As Boolean = False, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'This function calculates the liquid, ideal gas or PR1978 EOS entropy.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = True
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim DepartureS As Double
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim ig_Entropy As Double
    Dim passedTempK As Double
    Dim moleComp() As Double
    Dim CpRanges() As Integer
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim myErrorMsg As String
    Dim Phase As String
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim datasetErrMsgsOn As Boolean
    

    fcnName = "Entropy"
    myErrorMsg = ""
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, True)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    'Check user inputs
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    'Check moles
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    CpRanges = selectCpDataRanges(dataset, TempK, Phase, moleComp)
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        If CpRanges(i, UBound(CpRanges, 2)) = -400 Then
            myErrorMsg = "Some species in the dataset do no have valid data for " & TempC & " C"
        End If
    Next i
    
    Entropy = 0
    DepartureS = 0
    
    If LCase(Phase) = "vapor" Then
        If calcAsIdealGas = False Then
            If pbara > 1 Then
                DepartureS = calculate_S_Departure(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
                If DepartureS = 987654321.123457 Then                                                                                           '<=Error flag returned from calculate_S_Departure
                    myErrorMsg = "Entropy error: Departure function error: Divide by zero or log() of negative number. Check phase!"
                    GoTo myErrorHandler
                End If
            End If
        Else
            DepartureS = 0
        End If
        
        ig_Entropy = calculate_IdealGasEntropy(dataset, moleComp, TempK, pbara, CpRanges)
        
        If ig_Entropy = 987654321.123457 Then
            myErrorMsg = "calculate_IdealGasEntropy returned an error."
            GoTo myErrorHandler
        End If
                
        Entropy = DepartureS + ig_Entropy
        
    End If
    
    If LCase(Phase) = "liquid" Then
        Entropy = calculate_LiquidEntropy(dataset, moleComp, TempK, CpRanges)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
   
    Application.ScreenUpdating = True
    
    Exit Function
   
myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Entropy = 0
    
    Application.ScreenUpdating = True
    
   End Function
   
       Public Function return_kijTs(DataRange As Range, temperature As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, _
                                                        Optional errMsgsOn As Boolean = False) As Variant()
                                                        
    '***************************************************************************
    'This function calculates the calculates the return_kijTs of the Predictive PR1978 EOS.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim CvIG As Double
    Dim CpIG As Double
    Dim CpRanges() As Integer
    Dim CvResidual As Double
    Dim CpResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    
    Dim outputArray() As Double
    Dim returnArray() As Variant
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    fcnName = "return_kijTs"

    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    
    ReDim returnArray(dataset(0, iColumns.iSpecies), dataset(0, iColumns.iSpecies))
    For i = 0 To dataset(0, iColumns.iSpecies)
        For j = 0 To dataset(0, iColumns.iSpecies)
            returnArray(i, j) = kij0_Array(i, j)
        Next j
    Next i
    
    return_kijTs = returnArray
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
        
    return_kijTs = Application.Transpose(returnArray)
    
    Application.ScreenUpdating = True

    End Function
   
    Public Function Derivatives(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, _
                                                        Optional returnUnits As Boolean = False, Optional errMsgsOn As Boolean = False) As Variant()
                                                        
    '***************************************************************************
    'This function calculates the calculates the derivatives of the PR1978 EOS.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim CvIG As Double
    Dim CpIG As Double
    Dim CpRanges() As Integer
    Dim CvResidual As Double
    Dim CpResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    
    Dim outputArray() As Double
    Dim returnArray() As Variant
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    fcnName = "Derivatives"
    Phase = "vapor"
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                myErrorMsg = "Warning: User has provided decomposition and predictive binaries is specified but BinariesUsed is false! Calculating without binaries."
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If

    
    outputArray() = calculate_Derivatives(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)

    If returnUnits = True Then
        ReDim returnArray(1, 12)
        For i = 0 To 12
            returnArray(0, i) = outputArray(i)
        Next i
        returnArray(1, 0) = "dadT at constant V (m^6 bar/mol^2-K)"
        returnArray(1, 1) = "dPdv at constant T (bar/(m3/mol))"
        returnArray(1, 2) = "dPdT at constant V (bar/K)"
        returnArray(1, 3) = "dadT at constant P (m3/mol-K)"
        returnArray(1, 4) = "dBdT at constant P (1/K)"
        returnArray(1, 5) = "dZdT at constant P (1/K)"
        returnArray(1, 6) = "dVdT at constant P (m3/mol-K)"
        returnArray(1, 7) = "sum_b (m^3/mol)"
        returnArray(1, 8) = "sum_a (m^6 bara/mol^2)"
        returnArray(1, 9) = "A"
        returnArray(1, 10) = "B"
        returnArray(1, 11) = "Z"
        returnArray(1, 12) = "v (m3/mol)"
        
        Derivatives = Application.Transpose(returnArray)
    Else
        Derivatives = Application.Transpose(outputArray)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    If returnUnits = False Then
    ReDim returnArray(12)
        For i = 0 To 12
            returnArray(i) = 0
        Next i
    Else
        ReDim returnArray(1, 12)
        For i = 0 To 1
            For j = 0 To 12
                returnArray(i, j) = 0
            Next j
        Next i
    End If
        
    Derivatives = Application.Transpose(returnArray)
    
    Application.ScreenUpdating = True

    End Function
    Public Function phaseCp(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, vaporOrLiquid As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Variant()
    
    '***************************************************************************
    'This function calculates the liquid, ideal gas or PR1978 EOS Cp.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim Derivatives() As Double
    Dim CvIG As Double
    Dim CpIG As Double
    Dim CpRanges() As Integer
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim CvResidual As Double
    Dim CpResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    Dim ouputArray() As Variant
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    Application.EnableEvents = False
    fcnName = "Cp"
    
    ReDim ouputArray(2)
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, True)
    Else
        myErrorMsg = "No dataset provided."
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    CpRanges = selectCpDataRanges(dataset, TempK, Phase, moleComp)
    
    CpIG = calculate_Cp_IGorLiquid(dataset, moleComp, TempK, CpRanges, Phase)
       
    If LCase(Phase) = "liquid" Then
        ouputArray(0) = CpIG
        phaseCp = ouputArray
    Else
        If LCase(Phase) = "vapor" And pbara <= 1 Then
        
            ouputArray(0) = CpIG
            ouputArray(1) = CpIG
            ouputArray(2) = 0
            phaseCp = ouputArray
            phaseCp = Application.Transpose(phaseCp)
            
        Else
        
            If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
                If BinariesUsed = True And IsMissing(kij0) = True Then
                
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                    myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
                End If
                
                kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
                kijT_Array = createKijTArray(dataset, kijT)
            Else                                                                                        '<predictive
                If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
                    If BinariesUsed = True Then
                        myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
                    End If
                    
                    BinariesUsed = False
                    dataset(0, iColumns.predictive) = 0
                    kij0_Array = createKij0Array(dataset, BinariesUsed)
                    kijT_Array = createKijTArray(dataset)
                Else
                    If BinariesUsed = False Then
                        dataset(0, iColumns.predictive) = 0
                        kij0_Array = createKij0Array(dataset, BinariesUsed)
                        kijT_Array = createKijTArray(dataset)
                    End If
                    passedTempK = TempK
                    alpha_aiArray = create_alphaiArray(dataset, TempK)
                    aiArray = create_aiArray(dataset, alpha_aiArray)
                    kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
                    kijT_Array = createKijTArray(dataset)
                End If
            End If
            
            Derivatives() = calculate_Derivatives(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
            CvIG = CpIG - GasLawR * 100000
            
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            
            aiArray = create_aiArray(dataset, alpha_aiArray)
            
            d2aidT2Array = create_d2aidT2Array(dataset, TempK)
            
            daidTArray = create_daidTArray(dataset, aiArray, TempK)
            
            d2adT2 = calculate_d2adT2(dataset, moleComp, TempK, aiArray, daidTArray, d2aidT2Array, BinariesUsed, kij0_Array, kijT_Array)
            
            If (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5)) <= 0 Then
                myErrorMsg = "The natural log term in the Cv residual equation retun an error. Check phase."
                GoTo myErrorHandler
            End If
            
            If Derivatives(iColumns.sumb) <= 0 Then
                myErrorMsg = "The term sum_b is less than or equal to zero. Check phase."
                GoTo myErrorHandler
            End If
            
            CvResidual = 100000 * TempK * d2adT2 * Log((Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) _
                                    / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5))) / (Derivatives(iColumns.sumb) * 8 ^ 0.5)
                                    
            CpResidual = CvResidual + (TempK * (Derivatives(iColumns.dPdT_constV)) * Derivatives(iColumns.dVdT_constP) - GasLawR) * 100000
            
            
            
            ouputArray(0) = CpIG + CpResidual
            ouputArray(1) = CpIG
            ouputArray(2) = CpResidual
            phaseCp = ouputArray
            phaseCp = Application.Transpose(phaseCp)
        End If
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    ouputArray(0) = 0
    ouputArray(1) = 0
    ouputArray(2) = 0
    phaseCp = ouputArray
    phaseCp = Application.Transpose(phaseCp)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    
    End Function
    Public Function SpeedOfSound(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Double
                                    
    '***************************************************************************
    'This function calculates the vapor speed of sound for the PR1978 EOS.
    '***************************************************************************
                                    
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim Derivatives() As Double
    Dim CvIG As Double
    Dim CpIG As Double
    Dim Cp As Double
    Dim Cv As Double
    Dim CpRanges() As Integer
    Dim CvResidual As Double
    Dim CpResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    Dim aveMW As Double
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    Phase = "vapor"
    fcnName = "SpeedOfSound"
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    Derivatives() = calculate_Derivatives(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
    CpRanges = selectCpDataRanges(dataset, TempK, "vapor", moleComp)
    
    CpIG = calculate_Cp_IGorLiquid(dataset, moleComp, TempK, CpRanges, "vapor")
    
    CvIG = CpIG - GasLawR * 100000
    
    alpha_aiArray = create_alphaiArray(dataset, TempK)
    
    aiArray = create_aiArray(dataset, alpha_aiArray)
    
    d2aidT2Array = create_d2aidT2Array(dataset, TempK)
    
    daidTArray = create_daidTArray(dataset, aiArray, TempK)
    
    d2adT2 = calculate_d2adT2(dataset, moleComp, TempK, aiArray, daidTArray, d2aidT2Array, BinariesUsed, kij0_Array, kijT_Array)
    
        If (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5)) <= 0 Then
            myErrorMsg = "The natural log term in the Cv residual equation retun an error. Check phase."
            GoTo myErrorHandler
        End If
        
        If Derivatives(iColumns.sumb) <= 0 Then
            myErrorMsg = "The term sum_b is less than or equal to zero. Check phase."
            GoTo myErrorHandler
        End If
    
    CvResidual = 100000 * TempK * d2adT2 * Log((Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) _
                            / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5))) / (Derivatives(iColumns.sumb) * 8 ^ 0.5)
                            
    CpResidual = CvResidual + (TempK * (Derivatives(iColumns.dPdT_constV)) * Derivatives(iColumns.dVdT_constP) - GasLawR) * 100000
    
    Cv = CvIG + CvResidual
    Cp = CpIG + CpResidual
    
    SpeedOfSound = Derivatives(iColumns.vol) * (-(Cp / Cv) * (Derivatives(iColumns.dPdv_constT))) ^ 0.5
    
    aveMW = 0
    For i = 0 To dataset(0, iColumns.iSpecies)
        aveMW = aveMW + moleComp(i) * dataset(i, iColumns.MW)
    Next i
    
    SpeedOfSound = (SpeedOfSound ^ 2 * (100000) * (1000) * (1 / aveMW)) ^ 0.5           '<= m/s
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Cp = 0
    
    Application.ScreenUpdating = True

    
    End Function
    Public Function JT_Coef(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Double
    
    '***************************************************************************
    'This function calculates the Joule-Thompson coefficient for the PR1978 EOS.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim Derivatives() As Double
    Dim CvIG As Double
    Dim CpIG As Double
    Dim Cp As Double
    Dim Cv As Double
    Dim CpRanges() As Integer
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim CvResidual As Double
    Dim CpResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    Phase = "vapor"
    fcnName = "JT_Coef"
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    Derivatives() = calculate_Derivatives(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
    CpRanges = selectCpDataRanges(dataset, TempK, "vapor", moleComp)
    
    CpIG = calculate_Cp_IGorLiquid(dataset, moleComp, TempK, CpRanges, "vapor")
    
    CvIG = CpIG - GasLawR * 100000
    
    alpha_aiArray = create_alphaiArray(dataset, TempK)
    
    aiArray = create_aiArray(dataset, alpha_aiArray)
    
    d2aidT2Array = create_d2aidT2Array(dataset, TempK)
    
    daidTArray = create_daidTArray(dataset, aiArray, TempK)
    
    d2adT2 = calculate_d2adT2(dataset, moleComp, TempK, aiArray, daidTArray, d2aidT2Array, BinariesUsed, kij0_Array, kijT_Array)
    
        If (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5)) <= 0 Then
            myErrorMsg = "The natural log term in the Cv residual equation retun an error. Check phase."
            GoTo myErrorHandler
        End If
        
        If Derivatives(iColumns.sumb) <= 0 Then
            myErrorMsg = "The term sum_b is less than or equal to zero. Check phase."
            GoTo myErrorHandler
        End If
    
    CvResidual = 100000 * TempK * d2adT2 * Log((Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) _
                            / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5))) / (Derivatives(iColumns.sumb) * 8 ^ 0.5)
                            
    CpResidual = CvResidual + (TempK * (Derivatives(iColumns.dPdT_constV)) * Derivatives(iColumns.dVdT_constP) - GasLawR) * 100000
    
    Cv = CvIG + CvResidual
    Cp = CpIG + CpResidual
    
    JT_Coef = (1 / Cp) * (TempK * Derivatives(iColumns.dVdT_constP) - Derivatives(iColumns.vol)) * 100000
    
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Cp = 0
    
    Application.ScreenUpdating = True

    
    End Function
    

    Public Function Enthalpy(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                vaporOrLiquid As Variant, BinariesUsed As Boolean, Optional kij0 As Variant, _
                                Optional kijT As Variant, Optional deComp As Variant, Optional calcAsIdealGas As Boolean, Optional errMsgsOn As Boolean = False) As Double
                                
    '***************************************************************************
    'This function calculates the liquid, ideal gas or PR1978 EOS enthalpy.
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = True
    
    Dim i As Integer
    Dim j As Integer
    Dim DepartureH As Double
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim ig_Enthalpy As Double
    Dim passedTempK As Double
    Dim moleComp() As Double
    Dim CpRanges() As Integer
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim dataset() As Variant
    Dim myErrorMsg As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim Phase As String
    Dim datasetErrMsgsOn As Boolean
    

    fcnName = "Enthalpy"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset() = validateDataset(DataRange, Phase, True)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If

    If IsMissing(moles) = False Then
        moleComp() = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    CpRanges() = selectCpDataRanges(dataset, TempK, Phase, moleComp)
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            If CpRanges(i, UBound(CpRanges, 2)) = -400 Then
                myErrorMsg = "Some species in the dataset do no have valid data for " & TempC & " C"
            End If
        Next i
    
    Enthalpy = 0
    DepartureH = 0

    If LCase(Phase) = "vapor" Then
        If pbara > 1 Then
        
            DepartureH = calculate_H_Departure(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)    '<= vapor phase constant pressure (Cp) enthalpy change from ideal gas conditions (25 C @ 1 bara)
            
            If DepartureH = 987654321.123457 Then                                                                               ' Error flag returned from Calculate_H_Departure
                myErrorMsg = "Enthalpy error: Departure function error: Divide by zero of log() of negative number. Check phase!"
                GoTo myErrorHandler
            End If
        Else
            DepartureH = 0
        End If
                
        ig_Enthalpy = calculate_IdealGasEnthalpy(dataset, moleComp, TempK, CpRanges)
        
        If ig_Enthalpy = 987654321.123457 Then
            myErrorMsg = "calculate_IdealGasEnthalpy returned an error."
            GoTo myErrorHandler
        End If
                
        Enthalpy = DepartureH + ig_Enthalpy

    End If
    
    If LCase(Phase) = "liquid" Then
        Enthalpy = calculate_LiquidEnthalpy(dataset, moleComp, TempK, CpRanges)
        
        If Enthalpy = 987654321.123457 Then
            myErrorMsg = "calculate_LiquidEnthalpy returned an error."
            GoTo myErrorHandler
        End If
        
    End If
    

    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
   
myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Enthalpy = 0
    
    Application.ScreenUpdating = True
    
    End Function
    
    Public Function PhasePhi(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, vaporOrLiquid As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Variant()
                                    
    '***************************************************************************
    'This function calculates the PR1978 EOS fugacity coefficients.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim moleComp() As Double
    Dim outputArray() As Variant
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim myErrorMsg As String
    Dim passedTempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim Phase As String
    
    Dim datasetErrMsgsOn As Boolean
    
    
    

    fcnName = "PhasePhi"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    ReDim outputArray(dataset(0, iColumns.iSpecies) + 2)
    

    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    outputArray = calculate_Phi(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array, aiArray)
    
    If outputArray(dataset(0, iColumns.iSpecies) + 2) = -500 Then
        GoTo myErrorHandler
    End If
    
    PhasePhi = outputArray
    
    PhasePhi = Application.Transpose(PhasePhi)                  '<This makes the default horizontal output array into a vertical output array - the is better for heat and mass balances
    
 
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)            '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:
    
    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(4)
        For i = 0 To UBound(outputArray, 1)
            outputArray(i) = 0
        Next i
    Else
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i) = 0
        Next i
    End If
    
    PhasePhi = outputArray
    
    PhasePhi = Application.Transpose(PhasePhi)
    
    If dataset(0, iColumns.globalErrmsg) <> "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function

    Public Function vaporCv(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Variant()

    '***************************************************************************
    'This function calculates the liquid, ideal gas or PR1978 EOS Cv.
    '***************************************************************************

    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim Derivatives() As Double
    Dim CvIG As Double
    Dim CpIG As Double
    Dim CpRanges() As Integer
    Dim CvResidual As Double
    Dim d2aidT2Array() As Double
    Dim d2adT2 As Double
    Dim daidTArray() As Double
    Dim sum_b As Double
    Dim outputArray() As Variant
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    fcnName = "Cv"
    Phase = "vapor"
    ReDim outputArray(2)
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    Derivatives() = calculate_Derivatives(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
    If Derivatives(0) <> 987654321.12345 Then
    
        CpRanges = selectCpDataRanges(dataset, TempK, "vapor", moleComp)
        
        CpIG = calculate_Cp_IGorLiquid(dataset, moleComp, TempK, CpRanges, "vapor")
        
        CvIG = CpIG - GasLawR * 100000
        
        alpha_aiArray = create_alphaiArray(dataset, TempK)
        
        aiArray = create_aiArray(dataset, alpha_aiArray)
        
        d2aidT2Array = create_d2aidT2Array(dataset, TempK)
        
        daidTArray = create_daidTArray(dataset, aiArray, TempK)
        
        d2adT2 = calculate_d2adT2(dataset, moleComp, TempK, aiArray, daidTArray, d2aidT2Array, BinariesUsed, kij0_Array, kijT_Array)
        
        If (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5)) <= 0 Then
            myErrorMsg = "The natural log term in the Cv residual equation retun an error. Check phase."
            GoTo myErrorHandler
        End If
        
        If Derivatives(iColumns.sumb) <= 0 Then
            myErrorMsg = "The term sum_b is less than or equal to zero. Check phase."
            GoTo myErrorHandler
        End If
        
        CvResidual = 100000 * TempK * d2adT2 * Log((Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 + 2 ^ 0.5)) _
                                / (Derivatives(iColumns.Z) + Derivatives(iColumns.b) * (1 - 2 ^ 0.5))) / (Derivatives(iColumns.sumb) * 8 ^ 0.5)
        
        outputArray(0) = CvIG + CvResidual
        outputArray(1) = CvIG
        outputArray(2) = CvResidual
        vaporCv = outputArray
        vaporCv = Application.Transpose(vaporCv)
    Else
        myErrorMsg = "Calculate_Derivatives returned an error."
        GoTo myErrorHandler
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    outputArray(0) = 0
    outputArray(1) = 0
    outputArray(2) = 0
    vaporCv = outputArray
    vaporCv = Application.Transpose(vaporCv)
    
    Application.ScreenUpdating = True
    
    End Function
    
    Public Function Volume(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, vaporOrLiquid As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Double
                                    
    '***************************************************************************
    'This function calculates the liquid or vapor PR1978 EOS molar volume.
    '***************************************************************************

    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    Dim Z As Double
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    fcnName = "Volume"
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    Z = calculate_PhaseZ(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
    If Z = -500 Then
        myErrorMsg = "Calculate_PhaseZ returned an error."
    End If
    
    Volume = Z * GasLawR * TempK / pbara
    
    
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Volume = 0
    
    Application.ScreenUpdating = True
    
    End Function
    


    Public Function PhaseZ(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, vaporOrLiquid As Variant, _
                                    BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Double
                                    
    '***************************************************************************
    'This function calculates the PR1978 EOS compressibilty.
    '***************************************************************************

    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim j As Integer
    Dim dataset() As Variant
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim passedTempK As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim Phase As String
    
    Dim datasetErrMsgsOn As Boolean

    Dim UDF_Range As Range
    
    fcnName = "PhaseZ"
    
    Set UDF_Range = Application.Caller
    
    Phase = checkInputPhase(dataset, vaporOrLiquid)
    If LCase(Phase) <> "vapor" And LCase(Phase) <> "liquid" Then
        myErrorMsg = Phase
        GoTo myErrorHandler
    End If
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    myErrorMsg = ""
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    PhaseZ = calculate_PhaseZ(Phase, dataset, moleComp, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
    
    If PhaseZ = -500 Then
        myErrorMsg = "Calculate_PhaseZ returned an error."
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    PhaseZ = 0
    
    Application.ScreenUpdating = True
    
    End Function

    Public Function DewT(DataRange As Range, pressure As Variant, moles As Variant, _
                                BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional Guess As Double = 0, _
                                 Optional errMsgsOn As Boolean = False, Optional CallFromFlash As Boolean = False) As Variant()
                                 
    '***************************************************************************
    'This function calculates the PR1978 EOS dew point.
    '***************************************************************************
                                
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = True
    
    Dim Iter1 As Integer
    Dim Iter2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim HiLoTemp_Count  As Integer
    Dim Upper_xi_Dew_Count  As Integer
    Dim DewT_Count As Integer
    Dim Lower_xi_Dew_Count  As Integer
    Dim TotalIterations As Integer
    
    Dim ConvergenceSum As Double
    Dim xi_Sum As Double
    Dim xi_Low_Sum As Double
    Dim xi_Hi_Sum As Double
    Dim xi_New_Sum As Double
    Dim T_Low As Double
    Dim T_Hi As Double
    Dim T_New As Double
    Dim T_Bub_Est As Double
    Dim xi_Initial_Sum As Double
    Dim T_Dew As Double
    Dim T_Dew_C As Double
    Dim T_Bub As Double
    Dim TempK As Double
    Dim TempC As Double
    Dim VaporFugacity As Double
    Dim LiquidFugacity As Double
    Dim FugacityTest As Double
    Dim pbara As Double
    
    Dim Ki_Old() As Double
    Dim Ki_PR() As Double
    Dim xi_Dew() As Double
    Dim xi_Old() As Double
    Dim xi_temp() As Double
    Dim New_Ki() As Double
    Dim xi_Array() As Double
    Dim yi_Array() As Double
    Dim Ki() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim InitialTemps() As Double
    Dim passedTempK As Double
    
    Dim dataset() As Variant
    Dim Phi_Vap() As Variant
    Dim Phi_Liq() As Variant
    Dim outputArray() As Variant
    
    Dim myErrorMsg As String
    Dim Saturated As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim datasetErrMsgsOn As Boolean
    Dim xiDew_Equals_xiOld As Boolean
    Dim HiAndLoTempsFound As Boolean
    Dim DewT_Found As Boolean

    fcnName = "DewT"

    myErrorMsg = ""
    
    ReDim InitialTemps(1)
    
    InitialTemps(0) = 0
    InitialTemps(1) = 0

    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "Vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    ReDim Phi_Vap(dataset(0, iColumns.iSpecies) + 2)
    ReDim Phi_Liq(dataset(0, iColumns.iSpecies) + 2)
    ReDim Ki_Old(dataset(0, iColumns.iSpecies))
    ReDim Ki(dataset(0, iColumns.iSpecies))
    ReDim Ki_PR(dataset(0, iColumns.iSpecies))
    ReDim xi_Dew(dataset(0, iColumns.iSpecies))
    ReDim xi_Old(dataset(0, iColumns.iSpecies))
    ReDim xi_temp(dataset(0, iColumns.iSpecies))
    ReDim outputArray(3 + dataset(0, iColumns.iSpecies))
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    


    Upper_xi_Dew_Count = 0
    DewT_Count = 0
    Lower_xi_Dew_Count = 0
    HiLoTemp_Count = 0
    TotalIterations = 0
    
    If Guess <> 0 Then
        T_Dew = Guess + 273.15
    Else
        InitialTemps = calculate_T_BubDew_Est(dataset, moleComp, pbara)
        If InitialTemps(1) <> 0 Then
            T_Dew = InitialTemps(1) 'Index 0 = Bubble Point and Index 1 = Dew Point
        Else
            T_Dew = 0
            myErrorMsg = "Calculate_T_BubDew_Est calculation failed to provide dew T estimate."
            GoTo myErrorHandler
        End If
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = T_Dew
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    xi_Initial_Sum = 0
    
    For i = 0 To dataset(0, iColumns.iSpecies)
       Ki_Old(i) = dataset(i, iColumns.pc) ^ ((1 / T_Dew - 1 / dataset(i, iColumns.tb)) / _
       (1 / dataset(i, iColumns.tc) - 1 / dataset(i, iColumns.tb))) / pbara
       
       xi_Dew(i) = moleComp(i) / Ki_Old(i)
       
       xi_Initial_Sum = xi_Initial_Sum + xi_Dew(i)
    Next i
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        xi_Old(i) = xi_Dew(i) / xi_Initial_Sum
    Next i
    
    Do Until HiAndLoTempsFound = True

        HiLoTemp_Count = HiLoTemp_Count + 1
    
        If HiLoTemp_Count > 2000 Then
            myErrorMsg = "Warning: HiLoTemp_Count  > 2000."
            GoTo myErrorHandler
        End If
            
        For i = 0 To dataset(0, iColumns.iSpecies)
            xi_Dew(i) = moleComp(i) / Ki_Old(i)
        Next i
    

        Do Until xiDew_Equals_xiOld = True
        
            Upper_xi_Dew_Count = Upper_xi_Dew_Count + 1
            
            If Upper_xi_Dew_Count > 1000 Then
                myErrorMsg = "Upper_xi_Dew_Count > 1000"
            End If
        
            xi_Old() = xi_Dew()
            xi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Sum = xi_Sum + xi_Dew(i)
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = xi_Old(i) / xi_Sum
            Next i
            
            Phi_Liq = calculate_Phi("Liquid", dataset, xi_Dew, T_Dew, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Liq(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Liquid phase Phi error in inner upper loop."
                GoTo myErrorHandler
            End If
            
            Phi_Vap = calculate_Phi("Vapor", dataset, moleComp, T_Dew, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Vap(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Vapor phase Phi error in inner upper loop."
                GoTo myErrorHandler
            End If
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If Phi_Vap(i) > 10 ^ -35 Then
                    Ki_PR(i) = Phi_Liq(i) / Phi_Vap(i)
                Else
                    Ki_PR(i) = 10 ^ 35
                End If
            Next i
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = moleComp(i) / Ki_PR(i)
            Next i
                
            xi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Sum = xi_Sum + xi_Dew(i)
            Next i
            
            xi_temp() = xi_Dew()
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = xi_Dew(i) / xi_Sum
            Next i
            
            ConvergenceSum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies) - 1
                If xi_Old(i) <> 0 Then
                    ConvergenceSum = ConvergenceSum + Abs((xi_Dew(i) / xi_Old(i)) - 1)
                End If
            Next i
            
            If ConvergenceSum < 10 ^ -5 Then
                xiDew_Equals_xiOld = True
            End If
        Loop
        
        xiDew_Equals_xiOld = False
        
        If HiLoTemp_Count = 1 Then
             xi_Low_Sum = xi_Sum - 1#
             xi_Hi_Sum = xi_Sum - 1#
             T_Low = T_Dew
             T_Hi = T_Dew
        End If
        
         If xi_Sum < 1# Then
             T_Hi = T_Dew
             xi_Hi_Sum = xi_Sum - 1#
             T_Dew = T_Dew / 1.01
         Else
             T_Low = T_Dew
             xi_Low_Sum = xi_Sum - 1#
             T_Dew = T_Dew * 1.0101
         End If
        
         If xi_Low_Sum * xi_Hi_Sum > 0 Or xi_Low_Sum = 0 Or xi_Hi_Sum = 0 Then
          Else
             HiAndLoTempsFound = True
          End If
    Loop
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
        
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = (T_Hi + T_Low) / 2
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    Do Until DewT_Found = True

        DewT_Count = DewT_Count + 1
        
        If DewT_Count = 2000 Then
            myErrorMsg = "Warning - DewT_Count > 2000."
            GoTo myErrorHandler
        End If
        
        T_New = (T_Hi + T_Low) / 2
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            xi_Dew(i) = moleComp(i) / Ki_Old(i)
         Next i
        
        xiDew_Equals_xiOld = False
        
        Do Until xiDew_Equals_xiOld = True
        
            Lower_xi_Dew_Count = Lower_xi_Dew_Count + 1
            
            If Lower_xi_Dew_Count > 2000 Then
                myErrorMsg = "Warning - Lower_xi_Dew_Count  > 2000."
                GoTo myErrorHandler
            End If
        
            xi_Old() = xi_Dew()
            
            xi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Sum = xi_Sum + xi_Dew(i)
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = xi_Old(i) / xi_Sum
            Next i
                  
            Phi_Liq = calculate_Phi("Liquid", dataset, xi_Dew, T_New, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Liq(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Liquid phase Phi error in inner lower loop."
                GoTo myErrorHandler
            End If
            
            Phi_Vap = calculate_Phi("Vapor", dataset, moleComp, T_New, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Vap(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Vapor phase Phi error in inner lower loop."
                GoTo myErrorHandler
            End If
            
            If Abs(Phi_Vap(dataset(0, iColumns.iSpecies) + 1) - Phi_Liq(dataset(0, iColumns.iSpecies) + 1)) < 10 ^ -5 Then
                myErrorMsg = "Trival solution convergence in lower inner loop."
                GoTo myErrorHandler
            End If
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If Phi_Vap(i) > 10 - 35 Then
                    Ki_PR(i) = Phi_Liq(i) / Phi_Vap(i)
                Else
                    Ki_PR(i) = 10 ^ 35
                End If
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = moleComp(i) / Ki_PR(i)
            Next i
                
            xi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Sum = xi_Sum + xi_Dew(i)
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                xi_Dew(i) = xi_Dew(i) / xi_Sum
            Next i
            
            ConvergenceSum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If xi_Old(i) <> 0 Then
                    ConvergenceSum = ConvergenceSum + Abs((xi_Dew(i) / xi_Old(i)) - 1)
                End If
            Next i
            
            If ConvergenceSum < 10 ^ -10 Then
                xiDew_Equals_xiOld = True
            End If
        Loop
        
        xi_New_Sum = xi_Sum - 1
        
        If xi_Low_Sum * xi_New_Sum > 0 Then
            xi_Low_Sum = xi_New_Sum
            T_Low = T_New
        Else
            xi_Hi_Sum = xi_New_Sum
            T_Hi = T_New
        End If
        
        If Abs(T_Hi - T_Low) < 0.01 Then
            DewT_Found = True
        End If
 
    Loop
    
    VaporFugacity = 0
    LiquidFugacity = 0
    FugacityTest = 0
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        VaporFugacity = VaporFugacity + Phi_Vap(i) * moleComp(i)
        LiquidFugacity = LiquidFugacity + Phi_Liq(i) * xi_Dew(i)
    Next i
    
    FugacityTest = VaporFugacity - LiquidFugacity
    
    If FugacityTest > 10 ^ -2 Then                                       '<= Not holding a very tight tolerance for fugacity at the dew point.
        Debug.Print "Dew temperature equilibrium test failed. Returned dew temperature may be inacurate"
    End If

    outputArray(0) = T_New - 273.15
    
    If Guess = 0 Then
        outputArray(1) = InitialTemps(1) - 273.15
    Else
        outputArray(1) = Guess
    End If
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i + 2) = xi_Dew(i)
    Next i
    
    If CallFromFlash = True Then
        DewT = outputArray
    Else
        DewT = outputArray
        DewT = Application.Transpose(DewT)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If Err.Number = 0 And T_Hi <> 0 And T_Low <> 0 Then
        If Abs((T_Hi - T_Low) / T_Low) < 0.02 Then                          '<= On error test to see if t_hi and t_low are close then deliver result if test passes.
            outputArray(0) = T_New - 273.15
            If Guess = 0 Then
                outputArray(1) = InitialTemps(1) - 273.15
            Else
                outputArray(1) = Guess
            End If
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                outputArray(i + 2) = xi_Dew(i)
            Next i
            
            If CallFromFlash = True Then
                DewT = outputArray
            Else
                DewT = outputArray
                DewT = Application.Transpose(DewT)
                Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
            End If
            Application.ScreenUpdating = True
            Exit Function
        Else
            If dataset(0, iColumns.globalErrmsg) & myErrorMsg = "" Then
                myErrorMsg = "Convergence failed."
            End If
        End If
    End If
        
    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(4)
         For i = 0 To UBound(outputArray, 1) - 2
            outputArray(i + 2) = 0
        Next i
    Else
         For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 2) = 0
        Next i
    End If
    
    outputArray(0) = -273.15

    If Guess = 0 Then
        outputArray(1) = InitialTemps(1) - 273.15
    Else
        outputArray(1) = Guess
    End If
    

    
    If CallFromFlash = True Then
        DewT = outputArray
    Else
        DewT = outputArray
        DewT = Application.Transpose(DewT)
    End If
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function
    Private Function checkInputPhase(dataset As Variant, vaporOrLiquid As Variant) As String
    
    '***************************************************************************
    'This function checks for user input errors
    '***************************************************************************
    
    Dim myErrorMsg As String
    Dim testString As String
    Dim fcnName As String
    
    fcnName = "checkInputPhase"
    
    
    
    If TypeName(vaporOrLiquid) = "Range" Then
        If vaporOrLiquid.Rows.Count <> 1 Or vaporOrLiquid.Columns.Count <> 1 Then
            myErrorMsg = "The supplied Phase should be a string or a single cell reference to a string equal to 'Vapor' or 'Liquid'."
            GoTo myErrorHandler
        End If
    End If
    
    testString = CStr(vaporOrLiquid)
    
    If LCase(testString) <> "vapor" And LCase(testString) <> "liquid" Then
                myErrorMsg = "the supplied phase is not a string value of 'Vapor' or 'Liquid'."
            GoTo myErrorHandler
    End If
    
    If LCase(testString) = "vapor" Then
        checkInputPhase = "vapor"
    End If
    
    If LCase(testString) = "liquid" Then
        checkInputPhase = "liquid"
    End If
    
    Exit Function
        
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    checkInputPhase = myErrorMsg
    
    End Function

    Private Function checkInputPressure(pressure As Variant, Optional dataset As Variant) As Double
    
    '***************************************************************************
    'This function checks for user input errors
    '***************************************************************************
    
    Dim myErrorMsg As String
    Dim fcnName As String
    
    fcnName = "checkInputPressure"
    
    If TypeName(pressure) = "Range" Then
        If pressure.Rows.Count <> 1 Or pressure.Columns.Count <> 1 Then
            myErrorMsg = "The Supplied pressure should be a number or a reference to a single cell."
            GoTo myErrorHandler
        End If
    End If
        
    If IsNumeric(pressure) = False Then
        myErrorMsg = "The supplied pressure is not numeric"
        GoTo myErrorHandler
    End If
    
    If pressure <= 0 Then
        myErrorMsg = "The supplied pressure cannot be less than or equal to zero."
        GoTo myErrorHandler
    End If
    
    checkInputPressure = CDbl(pressure)
    
    Exit Function

myErrorHandler:

    If IsMissing(dataset) = False Then
        dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    End If
    
    checkInputPressure = -1
    
    End Function
    

    Public Function BubbleT(DataRange As Range, pressure As Variant, moles As Variant, BinariesUsed As Boolean, _
                                    Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional Guess As Double = 0, _
                                     Optional errMsgsOn As Boolean = False, Optional CallFromFlash As Boolean = False) As Variant()
                                     
    '***************************************************************************
    'This function calculates the PR1978 EOS bubble point.
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    Dim UsingWilson As Boolean
    
    Dim Iter1 As Integer
    Dim Iter2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim HiLoTemp_Count As Integer
    Dim Upper_yi_Bub_Count As Integer
    Dim TBub_Count As Integer
    Dim Lower_yi_Bub_Count As Integer
    
    Dim ConvergenceSum As Double
    Dim yi_Sum As Double
    Dim yi_Low_Sum As Double
    Dim yi_Hi_Sum As Double
    Dim yi_New_Sum As Double
    Dim T_Low As Double
    Dim T_Hi As Double
    Dim T_New As Double
    Dim T_Bub_Est As Double
    Dim yi_Initial_Sum As Double
    Dim Initial_Psi As Double
    Dim TempK As Double
    Dim TempC As Double
    Dim T_Bub As Double
    Dim T_Bub_C As Double
    Dim PsiTest As Double
    Dim LiquidFugacity As Double
    Dim FugacityTest As Double
    Dim pbara As Double
    
    Dim Ki() As Double
    Dim New_Ki() As Double
    Dim xi_Array() As Double
    Dim yi_Array() As Double
    Dim InitialTemps() As Double
    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim Ki_Old() As Double
    Dim Ki_PR() As Double
    Dim yi_Bub() As Double
    Dim yi_Old() As Double
    Dim yi_Temp() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double

    Dim Phi_Vap() As Variant
    Dim Phi_Liq() As Variant
    Dim dataset() As Variant
    Dim outputArray() As Variant
    
    Dim Saturated As String
    Dim myErrorMsg As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim datasetErrMsgsOn As Boolean
    Dim yiBub_Equals_yiOld As Boolean
    Dim HiAndLoTempsFound As Boolean
    Dim BubT_Found As Boolean
    Dim decompArray() As Double
    Dim passedTempK As Double
    
    fcnName = "BubbleT"
    myErrorMsg = ""
    
    ReDim InitialTemps(1)
    
    InitialTemps(0) = 0
    InitialTemps(1) = 0
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "Vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If

    ReDim Phi_Vap(dataset(0, iColumns.iSpecies) + 2)
    ReDim Phi_Liq(dataset(0, iColumns.iSpecies) + 2)
    ReDim Ki_Old(dataset(0, iColumns.iSpecies))
    ReDim Ki(dataset(0, iColumns.iSpecies))
    ReDim Ki_PR(dataset(0, iColumns.iSpecies))
    ReDim yi_Bub(dataset(0, iColumns.iSpecies))
    ReDim yi_Old(dataset(0, iColumns.iSpecies))
    ReDim yi_Temp(dataset(0, iColumns.iSpecies))
    ReDim outputArray(3 + dataset(0, iColumns.iSpecies))
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If

    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If
    
    UsingWilson = False
    
    Upper_yi_Bub_Count = 0
    TBub_Count = 0
    Lower_yi_Bub_Count = 0
    HiLoTemp_Count = 0

    If Guess <> 0 Then
        T_Bub = Guess + 273.15
    Else
        InitialTemps = calculate_T_BubDew_Est(dataset, moleComp, pbara)
    
        If InitialTemps(0) <> 0 Then
            T_Bub = InitialTemps(0) 'Index 0 = Bubble Point and Index 1 = Dew Point
        Else
            T_Bub = 0
            myErrorMsg = "Calculate_T_BubDew_Est returned zero for initial bubble T estimate."
            GoTo myErrorHandler
        End If
    End If
    
    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
        
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = T_Bub
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If
    
    yi_Initial_Sum = 0
    
    For i = 0 To dataset(0, iColumns.iSpecies)
       Ki_Old(i) = dataset(i, iColumns.pc) ^ ((1 / T_Bub - 1 / dataset(i, iColumns.tb)) / _
       (1 / dataset(i, iColumns.tc) - 1 / dataset(i, iColumns.tb))) / pbara

       yi_Bub(i) = moleComp(i) * Ki_Old(i)
       
       yi_Initial_Sum = yi_Initial_Sum + yi_Bub(i)
    Next i
    
    For i = 0 To dataset(0, iColumns.iSpecies)
            yi_Old(i) = yi_Bub(i) / yi_Initial_Sum
    Next i
        
    Do Until HiAndLoTempsFound = True
     
        HiLoTemp_Count = HiLoTemp_Count + 1
        
        If HiLoTemp_Count > 2000 Then
            myErrorMsg = "HiLoTemp_Count  > 2000."
            GoTo myErrorHandler
        End If
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            yi_Bub(i) = moleComp(i) * Ki_Old(i)
        Next i
        
        Do Until yiBub_Equals_yiOld = True

            Upper_yi_Bub_Count = Upper_yi_Bub_Count + 1
            
            If Upper_yi_Bub_Count > 2000 Then
                myErrorMsg = "Upper_yi_Bub_Count > 2000."
                GoTo myErrorHandler
            End If
            yi_Old() = yi_Bub()
            
            yi_Sum = 0
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Sum = yi_Sum + yi_Bub(i)
            Next i
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = yi_Old(i) / yi_Sum
            Next i
                 
            Phi_Liq = calculate_Phi("Liquid", dataset, moleComp, T_Bub, pbara, BinariesUsed, kij0_Array, kijT_Array)
            
            If Phi_Liq(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Liquid phase Phi error in upper inner loop."
                GoTo myErrorHandler
            End If
        
            Phi_Vap = calculate_Phi("Vapor", dataset, yi_Bub, T_Bub, pbara, BinariesUsed, kij0_Array, kijT_Array)
            
            If Phi_Vap(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                    myErrorMsg = "Vapor vapor Phi error in upper inner loop."
                GoTo myErrorHandler
            End If
 
            For i = 0 To dataset(0, iColumns.iSpecies)
                If Phi_Vap(i) > 10 ^ -35 Then
                    Ki_PR(i) = Phi_Liq(i) / Phi_Vap(i)
                Else
                    Ki_PR(i) = 10 ^ 35
                End If
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = moleComp(i) * Ki_PR(i)
            Next i
                
            yi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Sum = yi_Sum + yi_Bub(i)
            Next i
            
            yi_Temp() = yi_Bub()
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = yi_Temp(i) / yi_Sum
            Next i
            
            ConvergenceSum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If yi_Old(i) <> 0 Then
                    ConvergenceSum = ConvergenceSum + Abs((yi_Bub(i) / yi_Old(i)) - 1)
                End If
            Next i
            
            If ConvergenceSum < 10 ^ -5 Then
                yiBub_Equals_yiOld = True
            End If
        Loop    '<= HiAndLoTempsFound Loop
        yiBub_Equals_yiOld = False
        
        If HiLoTemp_Count = 1 Then
             yi_Low_Sum = yi_Sum - 1#
             yi_Hi_Sum = yi_Sum - 1#
             T_Low = T_Bub
             T_Hi = T_Bub
        End If
       
        If yi_Sum > 1 Then
            T_Hi = T_Bub
            yi_Low_Sum = yi_Sum - 1#
            T_Bub = T_Bub / 1.01
        Else
            T_Low = T_Bub
            yi_Hi_Sum = yi_Sum - 1#
            T_Bub = T_Bub * 1.0101
        End If
       
        If yi_Low_Sum * yi_Hi_Sum > 0 Or yi_Hi_Sum = 0 Or yi_Low_Sum = 0 Then
        Else
            HiAndLoTempsFound = True
        End If
    Loop   '<= yiBub_Equals_yiOld

    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
                BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
                myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
        
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If

            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = (T_Hi + T_Low) / 2
            alpha_aiArray = create_alphaiArray(dataset, passedTempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If

    Do Until BubT_Found = True
    
        TBub_Count = TBub_Count + 1
        
        If TBub_Count > 2000 Then
            myErrorMsg = "TBub_Count > 2000."
            GoTo myErrorHandler
        End If
        
        T_New = (T_Hi + T_Low) / 2

        yiBub_Equals_yiOld = False
        
        Do Until yiBub_Equals_yiOld = True
        
            Lower_yi_Bub_Count = Lower_yi_Bub_Count + 1
            
            If Lower_yi_Bub_Count > 2000 Then
                myErrorMsg = "Lower_yi_Bub_Count > 2000."
                GoTo myErrorHandler
            End If
            
            yi_Old() = yi_Bub()
            
            yi_Sum = 0
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Sum = yi_Sum + yi_Bub(i)
            Next i
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = yi_Old(i) / yi_Sum
            Next i
        
            yi_Old() = yi_Bub()
            yi_Sum = 0
               
            Phi_Liq = calculate_Phi("Liquid", dataset, moleComp, T_New, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Liq(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = " Liquid phase Phi error in lower inner loop."
                GoTo myErrorHandler
            End If
            
            Phi_Vap = calculate_Phi("Vapor", dataset, yi_Bub, T_New, pbara, BinariesUsed, kij0_Array, kijT_Array)
            If Phi_Vap(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                myErrorMsg = "Vapor phase Phi error in lower inner loop."
                GoTo myErrorHandler
            End If
            
            If Abs(Phi_Vap(dataset(0, iColumns.iSpecies) + 1) - Phi_Liq(dataset(0, iColumns.iSpecies) + 1)) < 10 ^ -5 Then
                myErrorMsg = "Trival solution convergence in lower inner loop."
                GoTo myErrorHandler
            End If
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If Phi_Vap(i) > 10 - 35 Then
                    Ki_PR(i) = Phi_Liq(i) / Phi_Vap(i)
                Else
                    Ki_PR(i) = 10 ^ 35
                End If
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = moleComp(i) * Ki_PR(i)
            Next i
                
            yi_Sum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Sum = yi_Sum + yi_Bub(i)
            Next i
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                yi_Bub(i) = yi_Bub(i) / yi_Sum
            Next i
            
            ConvergenceSum = 0
            
            For i = 0 To dataset(0, iColumns.iSpecies)
                If yi_Old(i) <> 0 Then
                    ConvergenceSum = ConvergenceSum + Abs((yi_Bub(i) / yi_Old(i)) - 1)
                End If
            Next i
            
            If ConvergenceSum < 10 ^ -10 Then
                yiBub_Equals_yiOld = True
            End If
        Loop  '<= yiBub_Equals_yiOld Loop
        yi_New_Sum = yi_Sum - 1
        
        If yi_Low_Sum * yi_New_Sum > 0 Then
            yi_Low_Sum = yi_New_Sum
            T_Hi = T_New
        Else
            yi_Hi_Sum = yi_New_Sum
            T_Low = T_New
        End If
        
        If Abs(T_Hi - T_Low) < 0.01 Then
            BubT_Found = True
        End If
    Loop  '<= BubT_Found Loop

    FugacityTest = 0
     
    For i = 0 To dataset(0, iColumns.iSpecies)
        FugacityTest = FugacityTest + Abs(Phi_Liq(i) * moleComp(i) - yi_Bub(i) * Phi_Vap(i))
     Next i

    If FugacityTest > 10 ^ -2 Then                                                      '<= Not holding a tolerance for fugacity test at bubble point.
        myErrorMsg = "Equlibrium test failed. Bubble temperature result may be inacurrate."
    End If
    
    outputArray(0) = T_New - 273.15
    
    If Guess = 0 Then
        outputArray(1) = InitialTemps(0) - 273.15
    Else
        outputArray(1) = Guess
    End If
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i + 2) = yi_Bub(i)
    Next i
    
    If CallFromFlash = True Then
        BubbleT = outputArray
    Else
        BubbleT = outputArray
        BubbleT = Application.Transpose(BubbleT)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)   '<= Used for warning messages
        
    Application.ScreenUpdating = True
        
    Exit Function
    
myErrorHandler:

    If Err.Number = 0 And T_Hi <> 0 And T_Low <> 0 Then
         
        If Abs((T_Hi - T_Low) / T_Low) < 0.02 And InitialTemps(0) > T_Low And InitialTemps(1) < T_Hi Then                                           '<= On error test to see if t_hi and t_low are close then deliver result if test passes.
            outputArray(0) = T_New - 273.15
        
            If Guess = 0 Then
                outputArray(1) = InitialTemps(0) - 273.15
            Else
                outputArray(1) = Guess
            End If
        
            For i = 0 To dataset(0, iColumns.iSpecies)
                outputArray(i + 2) = yi_Bub(i)
            Next i
            If CallFromFlash = True Then
                BubbleT = outputArray
            Else
                BubbleT = outputArray
                BubbleT = Application.Transpose(BubbleT)
                Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
            End If
            Application.ScreenUpdating = True
            Exit Function
        Else
            If dataset(0, iColumns.globalErrmsg) & myErrorMsg = "" Then
                myErrorMsg = "Convergence failed."
            End If
        End If
    End If
        
    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(4)
    End If
        
    outputArray(0) = -273.15

    If Guess = 0 Then
        outputArray(1) = InitialTemps(1) - 273.15
    Else
        outputArray(1) = Guess
    End If
    
    For i = 0 To UBound(outputArray) - 2
        outputArray(i + 2) = 0
    Next i
    
    If CallFromFlash = True Then
        BubbleT = outputArray
    Else
        BubbleT = outputArray
        BubbleT = Application.Transpose(BubbleT)
    End If
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function

    Public Function FlashTP(DataRange As Range, temperature As Variant, pressure As Variant, moles As Variant, _
                                BinariesUsed As Boolean, Optional kij0 As Variant, Optional kijT As Variant, Optional deComp As Variant, Optional errMsgsOn As Boolean = False) As Variant()
                                
    '***************************************************************************
    'This function calculates the PR1978 EOS vapor fraction, vapor composition
    'and liquid composition for a constant T & P flash.
    '***************************************************************************
                                
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    '*****************************
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    
    'Remember time when macro starts
      'StartTime = Timer                                                             '<= Timer! *********************************
    'Insert Your Code Here...
    '*****************************
    
    On Error GoTo myErrorHandler:
    
    Dim UsingDewOrBubbleFunction As Boolean

    Dim CounterLimit1 As Integer
    Dim CounterLimit2 As Integer
    Dim Counter1 As Integer
    Dim Counter2 As Integer
    Dim i As Integer
    Dim j As Integer

    Dim Initial_Psi As Double
    Dim SumRedfordRiceEq As Double          '<= Rachford Rice
    Dim SumRedfordRiceEqPrime As Double     '<= derivitive of Rachford Rice
    Dim LiquidFugacity As Double
    Dim VaporFugacity As Double
    Dim FugacityCheck As Double
    Dim SUMx As Double
    Dim SUMy As Double
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim passedTempK As Double

    Dim moleComp() As Double
    Dim aiArray() As Double
    Dim alpha_aiArray() As Double
    Dim kij0_Array() As Double
    Dim kijT_Array() As Double
    Dim xi_Array() As Double
    Dim yi_Array() As Double
    Dim Ki() As Double
    
    
    Dim Psi_New As Double
    Dim Psi_Old As Double
    Dim dewTemp As Variant
    Dim bubTemp As Variant
    
    Dim Phi_Vap() As Variant
    Dim Phi_Liq() As Variant
    Dim DewBubKi() As Variant
    Dim outputArray As Variant
    Dim dataset() As Variant
    
    Dim StreamCondition As String
    Dim myErrorMsg As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    
    Dim datasetErrMsgsOn As Boolean
    Dim VaporFractionFound As Boolean
    Dim EquilibriumFound As Boolean

    
    fcnName = "FlashTP"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    myErrorMsg = ""
    
    UsingDewOrBubbleFunction = False
    dewTemp = -273.15
    bubTemp = -273.15
    
    If IsMissing(moles) = False Then
        moleComp = validateMoles(dataset, moles)
    Else
        myErrorMsg = "No moles provided."
        GoTo myErrorHandler
    End If
    
    If moleComp(0) = -1 Then
        GoTo myErrorHandler
    End If

    If dataset(0, iColumns.predictive) = 0 Then                                             '<predictive
        If BinariesUsed = True And IsMissing(kij0) = True Then
        
        BinariesUsed = False   '<must have Kij0 range at minimum to calculate binaries
            myErrorMsg = "BinariesUsed = True but no binaries are provided. BinariesUsed changed to False."
        End If
        
        kij0_Array = createKij0Array(dataset, BinariesUsed, kij0)
        kijT_Array = createKijTArray(dataset, kijT)
    Else                                                                                        '<predictive
        If dataset(0, iColumns.predictive) = 1 And IsMissing(deComp) = True Then
            If BinariesUsed = True Then
                myErrorMsg = "Predictive PR1978 is specified but no decomposition is provided. Calculating without binaries!"
            End If
            
            BinariesUsed = False
            dataset(0, iColumns.predictive) = 0
            kij0_Array = createKij0Array(dataset, BinariesUsed)
            kijT_Array = createKijTArray(dataset)
        Else
            If BinariesUsed = False Then
                dataset(0, iColumns.predictive) = 0
                kij0_Array = createKij0Array(dataset, BinariesUsed)
                kijT_Array = createKijTArray(dataset)
            End If
            passedTempK = TempK
            alpha_aiArray = create_alphaiArray(dataset, TempK)
            aiArray = create_aiArray(dataset, alpha_aiArray)
            kij0_Array = createPredictivekijTArray(dataset, deComp, passedTempK, aiArray)
            kijT_Array = createKijTArray(dataset)
        End If
    End If

    ReDim xi_Array(dataset(0, iColumns.iSpecies))
    ReDim yi_Array(dataset(0, iColumns.iSpecies))
    ReDim Phi_Vap(dataset(0, iColumns.iSpecies) + 2)
    ReDim Phi_Liq(dataset(0, iColumns.iSpecies) + 2)
    ReDim DewBubKi(dataset(0, iColumns.iSpecies) + 2)
    
    CounterLimit1 = 2000
    CounterLimit2 = 2000
    
    ReDim outputArray(2 + 2 * (dataset(0, iColumns.iSpecies) + 1))
    
    Ki() = calculate_Wilson_K(dataset, TempK, pbara)                    '<Initialize Ki values with Wilson method

    Initial_Psi = 0.5
    Psi_New = Initial_Psi
    
        Do Until EquilibriumFound = True
            Do Until VaporFractionFound = True

                    SumRedfordRiceEq = 0
                    SumRedfordRiceEqPrime = 0

                    For i = 0 To dataset(0, iColumns.iSpecies)
                        If moleComp(i) > 0 Then
                            SumRedfordRiceEq = SumRedfordRiceEq + moleComp(i) * (Ki(i) - 1) / (Psi_New * Ki(i) + 1 - Psi_New)
                            SumRedfordRiceEqPrime = SumRedfordRiceEqPrime + moleComp(i) * (Ki(i) - 1) ^ 2 / (Psi_New * (Ki(i) - 1) + 1) ^ 2
                        End If
                    Next i

                    Psi_Old = Psi_New
                    Psi_New = Psi_New + SumRedfordRiceEq / SumRedfordRiceEqPrime

                    If Psi_New < 0 Then
                        Psi_New = Psi_Old / 2
                    End If

                    If Psi_New > 1 Then
                        Psi_New = (Psi_Old + 1) / 2
                    End If

                    If Psi_New < 10 ^ -8 And Psi_New >= 0 Then
                        If UsingDewOrBubbleFunction = False Then
                            UsingDewOrBubbleFunction = True
                            DewBubKi = BubbleT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)    '<= Get Ki near Bubble point temperature then try again.
                            If TempC <= DewBubKi(0) Or DewBubKi(0) = -273.15 Then           'DewBubKi(0) = bubble temperature, DewBubKi(1)...DewBubKi(n) are Ki's
                                StreamCondition = "At Or Below Bubble Point"
                                If DewBubKi(0) = -273.15 Then
                                    myErrorMsg = "FlashTP bubble point test failed. Vapor fraction does not appear to exist."
                                End If
                                VaporFractionFound = True
                                EquilibriumFound = True
                            End If
                            If StreamCondition = "" Then
                                For i = 0 To dataset(0, iColumns.iSpecies)                          '<Flash with Wilson Ki initialization failed to converge so retry with Ki's from Dew or Bubble point calculation.
                                    If moleComp(i) > 10 ^ -35 Then
                                        Ki(i) = moleComp(i) / DewBubKi(i + 2)
                                    Else
                                        Ki(i) = 10 ^ 35                                             '<= This prevent overflow errors.
                                    End If
                                Next i
                            End If
                        Else
                            StreamCondition = "At Or Below Bubble Point"
                            VaporFractionFound = True
                            EquilibriumFound = True
                        End If
                    End If

                    If 1 - Psi_New < 10 ^ -8 And Psi_New <= 1 Then
                        If UsingDewOrBubbleFunction = False Then
                            UsingDewOrBubbleFunction = True
                            DewBubKi = DewT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)  '<= Get Ki near Dew point temperature then try again.
                            If TempC >= DewBubKi(0) Or DewBubKi(0) = -273.15 Then                           'DewBubKi(0) = dew temperature, DewBubKi(1)...DewBubKi(n) are Ki's
                                StreamCondition = "At or Above Dew Point"
                                If DewBubKi(0) = -273.15 Then
                                    myErrorMsg = "FlashTP dew point test failed. Liquid phase does not appear to exist."
                                End If
                                VaporFractionFound = True
                                EquilibriumFound = True
                            End If
                            If StreamCondition = "" Then
                                For i = 0 To dataset(0, iColumns.iSpecies)                          '<Flash with Wilson Ki initialization failed to converge so retry with Ki's from Dew or Bubble point calculation.
                                    If moleComp(i) > 10 ^ -35 Then
                                        Ki(i) = moleComp(i) / DewBubKi(i + 2)
                                    Else
                                        Ki(i) = 10 ^ 35                                             '<= This prevent overflow errors.
                                    End If
                                Next i
                            End If
                        Else
                            StreamCondition = "At or Above Dew Point"
                            VaporFractionFound = True
                            EquilibriumFound = True
                        End If
                    End If

                    If Abs(1 - Abs(Psi_New / Psi_Old)) < 10 ^ -8 And StreamCondition = "" Then
                        VaporFractionFound = True
                    End If
                    
                    Counter1 = Counter1 + 1

                    If Counter1 + 1 > CounterLimit1 Then
                        myErrorMsg = "Flash failed to converge. Too many iterations in Psi calculation."
                        GoTo myErrorHandler
                    End If
            Loop
        
            VaporFractionFound = False
            
            If StreamCondition = "" Then

                SUMx = 0
                SUMy = 0

                For i = 0 To dataset(0, iColumns.iSpecies)
                    If moleComp(i) > 0 Then

                        On Error Resume Next

                        xi_Array(i) = moleComp(i) / (Psi_New * Ki(i) + 1 - Psi_New)                         '<= if xi or yi calculate to below zero then need to set xi or yi to zero

                        If Err.Number = 6 Then                                                              ' then renormalize the other to accept all of the component in the non-zero phase                                                                                  ' need to add this code her
                            xi_Array(i) = 0                                                                 '<=This prevents overflow error if number in denomenator is very small
                            Err.Clear
                        Else
                            If Err.Number <> 0 Then
                                GoTo myErrorHandler
                            End If
                        End If

                        yi_Array(i) = moleComp(i) * Ki(i) / (Psi_New * Ki(i) + 1 - Psi_New)

                        If Err.Number = 6 Then
                            yi_Array(i) = 0
                            Err.Clear
                        Else
                            If Err.Number <> 0 Then
                                GoTo myErrorHandler
                            End If
                        End If

                        On Error GoTo myErrorHandler

                        SUMx = SUMx + xi_Array(i)
                        SUMy = SUMy + yi_Array(i)
                    End If
                Next i

                For i = 0 To dataset(0, iColumns.iSpecies)
                    If moleComp(i) > 0 Then
                        xi_Array(i) = xi_Array(i) / SUMx
                        yi_Array(i) = yi_Array(i) / SUMy
                    End If
                Next i

                LiquidFugacity = 0
                VaporFugacity = 0
                FugacityCheck = 0

                Phi_Vap = calculate_Phi("Vapor", dataset, yi_Array, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
                
                If Phi_Vap(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                    If UsingDewOrBubbleFunction = False Then
                        dewTemp = DewT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)
                        bubTemp = BubbleT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)

                        If dewTemp(0) <> -273.15 Then
                           If TempC >= dewTemp Then
                                StreamCondition = "At or Above Dew Point"
                                EquilibriumFound = True
                            End If
                        End If

                        If bubTemp(0) <> -273.15 Then
                            If TempC <= bubTemp Then
                                StreamCondition = "At Or Below Bubble Point"
                                EquilibriumFound = True
                            End If
                        End If

                        If StreamCondition = "" Then
                            myErrorMsg = "Flash failed to converge. Error returned from calculate_Phi."
                            GoTo myErrorHandler
                        End If
                    End If
                End If

                Phi_Liq = calculate_Phi("Liquid", dataset, xi_Array, TempK, pbara, BinariesUsed, kij0_Array, kijT_Array)
                
                If Phi_Liq(dataset(0, iColumns.iSpecies) + 2) = -500 Then
                    If UsingDewOrBubbleFunction = False Then
                        dewTemp = DewT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)
                        bubTemp = BubbleT(DataRange, pbara, moles, BinariesUsed, kij0, kijT, deComp, , , 1)
                        
                        If dewTemp(0) <> -273.15 Then
                            If TempC >= dewTemp Then
                                StreamCondition = "At or Above Dew Point"
                                EquilibriumFound = True
                            End If
                        End If
    
                        If bubTemp(0) = -273.15 Then
                            StreamCondition = "At Or Below Bubble Point"
                            EquilibriumFound = True
                        End If
    
                        If bubTemp(0) <> -273.15 Then
                            If TempC <= bubTemp Then
                                StreamCondition = "At Or Below Bubble Point"
                                EquilibriumFound = True
                            End If
                        End If
                    End If
                    
                    If StreamCondition = "" Then
                        myErrorMsg = "Flash failed to converge. An error returned from calculate_Phi."
                        GoTo myErrorHandler
                    End If
                End If

                If StreamCondition = "" Then
                    For i = 0 To dataset(0, iColumns.iSpecies)
                        If moleComp(i) > 0 Then
    
                            LiquidFugacity = LiquidFugacity + pbara * xi_Array(i) * Phi_Liq(i)
    
                            VaporFugacity = VaporFugacity + pbara * yi_Array(i) * Phi_Vap(i)
    
                            If (pbara * yi_Array(i) * Phi_Vap(i)) > 10 ^ -35 Then  '<= got errors in very low temp vapor phase calculations until decreased this from 10^-12 to 10^-35
                                Ki(i) = Ki(i) * pbara * xi_Array(i) * Phi_Liq(i) / (pbara * yi_Array(i) * Phi_Vap(i))
                            Else
                                Ki(i) = 10 ^ 35
                            End If
                        End If
                    Next i

                    FugacityCheck = Abs(1 - LiquidFugacity / VaporFugacity)
    
                    If Abs(FugacityCheck) < 10 ^ -12 Then
                    
                        If Psi_New < 10 ^ -7 Then
                            StreamCondition = "At Or Below Bubble Point"
                            EquilibriumFound = True
                        End If
                        
                        If 1 - Psi_New < 10 ^ -7 And Psi_New <= 1 Then
                            StreamCondition = "At or Above Dew Point"
                            EquilibriumFound = True
                        End If
                        
                        If StreamCondition = "" Then
                            StreamCondition = "Vapor and liquid phases exist."
                            EquilibriumFound = True
                        End If
                    End If
                End If
            End If
            
            Counter2 = Counter2 + 1

            If Counter2 + 1 > CounterLimit2 Then
                myErrorMsg = "Flash failed to converge. Too many iterations in main flash loop."
                GoTo myErrorHandler
            End If
        Loop
        
    If StreamCondition = "" Then
        If myErrorMsg = "" Then
            myErrorMsg = "FlashTP failed to converge'"
        End If
        GoTo myErrorHandler
    End If
    
    If StreamCondition = "Vapor and liquid phases exist." Then

        outputArray(0) = Psi_New
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 1) = yi_Array(i)
        Next i
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(dataset(0, iColumns.iSpecies) + 2 + i) = xi_Array(i)
        Next i
    End If
    
    If StreamCondition = "At Or Below Bubble Point" Then
    
        outputArray(0) = 0
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 1) = 0
        Next i
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 2 + dataset(0, iColumns.iSpecies)) = moleComp(i)
        Next i
    
    End If
    
    If StreamCondition = "At or Above Dew Point" Then
        outputArray(0) = 1
    
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 1) = moleComp(i)
        Next i
        
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i + 2 + dataset(0, iColumns.iSpecies)) = 0
        Next i
    End If
    

    
        FlashTP = Application.Transpose(outputArray)

    '*********************************************************
    'Determine how many seconds code took to run
      'SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    'Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds" ', vbInformation                     '<= Timer! *********************************
    '*********************************************************
    
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)            '<=Used for warnings and to clear comments when errors are eliminated.
    
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True



    Exit Function
    
myErrorHandler:

    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(5)
        datasetErrMsgsOn = True
    Else
        ReDim outputArray(2 + 2 * (dataset(0, iColumns.iSpecies) + 1))
    End If
    
    'Returning all zero's if flash fails

    For i = 0 To UBound(outputArray)
        outputArray(i) = 0
    Next i

    FlashTP = Application.Transpose(outputArray)
    
    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
    End Function


    Private Function calculate_Wilson_K(dataset As Variant, TempK As Double, pbara As Double) As Double()
    
    '***************************************************************************
    'This function is called from the FlashTP function.
    'This function calculates the Wilson relative volitility
    '***************************************************************************
    
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    ' This function uses the Wilson equation to estimate the Ki values given temperature in degrees C
    'and the pressure in bara. It returns a vertical one dimensional array.

    Dim i As Integer
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim outputArray() As Double
    Dim myErrorMsg As String
    Dim fcnName As String

    fcnName = "calculate_Wilson_K"
   
    ReDim outputArray(0 To dataset(0, iColumns.iSpecies) + 1)
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = Exp(5.37 * (1 + dataset(i, iColumns.omega)) * (1 - dataset(i, iColumns.tc) / TempK)) / (pbara / dataset(i, iColumns.pc))
    Next i
        
    calculate_Wilson_K = outputArray()
    
    Application.ScreenUpdating = True
    
    Exit Function

myErrorHandler:

    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = 0
    Next i
    
    outputArray(dataset(0, iColumns.iSpecies) + 1) = -500                               '<=error flag
    
    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg
    
    Application.ScreenUpdating = True
    
    End Function
    Public Function Wilson_K(DataRange As Range, temperature As Variant, pressure As Variant, Optional errMsgsOn As Boolean = False) As Variant()
    
    '***************************************************************************
    'This function calculates the Wilson relative volitility
    '***************************************************************************
    
    On Error GoTo myErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim dataset() As Variant
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim TempK As Double
    Dim TempC As Double
    Dim pbara As Double
    Dim dblsArray() As Double
    Dim outputArray() As Variant
    Dim myErrorMsg As String
    Dim Phase As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    
    Dim datasetErrMsgsOn As Boolean
    
    

    fcnName = "Wilson_K"
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, Phase, False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    TempC = checkInputTemperature(temperature, dataset)
    If TempC = -273.15 Then
        GoTo myErrorHandler
    End If
    TempK = TempC + 273.15
    
    pbara = checkInputPressure(pressure, dataset)
    If pbara = -1 Then
        GoTo myErrorHandler
    End If
    
    dblsArray = calculate_Wilson_K(dataset, TempK, pbara)
    
    outputArray = convertDoubleToVariant(dblsArray)
    
    If outputArray(dataset(0, iColumns.iSpecies) + 1) = -500 Then
        myErrorMsg = "calculate_Wilson_K returned an error"
        GoTo myErrorHandler
    End If
    
    Wilson_K = outputArray
    
    Wilson_K = Application.Transpose(Wilson_K)
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(4)
        For i = 0 To UBound(outputArray, 1)
            outputArray(i) = 0
        Next i
    Else
        ReDim outputArray(dataset(0, iColumns.iSpecies))
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i) = 0
        Next i
    End If
    
    Wilson_K = outputArray
    Wilson_K = Application.Transpose(outputArray)
    
    Application.ScreenUpdating = True
    
    End Function
    Private Function convertDoubleToVariant(oneDimArray() As Double) As Variant()
    
    '***************************************************************************
    'This function converts and array of doubles to an array of variants
    '***************************************************************************
    
    Dim outputArray() As Variant
    Dim i As Integer
    
    ReDim outputArray(UBound(oneDimArray, 1))
    
    For i = 0 To UBound(oneDimArray, 1)
        If IsNumeric(oneDimArray(i)) = True Then
            outputArray(i) = oneDimArray(i)
        Else
            
        End If
    Next i
    
    convertDoubleToVariant = outputArray

    End Function

    Private Function calculate_IdealGasEntropy(dataset As Variant, moleComp() As Double, TempK As Double, pbara As Double, CpRanges() As Integer) As Double
    
    '***************************************************************************
    'This function calculates the ideal gas entropy
    '***************************************************************************
    
    On Error GoTo myErrorHandler:

    Dim p As Integer
    Dim g As Integer
    Dim m As Integer
    Dim i As Integer
    Dim Denominator As Double
    Dim SpeciesArray() As String
    Dim MolsArray() As Double
    Dim strCpData As String
    Dim CPData() As Double
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim TMN1 As Integer
    Dim TMX1 As Integer
    Dim TMX2 As Integer
    Dim TMX3 As Integer
    Dim TMX4 As Integer
    Dim TMX5 As Integer
    Dim TMX6 As Integer
    Dim CpDataType() As String
    Dim outputArray() As Double
    Dim H_H298_Temp As Double
    Dim molesArray() As Double
    Dim LiquidCpDataRequired As Variant
    Dim CpEquation As Double
    Dim local_Tempk As Double
    Dim myErrorMsg As String

    
    Dim fcnName As String

    fcnName = "calculate_IdealGasEntropy"
    
    calculate_IdealGasEntropy = 0
                
        For i = 0 To dataset(0, iColumns.iSpecies)
            If CpRanges(i, UBound(CpRanges, 2)) <> -500 Then
        
                If CpRanges(i, UBound(CpRanges, 2)) = -400 Then
                    local_Tempk = dataset(i, CpRanges(i, iColumns.TempK) - 1)
                Else
                    local_Tempk = TempK
                End If
                
                If dataset(i, iColumns.CpDataType) = "NIST" Then '<= 0 means we have "NIST" type data
                    Denominator = 1000
                Else
                    Denominator = 1
                End If
                        
                        CpEquation = 0
                                              
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK)) * Log(local_Tempk / Denominator)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 1) * (local_Tempk / Denominator)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 2) * ((local_Tempk / Denominator) ^ 2) / 2
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 3) * ((local_Tempk / Denominator) ^ 3) / 3 '<= This is the only one with a minus sign
                        CpEquation = CpEquation - dataset(i, CpRanges(i, iColumns.TempK) + 4) / (2 * (local_Tempk / Denominator) ^ 2)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 6)
                        calculate_IdealGasEntropy = calculate_IdealGasEntropy + moleComp(i) * CpEquation
                        
                        CpEquation = 0
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298)) * Log(298.15 / Denominator)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 1) * (298.15 / Denominator)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 2) * ((298.15 / Denominator) ^ 2) / 2
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 3) * ((298.15 / Denominator) ^ 3) / 3 '<= This is the only one with a minus sign
                        CpEquation = CpEquation - dataset(i, CpRanges(i, iColumns.Vap298) + 4) / (2 * (298.15 / Denominator) ^ 2)
                        CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 6)
                        calculate_IdealGasEntropy = calculate_IdealGasEntropy - moleComp(i) * CpEquation

            End If
        Next i
        
            calculate_IdealGasEntropy = calculate_IdealGasEntropy * (1000 / 1000)                       '<= convert Cp data from j/g-mole/K to kJ/kg-mole/K
            
            calculate_IdealGasEntropy = (calculate_IdealGasEntropy) - 100000 * GasLawR * Log(pbara / 1)
                                                                        
'    NIST Data (units for H & S are different)
'    Cp = heat capacity (J/mol*K)
'    H° = standard enthalpy (kJ/mol)
'    S° = standard entropy (J/mol*K)
'    t = temperature(k) / 1000

'   HSC Data
'   Cp j/mol/K (units are the same for H & S)
'   T = temperaqture(k)
                                                                        
            Exit Function
    
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    
    calculate_IdealGasEntropy = 987654321.123457 '<=error flag
    
    End Function
    Private Function calculate_Cp_IGorLiquid(dataset As Variant, moleComp() As Double, TempK As Double, CpRanges() As Integer, Phase As String) As Double
                
    '***************************************************************************
    'This function calculates the Cp for an ideal gas or a liquid
    '***************************************************************************
                
    On Error GoTo myErrorHandler:
    
    Dim i As Integer
    Dim Denominator As Double
    Dim J_to_kJ As Double
    Dim CpEquation As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim m As Integer
    Dim local_Tempk As Double

    fcnName = "calculate_Cp"
    
    If Phase = LCase("vapor") Then
        m = 0
    Else
        m = dataset(0, iColumns.iSpecies) + 1
    End If
        
    calculate_Cp_IGorLiquid = 0
          
    For i = 0 To dataset(0, iColumns.iSpecies)
        If CpRanges(i, UBound(CpRanges, 2)) <> -500 Then
            
                If CpRanges(i, UBound(CpRanges, 2)) = -400 Then
                    local_Tempk = dataset(i + dataset(0, iColumns.iSpecies) + 1, CpRanges(i, iColumns.TempK) - 1)
                Else
                    local_Tempk = TempK
                End If
    
            If dataset(i + m, iColumns.CpDataType) = "NIST" Then                               '<= This means we have "NIST" type(i.e. Shomate equations & t = t/1000) formatted data.
                Denominator = 1000
                J_to_kJ = 1
            Else                                                                             '<= Here we can test for other types of data in the future - not used for now
                Denominator = 1
                J_to_kJ = 1
            End If
            
            CpEquation = 0
        
            CpEquation = CpEquation + dataset(i + m, CpRanges(i, iColumns.TempK))
            CpEquation = CpEquation + dataset(i + m, CpRanges(i, iColumns.TempK) + 1) * (local_Tempk / Denominator)
            CpEquation = CpEquation + dataset(i + m, CpRanges(i, iColumns.TempK) + 2) * (local_Tempk / Denominator) ^ 2
            CpEquation = CpEquation + dataset(i + m, CpRanges(i, iColumns.TempK) + 3) * (local_Tempk / Denominator) ^ 3
            CpEquation = CpEquation + dataset(i + m, CpRanges(i, iColumns.TempK) + 4) / (local_Tempk / Denominator) ^ 2
            
            calculate_Cp_IGorLiquid = calculate_Cp_IGorLiquid + moleComp(i) * CpEquation / J_to_kJ  '<= Cp = heat capacity (J/mol*K)
        
        End If
    Next i
            
'   NIST Data
'    Cp = heat capacity (J/mol*K)
'    H° = standard enthalpy (kJ/mol)
'    S° = standard entropy (J/mol*K)
'    t = temperature(k) / 1000

'   HSC Data
'   Cp j/mol/K
'    T = temperaqture(k)
    
    Exit Function
    
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    calculate_Cp_IGorLiquid = 987654321.123457
    
    End Function
    Private Function calculate_IdealGasEnthalpy(dataset As Variant, moleComp() As Double, TempK As Double, CpRanges() As Integer) As Double
                
    '***************************************************************************
    'This function calculates the ideal gas enthalpy
    '***************************************************************************
                
    On Error GoTo myErrorHandler:
                
    Dim i As Integer
    Dim Denominator As Double
    Dim J_to_kJ As Double
    Dim CpEquation As Double
    Dim myErrorMsg As String
    Dim fcnName As String
    Dim local_Tempk As Double

    fcnName = "calculate_IdealGasEnthalpy"
       
    calculate_IdealGasEnthalpy = 0
                
    For i = 0 To dataset(0, iColumns.iSpecies)
        If CpRanges(i, UBound(CpRanges, 2)) <> -500 Then
        
            If CpRanges(i, UBound(CpRanges, 2)) = -400 Then
                local_Tempk = dataset(i, CpRanges(i, iColumns.TempK) - 1)
            Else
                local_Tempk = TempK
            End If
            
            If dataset(i, iColumns.CpDataType) = "NIST" Then                                 '<= This means we have "NIST" type(i.e. Shomate equations & t = t/1000) formatted data.
                Denominator = 1000
                J_to_kJ = 1  'NIST data for enthalpy is kJ/mol
            Else                                                                                '<= Here we can prepare for HSC type data
                Denominator = 1
                J_to_kJ = 1000      'HSC data is j/mol
            End If
            
            CpEquation = 0
        
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK)) * local_Tempk / Denominator
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 1) * (local_Tempk / Denominator) ^ 2 / 2
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 2) * (local_Tempk / Denominator) ^ 3 / 3
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 3) * (local_Tempk / Denominator) ^ 4 / 4
            CpEquation = CpEquation - dataset(i, CpRanges(i, iColumns.TempK) + 4) / (local_Tempk / Denominator)
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.TempK) + 5) - dataset(i, CpRanges(i, iColumns.TempK) + 7)
            
            calculate_IdealGasEnthalpy = calculate_IdealGasEnthalpy + moleComp(i) * CpEquation / J_to_kJ
            
            CpEquation = 0
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298)) * 298.15 / Denominator
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 1) * (298.15 / Denominator) ^ 2 / 2
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 2) * (298.15 / Denominator) ^ 3 / 3
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 3) * (298.15 / Denominator) ^ 4 / 4
            CpEquation = CpEquation - dataset(i, CpRanges(i, iColumns.Vap298) + 4) / (298.15 / Denominator)
            CpEquation = CpEquation + dataset(i, CpRanges(i, iColumns.Vap298) + 5) - dataset(i, CpRanges(i, iColumns.Vap298) + 7)
            
            calculate_IdealGasEnthalpy = calculate_IdealGasEnthalpy - moleComp(i) * CpEquation / J_to_kJ
        End If
    Next i
    

        calculate_IdealGasEnthalpy = calculate_IdealGasEnthalpy * 1000  '<convert from kJmols to kJ/kg-moles

'    NIST Data (units for H & S are different)
'    Cp = heat capacity (J/mol*K)
'    H° = standard enthalpy (kJ/mol)
'    S° = standard entropy (J/mol*K)
'    t = temperature(k) / 1000

'   HSC Data
'   Cp j/mol/K (units are the same for H & S)
'   T = temperaqture(k)
    
    Exit Function
    
myErrorHandler:

    dataset(0, iColumns.globalErrmsg) = fcnName & ": " & myErrorMsg

    
    calculate_IdealGasEnthalpy = 987654321.123457 '<=error flag
    
    End Function

    Public Function Hf(DataRange As Range, Optional errMsgsOn As Boolean = False) As Variant
    
    '***************************************************************************
    'This function returns the heat of formation at 298K and 1 bara for a dataset
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = False
    
    'returns standard (298K) heat of formation in kJ/gram-mol, 'Mols = grams moles
    
    Dim i As Integer
    Dim dataset() As Variant
    Dim outputArray() As Double
    Dim myErrorMsg As String
    
    Dim UDF_Range As Range
    Dim fcnName As String
    Dim datasetErrMsgsOn As Boolean

    fcnName = "Hf"
    myErrorMsg = ""
    
    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    ReDim outputArray(dataset(0, iColumns.iSpecies))
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i) = dataset(i, iColumns.iHf298)  '<kJ/kg-mole
    Next i
    
    Hf = outputArray
    Hf = Application.Transpose(Hf)
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
myErrorHandler:

    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(3)
        For i = 0 To 3
            outputArray(i) = 0
        Next i
    Else
        For i = 0 To dataset(0, iColumns.iSpecies)
            outputArray(i) = 0
        Next i
    End If
    
    Hf = outputArray
    Hf = Application.Transpose(Hf)
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function

    Public Function HVap1Bara(DataRange As Range, Optional errMsgsOn As Boolean = False) As Variant
    
    '***************************************************************************
    'This function returns the normal heat vaporization and the normal boiling point for a dataset
    '***************************************************************************
    
    On Error GoTo myErrorHandler:
    
    Application.ScreenUpdating = True
    
    'returns standard (298K) heat of formation in kJ/gram-mol, 'Mols = grams moles
    
    Dim i As Integer
    Dim dataset() As Variant
    Dim outputArray() As Variant
    Dim myErrorMsg As String
    Dim UDF_Range As Range
    Dim fcnName As String
    
    Dim datasetErrMsgsOn As Boolean

    fcnName = "HVap1Bara"
    myErrorMsg = ""

    Set UDF_Range = Application.Caller
    
    If IsMissing(DataRange) = False Then
        dataset = validateDataset(DataRange, "vapor", False)
    Else
        myErrorMsg = "No dataset provided."
        GoTo myErrorHandler
    End If
    
    If IsNumeric(dataset(0, 0)) = False Then
        myErrorMsg = dataset(0, 0)
        datasetErrMsgsOn = True
        GoTo myErrorHandler
    End If
    
    datasetErrMsgsOn = dataset(0, iColumns.errMsgsOn)
    
    ReDim outputArray(dataset(0, iColumns.iSpecies), 1)
      
    For i = 0 To dataset(0, iColumns.iSpecies)
        outputArray(i, 0) = dataset(i, iColumns.hvap)
        outputArray(i, 1) = dataset(i, iColumns.tb) - 273.15
    Next i
    
    HVap1Bara = outputArray
    
    Call errorSub(UDF_Range, fcnName & " Warning: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)             '<=Used for warnings and to clear comments when errors are eliminated.
    
    Application.ScreenUpdating = True
    
    Exit Function
    
myErrorHandler:

    If IsArrayAllocated(dataset) = False Or UBound(dataset, 1) = 0 Or IsArrayAllocated(outputArray) = False Then
        ReDim outputArray(3, 1)
        For i = 0 To 3
            outputArray(i, 0) = 0
            outputArray(i, 1) = 0
        Next i
    Else
    For i = 0 To dataset(0, iColumns.iSpecies)
         outputArray(i, 0) = 0
         outputArray(i, 1) = 0
     Next i
    End If

    HVap1Bara = outputArray
    
    If myErrorMsg = "" Then
        myErrorMsg = dataset(0, iColumns.globalErrmsg)
    End If

    Call errorSub(UDF_Range, fcnName & " Error: ", myErrorMsg, errMsgsOn, datasetErrMsgsOn)
    
    Application.ScreenUpdating = True
    
    End Function
