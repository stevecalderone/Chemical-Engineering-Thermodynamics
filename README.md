# ChemicalEngineering
User Defined Functions for multicomponent thermodynamic calculations of the Predictive Peng-Robinson 1978 Equation of State

Clean VBA functions - no UI changes and no pop up messages. Use these functions in HMB workbooks.
Errors are reported in cell comments. When error clears then error comment also clears.
Requires Pure Component Property Data WorkSheet contained in PData.xlsx.
Import the Math.bas, ModArraySupport.bas and ChE_Functions.bas modules into PData.xlsx and save as xlsm.
Create a dataset and start calculations. Calculations can be made with or without binary interaction coefficients. User
binary interaction coefficients or predictive binaries can be used in calculations.

Functions List:

CreateDataset, CreateDecomposition (flatten these cells after creation to greatly improve calculation speed)

FlashTP (vapor fractions or single phase component composition can be specified with Solver)

DewT, BubbleT (dew and bubble pressures calculated with DewT and BubbleT and Solver)

Enthalpy - real & ideal gas mixtures

Entropy - real & ideal gas mixtures

Cv - real & ideal gas mixtures

Cp - real & ideal gas mixtures

Speed of Sound for real gas mixtures

Joule Thompson Coefficient for real gas mixtures

Derivatives of PPR1978 EOS

Compressibility (z) OF real gas mixtures

Fugacity coefficients of real gas mixtures

Volume of real gas mixtures

References 

1
Peng Robinson Equation of State
A New Two-Constant Equation of State
Ding-Yu Peng, Donald B. Robinson
Ind. Eng. Chem. Fundamen., 1976, 15 (1), pp 59–64
DOI: 10.1021/i160057a011
Publication Date: February 1976

2
Spreadsheet for Thermodynamics Instruction
Phillip Savage - University of Michigan
'ChE classroom'
Fall 1995
http://ufdcimages.uflib.ufl.edu/AA/00/00/03/83/00128/AA00000383_00128_00262.pdf

3
Flash Routine Reference:
CHEMENG 120 taught by Professor Musgrave during the Spring '04 term at Stanford.
http://documentslide.com/documents/lecture-5-isothermal-flash-calculations.html

4
Implementation of Departure Functions 
Dr. Phillip Savage - Penn State University
http://www.che.psu.edu/department/directory-detail.aspx?LandOn=Gen&q=pes15

5
Bubble and Dew Point Routines adapted from:
It’s not as easy as it looks - Revisiting Peng–Robinson equation of state convergence issues for dew point, bubble point and flash calculations

Vamshi Krishna Kandula,(a) John C. Telotte (b) and F. Carl Knopf (corresponding author) (a)
(a) Chemical Engineering Department, Louisiana State University, USA
E-mail: cknopf@southalabama.edu
(b) Chemical Engineering Department, Florida A&M University – Florida State University, USA

International Journal of Mechanical Engineering Education, Volume 41, Number 3 (July 2013), © Manchester University Press
 http://journals.sagepub.com/doi/pdf/10.7227/IJMEE.41.3.2

6
Cubic Equation Solver Code
Dr. Tomas B. Co
Michigan Tech University
https://www.mtu.edu/chemical/department/faculty/co/

7
Bicubic Interpolation Code (for LeeKeslerZ() function)
https://mathformeremortals.wordpress.com/

8
modArraySupport module and array handling techniques
Chip Pearson, chip@cpearson.com, www.cpearson.com

9
Derivation of the enthalpy departure function
https://shareok.org/bitstream/handle/11244/12606/Thesis-1996-R231e.pdf?sequence=1
by Abhishek Rastogi
APPENDIX A
A DETAILED DERIVATION OF PENG-ROBINSON EQUATION OF STATE ENTHALPY DEPARTURE FUNCTION 

10
PData worksheet physical properties adapted from:
Properties Databank 1.0
Pedro Fajardo
ppfk@yahoo.com
http://www.cheresources.com/invision/files/file/125-physical-properties-ms-excel-add-in/

11
Thermodynamic Properties Involving Derivatives
Using the Peng-Robinson Equation of State
R.M. Pratt Ph. D.
The National University of Malaysia
(now Associate Professor of Mathematics and Sciences, Fresno Pacific University)
http://ufdcimages.uflib.ufl.edu/AA/00/00/03/83/00150/AA00000383_00150_00112.pdf

12
Chemical Equilibrium by Gibbs Energy Minimization on Spreadsheets
Int. J. Engng Ed. Vol. 16, No. 4, pp. 335±339, 2000
Y. LWIN
Department of Chemical Engineering, Rangoon Institute of Technology, Insein P. O., Rangoon, Burma.
E-mail: ylwin@yahoo.com
http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.476.5931&rep=rep1&type=pdf

13
Getting a Handle on Advanced Cubic Equations of State
www.cepmagazine.org, November 2002, CEP
Chorng H. Twu, Wayne D. Sim and Vince Tassone, Aspen Technology, Inc.
http://people.clarkson.edu/~wwilcox/Design/adv-ceos.pdf

14
'VLE predictions with the Peng.Robinson equation of state and
'temperature dependent kij calculated through a group contribution method
'Jean-No¡§el Jaubert., Fabrice Mutelet
'Laboratoire de Thermodynamique des Milieux Polyphas¢¥es, Institut National Polytechnique de Lorraine, Ecole Nationale Sup¢¥erieure des Industries
'Chimiques, 1rue Grandville, 54000 Nancy, France
'Received 24 January 2004; accepted 25 June 2004

15
Prediction of Thermodynamic Properties of Alkyne-Containing
Mixtures with the E‑PPR78 Model
Xiaochun Xu,† Jean-Noël Jaubert,*,† Romain Privat,† and Philippe Arpentinier‡
†Ecole Nationale Supérieure des Industries Chimiques, Laboratoire Réactions et Génie des Procédés (UMR CNRS 7274),
Université de Lorraine, 1 rue Grandville, 54000 Nancy, France
‡Centre de Recherche Paris Saclay, Air Liquide, 1 chemin de la porte des loges, BP 126, 78354 Jouy-en-Josas, France

16
Efficient flash calculations for chemical process
design — extension of the Boston–Britt
‘‘Inside–out’’ flash algorithm to extreme
conditions and new flash types
Vipul S. Parekh and Paul M. Mathias
Computers Chem. Engng Vol. 22, No. 10, pp. 1371—1380, 1998
( 1998 Published by Elsevier Science Ltd.


