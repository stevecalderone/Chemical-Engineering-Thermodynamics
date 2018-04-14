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

FlashTP (vapor fractions can be specified with Solver)

DewT, BubbleT (dew and bubble pressures calculated with DewT and BubbleT and Solver)

Enthalpy - real & ideal gas

Entropy - real & ideal gas

Cv - real & ideal gas

Cp - real & ideal gas

Speed of Sound for real gas

Joule Thompson Coefficient for real gas

Derivatives of PPR1978 EOS

Compressibility (z) OF real gas

Fugacity coefficients of real gas

Volume of real gas


