Attribute VB_Name = "TEST_DefectProbe"
'==============================================================================
' MODULE: TEST_DefectProbe
'------------------------------------------------------------------------------
' PURPOSE
'   Deliberately defective module of the controlled enforcement exercise of
'   #12. Line 13 ends in trailing white space, which the vba-source gate
'   must reject. It is removed by the next commit and never merged.
'==============================================================================
Option Explicit
Option Private Module

Private Const PROBE_VALUE As Long = 1   'Not used 
