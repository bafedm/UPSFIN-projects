Attribute VB_Name = "Constants"
'@Folder("Main")
Option Explicit

'Global constants
Public Const WS_PAF_GEN         As String = "Allocation and Forecast Gen"   'Worksheet to launch the WB Generator
Public Const TBL_PROJECT_LIST   As String = "q_co_PafProjectList"           'Table containing project list (PQ generated)

'@Description "Returns an array pretending to be a constant"
Public Function ARRAY_DESC_GROUPS( _
                                    Optional ByVal intIndex As Integer) _
                                    As Variant
                                    
ARRAY_DESC_GROUPS = Array("Revenue", "Personnel Expenses", "External Services", "Travel Expenses", _
                        "Depreciation", "Other Expenses", "Allocation Indirect Expenses", "Split Overhead & Dir/Indir Costs")
    
End Function
