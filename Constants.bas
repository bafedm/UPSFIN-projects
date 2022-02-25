Attribute VB_Name = "Constants"
'@Folder("Main")
Option Explicit

'Global constants
Public Const WS_PAF_GEN         As String = "Allocation and Forecast Gen"   'Worksheet to launch the WB Generator
Public Const TBL_PROJECT_LIST   As String = "q_co_PafProjectList"           'Table containing project list (PQ generated)

'@Description "Returns an array of desc groups pretending to be a constant"
Public Function ARRAY_DESC_GROUPS( _
                                    Optional ByVal intIndex As Integer) _
                                    As Variant
                                    
ARRAY_DESC_GROUPS = Array("Revenue", "Personnel Expenses", "External Services", "Travel Expenses", _
                        "Depreciation", "Other Expenses", "Allocation Indirect Expenses", "Split Overhead & Dir/Indir Costs")
    
End Function

'@Description "A list of all desc groups that fall into the revenue category"
Public Function ARRAY_DESC_GROUPS_REV( _
                                        Optional ByVal intIndex As Integer) _
                                        As Variant
ARRAY_DESC_GROUPS_REV = Array("Revenue")

End Function

'@Description "A list of all desc groups that fall into the costs category"
Public Function ARRAY_DESC_GROUPS_COSTS( _
                                        Optional ByVal intIndex As Integer) _
                                        As Variant
ARRAY_DESC_GROUPS_COSTS = Array("Personnel Expenses", "External Services", "Travel Expenses", _
                        "Depreciation", "Other Expenses", "Allocation Indirect Expenses", "Split Overhead & Dir/Indir Costs")

End Function
