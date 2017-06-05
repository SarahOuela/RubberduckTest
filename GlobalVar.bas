Attribute VB_Name = "GlobalVar"
''' Variables diverses
Public g_Language As String
Public g_Language_Temp As String
Public g_Language_Sheet As String
Public g_listHeadCR() As String
Public g_WSerror As String
'Public g_nbMeta As Integer
'Public g_nbData As Integer
'Public g_nbInput As Integer
''' Liste des colonnes
Public SYNONYMOUS As Integer
''' Variable de workbook
Public g_WB_Extra As Workbook
Public g_WB_Modele As Workbook
Public g_WB_Modele_Path As String
Public g_WB_Modele_Name As String
Public g_WB_Source As Workbook
Public g_WB_Source_Path As String
Public g_WB_Source_Name As String
Public g_WB_Target As Workbook
Public g_WB_Target_Path As String
Public g_WB_Target_Name As String
''' Liste des synonymes
Public g_Synonymes() As String
Public g_lastValue As Boolean
Public g_withoutFormula As Boolean
''' Variables de sheet
Public Const g_CONTROL = "CONTROL"
Public Const g_CONTROL_PRE_L = 13
Public Const g_CONTROL_PRE_C = 30
Public Const g_CONTROL_META_L = 16
Public Const g_CONTROL_META_C = 11
Public Const g_CONTROL_META_CR_L = 16
Public Const g_CONTROL_META_CR_C = 12
Public Const g_CONTROL_DATA_L = 16
Public Const g_CONTROL_DATA_C = 19 '16
Public Const g_CONTROL_DATA_GEN_L = 16
Public Const g_CONTROL_DATA_GEN_C = 20 '12
Public Const g_CONTROL_DATA_DAT_L = 16
Public Const g_CONTROL_DATA_DAT_C = 25 '17
Public Const g_CONTROL_DATA_FOR_L = 16
Public Const g_CONTROL_DATA_FOR_C = 30 '22
Public Const g_CONTROL_DATA_COM_L = 16
Public Const g_CONTROL_DATA_COM_C = 38 '27
Public Const g_CONTROL_INPUT_L = 16
Public Const g_CONTROL_INPUT_C = 3
Public Const g_CONTROL_NOMENCLATURE_GEN_CR_L = 11
Public Const g_CONTROL_NOMENCLATURE_GEN_CR_C = 4
Public Const g_CONTROL_QUANTITES_GEN_CR_L = 11
Public Const g_CONTROL_QUANTITES_GEN_CR_C = 12
Public Const g_CRMETA = "CRMETA"
Public Const g_CRFORMULA = "CRFORMULA"
Public Const g_CRDATA = "CRDATA"
Public Const g_QUANTITY = "QUANTITY"
Public Const g_NOMENCLATURE = "NOMENCLATURE"
Public Const g_AREA = "AREA"
Public Const g_SCENARIO = "SCENARIO"
Public Const g_TIME = "TIME"

''' Libellés
Public Const g_noFile = "Aucun Fichier Sélectionné"
Public g_SOURCE As String
Public g_TARGET As String
Public g_PARAMETERS As String
'Public g_SOURCE__WB As Workbook
'Public g_TARGET__WB As Workbook
'Public g_PARAMETERS__WB As Workbook

Public Const g_feuille = "DATA"
Public Const g_alimentation = "ALIMENTATION"
Public Const g_equation = "EQUATION"
Public Const g_line_nom = 18
Public Const g_col_nom = 2
Public Const g_line_qua = 18
Public Const g_Col_qua = 6
Public Const g_line_feu = 33
Public Const g_col_feu = 4
Public Const g_col_feu_cible = 7
Public Const g_col_feu_source = 8
Public Const g_col_nontype = 10
Public Const g_line_nontype = 18

