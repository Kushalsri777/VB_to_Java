Attribute VB_Name = "modResConstants"
Option Explicit
Option Compare Binary

' The following constant is used to determine whether the cerhErrorHandler class (gerhApp)
' is empty (i.e. set to its initialized state) or not.
Public Const gclngERR_NUM_DEFAULT As Long = 999

' The following is used to string together a Screen Name with a Proc Name to form
' the context when a form encounters an error directly (i.e. non-raised)
Public Const gcstrDOT As String = "."

'------------------------------------------------------------------------
'    Public Constants re: Icons, Bitmaps and others items in CLAIM.RES
'------------------------------------------------------------------------
Public Const gcRES_ICON_MAINAPP As String = "_MAINAPP"
Public Const gcRES_ICON_INFO As Long = 101
Public Const gcRES_ICON_WARN As Long = 103
Public Const gcRES_ICON_ALRT As Long = 102
Public Const gcRES_ICON_ERR As Long = 104
Public Const gcRES_ICON_BINOCULARS As Long = 105


'------------------------------------------------------------------------
'       Public Constants re: Warning/Info/Alert/Error messages
'
'      The following ranges MUST be used and MUST correspond to IDs
'                     in the CLAIM.RES resource file.
'------------------------------------------------------------------------
Public Const gcRES_LOWEST_APP_ERROR As Integer = 1000        ' Lower Bounds for CLAIM.RES
Public Const gcRES_HIGHEST_APP_ERROR As Integer = 9999       ' Upper Bounds for CLAIM.RES

' -=-= Informational Messages =-=-
Public Const gcRES_INFO_START As Integer = 1000                            ' <Informational messages (1000-1999) start here>
Public Const gcRES_INFO_CANT_LAUNCH_URL As Integer = 1001                  ' Unable to open <@@1> in a browser window. Please ensure that the Internet Explorer and the Adobe Acrobat reader software are installed.
Public Const gcRES_INFO_ANOTHER_USER_UPDATED_DISCARDED As Integer = 1002   ' Another user (@@1) updated this record since you displayed it. Your changes have been discarded.
Public Const gcRES_INFO_ANOTHER_USER_DELETED_NOT_SAVED As Integer = 1003   ' Another user deleted this record since you displayed it. Your changes have not been saved.
Public Const gcRES_INFO_TAX_FILE_GEND As Integer = 1004                    ' @@1 record(s) were written to the @@2 tax file. The total interest (Box 1) amount was @@3 and the
                                                                           ' total Interest Withheld (Box 4) amount was @@4.@@CRLF
                                                                           ' @@5 record(s) were written to the @@6 tax file. The total interest (Box 1) amount was @@7 and the
                                                                           ' total Interest Withheld (Box 4) amount was @@8.
Public Const gcRES_INFO_ANOTHER_USER_DELETED As Integer = 1005             ' Another user deleted this record since you displayed it.
Public Const gcRES_INFO_TABLE_IS_EMPTY As Integer = 1006                   ' The @@1 table is empty.
Public Const gcRES_INFO_NO_AUTHENTICATED_ENVIRONMENTS As Integer = 1007    ' The specified @@1 is not authorized for any of the application's @@2s.
Public Const gcRES_INFO_DT_CHG_MAY_AFFECT_PAYEES As Integer = 1008         ' The @@1 has changed. This change may affect the calculations for existing Payees. Please review and, if necessary, recalculate each Payee.
Public Const gcRES_INFO_INPUT_WAS_TRUNCATED As Integer = 1009              ' Your input was truncated to @@1 character(s).
Public Const gcRES_INFO_CANT_OPEN_FILE As Integer = 1014                   ' Unable to open @@1. The file either does not exist or no application is associated with files of type @@2.
Public Const gcRES_INFO_END As Integer = 1999                              ' <Informational messages (1000-1999) end here>
' -=-= Warnings =-=-
Public Const gcRES_WARN_START As Integer = 2000                            ' <Warning messages (2000-2999) start here>
Public Const gcRES_WARN_CBO_IS_EMPTY As Integer = 2001                     ' The drop-down list for @@1 is empty. Since you will be unable to make a selection in this field, this screen may behave unpredictibly.
Public Const gcRES_WARN_LST_IS_EMPTY As Integer = 2002                     ' The list for @@1 is empty. Since you will be unable to make a selection in this field, this screen may behave unpredictibly.
Public Const gcRES_WARN_NO_CURR_INSURED As Integer = 2003                  ' There is no current Insured record. The Payee screen cannot be opened.
Public Const gcRES_WARN_GET_TIN_BEFORE_PAYING_INT As Integer = 2004        ' This claims requires a certified @@1 to avoid withholding. Make sure you don't pay interest until it has been received.
Public Const gcRES_WARN_DTHB_PMT_AMT_MAY_BE_TOO_HIGH As Integer = 2005     ' The @@1 exceeds @@2. Please verify this amount is correct.
Public Const gcRES_WARN_NONNUMERIC_RATE As Integer = 2006                  ' The Rate supplied or obtained from the STATE_RULE_T table is non-numeric (@@1). Please try again.
Public Const gcRES_WARN_RATE_IS_NEGATIVE As Integer = 2007                 ' The Rate supplied or derived from the supplied Rate is a negative number (@@1). Please try again.
Public Const gcRES_WARN_TOO_MANY_DECIMALS As Integer = 2008                ' The Rate supplied cannot have more than @@1 decimal positions specified. Please try again.
Public Const gcRES_WARN_POSSIBLE_FILESYS_PERM_PROBLEM  As Integer = 2010   ' An error was encountered while trying to @@1. This may be due to network unavailability or insufficient authorizations. @@2

' Case use 2003, 2004, 2005, 2006, 2007
Public Const gcRES_WARN_END As Integer = 2999                              ' <Warning messages (2000-2999) end here>
' -=-= Alerts =-=-
Public Const gcRES_ALRT_START As Integer = 3000                            ' <Alert messages (3000-3999) start here>
' Can use 3001
Public Const gcRES_ALRT_OK_TO_DELETE_RECORD As Integer = 3002              ' Are you sure you want to delete this record?
' Can use 3003
Public Const gcRES_ALRT_CHANGES_PENDING As Integer = 3004                  ' You have changes pending.  Do you want to lose them?
' Can use 3005, 3006, 3008
'?? Public Const gcRES_ALRT_OK_TO_DELETE_INFO As Integer = 3007                ' You are about to delete the information for @@1. Are you sure you want to delete this?
Public Const gcRES_ALRT_END As Integer = 3999                              ' <Alert messages (3000-3999) end here>
' -=-= Non-Fatal (i.e. Process Fatal) Errors =-=-
Public Const gcRES_NERR_START As Integer = 4000                            ' <Non-fatal Error messages (4000-4999) start here>
Public Const gcRES_NERR_CANTOPEN_REGKEY As Integer = 4001                  ' Unable to retrieve the value (@@1) of the requested registry key (@@2).
Public Const gcRES_NERR_CANTSAVE_REGKEY As Integer = 4002                  ' Unable to save the value (@@1) to the requested registry key (@@2).
Public Const gcRES_NERR_DRIVE_OR_PATH_NOT_FOUND As Integer = 4003          ' The drive or path specified does not exist. Please be sure to specify an existing drive and directory.
Public Const gcRES_NERR_NO_RECS_WERE_FOUND As Integer = 4004               ' No records were found with @@1.
Public Const gcRES_NERR_INTEREST_RATE_TOO_HIGH As Integer = 4005           ' The Rate supplied or derived from the supplied Rate is more than @@1%. This is only allowed when the @@2 is Maine. Please try again.
Public Const gcRES_NERR_STATE_RATES_NOT_FOUND As Integer = 4006            ' Neither Group nor Individual rates were found for the state of @@1 as of @@2. The calculations cannot be done.
Public Const gcRES_NERR_INDV_STATE_RATES_NOT_FOUND As Integer = 4007       ' Individual rates were not found for the state of @@1 as of @@2. The calculations cannot be done.
Public Const gcRES_NERR_NUMERIC_FLD_TOO_LARGE As Integer = 4008            ' One or more numeric fields are too large to be stored in the database. Your changes cannot be saved.
Public Const gcRES_NERR_CALC_WAS_CANCELLED As Integer = 4010               ' The calculation was halted since you clicked Cancel. Your changes have not been saved.
Public Const gcRES_NERR_RPTFILE_NOT_FOUND As Integer = 4011                ' The report definition file for the selected report (@@1) could not be found.
' MME START - WRUS 4999
Public Const gcRES_NERR_INVALID_ENTRY_RULE_TIER_T As Integer = 4012        ' Invalid record found on table STATE_RULE_TIER_T (4012) - ' for the state of @@1 as of @@2. The calculations cannot be done.
' MME END  - WRUS 4999
' Can use 4012, 4014, 4017, 4018, 4019, 4020, 4021, 4022
Public Const gcRES_NERR_LOGON_FAILURE As Integer = 4013                    ' The logon was unsuccessful. Please verify the correct User ID and Password were specified correctly, with appropriate case, that you have permissions to the database and the server on which it is located, and that Microsoft Data Access Components (MDAC) is installed. (RC=@@1)
Public Const gcRES_NERR_CONNECTION_FAILURE As Integer = 4015               ' A connection could not be established to the @@1 environment's database. (State=@@2)
Public Const gcRES_NERR_ENV_REG_ENTRIES_MISSING As Integer = 4016          ' One or more registry entries that define how to connect to the selected Environment (@@1) are missing. Without all of these entries, the app cannot connect to the database.
'?? Public Const gcRES_NERR_TABLE_IS_EMPTY As Integer = 4023                   ' The @@1 table is empty.
' Can use 4024
'?? Public Const gcRES_NERR_MEFS_EFF_DT_BEFORE_CASE_EFF_DT As Integer = 4025   ' The @@1 is prior to the @@2 (@@3).
Public Const gcRES_NERR_UNEXPECTED_VAL_SELECT_CASE As Integer = 4026       ' Programmer error:  An unexpected value was encountered in a SELECT CASE statement.
Public Const gcRES_NERR_REC_NOT_FOUND As Integer = 4027                    ' The specified record was not found in the database (@@1).
Public Const gcRES_NERR_ERR_WHILE_TRYING_TO As Integer = 4028              ' An error occurred while attempting to @@1 this record.
Public Const gcRES_NERR_DEPENDENT_RECS_EXIST As Integer = 4029             ' This @@1 is associated with one or more records on the @@2 table and cannot be deleted until those records themselves are deleted.
'?? Public Const gcRES_NERR_FUND_USED_AS_MKTVAL_FUND As Integer = 4030         ' Fund @@1 cannot be deleted because it is used as another fund's Market Value Fund Cd.
Public Const gcRES_NERR_ADD_WITH_NONUNIQUE_KEY As Integer = 4031           ' A record with the specified key (@@1) already exists. Please specify a unique key.
Public Const gcRES_NERR_KEY_NOT_FOUND As Integer = 4032                    ' The @@1 specified (@@2) is no longer defined in the @@3 table. Please choose a different value.
' Can use 4033, 4035, 4039, 4043, 4044
Public Const gcRES_NERR_CROSS_FLD_VALIDATIONS_FAILED As Integer = 4034     ' Cross-field validation errors were found. These must be corrected before @@1:@@2
'?? Public Const gcRES_NERR_THIS_FUNCTIONALITY_NOT_AVAIL As Integer = 4036     ' This functionality isn't available yet.
Public Const gcRES_NERR_INVALID_DATA As Integer = 4037                     ' The @@1 is invalid. @@2
'?? Public Const gcRES_NERR_NO_DATA_WAS_FOUND As Integer = 4038                ' No data was found @@1.
Public Const gcRES_NERR_REQD_FIELDS_MISSING As Integer = 4041              ' The following required fields must be supplied before your request can be processed:@@CRLF@@1
'?? Public Const gcRES_NERR_All_MUST_BE_INPUT As Integer = 4042                ' If any of the following fields are input, then all must be:@@1
Public Const gcRES_NERR_END As Integer = 4999                              ' <Non-fatal Error messages (4000-4999) end here>

' -=-= Fatal (i.e. App Fatal) Errors =-=-
Public Const gcRES_FERR_START As Integer = 9000                            ' <Fatal Error messages (9000-9999) start here>
Public Const gcRES_FERR_NO_ENVS As Integer = 9001                          ' No Environments have been defined in the registry. Without these entries (built by the install), the app will be unable to connect to the database.
Public Const gcRES_FERR_SPROC_NOT_FOUND As Integer = 9002                  ' The stored procedure "@@1" was not found.
Public Const gcRES_FERR_SQL_STMT_OBJECT_NOT_FOUND As Integer = 9003        ' The database object referenced in a SQL statement (@@1) was not found or you have insufficient permissions to access it.
Public Const gcRES_FERR_END As Integer = 9999                              ' <Fatal Error messages (9000-9999) end here>

