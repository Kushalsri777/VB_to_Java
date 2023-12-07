public class modRegistry {

  //
  // Module      : modRegistry
  // Description : This module Implements routines for manipulating the registry.
  // Source      : Total Visual SourceBook 2000
  //
  // Procedures  :
  //   Private
  //   Public      RegistryCreateNewKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
  //   Public      RegistryDeleteKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
  //   Public      RegistryDeleteValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
  //                                   strValueName As String)
  //   Public      RegistryEnumerateSubKeys(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
  //                                   astrKeys() As String, lngKeyCount As Long)
  //   Public      RegistryEnumerateValues(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
  //                                   astrValues() As String, lngValueCount As Long)
  //   Public      RegistryGetKeyValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
  //                                   strValueName As String) As Variant
  //   Public      RegistrySetKeyValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, _
  //                                   strValueName As String, varData As Variant, eDataType As EnumRegistryValueType)
  //
  // Modified    :
  // 03/03/02 BAW (Phase2A) Added support for new global error handler
  //
//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modRegistry.";

//*TODO:** type is translated as a new class at the end of the file Private Type FILETIME

//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumRegistryRootKeys

//*TODO:** enum is translated as a new class at the end of the file Public Enum EnumRegistryValueType

*TODO: API Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngHKey As Long) As Long

*TODO: API Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

*TODO: API Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngHKey As Long, ByVal lpSubKey As String) As Long

*TODO: API Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngHKey As Long, ByVal lpValueName As String) As Long

*TODO: API Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

*TODO: API Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long

*TODO: API Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

*TODO: API Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal lngHKey As Long, ByVal lpClass As String, ByVal lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, ByVal lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, ByVal lpcbMaxValueLen As Long, ByVal lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

*TODO: API Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

*TODO: API Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

*TODO: API Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

*TODO: API Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

*TODO: API Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As Long, ByVal cbData As Long) As Long

*TODO: API Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

*TODO: API Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long


  private static final int MCREGOPTIONNONVOLATILE = 0;

  private static final int MCREGERRORNONE = 0;
  private static final int MCREGERRORBADDB = 1;
  private static final int MCREGERRORBADKEY = 2;
  private static final int MCREGERRORCANTOPEN = 3;
  private static final int MCREGERRORCANTREAD = 4;
  private static final int MCREGERRORCANTWRITE = 5;
  private static final int MCREGERROROUTOFMEMORY = 6;
  private static final int MCREGERRORINVALIDPARAMETER = 7;
  private static final int MCREGERRORACCESSDENIED = 8;
  private static final int MCREGERRORINVALIDPARAMETERS = 87;
  private static final int MCREGERRORNOMOREITEMS = 259;

  private static final int READ_CONTROL = 0x20000;
  private static final int STANDARD_RIGHTS_ALL = 0x1F0000;
  *TODO:** (the data type can't be found for the value [(READ_CONTROL)])Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
  *TODO:** (the data type can't be found for the value [(READ_CONTROL)])Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
  private static final int SYNCHRONIZE = 0x100000;

  private static final int KEY_CREATE_LINK = 0x20;
  private static final int KEY_CREATE_SUB_KEY = 0x4;
  private static final int KEY_ENUMERATE_SUB_KEYS = 0x8;
  private static final int KEY_NOTIFY = 0x10;
  private static final int KEY_QUERY_VALUE = 0x1;
  private static final int KEY_SET_VALUE = 0x2;
  private static final ((STANDARD_RIGHTS_ALL KEY_ALL_ACCESS = KEY_QUERY_VALUE;
  private static final ((STANDARD_RIGHTS_READ KEY_READ = KEY_QUERY_VALUE;
  private static final ((STANDARD_RIGHTS_WRITE KEY_WRITE = KEY_SET_VALUE;



  ///\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  //\                                                                  /
  //|                        PUBLIC  Procedures                        |
  ///                                                                  \
  //\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registryCreateNewKey(EnumRegistryRootKeys eRootKey, String strKeyName) {
    // Comments  : Creates a new key in the system registry
    // Parameters: eRootKey - The root key
    //             strKeyName - The name of the key to create
    // Returns   : Nothing
    //
    // Called by :
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryCreateNewKey"
.equals(Const cstrCurrentProc As String);
    int lngRetVal = 0;
    int lngHKey = 0;

    try {

      // Create the key
      lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, "", MCREGOPTIONNONVOLATILE, KEY_WRITE, 0&, lngHKey, 0&);

      // if the key was created, then close it
      if (lngRetVal == MCREGERRORNONE) {
        RegCloseKey(lngHKey);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registryDeleteKey(EnumRegistryRootKeys eRootKey, String strKeyName) {
    // Comments  : Deletes a key from the system registry
    // Parameters: eRootKey - The root key
    //             strKeyName - The name of the key to delete
    // Returns   : Nothing
    //
    // Called by :
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryDeleteKey"
.equals(Const cstrCurrentProc As String);
    int lngRetVal = 0;

    try {

      // Delete the key
      lngRetVal = RegDeleteKey(eRootKey, strKeyName);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registryDeleteValue(EnumRegistryRootKeys eRootKey, String strKeyName, String strValueName) {
    // Comments  : Deletes a value from the system registry
    // Parameters: eRootKey - The root key
    //             strKeyName - The name of the key to delete
    //             strValueName - The name of the value to delete
    // Returns   : Nothing
    //
    // Called by :
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryDeleteValue"
.equals(Const cstrCurrentProc As String);
    int lngRetVal = 0;
    int lngHKey = 0;

    try {

      // Open the key
      lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_WRITE, lngHKey);

      // If the key was opened successfully, then delete it
      if (lngRetVal == MCREGERRORNONE) {
        lngRetVal = RegDeleteValue(lngHKey, strValueName);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registryEnumerateSubKeys(EnumRegistryRootKeys eRootKey, String strKeyName, String[] astrKeys, int lngKeyCount) { // TODO: Use of ByRef founded Public Sub RegistryEnumerateSubKeys(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String, ByRef astrKeys() As String, ByRef lngKeyCount As Long)
    // Comments  : Enumerates the sub keys of the specified key
    // Parameters: eRootKey - The root key
    //             strKeyName - The name of the key to enumerate
    //             astrKeys - An array of strings to fill with sub key names
    //             lngKeyCount - The number of sub keys returned in the parameter
    //             astrKeys
    // Returns   : Nothing
    //
    // Called by :
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryEnumerateSubKeys"
.equals(Const cstrCurrentProc As String);
    int lngHKey = 0;
    int lngKeyIndex = 0;
    int lngMaxKeyLen = 0;
    int lngRetVal = 0;
    int lngSubkeyCount = 0;
    String strSubKeyName = "";
    FILETIME typFT = null;

    try {

      // Open the key
      lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_READ, lngHKey);

      if (lngRetVal == MCREGERRORNONE) {
        //find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, "", 0, 0, lngSubkeyCount, lngMaxKeyLen, 0, 0, 0, 0, 0, typFT);
        if (MCREGERRORNONE == lngRetVal) {
          if (lngSubkeyCount > 0) {
            G.redimPreserve(lngSubkeyCount - 1,  );

            //set up the while loop
            lngKeyIndex = 0;
            // Pad the string to the maximum length of a sub key, plus 1 for null
            // termination
            lngMaxKeyLen = lngMaxKeyLen + 1;
            strSubKeyName = Space$(lngMaxKeyLen);

            while (RegEnumKey(lngHKey, lngKeyIndex, strSubKeyName, lngMaxKeyLen + 1) == 0) {

              // Set the string array to the key name, removing null termination
              if (strSubKeyName.indexOf(vbNullChar, 1) > 0) {
                astrKeys(lngKeyIndex) = strSubKeyName.substring(0, strSubKeyName.indexOf(vbNullChar, 1) - 1);
              }
              // Increment the key index for the return string array
              lngKeyIndex = lngKeyIndex + 1;

            }
          }
          // return the new dimension of the return string array
          lngKeyCount = lngSubkeyCount;
        }

        // Close the key
        RegCloseKey(lngHKey);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registryEnumerateValues(EnumRegistryRootKeys eRootKey, String strKeyName, String[] astrValues, int lngValueCount) { // TODO: Use of ByRef founded Public Sub RegistryEnumerateValues(ByVal eRootKey As EnumRegistryRootKeys, ByVal strKeyName As String, ByRef astrValues() As String, ByRef lngValueCount As Long)
    // Comments  : Enumerates the values of the specified key
    // Parameters: eRootKey - The root key
    //             strKeyName - The name of the key to enumerate
    //             astrValues - An array of strings to fill with value names
    //             lngValueCount - The number of values returned in the parameter astrValues
    // Returns   : Nothing
    //
    // Called by :
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryEnumerateValues"
.equals(Const cstrCurrentProc As String);
    int lngHKey = 0;
    int lngKeyIndex = 0;
    int lngMaxValueLen = 0;
    int lngRetVal = 0;
    int lngTempValueCount = 0;
    String strValueName = "";
    FILETIME typFT = null;

    try {

      // Open the key
      lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, KEY_READ, lngHKey);

      if (lngRetVal == MCREGERRORNONE) {
        //find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, "", 0, 0, 0, 0, 0, lngTempValueCount, lngMaxValueLen, 0, 0, typFT);
        if (MCREGERRORNONE == lngRetVal) {
          if (lngTempValueCount > 0) {
            G.redimPreserve(lngTempValueCount - 1,  );

            //set up the while loop
            lngKeyIndex = 0;
            // Pad the string to the maximum length of a sub key, plus 1 for null
            // termination
            lngMaxValueLen = lngMaxValueLen + 1;
            strValueName = Space$(lngMaxValueLen);

            while (RegEnumValue(lngHKey, lngKeyIndex, strValueName, lngMaxValueLen + 1, 0, 0, 0, 0) == 0) {

              // Set the string array to the key name, removing null termination
              if (strValueName.indexOf(vbNullChar, 1) > 0) {
                astrValues(lngKeyIndex) = strValueName.substring(0, strValueName.indexOf(vbNullChar, 1) - 1);
              }
              // Increment the key index for the return string array
              lngKeyIndex = lngKeyIndex + 1;

            }
          }
          // return the new dimension of the return string array
          lngValueCount = lngTempValueCount;
        }

        // Close the key
        RegCloseKey(lngHKey);
      }
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static Object registryGetKeyValue(EnumRegistryRootKeys eHKeyRoot, String strKeyName, String strValueName) {
    Object _rtn = null;
    // Comments  : Returns a value from the system registry
    // Parameters: eHKeyRoot - The root key
    //             strKeyName - The name of the key
    //             strValueName - The name of the value
    // Returns   : The data in the registry value
    //
    // Called by : ReadEntry( ) in CAppSettings.cls
    //             RestoreForm( ) in CAppSettings.cls
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistryGetKeyValue"
.equals(Const cstrCurrentProc As String);
    byte[] abytValueData() = null;
    int lngDataSize = 0;
    int lngHKey = 0;
    int lngRetVal = 0;
    int lngValueData = 0;
    int lngValueType = 0;
    String strValueData = "";
    Object varValue = null;

    try {

      varValue = Empty;

      lngRetVal = RegOpenKeyEx(eHKeyRoot, strKeyName, 0&, KEY_READ, lngHKey);

      if (MCREGERRORNONE == lngRetVal) {

        lngRetVal = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, 0&, lngDataSize);

        if (lngRetVal == MCREGERRORNONE) {

          switch (lngValueType) {

              // String type

            case  EnumRegistryValueType.rRKREGSZ::
              if (lngDataSize > 0) {
                strValueData = String(lngDataSize, 0);
                lngRetVal = RegQueryValueExString(lngHKey, strValueName, 0&, lngValueType, strValueData, lngDataSize);
                if (strValueData.indexOf(vbNullChar) > 0) {
                  strValueData = strValueData.substring(1, strValueData.indexOf(vbNullChar) - 1);
                }
              }
              if (MCREGERRORNONE == lngRetVal) {
                varValue = strValueData.substring(0, lngDataSize);
              } 
              else {
                varValue = Empty;
              }

              // Long type
              break;

            case  EnumRegistryValueType.rRKREGDWORD::
              lngRetVal = RegQueryValueExLong(lngHKey, strValueName, 0&, lngValueType, lngValueData, lngDataSize);
              if (MCREGERRORNONE == lngRetVal) {
                varValue = lngValueData;
              }

              // Binary type
              break;

            case  EnumRegistryValueType.rRKREGBINARY:
              if (lngDataSize > 0) {
                G.redimPreserve(lngDataSize - 1,  );
                lngRetVal = RegQueryValueExBinary(lngHKey, strValueName, 0&, lngValueType, VarPtr(abytValueData[0]), lngDataSize);
              }
              if (MCREGERRORNONE == lngRetVal) {
                varValue = abytValueData;
              } 
              else {
                varValue = Empty;
              }

              break;

            default:
              //No other data types supported
              lngRetVal = -1;
              modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_UNEXPECTED_VAL_SELECT_CASE, MCSTRNAME+ cstrCurrentProc);
              // **TODO:** goto found: GoTo PROC_EXIT;
              break;
          }
        }

        RegCloseKey(lngHKey);
      }

      //Return varValue
      _rtn = varValue;
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return _rtn;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CANTOPEN_REGKEY, MCSTRNAME+ cstrCurrentProc, strValueName, strKeyName);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static void registrySetKeyValue(EnumRegistryRootKeys eHKeyRoot, String strKeyName, String strValueName, Object varData, EnumRegistryValueType eDataType) {
    // Comments  : This procedure sets a key value, creating the key if it doesn't exist.
    // Parameters: eHKeyRoot - The root key
    //             strKeyName - The name of the key to open
    //             strValueName - The name of the value to open (vbNulLString will open the default value).
    //             varData - The data to store in the value
    //             eDataType - The type of data to store in the value
    // Returns   : Nothing
    //
    // Called by : SaveForm( ) in CAppSettings.cls
    //             WriteEntry( ) in CAppSettings.cls
    //
    // Source    : Total Visual SourceBook 2000
    //
    "RegistrySetKeyValue"
.equals(Const cstrCurrentProc As String);
    byte[] abytData() = null;
    int lngData = 0;
    int lngHKey = 0;
    int lngRetVal = 0;
    String strData = "";

    try {

      // Open the specified key. If it does not exist, then create it
      lngRetVal = RegCreateKeyEx(eHKeyRoot, strKeyName, 0&, "", MCREGOPTIONNONVOLATILE, KEY_READ || KEY_WRITE, 0&, lngHKey, 0&);

      // Determine the data type of the key
      switch (eDataType) {
        //' String
        case  EnumRegistryValueType.rRKREGSZ      :
          strData = varData+ vbNullChar;
          lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, eDataType, strData, strData.length());
        //' DWord
          break;

        case  EnumRegistryValueType.rRKREGDWORD   :
          lngData = varData;
          lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, eDataType, lngData, lngData.length());
        //' Binary
          break;

        case  EnumRegistryValueType.rRKREGBINARY  :
          abytData = varData;
          lngRetVal = RegSetValueExBinary(lngHKey, strValueName, 0&, eDataType, VarPtr(abytData[0]), abytData.length + 1);
          //Case Else
          // Do nothing
          break;
      }

      RegCloseKey(lngHKey);
      // **TODO:** label found: PROC_EXIT:;
  //' disable error handler
  }
  //*TODO:** the error label PROC_ERR: couldn't be found
    try {
    // Clean-up statements go here
    if (modGeneral.gerhApp.getErrNum() != modResConstants.gCLNGERR_NUM_DEFAULT) {
      modGeneral.gerhApp.propagateError(MCSTRNAME+ cstrCurrentProc);
    }
    return;
    // **TODO:** label found: PROC_ERR:;
    switch (VBA.ex.Number) {
        //Case statements for expected errors go here
      case  Else:
        modGeneral.gerhApp.saveAppSpecificErr(vbObjectError + modResConstants.gCRES_NERR_CANTSAVE_REGKEY, MCSTRNAME+ cstrCurrentProc, strValueName+ "="+ CStr(varData), strKeyName);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }
}
}

private class FILETIME {
    public Long dwLowDateTime;
    public Long dwHighDateTime;
}


public class EnumRegistryRootKeys {
    public static final int RRKHKEY_CLASSES_ROOT = 0x80000000;
    public static final int RRKHKEY_CURRENT_USER = 0x80000001;
    public static final int RRKHKEY_LOCAL_MACHINE = 0x80000002;
    public static final int RRKHKEY_USERS = 0x80000003;
}


public class EnumRegistryValueType {
    public static final int RRKREGSZ = 1;
    public static final int RRKREGBINARY = 3;
    public static final int RRKREGDWORD = 4;
}




// Controller 

package controllers.logged.modules.general

import controllers._
import play.api.mvc._
import play.api.data._
import play.api.data.Forms._
import actions._
import play.api.Logger
import play.api.libs.json._
import models.cairo.modules.general._
import models.cairo.system.security.CairoSecurity
import models.cairo.system.database.DBHelper


case class OdregistryData(
              id: Option[Int],

              )

object Odregistrys extends Controller with ProvidesUser {

  val odregistryForm = Form(
    mapping(
      "id" -> optional(number),

  )(OdregistryData.apply)(OdregistryData.unapply))

  implicit val odregistryWrites = new Writes[Odregistry] {
    def writes(odregistry: Odregistry) = Json.obj(
      "id" -> Json.toJson(odregistry.id),
      C.ID -> Json.toJson(odregistry.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODREGISTRY), { user =>
      Ok(Json.toJson(Odregistry.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in odregistrys.update")
    odregistryForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odregistry => {
        Logger.debug(s"form: ${odregistry.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODREGISTRY), { user =>
          Ok(
            Json.toJson(
              Odregistry.update(user,
                Odregistry(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in odregistrys.create")
    odregistryForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      odregistry => {
        Logger.debug(s"form: ${odregistry.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODREGISTRY), { user =>
          Ok(
            Json.toJson(
              Odregistry.create(user,
                Odregistry(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in odregistrys.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODREGISTRY), { user =>
      Odregistry.delete(user, id)
      // Backbonejs requires at least an empty json object in the response
      // if not it will call errorHandler even when we responded with 200 OK :P
      Ok(JsonUtil.emptyJson)
    })
  }

}

// Model

package models.cairo.modules.general

import java.sql.{Connection, CallableStatement, ResultSet, Types, SQLException}
import anorm.SqlParser._
import anorm._
import services.DateUtil
import services.db.DB
import models.cairo.system.database.{DBHelper, Register, Field, FieldType, SaveResult}
import play.api.Play.current
import models.domain.CompanyUser
import java.util.Date
import play.api.Logger
import play.api.libs.json._
import scala.util.control.NonFatal

case class Odregistry(
              id: Int,
,
              createdAt: Date,
              updatedAt: Date,
              updatedBy: Int) {

  def this(
      id: Int,
) = {

    this(
      id,
,
      DateUtil.currentTime,
      DateUtil.currentTime,
      DBHelper.NoId)
  }

  def this(
) = {

    this(
      DBHelper.NoId,
)

  }

}

object Odregistry {

  lazy val emptyOdregistry = Odregistry(
)

  def apply(
      id: Int,
) = {

    new Odregistry(
      id,
)
  }

  def apply(
) = {

    new Odregistry(
)
  }

  private val odregistryParser: RowParser[Odregistry] = {
      SqlParser.get[Int](C.ID) ~
      SqlParser.get[Date](DBHelper.CREATED_AT) ~
      SqlParser.get[Date](DBHelper.UPDATED_AT) ~
      SqlParser.get[Int](DBHelper.UPDATED_BY) map {
      case
              id ~
 ~
              createdAt ~
              updatedAt ~
              updatedBy =>
        Odregistry(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, odregistry: Odregistry): Odregistry = {
    save(user, odregistry, true)
  }

  def update(user: CompanyUser, odregistry: Odregistry): Odregistry = {
    save(user, odregistry, false)
  }

  private def save(user: CompanyUser, odregistry: Odregistry, isNew: Boolean): Odregistry = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODREGISTRY}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODREGISTRY,
        C.ID,
        odregistry.id,
        false,
        true,
        true,
        getFields),
      isNew,
      C.CODE
    ) match {
      case SaveResult(true, id) => load(user, id).getOrElse(throwException)
      case SaveResult(false, id) => throwException
    }
  }

  def load(user: CompanyUser, id: Int): Option[Odregistry] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODREGISTRY} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(odregistryParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODREGISTRY} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODREGISTRY}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Odregistry = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOdregistry
    }
  }
}


// Router

GET     /api/v1/general/odregistry/:id              controllers.logged.modules.general.Odregistrys.get(id: Int)
POST    /api/v1/general/odregistry                  controllers.logged.modules.general.Odregistrys.create
PUT     /api/v1/general/odregistry/:id              controllers.logged.modules.general.Odregistrys.update(id: Int)
DELETE  /api/v1/general/odregistry/:id              controllers.logged.modules.general.Odregistrys.delete(id: Int)




/**/
