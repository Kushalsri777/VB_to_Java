public class modDataConversion {

  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=
  // Class       : modDataConversion
  // Description : Data Conversion-related routines
  // Source      : Total Visual SourceBook 2000
  //
  // Procedures  :
  //              fnASCIIByteToEBCDICByte(ByVal bytByte As Byte) As Byte
  //              fnAtLeast(ByVal varCheckVal As Variant, ByVal varMinVal As Variant) _
  //              fnAtMost(ByVal varCheckVal As Variant, ByVal varMaxVal As Variant) _
  //              fnBinaryStringToHexString(strBinary As String) As String
  //              fnConvertLongToBinaryString(ByVal lngNumber As Long, _
  //              fnConvertVarToRoundLong(varIn As Variant) As Long
  //              fnCurrencyToText(ByVal dblAmount As Double, _
  //              fnDelimitSendKeys(ByVal strIn As String) As String
  //              fnEBCDICByteToASCIIByte(ByVal bytByte As Byte) As Byte
  //              fnNullIfZero(ByVal dblTest As Double) As Variant
  //              fnNullIfZeroOrAll(ByVal varTest As Variant) As Variant
  //              fnNullIfZLS(ByVal varIn As Variant, _
  //              fnNullIfZLSOrAll(ByVal varIn As Variant, _
  //              fnNumberToRoman(ByVal intIn As Integer) As String
  //              fnOctalStringToDecimal(ByVal strOctal As String) As Long
  //              fnOverpunchedStringToNumber(ByVal strNum As String, _
  //              fnPhoneLetterToDigit(ByVal chrIn As String) As Integer
  //              fnRGBToHTMLColor(ByVal lngRGB As Long) As String
  //              fnVarToCurrency(ByVal varIn As Variant) As Currency
  //              fnVarToDouble(varIn As Variant) As Double
  //              fnVarToInteger(ByVal varIn As Variant) As Integer
  //              fnVarToLong(ByVal varIn As Variant) As Long
  //              fnVarToString(varIn As Variant) As String
  //              fnZeroIfNull(ByVal varTest As Variant) As Double
  //              fnZLSIfNull(ByVal varTest As Variant) As String'
  //
  // Modified:
  //
  //   Version Date     Who   What
  //   ------- -------- ---   -------------------------------------------------------------------
  //   1.0     04/04/02 BAW   (Phase2C) Added to project, updated comments & error handlers.
  // =-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=

//Option Explicit
  *Option Compare Binary
  private static final String MCSTRNAME = "modDataConversion.";


  //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static byte fnASCIIByteToEBCDICByte(byte bytByte) {
    byte _rtn = 0;
    // Comments  : Converts between ASCII character values and
    //             EBCDIC character values
    // Parameters: bytByte - The ASCII value to convert
    // Returns   : The converted value
    // Source    : Total Visual SourceBook 2000
    //
    byte[255] abytLookup(255) = null;
    "fnASCIIByteToEBCDICByte"
.equals(Const cstrCurrentProc As String);

    try {

      // There is no clean formula for iteratively calculating the byte
      // values for this conversion. Because of this, we build the
      // conversion lookup table into an array.
      abytLookup[0] = 0;
      abytLookup[1] = 1;
      abytLookup[2] = 2;
      abytLookup[3] = 3;
      abytLookup[4] = 55;
      abytLookup[5] = 45;
      abytLookup[6] = 46;
      abytLookup[7] = 47;
      abytLookup[8] = 22;
      abytLookup[9] = 5;
      abytLookup[10] = 37;
      abytLookup[11] = 11;
      abytLookup[12] = 12;
      abytLookup[13] = 13;
      abytLookup[14] = 14;
      abytLookup[15] = 15;
      abytLookup[16] = 16;
      abytLookup[17] = 17;
      abytLookup[18] = 18;
      abytLookup[19] = 19;
      abytLookup[20] = 60;
      abytLookup[21] = 61;
      abytLookup[22] = 50;
      abytLookup[23] = 38;
      abytLookup[24] = 24;
      abytLookup[25] = 25;
      abytLookup[26] = 63;
      abytLookup[27] = 39;
      abytLookup[28] = 28;
      abytLookup[29] = 29;
      abytLookup[30] = 30;
      abytLookup[31] = 31;
      abytLookup[32] = 64;
      abytLookup[33] = 79;
      abytLookup[34] = 127;
      abytLookup[35] = 123;
      abytLookup[36] = 91;
      abytLookup[37] = 108;
      abytLookup[38] = 80;
      abytLookup[39] = 125;
      abytLookup[40] = 77;
      abytLookup[41] = 93;
      abytLookup[42] = 92;
      abytLookup[43] = 78;
      abytLookup[44] = 107;
      abytLookup[45] = 96;
      abytLookup[46] = 75;
      abytLookup[47] = 97;
      abytLookup[48] = 240;
      abytLookup[49] = 241;
      abytLookup[50] = 242;
      abytLookup[51] = 243;
      abytLookup[52] = 244;
      abytLookup[53] = 245;
      abytLookup[54] = 246;
      abytLookup[55] = 247;
      abytLookup[56] = 248;
      abytLookup[57] = 249;
      abytLookup[58] = 122;
      abytLookup[59] = 94;
      abytLookup[60] = 76;
      abytLookup[61] = 126;
      abytLookup[62] = 110;
      abytLookup[63] = 111;
      abytLookup[64] = 124;
      abytLookup[65] = 193;
      abytLookup[66] = 194;
      abytLookup[67] = 195;
      abytLookup[68] = 196;
      abytLookup[69] = 197;
      abytLookup[70] = 198;
      abytLookup[71] = 199;
      abytLookup[72] = 200;
      abytLookup[73] = 201;
      abytLookup[74] = 209;
      abytLookup[75] = 210;
      abytLookup[76] = 211;
      abytLookup[77] = 212;
      abytLookup[78] = 213;
      abytLookup[79] = 214;
      abytLookup[80] = 215;
      abytLookup[81] = 216;
      abytLookup[82] = 217;
      abytLookup[83] = 226;
      abytLookup[84] = 227;
      abytLookup[85] = 228;
      abytLookup[86] = 229;
      abytLookup[87] = 230;
      abytLookup[88] = 231;
      abytLookup[89] = 232;
      abytLookup[90] = 233;
      abytLookup[91] = 74;
      abytLookup[92] = 224;
      abytLookup[93] = 90;
      abytLookup[94] = 95;
      abytLookup[95] = 109;
      abytLookup[96] = 121;
      abytLookup[97] = 129;
      abytLookup[98] = 130;
      abytLookup[99] = 131;
      abytLookup[100] = 132;
      abytLookup[101] = 133;
      abytLookup[102] = 134;
      abytLookup[103] = 135;
      abytLookup[104] = 136;
      abytLookup[105] = 137;
      abytLookup[106] = 145;
      abytLookup[107] = 146;
      abytLookup[108] = 147;
      abytLookup[109] = 148;
      abytLookup[110] = 149;
      abytLookup[111] = 150;
      abytLookup[112] = 151;
      abytLookup[113] = 152;
      abytLookup[114] = 153;
      abytLookup[115] = 162;
      abytLookup[116] = 163;
      abytLookup[117] = 164;
      abytLookup[118] = 165;
      abytLookup[119] = 166;
      abytLookup[120] = 167;
      abytLookup[121] = 168;
      abytLookup[122] = 169;
      abytLookup[123] = 192;
      abytLookup[124] = 106;
      abytLookup[125] = 208;
      abytLookup[126] = 161;
      abytLookup[127] = 7;
      abytLookup[128] = 32;
      abytLookup[129] = 33;
      abytLookup[130] = 34;
      abytLookup[131] = 35;
      abytLookup[132] = 36;
      abytLookup[133] = 21;
      abytLookup[134] = 6;
      abytLookup[135] = 23;
      abytLookup[136] = 40;
      abytLookup[137] = 41;
      abytLookup[138] = 42;
      abytLookup[139] = 43;
      abytLookup[140] = 44;
      abytLookup[141] = 9;
      abytLookup[142] = 10;
      abytLookup[143] = 27;
      abytLookup[144] = 48;
      abytLookup[145] = 49;
      abytLookup[146] = 26;
      abytLookup[147] = 51;
      abytLookup[148] = 52;
      abytLookup[149] = 53;
      abytLookup[150] = 54;
      abytLookup[151] = 8;
      abytLookup[152] = 56;
      abytLookup[153] = 57;
      abytLookup[154] = 58;
      abytLookup[155] = 59;
      abytLookup[156] = 4;
      abytLookup[157] = 20;
      abytLookup[158] = 62;
      abytLookup[159] = 225;
      abytLookup[160] = 65;
      abytLookup[161] = 66;
      abytLookup[162] = 67;
      abytLookup[163] = 68;
      abytLookup[164] = 69;
      abytLookup[165] = 70;
      abytLookup[166] = 71;
      abytLookup[167] = 72;
      abytLookup[168] = 73;
      abytLookup[169] = 81;
      abytLookup[170] = 82;
      abytLookup[171] = 83;
      abytLookup[172] = 84;
      abytLookup[173] = 85;
      abytLookup[174] = 86;
      abytLookup[175] = 87;
      abytLookup[176] = 88;
      abytLookup[177] = 89;
      abytLookup[178] = 98;
      abytLookup[179] = 99;
      abytLookup[180] = 100;
      abytLookup[181] = 101;
      abytLookup[182] = 102;
      abytLookup[183] = 103;
      abytLookup[184] = 104;
      abytLookup[185] = 105;
      abytLookup[186] = 112;
      abytLookup[187] = 113;
      abytLookup[188] = 114;
      abytLookup[189] = 115;
      abytLookup[190] = 116;
      abytLookup[191] = 117;
      abytLookup[192] = 118;
      abytLookup[193] = 119;
      abytLookup[194] = 120;
      abytLookup[195] = 128;
      abytLookup[196] = 138;
      abytLookup[197] = 139;
      abytLookup[198] = 140;
      abytLookup[199] = 141;
      abytLookup[200] = 142;
      abytLookup[201] = 143;
      abytLookup[202] = 144;
      abytLookup[203] = 154;
      abytLookup[204] = 155;
      abytLookup[205] = 156;
      abytLookup[206] = 157;
      abytLookup[207] = 158;
      abytLookup[208] = 159;
      abytLookup[209] = 160;
      abytLookup[210] = 170;
      abytLookup[211] = 171;
      abytLookup[212] = 172;
      abytLookup[213] = 173;
      abytLookup[214] = 174;
      abytLookup[215] = 175;
      abytLookup[216] = 176;
      abytLookup[217] = 177;
      abytLookup[218] = 178;
      abytLookup[219] = 179;
      abytLookup[220] = 180;
      abytLookup[221] = 181;
      abytLookup[222] = 182;
      abytLookup[223] = 183;
      abytLookup[224] = 184;
      abytLookup[225] = 185;
      abytLookup[226] = 186;
      abytLookup[227] = 187;
      abytLookup[228] = 188;
      abytLookup[229] = 189;
      abytLookup[230] = 190;
      abytLookup[231] = 191;
      abytLookup[232] = 202;
      abytLookup[233] = 203;
      abytLookup[234] = 204;
      abytLookup[235] = 205;
      abytLookup[236] = 206;
      abytLookup[237] = 207;
      abytLookup[238] = 218;
      abytLookup[239] = 219;
      abytLookup[240] = 220;
      abytLookup[241] = 221;
      abytLookup[242] = 222;
      abytLookup[243] = 223;
      abytLookup[244] = 234;
      abytLookup[245] = 235;
      abytLookup[246] = 236;
      abytLookup[247] = 237;
      abytLookup[248] = 238;
      abytLookup[249] = 239;
      abytLookup[250] = 250;
      abytLookup[251] = 251;
      abytLookup[252] = 252;
      abytLookup[253] = 253;
      abytLookup[254] = 254;
      abytLookup[255] = 255;

      // After the table is initialized, the return value is a simple lookup
      _rtn = abytLookup[bytByte];
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnAtLeast(Object varCheckVal, Object varMinVal) {
    Object _rtn = null;
    // Comments  : Returns the greater of two values
    // Parameters: varCheckVal - value to check
    //             varMinVal - minimum allowed value
    // Returns   : If the value is less than the minimum,
    //             the minimum is returned. Otherwise the
    //             original value is returned
    // Source    : Total Visual SourceBook 2000
    //
    "fnAtLeast"
.equals(Const cstrCurrentProc As String);
    try {

      if (varCheckVal < varMinVal) {
        _rtn = varMinVal;
      } 
      else {
        _rtn = varCheckVal;
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
    return _rtn;
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

  return _rtn;
}




//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnAtMost(Object varCheckVal, Object varMaxVal) {
    Object _rtn = null;
    // Comments  : Returns the lesser of two values
    // Parameters: varCheckVal - value to check
    //             varMaxVal - maximum allowed value
    // Returns   : If the value is greater than the maximum,
    //             the maximum is returned. Otherwise the
    //             original value is returned
    // Source    : Total Visual SourceBook 2000
    //
    "fnAtMost"
.equals(Const cstrCurrentProc As String);
    try {

      if (varCheckVal > varMaxVal) {
        _rtn = varMaxVal;
      } 
      else {
        _rtn = varCheckVal;
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnBinaryStringToHexString(String strBinary) {
    String _rtn = "";
    // Comments : Converts a string representation of a binary number
    // to a string representation of a hexadecimal number
    // Parameters: strBinary - string representation of the binary number
    // Returns : string representation of the hexadecimal equivalent
    // Source : Total Visual SourceBook 2000
    //
    "fnBinaryStringToHexString"
.equals(Const cstrCurrentProc As String);
    String strResult = "";
    int lngTmp = 0;
    String strTmp = "";
    int lngMask = 0;
    int lngCounter = 0;

    strTmp = (String$(16, "0") + strBinary).substring((String$(16, "0") + strBinary).length() - 16);

    if (strTmp.substring(0, 1).equals("1")) {
      lngTmp = &H8000;
    }

    lngMask = &H4000;

    for (lngCounter = 2; lngCounter <= 16; lngCounter++) {
      if (strTmp.substring(lngCounter, 1).equals("1")) {
        lngTmp = lngTmp || lngMask;
      }

      lngMask = lngMask \ 2;

    }

    strResult = Hex$(lngTmp);
    _rtn = strResult;
    // **TODO:** label found: PROC_EXIT:;
    //' disable error handler
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
          modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
          break;
      }
      /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnConvertLongToBinaryString(int lngNumber, int intDigits) {
    String _rtn = "";
    // Comments  : Converts a long number to a string representation
    //             of a binary number
    // Parameters: lngNumber - A long number to convert to binary
    //             intDigits - The minimum number of digits to return
    // Returns   : String representation of a binary number
    // Source    : Total Visual SourceBook 2000
    //
    "fnConvertLongToBinaryString"
.equals(Const cstrCurrentProc As String);
    int lngCounter = 0;
    String strTemp = "";

    try {

      // Check to see if number is negative
      if (lngNumber < 0) {
        strTemp = strTemp+ "1";
      } 
      else {
        strTemp = strTemp+ "0";
      }

      // Convert each bit into "0" or "1"
      for (lngCounter = 30; lngCounter >= 0; lngCounter--) {
        if (lngNumber && (2 ^ lngCounter)) {
          strTemp = strTemp+ "1";
        } 
        else {
          strTemp = strTemp+ "0";
        }
      }

      // Return the result
      _rtn = strTemp;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnConvertVarToRoundLong(Object varIn) {
    int _rtn = 0;
    // Comments   : Converts the passed variant to a long integer and
    //              rounds it using arithmetic rounding. Nulls are returned
    //              as 0.
    // Parameters : varIn - number to convert/round
    // Returns    : Long integer
    // Source     : Total Visual SourceBook 2000
    //
    "fnConvertVarToRoundLong"
.equals(Const cstrCurrentProc As String);
    double dbl = 0;

    try {

      if (varIn == null) {
        _rtn = 0;
      } 
      else {
        if (!IsNumeric(varIn)) {
          _rtn = 0;
        } 
        else {
          dbl = varIn + 0.5;
          _rtn = Long.parseLong(dbl);
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnCurrencyToText(double dblAmount, int intLength) {
    String _rtn = "";
    // Comments  : Converts a number to spelled out text with padding/length
    //             options.
    // Parameters: dblAmount - Dollar amount to convert
    //             intLength - Length of string to create (pads with trailing
    //             asterisks to fill to length). If text exceeds the given
    //             length, the numeric representation is given. If the
    //             intLength parameter is set to 0, the full string is
    //             returned without padding.
    // Returns   : String representation of the amount
    // Source    : Total Visual SourceBook 2000
    //
    "fnCurrencyToText"
.equals(Const cstrCurrentProc As String);
    String strText = "";
    String strCurrFormat = "";
    int intLowDigit = 0;
    String strCents = "";
    int intGroup = 0;
    int intDigit1 = 0;
    int intDigit2 = 0;
    int intDigit3 = 0;
    String strSubText = "";
    int intGroupCounter = 0;
    int intCounter = 0;

    try {

      // Create the arrays
      G.redimPreserve(1 To 19,  );
      G.redimPreserve(1 To 9,  );
      G.redimPreserve(1 To 3,  );

      // Fill the arrays.
      astrOnes(1) = "ONE";
      astrOnes(2) = "TWO";
      astrOnes(3) = "THREE";
      astrOnes(4) = "FOUR";
      astrOnes(5) = "FIVE";
      astrOnes(6) = "SIX";
      astrOnes(7) = "SEVEN";
      astrOnes(8) = "EIGHT";
      astrOnes(9) = "NINE";
      astrOnes(10) = "TEN";
      astrOnes(11) = "ELEVEN";
      astrOnes(12) = "TWELVE";
      astrOnes(13) = "THIRTEEN";
      astrOnes(14) = "FOURTEEN";
      astrOnes(15) = "FIFTEEN";
      astrOnes(16) = "SIXTEEN";
      astrOnes(17) = "SEVENTEEN";
      astrOnes(18) = "EIGHTTEEN";
      astrOnes(19) = "NINETEEN";
      astrTens(1) = "TEN";
      astrTens(2) = "TWENTY";
      astrTens(3) = "THIRTY";
      astrTens(4) = "FORTY";
      astrTens(5) = "FIFTY";
      astrTens(6) = "SIXTY";
      astrTens(7) = "SEVENTY";
      astrTens(8) = "EIGHTY";
      astrTens(9) = "NINETY";
      astrGroup(1) = "THOUSAND";
      astrGroup(2) = "MILLION";
      astrGroup(3) = "BILLION";

      // Prepare the temp variable
      strText = "";

      // Ensure amount is greater than zero
      if (dblAmount > 0) {

        // Format the string
        strCurrFormat = Format$(dblAmount, "#,###.00");

        // Get the lower digit part
        intLowDigit = strCurrFormat.indexOf(".") - 1;

        // Get the cents
        strCents = strCurrFormat.substring(intLowDigit + 2, 2);

        intGroup = 0;

        // Loop through lower digit part
        while (intLowDigit > 0) {

          intDigit3 = Integer.parseInt(strCurrFormat.substring(intLowDigit, 1));

          if (intLowDigit > 1) {
            intDigit2 = Integer.parseInt(strCurrFormat.substring(intLowDigit - 1, 1));
          } 
          else {
            intDigit2 = 0;
          }

          if (intLowDigit > 2) {
            intDigit1 = Integer.parseInt(strCurrFormat.substring(intLowDigit - 2, 1));
          } 
          else {
            intDigit1 = 0;
          }

          strSubText = "";

          // Get the hundreds
          if (intDigit1 > 0) {
            strSubText = astrOnes(intDigit1)+ " HUNDRED ";
          }

          if (intDigit2 > 0) {

            // Get the ones
            if (intDigit2 == 1) {
              strSubText = strSubText+ String.valueOf(astrOnes(intDigit3 + 10))+ " ";
            } 
            else {
              strSubText = strSubText+ astrTens(intDigit2);

              if (intDigit3 > 0) {
                strSubText = strSubText+ "-"+ astrOnes(intDigit3);
              }

              strSubText = strSubText+ " ";
            }

          } 
          else {

            if (intDigit3 > 0) {
              strSubText = strSubText+ astrOnes(intDigit3)+ " ";
            }

          }

          // Get the grouping
          if (!(strSubText.isEmpty()) && intGroupCounter != 0) {
            strSubText = strSubText+ astrGroup(intGroupCounter)+ " ";
          }

          // Concatenate the temp vars
          strText = strSubText+ strText;

          // Move back through the number
          intLowDigit = intLowDigit - 4;

          // Increment the counter
          intGroupCounter = intGroupCounter + 1;

        }

        // Finalize the text
        strText = strText + "& " + strCents + "/100";

        // Replace the place holder with "NO" cents string
        if (strText.substring(0, 1).equals("&")) {
          strText = "NO " + strText;
        }

        // Cleanup and pad
        if (intLength > 0) {
          if (strText.length() > intLength) {
            strText = strCurrFormat;
          } 
          else {
            for (intCounter = 1; intCounter <= (intLength - strText.length()); intCounter++) {
              strText = strText + "*";
            }
          }
        }
      }

      // Return the result
      _rtn = strText;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnDelimitSendKeys(String strIn) {
    String _rtn = "";
    // Comments  : Fixes sendkeys statements by delimiting magic characters
    // Parameters: strIn - string to fix
    // Returns   : Fixed string
    // Source    : Total Visual SourceBook 2000
    //
    "fnDelimitSendKeys"
.equals(Const cstrCurrentProc As String);
    String strTmp = "";
    int intCounter = 0;
    String chrTmp = ""; * 1

    try {

      for (intCounter = 1; intCounter <= strIn.length(); intCounter++) {

        chrTmp = strIn.substring(intCounter, 1);

        if (chrTmp.equals("+") || chrTmp.equals("^") || chrTmp.equals("%") || chrTmp.equals("~")) {
          strTmp = strTmp+ "{"+ chrTmp+ "}";
        } 
        else {
          strTmp = strTmp+ chrTmp;
        }

      }

      _rtn = strTmp;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static byte fnEBCDICByteToASCIIByte(byte bytByte) {
    byte _rtn = 0;
    // Comments  : Converts between EBCDIC character values and
    //             ASCII character values
    // Parameters: bytByte - The EBCDIC value to convert
    // Returns   : The converted value
    // Source    : Total Visual SourceBook 2000
    //
    // This is the lookup table used to do the conversion
    "fnEBCDICByteToASCIIByte"
.equals(Const cstrCurrentProc As String);
    byte[255] abytLookup(255) = null;

    try {

      // There is no clean formula for iteratively calculating the byte
      // values for this conversion. Because of this, we build the
      // conversion lookup table into an array.
      abytLookup[0] = 0;
      abytLookup[1] = 1;
      abytLookup[2] = 2;
      abytLookup[3] = 3;
      abytLookup[4] = 156;
      abytLookup[5] = 9;
      abytLookup[6] = 134;
      abytLookup[7] = 127;
      abytLookup[8] = 151;
      abytLookup[9] = 141;
      abytLookup[10] = 142;
      abytLookup[11] = 11;
      abytLookup[12] = 12;
      abytLookup[13] = 13;
      abytLookup[14] = 14;
      abytLookup[15] = 15;
      abytLookup[16] = 16;
      abytLookup[17] = 17;
      abytLookup[18] = 18;
      abytLookup[19] = 19;
      abytLookup[20] = 157;
      abytLookup[21] = 133;
      abytLookup[22] = 8;
      abytLookup[23] = 135;
      abytLookup[24] = 24;
      abytLookup[25] = 25;
      abytLookup[26] = 146;
      abytLookup[27] = 143;
      abytLookup[28] = 28;
      abytLookup[29] = 29;
      abytLookup[30] = 30;
      abytLookup[31] = 31;
      abytLookup[32] = 128;
      abytLookup[33] = 129;
      abytLookup[34] = 130;
      abytLookup[35] = 131;
      abytLookup[36] = 132;
      abytLookup[37] = 10;
      abytLookup[38] = 23;
      abytLookup[39] = 27;
      abytLookup[40] = 136;
      abytLookup[41] = 137;
      abytLookup[42] = 138;
      abytLookup[43] = 139;
      abytLookup[44] = 140;
      abytLookup[45] = 5;
      abytLookup[46] = 6;
      abytLookup[47] = 7;
      abytLookup[48] = 144;
      abytLookup[49] = 145;
      abytLookup[50] = 22;
      abytLookup[51] = 147;
      abytLookup[52] = 148;
      abytLookup[53] = 149;
      abytLookup[54] = 150;
      abytLookup[55] = 4;
      abytLookup[56] = 152;
      abytLookup[57] = 153;
      abytLookup[58] = 154;
      abytLookup[59] = 155;
      abytLookup[60] = 20;
      abytLookup[61] = 21;
      abytLookup[62] = 158;
      abytLookup[63] = 26;
      abytLookup[64] = 32;
      abytLookup[65] = 160;
      abytLookup[66] = 161;
      abytLookup[67] = 162;
      abytLookup[68] = 163;
      abytLookup[69] = 164;
      abytLookup[70] = 165;
      abytLookup[71] = 166;
      abytLookup[72] = 167;
      abytLookup[73] = 168;
      abytLookup[74] = 91;
      abytLookup[75] = 46;
      abytLookup[76] = 60;
      abytLookup[77] = 40;
      abytLookup[78] = 43;
      abytLookup[79] = 33;
      abytLookup[80] = 38;
      abytLookup[81] = 169;
      abytLookup[82] = 170;
      abytLookup[83] = 171;
      abytLookup[84] = 172;
      abytLookup[85] = 173;
      abytLookup[86] = 174;
      abytLookup[87] = 175;
      abytLookup[88] = 176;
      abytLookup[89] = 177;
      abytLookup[90] = 93;
      abytLookup[91] = 36;
      abytLookup[92] = 42;
      abytLookup[93] = 41;
      abytLookup[94] = 59;
      abytLookup[95] = 94;
      abytLookup[96] = 45;
      abytLookup[97] = 47;
      abytLookup[98] = 178;
      abytLookup[99] = 179;
      abytLookup[100] = 180;
      abytLookup[101] = 181;
      abytLookup[102] = 182;
      abytLookup[103] = 183;
      abytLookup[104] = 184;
      abytLookup[105] = 185;
      abytLookup[106] = 124;
      abytLookup[107] = 44;
      abytLookup[108] = 37;
      abytLookup[109] = 95;
      abytLookup[110] = 62;
      abytLookup[111] = 63;
      abytLookup[112] = 186;
      abytLookup[113] = 187;
      abytLookup[114] = 188;
      abytLookup[115] = 189;
      abytLookup[116] = 190;
      abytLookup[117] = 191;
      abytLookup[118] = 192;
      abytLookup[119] = 193;
      abytLookup[120] = 194;
      abytLookup[121] = 96;
      abytLookup[122] = 58;
      abytLookup[123] = 35;
      abytLookup[124] = 64;
      abytLookup[125] = 39;
      abytLookup[126] = 61;
      abytLookup[127] = 34;
      abytLookup[128] = 195;
      abytLookup[129] = 97;
      abytLookup[130] = 98;
      abytLookup[131] = 99;
      abytLookup[132] = 100;
      abytLookup[133] = 101;
      abytLookup[134] = 102;
      abytLookup[135] = 103;
      abytLookup[136] = 104;
      abytLookup[137] = 105;
      abytLookup[138] = 196;
      abytLookup[139] = 197;
      abytLookup[140] = 198;
      abytLookup[141] = 199;
      abytLookup[142] = 200;
      abytLookup[143] = 201;
      abytLookup[144] = 202;
      abytLookup[145] = 106;
      abytLookup[146] = 107;
      abytLookup[147] = 108;
      abytLookup[148] = 109;
      abytLookup[149] = 110;
      abytLookup[150] = 111;
      abytLookup[151] = 112;
      abytLookup[152] = 113;
      abytLookup[153] = 114;
      abytLookup[154] = 203;
      abytLookup[155] = 204;
      abytLookup[156] = 205;
      abytLookup[157] = 206;
      abytLookup[158] = 207;
      abytLookup[159] = 208;
      abytLookup[160] = 209;
      abytLookup[161] = 126;
      abytLookup[162] = 115;
      abytLookup[163] = 116;
      abytLookup[164] = 117;
      abytLookup[165] = 118;
      abytLookup[166] = 119;
      abytLookup[167] = 120;
      abytLookup[168] = 121;
      abytLookup[169] = 122;
      abytLookup[170] = 210;
      abytLookup[171] = 211;
      abytLookup[172] = 212;
      abytLookup[173] = 213;
      abytLookup[174] = 214;
      abytLookup[175] = 215;
      abytLookup[176] = 216;
      abytLookup[177] = 217;
      abytLookup[178] = 218;
      abytLookup[179] = 219;
      abytLookup[180] = 220;
      abytLookup[181] = 221;
      abytLookup[182] = 222;
      abytLookup[183] = 223;
      abytLookup[184] = 224;
      abytLookup[185] = 225;
      abytLookup[186] = 226;
      abytLookup[187] = 227;
      abytLookup[188] = 228;
      abytLookup[189] = 229;
      abytLookup[190] = 230;
      abytLookup[191] = 231;
      abytLookup[192] = 123;
      abytLookup[193] = 65;
      abytLookup[194] = 66;
      abytLookup[195] = 67;
      abytLookup[196] = 68;
      abytLookup[197] = 69;
      abytLookup[198] = 70;
      abytLookup[199] = 71;
      abytLookup[200] = 72;
      abytLookup[201] = 73;
      abytLookup[202] = 232;
      abytLookup[203] = 233;
      abytLookup[204] = 234;
      abytLookup[205] = 235;
      abytLookup[206] = 236;
      abytLookup[207] = 237;
      abytLookup[208] = 125;
      abytLookup[209] = 74;
      abytLookup[210] = 75;
      abytLookup[211] = 76;
      abytLookup[212] = 77;
      abytLookup[213] = 78;
      abytLookup[214] = 79;
      abytLookup[215] = 80;
      abytLookup[216] = 81;
      abytLookup[217] = 82;
      abytLookup[218] = 238;
      abytLookup[219] = 239;
      abytLookup[220] = 240;
      abytLookup[221] = 241;
      abytLookup[222] = 242;
      abytLookup[223] = 243;
      abytLookup[224] = 92;
      abytLookup[225] = 159;
      abytLookup[226] = 83;
      abytLookup[227] = 84;
      abytLookup[228] = 85;
      abytLookup[229] = 86;
      abytLookup[230] = 87;
      abytLookup[231] = 88;
      abytLookup[232] = 89;
      abytLookup[233] = 90;
      abytLookup[234] = 244;
      abytLookup[235] = 245;
      abytLookup[236] = 246;
      abytLookup[237] = 247;
      abytLookup[238] = 248;
      abytLookup[239] = 249;
      abytLookup[240] = 48;
      abytLookup[241] = 49;
      abytLookup[242] = 50;
      abytLookup[243] = 51;
      abytLookup[244] = 52;
      abytLookup[245] = 53;
      abytLookup[246] = 54;
      abytLookup[247] = 55;
      abytLookup[248] = 56;
      abytLookup[249] = 57;
      abytLookup[250] = 250;
      abytLookup[251] = 251;
      abytLookup[252] = 252;
      abytLookup[253] = 253;
      abytLookup[254] = 254;
      abytLookup[255] = 255;

      // After the table is initialized, the return value is a simple lookup
      _rtn = abytLookup[bytByte];
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnNullIfZero(double dblTest) {
    Object _rtn = null;
    // Comments  : Returns Null if the passed value is zero, otherwise returns
    //             the passed value.
    // Parameters: lngTest - Value to test
    // Returns   : Null or passed value
    // Source    : Total Visual SourceBook 2000
    //
    "fnNullIfZero"
.equals(Const cstrCurrentProc As String);
    try {

      // CMP Modified this since it wasn't working.
      //If Len(dblTest & "") = 0 Then
      if (dblTest == 0) {
        _rtn = Null;
      } 
      else {
        _rtn = dblTest;
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnNullIfZeroOrAll(Object varTest) {
    Object _rtn = null;
    // Comments  : Returns Null if the passed value is zero or "--All--", otherwise returns
    //             the passed value.
    // Parameters: lngTest - Value to test
    // Returns   : Null or passed value
    // Source    : Total Visual SourceBook 2000
    //
    "fnNullIfZeroOrAll"
.equals(Const cstrCurrentProc As String);
    double dblResult = 0;
    try {

      // CMP modified this 6/4/02 since the old way wasn't working.
      if (varTest == 0  || (modGeneral.gCSTRALLENTRY.equals(varTest))) {
        _rtn = Null;
      } 
      else {
        _rtn = Double.parseDouble(varTest);
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnNullIfZLS(Object varIn, boolean bHandleEmbeddedQuotes) {
    Object _rtn = null;
    // Comments  : Returns Null if the passed value is a zero-length
    //             string (""), otherwise returns the passed value.
    //
    //             NOTE: If working with data to send to SQL Server to
    //                   do an Insert or Update, then you should use
    //                   fnQuotedOrNull( ) in modGeneral.bas.
    //
    // Parameters:
    //       varIn (in)                 - Value to test
    //       bHandleEmbeddedQuotes (in) - Indicates whether to replace
    //                                    single quotes with two single quotes.
    //
    // Returns   : Null or passed value
    // Source    : Total Visual SourceBook 2000
    //
    "fnNullIfZLS"
.equals(Const cstrCurrentProc As String);
    Object varResult = null;

    try {

      if ((varIn+ "").length() == 0) {
        varResult = Null;
      } 
      else {
        if (bHandleEmbeddedQuotes) {
          varResult = varIn.replace("'", "''");
        }
        varResult = varIn;
      }

      _rtn = varResult;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnNullIfZLSOrAll(Object varIn, boolean bHandleEmbeddedQuotes) {
    Object _rtn = null;
    // Comments  : Returns Null if the passed value is a zero-length
    //             string ("") or "--All--", otherwise returns the passed value.
    //
    //             NOTE: If working with data to send to SQL Server to
    //                   do an Insert or Update, then you should use
    //                   fnQuotedOrNull( ) in modGeneral.bas.
    //
    // Parameters:
    //       varIn (in)                 - Value to test
    //       bHandleEmbeddedQuotes (in) - Indicates whether to replace
    //                                    single quotes with two single quotes.
    //
    // Returns   : Null or passed value
    // Source    : Total Visual SourceBook 2000
    //
    "fnNullIfZLSOrAll"
.equals(Const cstrCurrentProc As String);
    Object varResult = null;

    try {

      if ((varIn+ "").length() == 0  || (modGeneral.gCSTRALLENTRY.equals(varIn))) {
        varResult = Null;
      } 
      else {
        if (bHandleEmbeddedQuotes) {
          varResult = varIn.replace("'", "''");
        }
        varResult = varIn;
      }

      _rtn = varResult;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}


//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnNumberToRoman(int intIn) {
    String _rtn = "";
    // Comments  : Converts the passed integer to Roman numerals
    // Parameters: intIn - Value to convert
    // Returns   : String
    // Source    : Total Visual SourceBook 2000
    //
    "fnNumberToRoman"
.equals(Const cstrCurrentProc As String);
    int intCounter = 0;
    int intDigit = 0;
    String strTmp = "";
    "IVXLCDM"
.equals(Const cstrDigits As String);

    try {

      intCounter = 1;

      // Loop through values in input value
      while (intIn > 0) {

        // Get  the current digit
        intDigit = intIn Mod 10;

        intIn = intIn \ 10;

        // Build the temp string
        switch (intDigit) {

          case  1:
            strTmp = cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  2:
            strTmp = cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  3:
            strTmp = cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  4:
            strTmp = cstrDigits.substring(intCounter, 2)+ strTmp;

            break;

          case  5:
            strTmp = cstrDigits.substring(intCounter + 1, 1)+ strTmp;

            break;

          case  6:
            strTmp = cstrDigits.substring(intCounter + 1, 1)+ cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  7:
            strTmp = cstrDigits.substring(intCounter + 1, 1)+ cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  8:
            strTmp = cstrDigits.substring(intCounter + 1, 1)+ cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter, 1)+ strTmp;

            break;

          case  9:
            strTmp = cstrDigits.substring(intCounter, 1)+ cstrDigits.substring(intCounter + 2, 1)+ strTmp;

            break;
        }
        intCounter = intCounter + 2;
      }

      _rtn = strTmp;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  public static int fnOctalStringToDecimal(String strOctal) {
    int _rtn = 0;
    // Comments   : Converts the passed string representation of an octal
    //              number to a decimal long integer.
    // Parameters : strOctal - String representation of octal number
    // Returns    : Decimal value
    // Source     : Total Visual SourceBook 2000
    //
    "fnOctalStringToDecimal"
.equals(Const cstrCurrentProc As String);
    try {

      _rtn = Val("&O"+ strOctal);
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Object fnOverpunchedStringToNumber(String strNum, int intDecimals) {
    Object _rtn = null;
    // Comments  : Converts a "zoned overpunch" number to a regular number
    // Parameters: strNum - Zoned overpunch value to convert
    //             intDecimals - Number of decimal places
    // Returns   : Converted number
    // Source    : Total Visual SourceBook 2000
    //
    "fnOverpunchedStringToNumber"
.equals(Const cstrCurrentProc As String);
    int intLen = 0;
    int intSign = 0;
    double dblOut = 0;
    String strLast = ""; * 1
    int intLast = 0;

    try {

      // Get the length of the string
      intLen = strNum.trim().length();

      // Get the last character
      strLast = strNum.substring(intLen, 1);

      // Decide how to convert the last character
      switch (strLast) {
        case  "A" To "I":
          intSign = 1;
          intLast = Asc(strLast) - 65 + 1;
          break;

        case  "J" To "R":
          intSign = -1;
          intLast = Asc(strLast) - 74 + 1;
          break;

        case  "{":
          intSign = 1;
          break;

        case  "}":
          intSign = -1;
          break;

        default:
          intSign = 1;
          intLast = 9;
          strNum = "9999999999999";
          break;
      }

      dblOut = Val(strNum.substring(0, intLen - 1)+ ((Integer) intLast).toString()) * intSign;
      _rtn = dblOut * (10 ^ -(intDecimals));
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnPhoneLetterToDigit(String chrIn) {
    int _rtn = 0;
    // Comments  : Converts a phone number letter to a number
    // Parameters: chrIn - Letter to check. Must be in the range a-p
    //             or r-y. Q and Z are not valid phone letters.
    // Returns   : Integer number
    // Source    : Total Visual SourceBook 2000
    //
    "fnPhoneLetterToDigit"
.equals(Const cstrCurrentProc As String);
    int intDigit = 0;
    String chrTmp = ""; * 1

    try {

      if (!(chrIn.isEmpty())) {

        // Trim any excess characters
        chrTmp = chrIn.substring(0, 1).toLowerCase();

        // Make sure its a letter
        if (chrTmp >= "a" && chrTmp <= "z") {

          // For historical reasons, Q is not a valid letter on a phone.
          // Z is also left out.
          if (!(chrTmp.equals("q")) && !(chrTmp.equals("z"))) {

            intDigit = Asc(chrTmp.toUpperCase());

            if (intDigit > Asc("Q")) {
              intDigit = intDigit - 1;
            }

            intDigit = (intDigit - Asc("A")) \ 3 + 2;

            _rtn = CStr(intDigit).trim();
          }
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnRGBToHTMLColor(int lngRGB) {
    String _rtn = "";
    // Comments  : Formats an RGB value into the hex format standard used
    //             in HTML.
    // Parameters: lngRGB - the RGB value, or a VB-defined constant such
    //             as 'vbRed' that evaluates to an RGB value
    // Returns   : The formatted hex value of the RGB color
    // Source    : Total Visual SourceBook 2000
    //
    "fnRGBToHTMLColor"
.equals(Const cstrCurrentProc As String);
    String strValue = "";

    try {

      // Break out individual color portions of the RGB value, and then
      // get the hex value in the format HTML expects (rrggbb)
      strValue = Hex$((lngRGB && &HFF&)&H10000 || (lngRGB && &HFF00&) || (lngRGB && &HFF0000) \ &H10000);

      // Force leading zeroes, which VB's hex function drops
      strValue = String.valueOf(String(6 - strValue.length(), "0"))+ strValue;

      _rtn = strValue;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static Currency fnVarToCurrency(Object varIn) {
    Currency _rtn = null;
    // Comments   : Converts the passed variant to a currency value,
    //              0 if the passed value is Null.
    // Parameters : varIn - Value to convert
    // Returns    : Currency
    // Source     : Total Visual SourceBook 2000
    //
    "fnVarToCurrency"
.equals(Const cstrCurrentProc As String);
    try {

      if (varIn == null) {
        _rtn = 0;
      } 
      else {
        if (!IsNumeric(varIn)) {
          _rtn = 0;
        } 
        else {
          _rtn = Double.parseDouble(varIn);
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static double fnVarToDouble(Object varIn) {
    double _rtn = 0;
    // Comments   : Converts the passed variant to a double, returning
    //              0 if the passed value is Null.
    // Parameters : varIn - Value to convert
    // Returns    : Double
    // Source     : Total Visual SourceBook 2000
    //
    "fnVarToDouble"
.equals(Const cstrCurrentProc As String);
    try {

      if (varIn == null) {
        _rtn = 0;
      } 
      else {
        if (!IsNumeric(varIn)) {
          _rtn = 0;
        } 
        else {
          _rtn = varIn;
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnVarToInteger(Object varIn) {
    int _rtn = 0;
    // Comments   : Converts the passed variant to an integer, returning
    //              0 if the passed value is Null
    // Parameters : varIn - Value to convert
    // Returns    : Integer
    // Source     : Total Visual SourceBook 2000
    //
    "fnVarToInteger"
.equals(Const cstrCurrentProc As String);
    try {

      if (varIn == null) {
        _rtn = 0;
      } 
      else {
        if (!IsNumeric(varIn)) {
          _rtn = 0;
        } 
        else {
          _rtn = varIn;
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static int fnVarToLong(Object varIn) {
    int _rtn = 0;
    // Comments   : Converts the passed variant to a long integer, returning
    //              0 if the passed value is Null
    // Parameters : varIn - Value to convert
    // Returns    : Long integer
    // Source     : Total Visual SourceBook 2000
    //
    "fnVarToLong"
.equals(Const cstrCurrentProc As String);
    try {

      if (varIn == null) {
        _rtn = 0;
      } 
      else {
        if (!IsNumeric(varIn)) {
          _rtn = 0;
        } 
        else {
          _rtn = varIn;
        }
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnVarToString(Object varIn) {
    String _rtn = "";
    // Comments  : Converts the supplied variant to a string. Nulls are returned
    //             as a zero-length string ("")
    // Parameters: varIn - Variant to convert
    // Returns   : String
    // Source    : Total Visual SourceBook 2000
    //
    "fnVarToString"
.equals(Const cstrCurrentProc As String);
    try {

      if (varIn == null) {
        _rtn = "";
      } 
      else {
        _rtn = varIn;
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
    return _rtn;
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

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static double fnZeroIfNull(Object varTest) {
    double _rtn = 0;
    // Comments  : Returns zero if Null is passed, otherwise returns the
    //             passed value.
    // Parameters: varTest - Value to test
    // Returns   : Zero or Null
    // Source    : Total Visual SourceBook 2000
    //
    "fnZeroIfNull"
.equals(Const cstrCurrentProc As String);
    double dblResult = 0;

    try {

      if (varTest == null) {
        dblResult = 0;
      } 
      else {
        dblResult = varTest;
      }

      _rtn = dblResult;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}



//////////////////////////////////////////////////////////////////////////////////////////////////
  public static String fnZLSIfNull(Object varTest) {
    String _rtn = "";
    // Comments  : Returns a zero-length string ("") if Null is passed,
    //             otherwise returns the passed value.
    // Parameters: varTest - Value to test
    // Returns   : If the value is Null, it returns a zero-length string,
    //             otherwise it returns the passed value.
    // Source    : Total Visual SourceBook 2000
    //
    "fnZLSIfNull"
.equals(Const cstrCurrentProc As String);
    Object varResult = null;

    try {

      if (varTest == null) {
        varResult = "";
      } 
      else {
        varResult = varTest;
      }

      _rtn = varResult;
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
        modGeneral.gerhApp.saveErrObjectData(MCSTRNAME+ cstrCurrentProc);
        break;
    }
    /**TODO:** resume found: Resume(PROC_EXIT)*/;
//*TODO:** the error label 0: couldn't be found
  }

  return _rtn;
}
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


case class OddataconversionData(
              id: Option[Int],

              )

object Oddataconversions extends Controller with ProvidesUser {

  val oddataconversionForm = Form(
    mapping(
      "id" -> optional(number),

  )(OddataconversionData.apply)(OddataconversionData.unapply))

  implicit val oddataconversionWrites = new Writes[Oddataconversion] {
    def writes(oddataconversion: Oddataconversion) = Json.obj(
      "id" -> Json.toJson(oddataconversion.id),
      C.ID -> Json.toJson(oddataconversion.id),

    )
  }

  def get(id: Int) = GetAction { implicit request =>
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.LIST_ODDATACONVERSION), { user =>
      Ok(Json.toJson(Oddataconversion.get(user, id)))
    })
  }

  def update(id: Int) = PostAction { implicit request =>
    Logger.debug("in oddataconversions.update")
    oddataconversionForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      oddataconversion => {
        Logger.debug(s"form: ${oddataconversion.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.EDIT_ODDATACONVERSION), { user =>
          Ok(
            Json.toJson(
              Oddataconversion.update(user,
                Oddataconversion(
                       id,

                ))))
        })
      }
    )
  }

  def create = PostAction { implicit request =>
    Logger.debug("in oddataconversions.create")
    oddataconversionForm.bindFromRequest.fold(
      formWithErrors => {
        Logger.debug(s"invalid form: ${formWithErrors.toString}")
        BadRequest
      },
      oddataconversion => {
        Logger.debug(s"form: ${oddataconversion.toString}")
        LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.NEW_ODDATACONVERSION), { user =>
          Ok(
            Json.toJson(
              Oddataconversion.create(user,
                Oddataconversion(

                ))))
        })
      }
    )
  }

  def delete(id: Int) = PostAction { implicit request =>
    Logger.debug("in oddataconversions.delete")
    LoggedIntoCompanyResponse.getAction(request, CairoSecurity.hasPermissionTo(S.DELETE_ODDATACONVERSION), { user =>
      Oddataconversion.delete(user, id)
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

case class Oddataconversion(
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

object Oddataconversion {

  lazy val emptyOddataconversion = Oddataconversion(
)

  def apply(
      id: Int,
) = {

    new Oddataconversion(
      id,
)
  }

  def apply(
) = {

    new Oddataconversion(
)
  }

  private val oddataconversionParser: RowParser[Oddataconversion] = {
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
        Oddataconversion(
              id,
,
              createdAt,
              updatedAt,
              updatedBy)
    }
  }

  def create(user: CompanyUser, oddataconversion: Oddataconversion): Oddataconversion = {
    save(user, oddataconversion, true)
  }

  def update(user: CompanyUser, oddataconversion: Oddataconversion): Oddataconversion = {
    save(user, oddataconversion, false)
  }

  private def save(user: CompanyUser, oddataconversion: Oddataconversion, isNew: Boolean): Oddataconversion = {
    def getFields = {
      List(

      )
    }
    def throwException = {
      throw new RuntimeException(s"Error when saving ${C.ODDATACONVERSION}")
    }

    DBHelper.saveEx(
      user,
      Register(
        C.ODDATACONVERSION,
        C.ID,
        oddataconversion.id,
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

  def load(user: CompanyUser, id: Int): Option[Oddataconversion] = {
    loadWhere(user, s"${C.ID} = {id}", 'id -> id)
  }

  def loadWhere(user: CompanyUser, where: String, args : scala.Tuple2[scala.Any, anorm.ParameterValue[_]]*) = {
    DB.withConnection(user.database.database) { implicit connection =>
      SQL(s"SELECT t1.*, t2.${C.FK_NAME} FROM ${C.ODDATACONVERSION} t1 INNER JOIN ${C.???} t2 ON t1.${C.FK_ID} = t2.${C.FK_ID} WHERE $where")
        .on(args: _*)
        .as(oddataconversionParser.singleOpt)
    }
  }

  def delete(user: CompanyUser, id: Int) = {
    DB.withConnection(user.database.database) { implicit connection =>
      try {
        SQL(s"DELETE FROM ${C.ODDATACONVERSION} WHERE ${C.ID} = {id}")
        .on('id -> id)
        .executeUpdate
      } catch {
        case NonFatal(e) => {
          Logger.error(s"can't delete a ${C.ODDATACONVERSION}. ${C.ID} id: $id. Error ${e.toString}")
          throw e
        }
      }
    }
  }

  def get(user: CompanyUser, id: Int): Oddataconversion = {
    load(user, id) match {
      case Some(p) => p
      case None => emptyOddataconversion
    }
  }
}


// Router

GET     /api/v1/general/oddataconversion/:id              controllers.logged.modules.general.Oddataconversions.get(id: Int)
POST    /api/v1/general/oddataconversion                  controllers.logged.modules.general.Oddataconversions.create
PUT     /api/v1/general/oddataconversion/:id              controllers.logged.modules.general.Oddataconversions.update(id: Int)
DELETE  /api/v1/general/oddataconversion/:id              controllers.logged.modules.general.Oddataconversions.delete(id: Int)




/**/
