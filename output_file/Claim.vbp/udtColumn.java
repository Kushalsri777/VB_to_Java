public class udtColumn {
    public String colName;//' Corresponds to COLUMN_NAME meta data from Column schema info
    public DataTypeEnum dataType;//' Corresponds to DATA_TYPE meta data from Column schema info
    public Boolean isKey;//' Corresponds to XXX from PrimaryKeys schema info
    public Boolean isNullable;//' Corresponds to IS_NULLABLE  meta data from Column schema info
    public Boolean hasDefault;//' Corresponds to COLUMN_HASDEFAULT meta data from Column schema info
    public Variant defaultValue;//' Corresponds to COLUMN_DEFAULT meta data from Column schema info
    public Integer dollarPositions;//' Calculated from PRECISION meta data from Column schema info, but which could be overriden
    public Integer decimalPositions;//' Corresponds to SCALE meta data from Column schema info, but which could be overriden
    public Integer precision;//' Corresponds to original PRECISION from DBMS. SHOULD NOT be overriden!
    public Integer numericScale;//' Corresponds to original SCALE from DBMS. SHOULD NOT be overriden!
    public Integer maxCharacters;//' Correspond to CHARACTER_MAXIMUM_LENGTH meta data from Column schema info
    public String format;//' Initially set based on DataType, but form can override
    public String mask;//' Initially set based on DataType, but form can override
    public String allowableCharacters;//' Initially set based on DataType and DecimalPositions, but form can override
    public Boolean shouldForceToUppercase;//' Does *not* correspond to DBMS meta data.
    public Variant value;//' Initially set based on DefaultValue, if present, but form can override
}

