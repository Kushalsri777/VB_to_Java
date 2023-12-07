public class StateInfo {
    public String lobCd;//' not null
    public String stCd;//' not null
    public Date strlEffDt;//' not null
    public String calcIdtypCd;//' not null
    public String reqdIdtypCd;//' not null
    public String iruleCd;//' not null
    public Variant strlEndDt;//' nullable
    public Currency strlIntRptgFlrAmt;//' not null  decimal(11,2)
    public Integer strlIntCalcOfstNum;//' not null  smallint
    public Integer strlIntReqdOfstNum;//' not null  smalling
    public Variant strlIntRuleAmt;//' nullable  decimal(11,5)
    public String strlSpclInstrTxt;//' nullable
    public Date figuredFromDate;
    public Date payablePeriodEndDate;
    public Integer nbrOfDaysToPayInterest;
    public Double interestRateToUse;
    public Currency claimInterestAmt;
    public Currency withheldAmt;
    public Currency totalForThisPayee;
    public String calculationInfo;
}

