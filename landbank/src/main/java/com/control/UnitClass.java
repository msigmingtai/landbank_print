package com.control;

/*
 *取得保單單位別 
 *分28個單位 
 * 
 */
public class UnitClass
{
    private final static String[] UnitArray= { "00台北", "01桃園", "01中壢", "02台中", "02豐原", "03台南", "04北高", "04鳳山", "04南高", "05彰化", "05員林", "06羅東", 
	                                       "13嘉義", "32新竹", "36屏東", "43苗栗", "45斗六", "46新營", "51板橋", "51土城", "51中和", "55新莊", 
 	                                       "55三重", "85中山", "85北投", "85南港", "85新店", "98花蓮", "99南投","55樹林" };

    public String GetUnit(String unit, String blunit)
    {

	if(unit.equals("00"))
	{
	    return UnitArray[0].toString();
	}
	else if(unit.equals("01"))
	{
	    if(blunit.equals("FP")||blunit.equals("40")||blunit.equals("FN"))
	    {
		return UnitArray[2].toString();
	    }
	    else
	    {
		return UnitArray[1].toString();
	    }
	}
	else if(unit.equals("02"))
	{
	    if(blunit.equals("GP")||blunit.equals("57"))
	    {
		return UnitArray[4].toString();
	    }
	    else
	    {
		return UnitArray[3].toString();
	    }
	}
	else if(unit.equals("03"))
	{
	    return UnitArray[5].toString();
	}
	else if(unit.equals("04"))
	{
	    if(blunit.equals("18")||blunit.equals("60")||blunit.equals("JT")||blunit.equals("5F")||blunit.equals("JW")||blunit.equals("JD")||blunit.equals("JA"))
	    {
		return UnitArray[6].toString();
	    }
	    else if(blunit.equals("JS"))
	    {
		return UnitArray[7].toString();
	    }
	    else
	    {
		return UnitArray[8].toString();
	    }
	}
	else if(unit.equals("05"))
	{
	    if(blunit.equals("KZ"))
	    {
		return UnitArray[10].toString();
	    }
	    else
	    {
		return UnitArray[9].toString();
	    }
	}
	else if(unit.equals("06"))
	{
	    return UnitArray[11].toString();
	}
	else if(unit.equals("13"))
	{
	    return UnitArray[12].toString();
	}
	else if(unit.equals("32"))
	{
	    return UnitArray[13].toString();
	}
	else if(unit.equals("36"))
	{
	    return UnitArray[14].toString();
	}
	else if(unit.equals("43"))
	{
	    return UnitArray[15].toString();
	}
	else if(unit.equals("45"))
	{
	    return UnitArray[16].toString();
	}
	else if(unit.equals("46"))
	{
	    return UnitArray[17].toString();
	}
	else if(unit.equals("51"))
	{
	    if(blunit.equals("RR"))
	    {
		return UnitArray[20].toString();
	    }
	    else if(blunit.equals("RF"))
	    {
		return UnitArray[18].toString();
	    }
	    else
	    {
		return UnitArray[19].toString();
	    }
	}
	else if(unit.equals("55"))
	{
	    if(blunit.equals("QF"))
	    {
		return UnitArray[21].toString();
	    }else if(blunit.equals("QN")){
		return UnitArray[29].toString();
	    }
	    else
	    {
		return UnitArray[22].toString();
	    }
	}
	else if(unit.equals("85"))
	{
	    if(blunit.equals("KW")||blunit.equals("KU"))
	    {
		return UnitArray[23].toString();
	    }
	    else
	    {
		if(blunit.equals("WT")||blunit.equals("WX")||blunit.equals("WH"))
		{
		    return UnitArray[24].toString();
		}
		else
		{
		    if(blunit.equals("UP")||blunit.equals("UQ")||blunit.equals("UW"))
		    {
			return UnitArray[25].toString();
		    }
		    else if(blunit.equals("YQ")||blunit.equals("YV"))
		    {
			return UnitArray[26].toString();
		    }
		}
	    }
	}
	else if(unit.equals("98"))
	{
	    return UnitArray[27].toString();
	}
	else if(unit.equals("99"))
	{
	    return UnitArray[28].toString();
	}

	return UnitArray[0].toString();
    }

}
