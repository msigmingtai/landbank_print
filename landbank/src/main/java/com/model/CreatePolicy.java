package com.model;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import com.control.SortDatas;
import com.control.UnitClass;



public class CreatePolicy
{
    protected List<DataEntity> datas=new ArrayList<DataEntity>();
    protected InputStreamReader inFile;
    protected BufferedReader brFile;
    protected String Line=new String();

    public List<?> GetDatas(File file) throws IOException
    {
	inFile=new InputStreamReader(new FileInputStream(file),"ISO8859_1");
	brFile=new BufferedReader(inFile);
	while ((Line=brFile.readLine())!=null)
	{
	    DataEntity ply=new DataEntity();
	    ply.setPLY_NO(Line.substring(0,10));
	    ply.setSEQ_NO(Line.substring(10,20));
	    ply.setAG_EMP(Line.substring(20,30));
	    ply.setINSURED(new String(Line.substring(30,60).getBytes("ISO8859_1"),"MS950"));
	    ply.setFASR_NAME(new String(Line.substring(60,90).getBytes("ISO8859_1"),"MS950"));
	    ply.setAREA_NO(Line.substring(90,95));
	    ply.setLOCATION_1(new String(Line.substring(95,195).getBytes("ISO8859_1"),"MS950"));
	    ply.setADDR_AREA(Line.substring(195,200));
	    ply.setADDR(new String(Line.substring(200,260).getBytes("ISO8859_1"),"MS950"));
	    ply.setCONST_CLASS(Line.substring(260,261));
	    ply.setCONST_FLOOR(Line.substring(261,264));
	    ply.setCONST_SIZE(Line.substring(264,271));
	    ply.setN_E_DAY(Line.substring(271,278));
	    ply.setEXP_DAY(Line.substring(278,285));
	    ply.setBUILD_AMT(Line.substring(285,296));
	    ply.setBUILD_PREM(Line.substring(296,305));
	    ply.setINB_AMT(Line.substring(305,316));
	    ply.setINB_PREM(Line.substring(316,325));
	    ply.setEQ_AMT(Line.substring(325,336));
	    ply.setEQ_PREM(Line.substring(336,345));
	    ply.setTOT_PREM(Line.substring(345,354));
	    ply.setISSU_DAY(Line.substring(354,361));
	    ply.setID_NO(Line.substring(361,371));
	    ply.setTEL(Line.substring(371,386));
	    ply.setST_UNIT(Line.substring(386,388));
	    ply.setMTG(Line.substring(388,395));
	    ply.setSTRU_YEAR(Line.substring(395,398));
	    ply.setCOMIS(Line.substring(398,405));
	    ply.setEMP(Line.substring(405,410));
	    ply.setS_EMP(Line.substring(410,415));
	    ply.setBROKER(Line.substring(415,417));
	    ply.setBANK_KEY1(Line.substring(417,427));
	    ply.setBANK_KEY2(Line.substring(427,443));
	    ply.setFIRE_RATE(Line.substring(443,448));
	    ply.setOCCUPY(new String(Line.substring(448,640).getBytes("ISO8859_1"),"MS950"));
	    ply.setOCCU_CODE(Line.substring(460,470));
	    ply.setCOMM_RATE(Line.substring(470,475));
	    ply.setB_AMT(Line.substring(475,486));
	    ply.setB_PREM(Line.substring(486,495));
	    ply.setORG_PLY(Line.substring(495,505));
	    ply.setFASR_ID(Line.substring(505,515));
	    ply.setMEMO(Line.substring(515,587));
	    ply.setProject_flag(Line.substring(587,630));
	    ply.setEmp_no_id(Line.substring(630,640));
	    ply.setEmp_no_name(new String(Line.substring(640,660).getBytes("ISO8859_1"),"MS950"));
	    ply.setUnit(new UnitClass().GetUnit(Line.substring(0,2),Line.substring(386,388)));
	    datas.add(ply);
	}
	SortDatas sort=new SortDatas();
	sort.setPolicys(datas);
	return sort.getPolicys();
    }
    /**
     * @param args
     */

}
