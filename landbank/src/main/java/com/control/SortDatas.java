package com.control;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import com.model.DataEntity;



/**
 * 
 * 將policy注入此類別後排序再取回來
 * 
 **/
public class SortDatas implements Comparator<Object>
{
    private List<DataEntity> policys=new ArrayList<DataEntity>();

    public List<DataEntity> getPolicys()
    {
	Collections.sort(policys,new SortDatas());
	return policys;
    }

    public void setPolicys(List<DataEntity> policys)
    {
	this.policys=policys;
    }

    @Override
    public int compare(Object o1, Object o2)
    {
	DataEntity FirePolicy0=(DataEntity)o1;
	DataEntity FirePolicy1=(DataEntity)o2;
	// 單位
	int i=FirePolicy0.getUnit().compareTo(FirePolicy1.getUnit());
	if(i!=0)
	    return i;
	// 保單號排序
	i=FirePolicy0.getPLY_NO().compareTo(FirePolicy1.getPLY_NO());
	if(i!=0)
	    return i;

	return i;
    }
}
