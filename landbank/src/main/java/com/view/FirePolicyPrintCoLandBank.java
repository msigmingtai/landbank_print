package com.view;

import java.awt.Container;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.DecimalFormat;
import java.text.Format;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;


import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfWriter;
import com.model.CreatePolicy;
import com.model.DataEntity;

/**
 * 
 * @author 50123
 * @version 1.0
 * @category 土地銀行轉要保書列印程式
 */

public class FirePolicyPrintCoLandBank extends JFrame implements ActionListener
{
    private static final long serialVersionUID=1L;
    List<DataEntity> plys=new ArrayList<DataEntity>();
    DataEntity ply;
    File obj=null;
    int errori=0;
    int ret=1;
    int wrong=0;
    String[] error=new String[5000];
    String SuccessFile="";
    String classes=null;
    JFileChooser fc=new JFileChooser();
    JButton bt=new JButton("開啟");
    JButton bt1=new JButton("產生列印檔");
    JButton bt2=new JButton("關閉");
    JLabel lb=new JLabel("請選擇要保書資料檔　　　　  ",2);
    JLabel lb1=new JLabel("請選擇要保書種類",2);
    JComboBox<String> cb=new JComboBox<String>();
    JComboBox<String> cb1=new JComboBox<String>();
    JComboBox<String> cb2=new JComboBox<String>();
    JLabel lb2=new JLabel("請選擇要保書是否分割",2);
    JLabel lb3=new JLabel("請選擇列印方式　　　　　",2);
    Document document=new Document();
    BaseFont bfChinese=null;
    Font FontChinese=null;
    Font FontChinese1=null;
    Font FontChinese2=null;
    Image jpeg=null;
    String str5=null;
    String str6=null;
    String[] str3=new String[14];
    String Unit_Old;
    String Unit_new;
    StringWriter sw=new StringWriter();
    PrintWriter pw=new PrintWriter(sw);

    public static void main(String[] args)
    {
	FirePolicyPrintCoLandBank app=new FirePolicyPrintCoLandBank();
	app.setTitle("明台產物保險要保書產生程式(合庫土銀版)");
	app.setSize(280,200);
	app.setLocationRelativeTo(null);
	app.setVisible(true);
	app.setDefaultCloseOperation(2);
	app.addWindowListener(new WindowAdapter()
	{
	    public void windowClosing(WindowEvent e)
	    {
		System.exit(0);
	    }
	});
    }

    public FirePolicyPrintCoLandBank()
    {
	Container cp=getContentPane();
	cp.setLayout(new FlowLayout());
	cp.add(this.lb);
	cp.add(this.bt);
	this.bt.addActionListener(this);
	cp.add(this.lb1);
	cb.addItem("基本住宅地震黑白");
	cp.add(this.cb);
	cp.add(this.lb2);
	cb1.addItem("單位分割列印");
	cb2.addItem("黑白列印");
	cp.add(this.cb1);
	cp.add(this.lb3);
	cp.add(this.cb2);
	cp.add(this.bt1);
	this.bt1.addActionListener(this);
	cp.add(this.bt2);
	this.bt2.addActionListener(this);
    }

    public void actionPerformed(ActionEvent e)
    {
	File TempPath=new File("C:\\");
	if(e.getSource()==this.bt)
	{
	    this.ret=this.fc.showOpenDialog(null);
	    this.obj=this.fc.getSelectedFile();
	}
	else if(e.getSource()==this.bt2)
	{
	    System.exit(0);
	}
	else if(this.ret==0)
	{
	    try
	    {
		// 產生保單檔
		plys=(List<DataEntity>)new CreatePolicy().GetDatas(this.obj);
		if(this.cb.getSelectedItem().equals("基本住宅地震黑白"))
		{
		    PrintBW();
		}

		if(this.errori>0)
		{
		    File fileE=new File("C:\\要保書資料錯誤清單.txt");
		    if(fileE.exists())
		    {
			fileE.delete();
		    }
		    fileE.createNewFile();
		    FileWriter fwE=new FileWriter(fileE);
		    for(int i=0; i<this.errori; i++)
		    {
			fwE.write(this.error[i]+"\r\n");
		    }
		    fwE.close();
		    JOptionPane.showMessageDialog(null,"請檢查錯誤清單");
		}
		if(!this.SuccessFile.equals(""))
		{
		    JOptionPane.showMessageDialog(null,this.SuccessFile);
		}
		JOptionPane.showMessageDialog(null,"執行完畢\n ");
	    }
	    catch (Exception e1)
	    {
		e1.printStackTrace(pw);
		JOptionPane.showMessageDialog(null,sw.toString());
	    }

	}
	else
	{
	    JOptionPane.showMessageDialog(null,"無匯入檔案");
	}
    }

    public void PrintBW()
    {

	try
	{

	    // document=new Document(PageSize.A4,72.0F,0.0F,175.0F,0.0F);
	    jpeg=Image.getInstance(FirePolicyPrintCoLandBank.class.getResource("/A.jpg"));
	    jpeg.setAlignment(8);
	    jpeg.scaleAbsolute(590.0F,840.0F);
	    jpeg.setAbsolutePosition(0.0F,0.0F);

	    Calendar cal=Calendar.getInstance();
	    int YEAR=cal.get(1);
	    int MONTH=cal.get(2)+1;
	    int DATE=cal.get(5);

	    // PdfWriter.getInstance(document,new
	    // FileOutputStream("C:\\火險要保書"+YEAR+"-"+MONTH+"-"+DATE+"-"+"-"+obj.getName().substring(0,4)+this.cb2.getSelectedItem().toString()+".pdf"));

	    bfChinese=BaseFont.createFont("C:\\windows\\fonts\\MINGLIU.TTC,1","Identity-H",false);
	    FontChinese=new Font(bfChinese,10.0F,0);
	    FontChinese1=new Font(bfChinese,7.0F,0);
	    FontChinese2=new Font(bfChinese,8.0F,0);

	    document.addTitle("ROCKP");
	    document.addAuthor("MINGTAI");
	    document.addSubject("MINGTAI");
	    document.addKeywords("MINGTAI");
	    document.addCreator("MINGTAI");
	    document.open();

	    for(DataEntity ply1 : plys)
	    {
		ply=ply1;
		Unit_new=ply.getUnit();
		// document.open();
		if(Unit_new.equals(Unit_Old))
		{
		    PrintPolicy();
		}
		else
		{
		    document.close();
		    // document=new
		    // Document(PageSize.A4,72.0F,0.0F,175.0F,0.0F);
		    document=new Document(PageSize.A4,72,36,173,-60);
		    PdfWriter.getInstance(document,new FileOutputStream("C:\\"+ply.getUnit()+"-火險要保書"+YEAR+"-"+MONTH+"-"+DATE+".pdf"));
		    document.open();
		    PrintPolicy();
		}

	    }
	    document.close();
	}
	catch (Exception e1)
	{
	    e1.printStackTrace(pw);
	    JOptionPane.showMessageDialog(null,sw.toString());
	}
	document.close();

    }

    public void PrintPolicy() throws DocumentException
    {

	document.add(jpeg);
	document.add(new Paragraph(0.0F,"                           "+ply.getPLY_NO().substring(0,2)+"       "+ply.getPLY_NO().substring(2,10)+"                                  "+ply.getORG_PLY()+"                                  ",FontChinese));
	String str2;
	String str1;
	if(ply.getADDR().trim().length()>10)
	{
	    str1=ply.getADDR().trim().substring(0,10);
	    str2=ply.getADDR().trim().substring(10);
	}
	else
	{
	    str1=ply.getADDR().trim();
	    str2="";
	}
	document.add(new Paragraph(18.0F,"                      "+ply.getFASR_NAME()+"                                                                                         ",FontChinese2));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                            "+ply.getADDR_AREA()+ply.getADDR().trim().substring(0,13).trim(),FontChinese2));
	document.add(new Paragraph(7.0F,"                                                                                                                                                                            　　　　　　　"+ply.getADDR().trim().substring(13).trim(),FontChinese1));
	document.add(new Paragraph(16.0F,"                                                 "+ply.getFASR_ID()+"                 "+ply.getTEL(),FontChinese));
	document.add(new Paragraph(12.0F,"                      "+ply.getINSURED()+"                                                                                         ",FontChinese2));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                            "+ply.getADDR_AREA()+ply.getADDR().trim().substring(0,13),FontChinese2));
	document.add(new Paragraph(7.0F,"                                                                                                                                                                            　　　　　　　"+ply.getADDR().trim().substring(13),FontChinese1));
	document.add(new Paragraph(16.0F,"                                                 "+ply.getID_NO()+"                 "+ply.getTEL(),FontChinese));
	document.add(new Paragraph(15.0F,"                                  "+CatZero(ZeroCut(ply.getBUILD_AMT())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                       "+CatZero(ZeroCut(ply.getBUILD_PREM())),FontChinese));
	document.add(new Paragraph(13.0F,"                                                                                                      "+CatZero(ZeroCut(ply.getTOT_PREM())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                       "+CatZero(ZeroCut(ply.getEQ_PREM())),FontChinese));
	document.add(new Paragraph(14,"                                  "+CatZero(ZeroCut(ply.getEQ_AMT()))+"                                                                                                                            0",FontChinese));
	document.add(new Paragraph(16.0F,"                    12                      "+ply.getN_E_DAY().substring(0,3)+"         "+ply.getN_E_DAY().substring(3,5)+"        "+ply.getN_E_DAY().substring(5)+"                                      "+ply.getEXP_DAY().substring(0,3)+"      "+ply.getEXP_DAY().substring(3,5)+"           "+ply.getEXP_DAY().substring(5),FontChinese));
	document.add(new Paragraph(26.0F,"                "+ply.getLOCATION_1().trim()+ply.getAREA_NO(),FontChinese));
	for(int i=0; i<=13; i++)
	{
	    str3[i]="  ";
	}
	int num1;
	try
	{
	    num1=Integer.parseInt(ply.getCONST_CLASS());
	    if(num1==1)
	    {
		this.classes="RCFR";
	    }
	    else if(num1==2)
	    {
		this.classes="BCER";
	    }
	    else if(num1==3)
	    {
		this.classes="ECFR";
	    }
	}
	catch (NumberFormatException e2)
	{
	    num1=0;
	}
	String str4;
	if(num1==1)
	{
	    str4="A1";
	}
	else
	{
	    if(num1==2)
	    {
		str4="A2";
	    }
	    else
	    {
		if(num1==3)
		{
		    str4="B";
		}
		else
		{
		    str4="C";
		}
	    }
	}
	str5=ZeroCut(ply.getCONST_SIZE()).substring(0,ZeroCut(ply.getCONST_SIZE()).length()-2);
	str6=ZeroCut(ply.getCONST_SIZE()).substring(ZeroCut(ply.getCONST_SIZE()).length()-2);
	document.add(new Paragraph(12.0F,"             "+str3[0]+"               "+str3[1]+str3[2],FontChinese1));
	document.add(new Paragraph("             "+str3[3]+"               "+str3[4]+str3[5]+str3[6]+str3[7],FontChinese1));
	document.add(new Paragraph("             "+str3[8]+"               "+str3[9]+str3[10]+str3[11]+"                                                                                                                                                                                    "+ply.getCONST_FLOOR()+"                              "+str4,FontChinese1));
	document.add(new Paragraph("                                                                              "+ply.getSTRU_YEAR()+"           "+this.classes+str3[12]+str3[13],FontChinese1));
	document.add(new Paragraph("                                                                              "+str5+"."+str6,FontChinese1));
	document.add(new Paragraph(88.0F,"   ",FontChinese));

	document.add(new Paragraph(56,"     01                                                                            "+CatZero(ZeroCut(ply.getBUILD_AMT())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                               0."+ply.getFIRE_RATE().substring(2)+"    1/1",FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                          "+CatZero(ZeroCut(ply.getBUILD_PREM())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                                     "+str4,FontChinese));

	document.add(new Paragraph(" "));
	document.add(new Paragraph(20.0F,"         自101年1月1日起，住宅地震基本保險最高賠償責任調整為新台幣150萬元。",FontChinese));

	document.add(new Paragraph(20.0F,"     02                                                                            "+CatZero(ZeroCut(ply.getEQ_AMT())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                               0.9        1/1           ",FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                          "+CatZero(ZeroCut(ply.getEQ_PREM())),FontChinese));
	document.add(new Paragraph(0.0F,"                                                                                                                                                                                     "+str4,FontChinese));
	document.add(new Paragraph(26,"                                                                                                              "+ply.getBANK_KEY1()+ply.getBANK_KEY2(),FontChinese));
	document.add(new Paragraph(20,"                                                                                                                                                           "+ply.getSEQ_NO(),FontChinese));
	document.add(new Paragraph(12.0F,"                                                                                                       "+ply.getMTG()+"                                     "+ply.getAG_EMP(),FontChinese));

	document.add(new Paragraph(" "));
	document.add(new Paragraph(" "));
	document.add(new Paragraph(72.0F,"            2                   "+ply.getMTG()+"                   "+ply.getBROKER()+"                   "+ply.getEMP()+"                   ",FontChinese));
	document.add(new Paragraph(15.0F,"16.9% + 6%                                                                                                                                  "+ply.getST_UNIT()+"                   "+ply.getS_EMP(),FontChinese));
	document.add(new Paragraph(" "));
	document.add(new Paragraph(24,"                                                                                                                                                                                                                  "+EmpNameProtect(ply.getEmp_no_name(),ply.getBROKER()),FontChinese1));
	document.add(new Paragraph(0,"                                                                                                                                                                                                                                                "+EmpSeqProtect(ply.getEmp_no_id(),ply.getBROKER()),FontChinese1));
	document.newPage();
	Unit_Old=Unit_new;

    }

    public String EmpNameProtect(String name, String broker)
    {
	String EmpName=name;
	if(!EmpName.equals(null))
	{

	    EmpName=EmpName.substring(0,1)+"＊"+EmpName.substring(2);

	}
	else
	{
	    EmpName="";
	}

	return EmpName;
    }

    public String EmpSeqProtect(String eseq, String broker)
    {
	String seq=eseq;
	if(!seq.equals(null))
	{
	    if(broker.equals("SK"))
	    {
		seq=seq.substring(0,2)+"****"+seq.substring(7);
	    }
	}
	else
	{
	    seq="";
	}
	return seq;
    }

    public String ZeroCut(String st)
    {
	String st2=null;
	for(int i=1; i<=st.length(); i++)
	{
	    if(!st.substring(i-1,i).equals("0"))
	    {
		String st1=st.substring(0,i).replace("0","  ");
		st2=st.substring(i,st.length());
		st=st1+st2;
		break;
	    }
	}
	return st;
    }

    public String CatZero(String str)
    {
	char[] charArr=str.toCharArray();
	int i=0;
	String Newstr="";
	for(i=0; i<charArr.length; i++)
	{
	    if(charArr[i]!='0')
	    {
		break;
	    }
	    charArr[i]=' ';
	}
	for(i=0; i<charArr.length; i++)
	{
	    Newstr=Newstr+charArr[i];
	}
	Newstr=Newstr.replaceAll(" ","");
	Newstr=Newstr.replaceAll("　","");

	Format fm=new DecimalFormat();
	try
	{
	    double Mnoey=Integer.parseInt(Newstr);
	    Newstr=fm.format(Double.valueOf(Mnoey));
	}
	catch (NumberFormatException localNumberFormatException)
	{
	    localNumberFormatException.printStackTrace(pw);
	    JOptionPane.showMessageDialog(null,sw.toString());
	}
	return Newstr;
    }

    public static String CatZero(int str)
    {
	Format fm=new DecimalFormat();

	String Mnoey=fm.format(Integer.valueOf(str));
	return Mnoey;
    }
}