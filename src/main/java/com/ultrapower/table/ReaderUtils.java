package com.ultrapower.table;

import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.UUID;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;


import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * @Description: poi实现word表格数据读取
 * @Author: HALEN(李智刚)
 * @CreateDate: 2018/7/1810:31
 * <p>Copyright: Copyright (c) 2018</p>
 */
public class ReaderUtils {

    public static void main(String[] args) {
        test1();
    }

    public static void testWord() {
        try {
            FileInputStream in = new FileInputStream("C:\\Users\\halen\\Desktop\\质量核查\\data.doc");// 加载文档
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();// 获取文档的读取范围
            TableIterator it = new TableIterator(range);
            // 迭代文档中的表格
            while (it.hasNext()) {
                Table tb = (Table) it.next();
                // 迭代行，默认从0开始
                for (int i = 0; i < tb.numRows(); i++) {
                    TableRow tr = tb.getRow(i);
                    // 迭代列，默认从0开始
                    for (int j = 0; j < tr.numCells(); j++) {
                        TableCell td = tr.getCell(j);
                        // System.out.println(td.text());
                        // 取得单元格的内容
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            Paragraph para = td.getParagraph(k);
                            String s = para.text();
                            System.out.println(s.replaceAll("\r", "").replaceAll(" ","")+":"+s.replaceAll("\r", "").replaceAll(" ",""));
                        }

                    }
                }
            }

            in.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 不规则行表格读取方法
     */
    public static void test(){
        try{
            FileInputStream in = new FileInputStream("C:\\Users\\halen\\Desktop\\质量核查\\data.doc");//载入文档
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);
            //获取当前时间（年月日）
            Date date = new Date();
            DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
            int currentTime = Integer.parseInt(dateFormat.format(date));

            String uuid = UUID.randomUUID().toString().replaceAll("-", "");
            //迭代文档中的表格
            while (it.hasNext()) {
                Table tb = it.next();
                String uuid1 = "";
                String uuid2 = "";
                //迭代行，默认从0开始
                for (int i = 1; i < tb.numRows(); i++) {
                    TableRow tr = tb.getRow(i);
                    String front = "INSERT INTO bs_t_sm_dictionaryconfig VALUES(";
                    String middle = "NULL,1,'halen',"+currentTime +",'halen',"+currentTime + ",null";
                    String back = ",parent-uuid);";
                    /*String front = "INSERT INTO bs_t_sm_dictionaryconfig VALUES(uuid,'报警方式',";
                    String middle = "编码,NULL,1,'halen',时间,'halen',时间,";
                    String back = "说明,parent-uuid);";*/
                    int sum = 0;

                    //迭代列，默认从0开始
                    for (int j = 0; j < tr.numCells(); j++) {
                        TableCell td = tr.getCell(j);//取得单元格
                        //取得单元格的内容
                        for(int k=0;k<td.numParagraphs();k++){
                            Paragraph para =td.getParagraph(k);
                            String text = para.text();
                            text = text.substring(0,text.length()-1);
                            if (text.length() < 2)
                                continue;
                            sum ++;
//                            System.out.print(text.length()+ " ");

                            if (sum % 2 == 0) {
//                                middle = ",'" + text + "'" + middle;
                                front += ",'"+ text +"'";
                                System.out.println(front + middle);
                                front = "INSERT INTO bs_t_sm_dictionaryconfig VALUES(";
                                middle = "NULL,1,'halen',"+ currentTime +",'halen',"+ currentTime + ",null";
                            }else {
                                try {
                                    Integer.parseInt(text);
                                } catch (NumberFormatException e) {
                                    continue;
                                }
                                //三级类型表格
                                if ("0000".equals(text.substring(text.length()-4))) {
                                    uuid1 = UUID.randomUUID().toString().replaceAll("-", "");
                                    front += "'" + uuid1 + "'";
                                    middle = ",'" + text + "'," + middle + ",'" + uuid + "');";
//                                    System.out.print("一级");
                                }else if ("00".equals(text.substring(text.length()-2))) {
//                                    System.out.print("uuid1:"+ uuid1);
                                    uuid2 = UUID.randomUUID().toString().replaceAll("-", "");
                                    front += "'" + uuid2 + "'";
                                    middle = ",'" + text + "'," + middle + ",'" + uuid1 + "');";
//                                    System.out.print("二级");
                                }else {
                                    String uuid3 = UUID.randomUUID().toString().replaceAll("-", "");
                                    front += "'" + uuid3 + "'";
                                    middle = ",'" + text + "'," + middle + ",'" + uuid2 + "');";
//                                    System.out.print("uuid2:"+ uuid2);
//                                    System.out.print("三级");
                                }



                                //三级
                                /*if (text.length() == 2) {
                                    uuid1 = UUID.randomUUID().toString().replaceAll("-", "");
                                    front += "'" + uuid1 + "'";
                                    middle = ",'" + text + "'," + middle + ",'" + uuid + "');";
//                                    System.out.print("一级");
                                }else if (text.length() == 4) {
//                                    System.out.print("uuid1:"+ uuid1);
                                    uuid2 = UUID.randomUUID().toString().replaceAll("-", "");
                                    front += "'" + uuid2 + "'";
                                    middle = ",'" + text + "'," + middle + ",'" + uuid1 + "');";
//                                    System.out.print("二级");
                                }*/
                            }
                        }
                    }
                }
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 规则行列表格读取方法
     */
    public static void test1(){
        try{
            FileInputStream in = new FileInputStream("C:\\Users\\halen\\Desktop\\质量核查\\data.doc");//载入文档
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);

            //获取当前时间（年月日）
            Date date = new Date();
            DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
            int currentTime = Integer.parseInt(dateFormat.format(date));

            String puuid = UUID.randomUUID().toString().replaceAll("-", "");
            String psql = "INSERT INTO bs_t_sm_dictionaryconfig VALUES('"+ puuid + "',''"+ ",''" +",NULL,1,NULL," + currentTime + ",NULL," + currentTime + ",NULL,NULL);";
            System.out.println(psql + "-- 一级");

            //迭代文档中的表格
            while (it.hasNext()) {
                Table tb = (Table) it.next();
                //迭代行，默认从0开始
                for (int i = 0; i < tb.numRows(); i++) {
                    TableRow tr = tb.getRow(i);
                    if (i == 0) {
                        continue;
                    }
                   /* String front = "INSERT INTO bs_t_sm_dictionaryconfig VALUES("+ UUID.randomUUID().toString().replaceAll("-","") +",";
                    String middle = ",NULL,1,'halen',"+new Date().getTime() +",'halen',"+new Date().getTime();
                    String back = ",parent-uuid);";*/
                    String front = "INSERT INTO bs_t_sm_dictionaryconfig VALUES(";
                    String middle = "NULL,1,NULL,"+currentTime +",NULL,"+currentTime;

                    //拼接uuid（pid）
                    String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                    front += "'" + uuid + "'";

                    //迭代列，默认从0开始
                    for (int j = 0; j < tr.numCells(); j++) {
                        TableCell td = tr.getCell(j);//取得单元格
                        //取得单元格的内容
                        for(int k=0;k<td.numParagraphs();k++){
                            Paragraph para =td.getParagraph(k);

                            //获取单元格数据并做截取
                            String text = para.text();
                            text = text.substring(0,text.length()-1);

                            switch (j) {
                                case 0://获取编码
                                    middle = ",'" + text + "'," + middle + "";
                                    break;
                                case 1://获取类别名称
                                    front += ",'" + text + "'";
                                    break;
                                case 2://获取说明（数据库中叫备注）
                                    if (text.isEmpty())
                                        middle += ",NULL";
                                    else {
                                        middle += ",'" + text + "'";
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }

                    }
                    //拼接uuid（parentid）
                    middle += ",'" + puuid + "');";
                    System.out.println(front + middle);
                }
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}
