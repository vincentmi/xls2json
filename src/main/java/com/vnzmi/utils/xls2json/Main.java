package com.vnzmi.utils.xls2json;

import java.io.File;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook ;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.CellType;

public class Main {

    public static void main(String[] args)
    {
       if(args.length == 0)
       {
           System.out.println("请设置要处理的文件目录");
           return ;
       }

       String path = args[0];
       System.out.println("RootPath="+path);

        File file = new File(path);

        if(!file.exists() ||  !file.isDirectory())
        {
            System.out.println("无效的目录");
            return ;
        }
        readDirectory(file);
    }


    public static void readDirectory(File dir)
    {
        if(!dir.exists() ||  !dir.isDirectory())
        {
            return ;
        }
        File[] fs = dir.listFiles();
        for(File f:fs){
            if(f.isDirectory())
            {
                readDirectory(f);
            }else{
                convert(f);
            }

        }
    }



    public static void convert(File f)
    {
        int labelIndex = 1;
        int keyIndex = 0;

        try {
            //System.out.println(f.getCanonicalFile());
            if(f.getName().endsWith(".xls") || f.getName().endsWith(".xlsx"))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(f);
                XSSFSheet sheet = workbook.getSheetAt(0);

                int lastRowIndex = sheet.getLastRowNum();

                if(lastRowIndex <=  labelIndex || lastRowIndex <= keyIndex)
                {
                    System.out.println("NO-Content");
                    return ;
                }


                XSSFRow keyRow = sheet.getRow(keyIndex);
                ArrayList<String> keys = new ArrayList<String>();
                String cellValue;
                XSSFCell cell;
                CellType cellType;

                ArrayList<HashMap<String,String>>  data= new ArrayList<HashMap<String,String>>();

                short lastCellNum = keyRow.getLastCellNum();
                for (int j = 0; j < lastCellNum; j++) {
                     cellValue = keyRow.getCell(j).getStringCellValue();
                    keys.add(cellValue);
                }

                for (int i = 2; i <= lastRowIndex; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row == null) { break; }
                    lastCellNum = row.getLastCellNum();
                    HashMap<String,String> rowData = new HashMap<String, String>();
                    for (int j = 0; j < lastCellNum; j++) {
                        cell = row.getCell(j);

                        cellType = cell.getCellTypeEnum();
                        if(cellType == CellType.STRING)
                        {
                            cellValue = cell.getStringCellValue();
                        }else if(cellType == CellType.NUMERIC)
                        {
                            cellValue = Double.toString(cell.getNumericCellValue());
                        }else if(cellType == CellType.BOOLEAN)
                        {
                            cellValue =cell.getBooleanCellValue() ? "1":"0";
                        }else if(cellType == CellType.BLANK)
                        {
                            cellValue ="";
                        }else if(cellType == CellType.FORMULA)
                        {
                            cellValue =cell.getCellFormula();
                        }else{
                            cellValue = "";
                        }



                        if(keys.size() > j)
                        {
                            rowData.put(keys.get(j),cellValue);
                        }
                    }
                    data.add(rowData);
                }
                String fPath = f.getCanonicalPath();
                fPath = fPath.replace(".xlsx",".json");
                fPath = fPath.replace(".xls",".json");


                File jsonFile = new File(fPath);

                if(!jsonFile.exists()){
                    jsonFile.createNewFile();
                }

                FileWriter fw = new FileWriter(jsonFile);
                fw.write(JSON.toJSONString(data, SerializerFeature.PrettyFormat).toString());
                fw.close();

                System.out.println("save:" + fPath);

            }else{
                //System.out.println("skip-"+f. getName());
            }

        }catch(Exception e)
        {
            e.printStackTrace();
        }

    }
}
