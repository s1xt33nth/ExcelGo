package com.excelgo.go_hashmap;

import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GoExcel {
	


    private Workbook wb;  
    private Sheet sheet;  
    private Row row;  
  
    /** 
     * 判断Excel文件后缀类型
     */ 
    public GoExcel(String filepath) {  
        if(filepath==null){  
            return;  
        }  
        String ext = filepath.substring(filepath.lastIndexOf("."));  
        try {  
            InputStream is = new FileInputStream(filepath);  
            if(".xls".equals(ext)){  
                wb = new HSSFWorkbook(is);  
            }else if(".xlsx".equals(ext)){  
                wb = new XSSFWorkbook(is);  
            }else{  
                wb=null;  
            }  
        } catch (FileNotFoundException e) {  
            
        } catch (IOException e) {  
          
        }  
    }  
      
    /** 
     * 读取Excel表格表头的内容 ，返回String类型
     */  
    public String[] readExcelTitle1() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        sheet = wb.getSheetAt(0);  
        row = sheet.getRow(0);  
        // 标题总列数  
        int colNum = row.getPhysicalNumberOfCells();  
        String[] title = new String[colNum];       
        for (int i = 0; i < colNum; i++) {  
//          title[i] = getStringCellValue(row.getCell((short) i));  
//          title[i] = row.getCell(i).getCellFormula();  
    	  title[i] = (String) getCellFormatValue(row.getCell(i));
        }  
        return title;  
    }  
    
    /** 
     * 读取Excel表格表头的内容 ，返回Map类型
     */  
    public Map<Integer,Object> readExcelTitle2() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        sheet = wb.getSheetAt(0);  
        row = sheet.getRow(0);  
        // 标题总列数  
        int colNum = row.getPhysicalNumberOfCells();  
        Map<Integer,Object> cellTitle = new HashMap<Integer, Object>();
        for (int i = 0; i < colNum; i++) {  
            Object obj = getCellFormatValue(row.getCell(i));
            if(!obj.toString().trim().equals("")){
                cellTitle.put(i, obj);	
            }
        }  
        return cellTitle;  
    }  
  
    /** 
     * 读取Excel数据内容 
     */  
    public Map<Integer, Map<Integer,Object>> readExcelContent() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();  
          
        sheet = wb.getSheetAt(0);  
        // 得到总行数  
        int rowNum = sheet.getLastRowNum();  
        row = sheet.getRow(0);  
       // 得到总列数  
        int colNum = row.getPhysicalNumberOfCells();  
        // 正文内容从第二行开始,第一行为表头的标题  
        for (int i = 1; i <= rowNum; i++) {  
            row = sheet.getRow(i);  
            int j = 0;  
            Map<Integer,Object> cellValue = new HashMap<Integer, Object>();  
            while (j < colNum) {  
                Object obj = getCellFormatValue(row.getCell(j));   
                	if(!obj.toString().trim().equals("")){
                        cellValue.put(j, obj);  
                	}
                    j++;  
            }  
            content.put(i, cellValue);  
        }  
        return content;  
    }  
  
    /** 
     * 根据Cell类型设置数据  
     */  
    private Object getCellFormatValue(Cell cell) {  
        Object cellvalue = "";  
        if (cell != null) {  
            // 判断当前Cell的Type  
            switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC:// 如果当前Cell的Type为NUMERIC  
            case Cell.CELL_TYPE_FORMULA: {  
                // 判断当前的cell是否为Date  
                if (DateUtil.isCellDateFormatted(cell)) {  
                	SimpleDateFormat sdf = null; 
                	sdf = new SimpleDateFormat("yyyy-MM-dd"); 
                    Date date = cell.getDateCellValue();  
                    cellvalue = sdf.format(date); 
                } else {// 如果是纯数字  
                    // 取得当前Cell的数值  
                    cellvalue = String.valueOf(cell.getNumericCellValue());  
                }  
                break;  
            }  
            case Cell.CELL_TYPE_STRING:// 如果当前Cell的Type为STRING  
                // 取得当前的Cell字符串  
                cellvalue = cell.getRichStringCellValue().getString();  
                break;  
            default:// 默认的Cell值  
                cellvalue = "";  
            }  
        } else {  
            cellvalue = "";  
        }  
        return cellvalue;  
    }
}

