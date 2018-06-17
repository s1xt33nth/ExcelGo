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
     * �ж�Excel�ļ���׺����
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
     * ��ȡExcel����ͷ������ ������String����
     */  
    public String[] readExcelTitle1() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook����Ϊ�գ�");  
        }  
        sheet = wb.getSheetAt(0);  
        row = sheet.getRow(0);  
        // ����������  
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
     * ��ȡExcel����ͷ������ ������Map����
     */  
    public Map<Integer,Object> readExcelTitle2() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook����Ϊ�գ�");  
        }  
        sheet = wb.getSheetAt(0);  
        row = sheet.getRow(0);  
        // ����������  
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
     * ��ȡExcel�������� 
     */  
    public Map<Integer, Map<Integer,Object>> readExcelContent() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook����Ϊ�գ�");  
        }  
        Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();  
          
        sheet = wb.getSheetAt(0);  
        // �õ�������  
        int rowNum = sheet.getLastRowNum();  
        row = sheet.getRow(0);  
       // �õ�������  
        int colNum = row.getPhysicalNumberOfCells();  
        // �������ݴӵڶ��п�ʼ,��һ��Ϊ��ͷ�ı���  
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
     * ����Cell������������  
     */  
    private Object getCellFormatValue(Cell cell) {  
        Object cellvalue = "";  
        if (cell != null) {  
            // �жϵ�ǰCell��Type  
            switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC:// �����ǰCell��TypeΪNUMERIC  
            case Cell.CELL_TYPE_FORMULA: {  
                // �жϵ�ǰ��cell�Ƿ�ΪDate  
                if (DateUtil.isCellDateFormatted(cell)) {  
                	SimpleDateFormat sdf = null; 
                	sdf = new SimpleDateFormat("yyyy-MM-dd"); 
                    Date date = cell.getDateCellValue();  
                    cellvalue = sdf.format(date); 
                } else {// ����Ǵ�����  
                    // ȡ�õ�ǰCell����ֵ  
                    cellvalue = String.valueOf(cell.getNumericCellValue());  
                }  
                break;  
            }  
            case Cell.CELL_TYPE_STRING:// �����ǰCell��TypeΪSTRING  
                // ȡ�õ�ǰ��Cell�ַ���  
                cellvalue = cell.getRichStringCellValue().getString();  
                break;  
            default:// Ĭ�ϵ�Cellֵ  
                cellvalue = "";  
            }  
        } else {  
            cellvalue = "";  
        }  
        return cellvalue;  
    }
}

