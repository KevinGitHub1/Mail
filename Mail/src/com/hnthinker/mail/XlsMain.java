package com.hnthinker.mail;

import java.io.FileInputStream;  
import java.io.IOException;  
import java.io.InputStream;  
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;  
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class XlsMain {  
	private static FormulaEvaluator evaluator;
//    public static void main(String[] args) {  
//        XlsMain xlsMain = new XlsMain();  
//        try {  
//            List<String []> list = xlsMain.readXls();  
////            System.err.println(list);  
//            System.err.println("--------------------------");  
//            int k = 0;  
//            for (Iterator<String[]> iterator = list.iterator(); iterator.hasNext();) {  
//                String[] strings = (String[]) iterator.next();  
//                for (int i = 0; i < strings.length; i++) {  
//                    if(strings[i] != null){  
//                        System.err.print(strings[i] + "  ");  
//                    }  
//                }  
//                System.out.print("\n");  
//                k++;  
//                if(k == 3){  
//                    break;  
//                }  
//            }  
//            System.err.println("--------------------------");  
//        } catch (IOException e) {  
//            e.printStackTrace();  
//        }  
//    }  
      
    public String getValue(XSSFCell hssfCell) {  
        if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {  
            // 返回布尔类型的值  
            return String.valueOf(hssfCell.getBooleanCellValue());  
        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {  
            // 返回数值类型的值  
        	String result = "";
        	short format = hssfCell.getCellStyle().getDataFormat(); 
        	  SimpleDateFormat sdf = null; 
        	  if(format == 14 || format == 31 || format == 57 || format == 58){ 
        	    //日期 
        	    sdf = new SimpleDateFormat("yyyy年M月"); 
        	  }else if (format == 20 || format == 32) { 
        	    //时间 
        	    sdf = new SimpleDateFormat("HH:mm"); 
        	  } 
        	  double value = hssfCell.getNumericCellValue(); 
        	  Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value); 
        	  if(sdf==null) {
        		  result = String.valueOf(value);
        	  }else {
        		  result = sdf.format(date); 
        	  }
            return   result;
        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_BLANK) {  
            // 返回空类型的值  
            return "";  
        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_FORMULA) {  
            // 返回公式类型的值  
            return getCellValue(evaluator.evaluate(hssfCell)); 
        }else {  
            // 返回字符串类型的值  
            return String.valueOf(hssfCell.getStringCellValue());  
        }  
    }  
      public String getCellValue(CellValue cell) {

          String cellValue = null;
          switch (cell.getCellType()) {
          case Cell.CELL_TYPE_STRING:
              cellValue=cell.getStringValue();
              break;

          case Cell.CELL_TYPE_NUMERIC:
              cellValue=String.valueOf(cell.getNumberValue());
              break;
          case Cell.CELL_TYPE_FORMULA:
              break;
          default:
              break;
          }
          
          return cellValue;
      }

    public List<String []> readXls() throws IOException {  
        InputStream is = new FileInputStream("工资.xlsx");  
        XSSFWorkbook hssfWorkbook = new XSSFWorkbook(is); 
        evaluator=hssfWorkbook.getCreationHelper().createFormulaEvaluator();
        List<String []> list = new ArrayList<String []>();  
        // 循环工作表Sheet  
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {  
              
            XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);  
            if (hssfSheet == null) {  
                continue;  
            }  
            // 循环行Row  
            for (int rowNum = 0; rowNum < hssfSheet.getLastRowNum(); rowNum++) {  
                String[] str = new String[1000];  
                System.err.print(rowNum + "\t");  
                XSSFRow hssfRow = hssfSheet.getRow(rowNum);  
                if (hssfRow == null) {  
                    continue;  
                }  
                // 循环列Cell  
                Iterator<Cell> cellIterator = hssfRow.cellIterator();  
                int k = 0;  
                while (cellIterator.hasNext()) {  
                	XSSFCell cell = (XSSFCell) cellIterator.next();  
                    System.out.print(getValue(cell)+"\t\t");  
                    str[k++] = getValue(cell);  
                }  
                System.out.print("\n");  
                  
                list.add(str);  
            }  
        }  
        return list;  
    }  
  
}  