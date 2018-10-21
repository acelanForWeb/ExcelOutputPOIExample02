package tw.acelan;

import java.io.FileOutputStream;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	//excel檔案輸出的路徑，可以自行喜好設定，請確認路徑相關資料夾已存在
	private static final String FILE_NAME = "D:/Demo/PhoneBook.xlsx";
		
	public static void main(String[] args){
		//建立一個excel work book物件，等同於建立一個Excel檔案
		XSSFWorkbook workbook = new XSSFWorkbook();
		System.out.println("建立excel檔案完成");
		
		
		//在excel work book物件中建立一個「sheet」，描述為「電話簿」
		XSSFSheet sheet = workbook.createSheet("電話簿");
		System.out.println("建立sheet完成");

		
		//====		設定每一欄的寬度，本範例總共有三欄		Begin		====//
		//設定姓名欄位寬度
		sheet.setColumnWidth(0,(int)((10 + 0.72) * 256));//在excel文件中該寬度為10
		//設定生日欄位寬度
		sheet.setColumnWidth(1,(int)((20 + 0.72) * 256));//在excel文件中該寬度為20
		//設定手機號碼欄位寬度
		sheet.setColumnWidth(2,(int)((20 + 0.72) * 256));//在excel文件中該寬度為20
		//====		設定每一欄的寬度，本範例總共有三欄		End			====//
		
		
		
		
		//====		設定主標題列		Begin		====//
		CellStyle styleSubject = workbook.createCellStyle();
		Font subjectFont = workbook.createFont();
		subjectFont.setBold(true);//設定粗體
		subjectFont.setColor(HSSFColor.DARK_GREEN.index);//設定字體顏色
		Short fontSize16 = 16;
		subjectFont.setFontHeightInPoints(fontSize16);//設定字體大小
		styleSubject.setFont(subjectFont);
		styleSubject.setAlignment(HorizontalAlignment.CENTER); // 水平置中
		styleSubject.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直置中
		styleSubject.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);//填滿顏色
		styleSubject.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		Row row0 = sheet.createRow(0);
		Cell cell0 = row0.createCell(0);
		cell0.setCellValue("我的電話簿");
		cell0.setCellStyle(styleSubject);//設定格式
		
		//設定合併儲存格;new CellRangeAddress（int firstRow, int lastRow, int firstCol, int lastCol)
		CellRangeAddress cellRangeAddress =new CellRangeAddress(0,0,0,2);
		sheet.addMergedRegion(cellRangeAddress);
		//====		設定主標題列		End			====//
	
		
		
		
		
		
		//====		將電話簿資料填充到sheet中		Begin		====//
		
		//字型設定用於欄位標題
		CellStyle styleRow1 = workbook.createCellStyle();
		Font head_font = workbook.createFont();
		head_font.setBold(true);//設定粗體
		head_font.setColor(HSSFColor.BLUE.index);//設定字體顏色
		Short fontSize14 = 14;
		head_font.setFontHeightInPoints(fontSize14);//設定字體大小
		styleRow1.setFont(head_font);
		
		
		
		//電話簿資料
		Object[][] phoneBook = {
			{"姓名", "生日", "手機號碼"},	
			{"A君", "1970/01/25", "0900-111-111"},
            {"B君", "1980/02/26", "0900-222-222"},
            {"C君", "1990/05/25", "0900-333-333"},
            {"D君", "2000/07/13", "0900-444-444"},
            {"E君", "2010/11/21", "0900-555-555"}
		};
		
		System.out.println("開始將資料寫入sheet...");
		int rowNum = 1;
		for (Object[] rowData : phoneBook) {
			 Row row = sheet.createRow(rowNum);
			 int colNum = 0;
			 for (Object field : rowData) {
				 Cell cell = row.createCell(colNum);
				 if(rowNum == 1){
					 //設定欄位標題格式
					 cell.setCellStyle(styleRow1);
				 }
				 
				 if (field instanceof String) {
					 cell.setCellValue((String) field);
				 }else if(field instanceof Integer){
					 cell.setCellValue((Integer) field);
				 }
				 
				 colNum++;
			 }
			 rowNum++;
		}
		System.out.println("資料寫入sheet完成");
		//====		將電話簿資料填充到sheet中		End			====//
		
		
		//====		輸出Excel檔案				Begin		====//
		System.out.println("開始將excel檔案進行輸出...");
		FileOutputStream outputStream = null;
		Boolean isOutputSuccess = false;
		try {
            outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            isOutputSuccess = true;
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
        	try{
        		if(workbook != null){
        			workbook.close();
        		}
        	}catch(Exception e){e.printStackTrace();}
        	
        	try{
        		if(outputStream != null){
        			outputStream.close();
        		}
        	}catch(Exception e){e.printStackTrace();}
        }
		
		if(isOutputSuccess){
			System.out.println("excel檔案輸出完成");
		}else{
			System.out.println("excel檔案輸出失敗");
		}
		
		//====		輸出Excel檔案				End			====//
	}
}
