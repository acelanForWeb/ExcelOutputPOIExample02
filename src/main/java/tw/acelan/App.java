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
	//excel�ɮ׿�X�����|�A�i�H�ۦ�ߦn�]�w�A�нT�{���|������Ƨ��w�s�b
	private static final String FILE_NAME = "D:/Demo/PhoneBook.xlsx";
		
	public static void main(String[] args){
		//�إߤ@��excel work book����A���P��إߤ@��Excel�ɮ�
		XSSFWorkbook workbook = new XSSFWorkbook();
		System.out.println("�إ�excel�ɮק���");
		
		
		//�bexcel work book���󤤫إߤ@�ӡusheet�v�A�y�z���u�q��ï�v
		XSSFSheet sheet = workbook.createSheet("�q��ï");
		System.out.println("�إ�sheet����");

		
		//====		�]�w�C�@�檺�e�סA���d���`�@���T��		Begin		====//
		//�]�w�m�W���e��
		sheet.setColumnWidth(0,(int)((10 + 0.72) * 256));//�bexcel��󤤸Ӽe�׬�10
		//�]�w�ͤ����e��
		sheet.setColumnWidth(1,(int)((20 + 0.72) * 256));//�bexcel��󤤸Ӽe�׬�20
		//�]�w������X���e��
		sheet.setColumnWidth(2,(int)((20 + 0.72) * 256));//�bexcel��󤤸Ӽe�׬�20
		//====		�]�w�C�@�檺�e�סA���d���`�@���T��		End			====//
		
		
		
		
		//====		�]�w�D���D�C		Begin		====//
		CellStyle styleSubject = workbook.createCellStyle();
		Font subjectFont = workbook.createFont();
		subjectFont.setBold(true);//�]�w����
		subjectFont.setColor(HSSFColor.DARK_GREEN.index);//�]�w�r���C��
		Short fontSize16 = 16;
		subjectFont.setFontHeightInPoints(fontSize16);//�]�w�r��j�p
		styleSubject.setFont(subjectFont);
		styleSubject.setAlignment(HorizontalAlignment.CENTER); // �����m��
		styleSubject.setVerticalAlignment(VerticalAlignment.CENTER); // �����m��
		styleSubject.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);//���C��
		styleSubject.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		Row row0 = sheet.createRow(0);
		Cell cell0 = row0.createCell(0);
		cell0.setCellValue("�ڪ��q��ï");
		cell0.setCellStyle(styleSubject);//�]�w�榡
		
		//�]�w�X���x�s��;new CellRangeAddress�]int firstRow, int lastRow, int firstCol, int lastCol)
		CellRangeAddress cellRangeAddress =new CellRangeAddress(0,0,0,2);
		sheet.addMergedRegion(cellRangeAddress);
		//====		�]�w�D���D�C		End			====//
	
		
		
		
		
		
		//====		�N�q��ï��ƶ�R��sheet��		Begin		====//
		
		//�r���]�w�Ω������D
		CellStyle styleRow1 = workbook.createCellStyle();
		Font head_font = workbook.createFont();
		head_font.setBold(true);//�]�w����
		head_font.setColor(HSSFColor.BLUE.index);//�]�w�r���C��
		Short fontSize14 = 14;
		head_font.setFontHeightInPoints(fontSize14);//�]�w�r��j�p
		styleRow1.setFont(head_font);
		
		
		
		//�q��ï���
		Object[][] phoneBook = {
			{"�m�W", "�ͤ�", "������X"},	
			{"A�g", "1970/01/25", "0900-111-111"},
            {"B�g", "1980/02/26", "0900-222-222"},
            {"C�g", "1990/05/25", "0900-333-333"},
            {"D�g", "2000/07/13", "0900-444-444"},
            {"E�g", "2010/11/21", "0900-555-555"}
		};
		
		System.out.println("�}�l�N��Ƽg�Jsheet...");
		int rowNum = 1;
		for (Object[] rowData : phoneBook) {
			 Row row = sheet.createRow(rowNum);
			 int colNum = 0;
			 for (Object field : rowData) {
				 Cell cell = row.createCell(colNum);
				 if(rowNum == 1){
					 //�]�w�����D�榡
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
		System.out.println("��Ƽg�Jsheet����");
		//====		�N�q��ï��ƶ�R��sheet��		End			====//
		
		
		//====		��XExcel�ɮ�				Begin		====//
		System.out.println("�}�l�Nexcel�ɮ׶i���X...");
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
			System.out.println("excel�ɮ׿�X����");
		}else{
			System.out.println("excel�ɮ׿�X����");
		}
		
		//====		��XExcel�ɮ�				End			====//
	}
}
