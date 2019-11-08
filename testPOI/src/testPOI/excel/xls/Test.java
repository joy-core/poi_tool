package testPOI.excel.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddressList;

public class Test {
    public static void main(String[] argv) throws Exception {
    	
    	// 创建Excel
    	writeExcel();
    	
    	// 读取Excel
    	// readExcel();
    	
    }
    
    public static void readExcel() {
    	
    }
	
    /**
     * 写入Excel
     * @param file
     */
    public static void writeExcel() {
    	File file = new File("../testPOI/excel/测试.xls");
        if (!file.exists()) {
        	try {
				file.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		FileInputStream readFile = null;
		try {
			readFile = new FileInputStream(file);
			@SuppressWarnings("resource")
			HSSFWorkbook wb = new HSSFWorkbook();// 创建Excel工作簿对象
			HSSFSheet sheet = wb.createSheet("流程清单");// 创建工作表对象
			
			// 第一行标题行
			// 内容：序号，姓名，性别（下拉列表），年龄，性格（下拉列表），出生日期
			HSSFRow titlerow = sheet.createRow((short)0); //创建Excel工作表的行 
			titlerow.setHeightInPoints(30);// 设置行高
			HSSFCellStyle cellStyle = wb.createCellStyle();//创建单元格样式 
			cellStyle.setAlignment(HorizontalAlignment.CENTER);// 水平居中  
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中
			
			List<HSSFCell> titleList = new ArrayList<>();
			
			int index = 0;
			//  序号
			titleList.add(titlerow.createCell((short)index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("序号");
			sheet.setColumnWidth((short)index, 3000);// 设置当前列宽度
			
			index++;
			
			// 姓名
			titleList.add(titlerow.createCell((short)index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("姓名");
			sheet.setColumnWidth((short)index, 5000);// 设置当前列宽度
			
			index++;
			
			// 性别
			titleList.add(titlerow.createCell(index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("性别");
			sheet.setColumnWidth((short)index, 2000);// 设置当前列宽度
			
			index++;
			
			// 年龄
			titleList.add(titlerow.createCell(index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("年龄");
			sheet.setColumnWidth((short)index, 2000);// 设置当前列宽度
			
			index++;
			
			// 性格
			titleList.add(titlerow.createCell(index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("性格");
			sheet.setColumnWidth((short)index, 2000);// 设置当前列宽度
			
			index++;
			
			// 出生日期
			titleList.add(titlerow.createCell(index));
			titleList.get((short)index).setCellStyle(cellStyle);
			titleList.get((short)index).setCellValue("出生日期");
			sheet.setColumnWidth((short)index, 6000);// 设置当前列宽度
			
			// 数据区域
			List<HSSFRow> datas = new ArrayList<HSSFRow>();
			/*
			 * 序号：文本
			 * 姓名：文本
			 * 性别：下拉框，男/女/保密
			 * 年龄：文本
			 * 性格：下拉框，外向/内向
			 * 出生日期：日期类型，xxxx年x月日
			 */
			
			int count = 10;// 行数
			
			// 性别下拉框数据
			for (int i = 0; i < count; i++) {// 设置以前行的格式
				HSSFRow dataRow = sheet.createRow(i+1);
				// 序号
				HSSFCell cell1 = dataRow.createCell((short)0);
				cell1.setCellValue(i + 1);
				
				// 姓名
				HSSFCell cell2 = dataRow.createCell((short)1);
				cell2.setCellValue("张三" + (i + 1));
				
				// 性别
				HSSFCell cell3 = dataRow.createCell((short)2);
				// cell3.setCellValue("男");// 设置默认值
				
				// 年龄
				HSSFCell cell4 = dataRow.createCell((short)3);
				cell4.setCellValue(i + 15);
				
				// 性格
				HSSFCell cell5 = dataRow.createCell((short)4);
				
				// 出生日期
				HSSFCell cell6 = dataRow.createCell((short)5);
				// 日期设置
				CreationHelper createHelper=wb.getCreationHelper();
				HSSFCellStyle dateCellStyle = wb.createCellStyle(); //单元格样式类
				dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy年mm月dd日"));
				// cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
				cell6.setCellValue(new Date());
				cell6.setCellStyle(dateCellStyle);
				
				datas.add(dataRow);
			}
			
			// 性别下拉框
			String [] sexs = new String[] {"男", "女", "保密"};
			CellRangeAddressList range = new CellRangeAddressList(1, count, 2, 2);
			DVConstraint constraint = DVConstraint.createExplicitListConstraint(sexs);
			HSSFDataValidation dataValidation = new HSSFDataValidation(range, constraint);
			sheet.addValidationData(dataValidation);
			
			
			// 性格下拉框
			String [] charas = new String[] {"内向", "外向"};
			CellRangeAddressList charaRange = new CellRangeAddressList(1, count, 4, 4);
			DVConstraint constraint2 = DVConstraint.createExplicitListConstraint(charas);
			HSSFDataValidation dataValidation2 = new HSSFDataValidation(charaRange, constraint2);
			sheet.addValidationData(dataValidation2);
			
			// 写入文件
			wb.write(new FileOutputStream(file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (readFile != null) {
				try {
					readFile.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
}
