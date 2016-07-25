package demo.lisong.poi.read;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DemoRead {

	private static final Logger log = LoggerFactory.getLogger(DemoRead.class);
	@Test
	public void testReadExcelFile() throws Exception {
		InputStream file = DemoRead.class.getClassLoader().getResourceAsStream("学员导入模板.xlsx");
		
		
		//Get the workbook instance for XLS file 
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		//Get first sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);

		String[] keys = new String[]{
				"memname", 
				"sex", "age", "district", "identification", 
				"education", "work", "profession", "language", 
				"mobile", "wechat", "mail", "qq", "adress", 
				"sect", "setting", "weektime", "monthtime", 
				"selftime", "alltime", "selfalltime"};
		//Get iterator to all the rows in current sheet
		List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
		for (Iterator<Row> rowIterator = sheet.iterator(); rowIterator.hasNext();) {
			Row row = rowIterator.next();
			//Get iterator to all cells of current row
			int index = 0;
			Map<String, Object> item = new HashMap<String, Object>();
			for (Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); index++) {
				Cell cell = cellIterator.next();
				String key = keys[index];
//				String type = keys[index][1];
				Object obj = null;
				int cellType = cell.getCellType();
				log.debug("key: " + key);
				log.debug("type: " + cellType);
				switch(cellType) {
				case Cell.CELL_TYPE_BLANK:
				case Cell.CELL_TYPE_ERROR:	
				case Cell.CELL_TYPE_FORMULA:
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					obj = cell.getBooleanCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC:
					obj = cell.getNumericCellValue();
					break;
				case Cell.CELL_TYPE_STRING:
					obj = cell.getStringCellValue();
					break;
				}
				item.put(key, obj);
				
			}
			dataList.add(item);
		}
		log.debug(dataList.toString());
		workbook.close();
	}
}