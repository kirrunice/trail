# trail

package local.util;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class excelDataExtractor {	
	static DataFormatter dataFormatter = new DataFormatter();
	
	public static Object[][] getExcelData(XSSFSheet sh , String methodName){
	
		int rowCount = sh.getLastRowNum();
		int columnCount = sh.getRow(0).getLastCellNum();
		int testCaserowCount = getTestCaseRowCount(methodName,rowCount,sh);
		Object[][] dataObject = new Object[testCaserowCount][1];
		
			
		for(int i=0;i<testCaserowCount;i++) {
			Map<String, String> dataMap = new HashMap<String, String>();
			if(dataFormatter.formatCellValue(sh.getRow(i+1).getCell(0)).equalsIgnoreCase(methodName)) {
				for(int j=0 ; j<columnCount;j++) {					
					dataMap.put(dataFormatter.formatCellValue(sh.getRow(0).getCell(j)), dataFormatter.formatCellValue(sh.getRow(i+1).getCell(j)));
			}
			}
			dataObject[i][0] = dataMap;
		}		
		
		return dataObject;
	}
	
	private static int getTestCaseRowCount(String testCaseName, int overallRowCount , XSSFSheet sh) {
		int counter = 0;
		for(int i=0;i<overallRowCount;i++) {
			if(dataFormatter.formatCellValue(sh.getRow(i+1).getCell(0)).equalsIgnoreCase(testCaseName)) {
				counter = counter + 1;
			}
		}		
		return counter;		
	}

}
