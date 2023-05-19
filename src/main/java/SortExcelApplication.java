
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class SortExcelApplication {
	Sheet sheet;

	public static void main(String[] args) throws IOException {
		SortExcelApplication sea = new SortExcelApplication();
		sea.sortColumn("testsheet.xlsx", "B",3);
	}

	private boolean sortColumn(String path, String col, int stRow) {

        try {
        	File fle = new File(path);
    		Workbook wb = WorkbookFactory.create(fle);
	        List<String> sList = new ArrayList<String>();
        	sheet = wb.getSheetAt(0);
	        boolean cntinue = true;
	        int rNum=stRow;
	        
        	while (cntinue) {
		        Cell cell = getCell(col + rNum);
		        String contents = getCell(col + rNum).getStringCellValue();
		        if (contents == null || contents.isEmpty() || contents.equals("")) {
		        	cntinue=false;
		        } else {
		        	sList.add(cell.getStringCellValue());
		        }
		        rNum++;
        	}
        	
	        Collections.sort(sList);
	        
	        for (int i=1;i<sList.size()+1;i++) {
	        	getCell(col + (i+stRow-1)).setCellValue(sList.get(i-1));
	        }
	        try (FileOutputStream fos = new FileOutputStream("sorted_" + path)) {
                wb.write(fos);
                System.out.println("Excel file sorted and saved successfully!");
            }
        } catch (Exception e) {
        	e.printStackTrace();
        	return false;
        }
        return true;
	}
	
	private Cell getCell(String alphanumber) {
        CellReference cellReference = new CellReference(alphanumber); 
        Row row = sheet.getRow(cellReference.getRow());
        return row.getCell(cellReference.getCol()); 
	}
}
