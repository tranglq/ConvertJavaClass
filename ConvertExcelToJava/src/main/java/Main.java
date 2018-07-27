
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;


public class Main {
    private StructData strucdata = new StructData();

    public void main(String[] argv) throws IOException {

        FileInputStream file = new FileInputStream("Calculation.xlsx");
        HSSFWorkbook workbook = new HSSFWorkbook(file);
        HSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();

        while(rowIterator.hasNext()){
            Row row = rowIterator.next();
            int rowNum = row.getRowNum();

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                System.out.println(cell.getRichStringCellValue());

                if (rowNum == 0){
                strucdata.classname = cell.getStringCellValue();
                }else if (rowNum == 1){
                    strucdata.varname = cell.getStringCellValue();
                }else if (rowNum == 2){
                    strucdata.methodname = cell.getStringCellValue();
                }
            }
        }

    }
}
