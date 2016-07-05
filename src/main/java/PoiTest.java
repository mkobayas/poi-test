import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.jakewharton.fliptables.FlipTable;

public class PoiTest {

    public static void main(String[] args) throws Exception {

        File f = new File("hoge.xlsx");

        try (Workbook workbook = WorkbookFactory.create(f, null, true)) {

            for (Sheet sheet : workbook) {
                System.out.println("sheetName=" + sheet.getSheetName());

                List<String> header = new ArrayList<>();
                List<List<String>> tables = new ArrayList<>();
                List<String> tRow;
                boolean firstRow = true;

                DataFormatter formatter = new DataFormatter();
                FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

                for (Row row : sheet) {
                    tRow = new ArrayList<>();
                    for (Cell cell : row) {

                        Object value = null;
                        switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_FORMULA:
                            value = formatter.formatCellValue(cell, formulaEvaluator);
                            break;
                        default:
                            value = formatter.formatCellValue(cell);
                        }
                        if (firstRow) {
                            header.add("" + value);
                        } else {
                            tRow.add("" + value);
                        }
                    }

                    if (!firstRow) {
                        tables.add(tRow);
                    }
                    firstRow = false;
                }

                if (!header.isEmpty()) {
                    String[] headerArray = header.toArray(new String[0]);
                    String[][] tableArray = new String[tables.size()][headerArray.length];
                    for (int i = 0; i < tableArray.length; i++) {
                        List<String> r = tables.get(i);
                        for (int j = 0; j < headerArray.length; j++) {
                            tableArray[i][j] = r.get(j);
                        }
                    }
                    System.out.println(FlipTable.of(headerArray, tableArray));
                }
            }
        }
    }
}
