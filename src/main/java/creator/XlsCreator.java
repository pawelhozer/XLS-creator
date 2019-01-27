package creator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class XlsCreator<T> {

    private Class<T> myClass;

    public XlsCreator(Class<T> myClass) {
        this.myClass = myClass;
    }

    public void createFile(List<T> series, String fileName, String path) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet xlsSheet = workbook.createSheet(fileName);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)10);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(headerFont);

        List<String> columns = new ArrayList<>();

        for(Field f : myClass.getDeclaredFields()){
            columns.add(f.getName());
        }

        Row headerRow = xlsSheet.createRow(0);
        for(int i=0; i<columns.size();i++){
            Cell cell =  headerRow.createCell(i);
            cell.setCellValue(columns.get(i));
            cell.setCellStyle(cellStyle);

            xlsSheet.autoSizeColumn(i);
        }

        for(int i=0;i<series.size();i++){

            HSSFRow row = xlsSheet.createRow(i+1);

            for(int j=0;j<columns.size();j++){
                HSSFCell cell = row.createCell(j);

                Method method = series.get(i)
                        .getClass()
                        .getMethod("get"+columns.get(j)
                        .substring(0,1)
                        .toUpperCase()+columns.get(j)
                        .substring(1));

                Object result = method.invoke(series.get(i));
                cell.setCellValue(String.valueOf(result));
            }
        }

        long time = System.currentTimeMillis();
        String file = fileName+"_"+time+".xls";

        workbook.write(new File(file));
        workbook.close();

    }
}
