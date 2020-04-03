package sk.mimac.properties2excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelExportProcessor extends AbstractProcessor {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelExportProcessor.class);

    public ExcelExportProcessor(String excelFile, String propertiesFolder, String propertiesExtension) {
        super(excelFile, propertiesFolder, propertiesExtension);
    }

    @Override
    protected void validateInternal() throws ProcessingException {
        if (!excelFile.exists()) {
            throw new ProcessingException("Excel file '" + excelFile.getAbsolutePath() + "' doesn't exist");
        }
    }

    @Override
    public void processInternal() throws IOException, ProcessingException {
        LOG.info("Processing export from '{}' to '{}'", excelFile.getAbsolutePath(), propertiesFolder.getAbsolutePath());

        Workbook workbook = WorkbookFactory.create(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
        Row headerRow = sheet.getRow(0);

        int columnsCount = headerRow.getLastCellNum() - 1;
        int rowsCount = sheet.getLastRowNum() - 1;

        for (int columnNo = 0; columnNo <= columnsCount; columnNo++) {
            Cell cell = headerRow.getCell(columnNo + 1);
            if (cell == null) {
                return; // End of file
            }
            String columnName = headerRow.getCell(columnNo + 1).getStringCellValue();
            File propertiesFile = new File(propertiesFolder, columnName + "." + propertiesExtension);

            writeToFile(propertiesFile, rowsCount, sheet, columnNo);
        }
    }

    public void writeToFile(File propertiesFile, int rowsCount, Sheet sheet, int columnNo) throws IOException, ProcessingException {
        LOG.info("Exporting column no. {} to file '{}'", columnNo, propertiesFile.getName());
        try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
                new FileOutputStream(propertiesFile), StandardCharsets.UTF_8))) {
            for (int rowNo = 0; rowNo <= rowsCount; rowNo++) {
                Row row = sheet.getRow(rowNo + 1);
                String key = PropertiesUtils.saveConvert(row.getCell(0).getStringCellValue(), true, false);

                Cell cell = row.getCell(columnNo + 1);
                if (cell == null) {
                    throw new ProcessingException("Missing key '" + key + "' for file '" + propertiesFile.getName());
                }
                String value = PropertiesUtils.saveConvert(cell.getStringCellValue(), false, false);
                writer.append(key).append("=").append(value).append('\n');
            }
        }
    }
}
