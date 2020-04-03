package sk.mimac.properties2excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelImportProcessor extends AbstractProcessor {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelExportProcessor.class);

    private File[] propertiesFiles;
    private Sheet sheet;
    private CellStyle boldCellStyle;

    public ExcelImportProcessor(String excelFile, String propertiesFolder, String propertiesExtension) {
        super(excelFile, propertiesFolder, propertiesExtension);
    }

    @Override
    protected void validateInternal() throws ProcessingException {
        propertiesFiles = propertiesFolder.listFiles((dir, name) -> name.toLowerCase().endsWith("." + propertiesExtension));
        if (propertiesFiles == null || propertiesFiles.length == 0) {
            throw new ProcessingException("Properties folder '" + propertiesFolder.getAbsolutePath() + "' doesn't contain any property file");
        }
        if (excelFile.exists()) {
            throw new ProcessingException("Excel file '" + excelFile.getAbsolutePath() + "' already exists");
        }
    }

    @Override
    protected void processInternal() throws IOException, ProcessingException {
        LOG.info("Processing import from '{}' to '{}'", propertiesFolder.getAbsolutePath(), excelFile.getAbsolutePath());

        Workbook workbook = excelFile.getName().endsWith(".xls") ? new HSSFWorkbook() : new XSSFWorkbook();
        sheet = workbook.createSheet("properties");
        createBoldCellStyle();

        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue("keys");
        cell.setCellStyle(boldCellStyle);

        Map<String, Integer> orderMap = fillKeys();

        int i = 0;
        for (File file : propertiesFiles) {
            LOG.info("Importing file '{}'", file.getAbsolutePath());
            fillValues(file, orderMap, i++);
        }

        sheet.autoSizeColumn(0);
        try (OutputStream fileOut = new FileOutputStream(excelFile)) {
            workbook.write(fileOut);
        }
    }

    private void createBoldCellStyle() {
        boldCellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        boldCellStyle.setFont(font);
    }

    private Map<String, Integer> fillKeys() throws IOException, ProcessingException {
        List<String[]> lines = loadLines(propertiesFiles[0]);
        int i = 0;
        Map<String, Integer> orderMap = new HashMap<>();
        for (String[] line : lines) {
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(line[0]);
            Integer oldKey = orderMap.put(line[0], i);
            if (oldKey != null) {
                throw new ProcessingException("Duplicate key '" + line[0] + "' on lines " + (oldKey + 1) + " and " + (i + 1)
                        + " in file '" + propertiesFiles[0].getName() + "'");
            }
            i++;
        }
        return orderMap;
    }

    private void fillValues(File propertiesFile, Map<String, Integer> orderMap, int column) throws IOException, ProcessingException {
        String name = propertiesFile.getName();
        name = name.substring(0, name.lastIndexOf('.')); // without extension

        Cell cell = sheet.getRow(0).createCell(column + 1);
        cell.setCellValue(name);
        cell.setCellStyle(boldCellStyle);

        List<String[]> lines = loadLines(propertiesFile);

        if (orderMap.size() != lines.size()) {
            throw new ProcessingException("File '" + propertiesFile.getName() + "' contains invalid number of lines: "
                    + lines.size() + " instead of " + orderMap.size());
        }

        for (String[] line : lines) {
            Integer rowNo = orderMap.get(line[0]);
            if (rowNo == null) {
                throw new ProcessingException("File '" + propertiesFile.getName() + "' contains unknown extra key '" + line[0] + "'");
            }
            sheet.getRow(rowNo + 1).createCell(column + 1).setCellValue(line[1]);
        }
    }

    private List<String[]> loadLines(File propertiesFile) throws IOException {
        try (Reader reader = new InputStreamReader(new FileInputStream(propertiesFile), StandardCharsets.UTF_8)) {
            return PropertiesUtils.load(new LineReader(reader));
        }
    }

}
