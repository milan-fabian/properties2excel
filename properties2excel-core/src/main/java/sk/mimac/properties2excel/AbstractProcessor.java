package sk.mimac.properties2excel;

import java.io.File;
import java.io.IOException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public abstract class AbstractProcessor {

    private static final Logger LOG = LoggerFactory.getLogger(AbstractProcessor.class);

    protected final File excelFile;
    protected final File propertiesFolder;
    protected final String propertiesExtension;

    public AbstractProcessor(String excelFile, String propertiesFolder, String propertiesExtension) {
        this.excelFile = new File(excelFile);
        this.propertiesFolder = new File(propertiesFolder);
        this.propertiesExtension = propertiesExtension;
    }

    public void process() throws IOException, ProcessingException {
        validate();
        processInternal();
    }

    private void validate() throws ProcessingException {
        if (!propertiesFolder.exists()) {
            throw new ProcessingException("Properties folder '" + propertiesFolder.getAbsolutePath() + "' doesn't exist");
        }
        if (!propertiesFolder.isDirectory()) {
            throw new ProcessingException("Properties folder '" + propertiesFolder.getAbsolutePath() + "' is not folder");
        }
        String excelFileName = excelFile.getName();
        if (!excelFileName.endsWith(".xls") && !excelFileName.endsWith(".xlsx")) {
            throw new ProcessingException("Only .xls and .xlsx extensions are supported for Excel file");
        }
        validateInternal();
    }

    protected abstract void validateInternal() throws ProcessingException;

    protected abstract void processInternal() throws IOException, ProcessingException;

}
