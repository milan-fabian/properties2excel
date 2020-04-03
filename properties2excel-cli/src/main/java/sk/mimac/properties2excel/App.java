package sk.mimac.properties2excel;

import java.io.IOException;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class App {

    private static final Logger LOG = LoggerFactory.getLogger(App.class);

    private static final Options options = new Options();

    static {
        options.addOption("t", "to-excel", false, "Convert from properties files to Excel file");
        options.addOption("f", "from-excel", false, "Convert from Excel file to properties files");
        options.addRequiredOption("e", "excel-file", true, "Path to Excel file");
        options.addRequiredOption("p", "properties-folder", true, "Path to folder with properties files");
        options.addOption("h", "help", false, "Display usage");
        options.addOption("x", "extension", true, "Property file extension (default is 'properties')");
    }

    public static void main(String[] args) {
        CommandLineParser parser = new DefaultParser();
        CommandLine cmd;
        try {
            cmd = parser.parse(options, args);
        } catch (ParseException ex) {
            System.out.println("Error: " + ex.getMessage());
            printHelp();
            return;
        }

        if (cmd.hasOption("to-excel") == cmd.hasOption("from-excel")) { // Both true or both false
            System.out.println("Error: Exactly one parameter from 't', 'f' should be used");
            printHelp();
            return;
        }

        if (cmd.hasOption("help")) {
            printHelp();
            return;
        }

        run(cmd);
    }

    private static void run(CommandLine cmd) {
        String excelFile = cmd.getOptionValue("excel-file");
        String propertiesFolder = cmd.getOptionValue("properties-folder");
        String propertiesExtension = cmd.getOptionValue("extension", "properties");
        try {
            if (cmd.hasOption("to-excel")) {
                new ExcelImportProcessor(excelFile, propertiesFolder, propertiesExtension).process();
            } else {
                new ExcelExportProcessor(excelFile, propertiesFolder, propertiesExtension).process();
            }
            LOG.info("Done");
        } catch (ProcessingException ex) {
            LOG.error("Can't process: " + ex.getMessage());
        } catch (IOException ex) {
            LOG.error("Can't process", ex);
        }
    }

    private static void printHelp() {
        System.out.println();
        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp("java -jar properties-to-excel.jar", options, true);
    }

}
