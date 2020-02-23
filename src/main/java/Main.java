import org.jodconverter.LocalConverter;
import org.jodconverter.document.DefaultDocumentFormatRegistry;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.nio.file.Paths;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Main {

    public static File getDefaultOfficeHome() {


        return Paths.get("/Applications/OpenOffice.app/Contents").toFile();
    }

    public static void main(String[] args) {

        String inputFile = "/Users/macmac3/Documents/word2pdf/SPA Template-converted.docx";
        String outputFile = "/Users/macmac3/Documents/word2pdf/SPA Template-converted.pdf";


        LocalOfficeManager localOfficeManager = LocalOfficeManager.builder()
                .install()
                .officeHome(getDefaultOfficeHome()) //your path to openoffice
                .build();
        try {
            localOfficeManager.start();
            final DocumentFormat format
                    = DocumentFormat.builder()
                    .from(DefaultDocumentFormatRegistry.PDF)
                    .build();

            LocalConverter
                    .make()
                    .convert(new FileInputStream(new File(inputFile)))
                    .as(DefaultDocumentFormatRegistry.DOCX)
                    .to(new File(outputFile))
                    .as(format)
                    .execute();

        } catch (
                OfficeException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (
                FileNotFoundException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            OfficeUtils.stopQuietly(localOfficeManager);
        }
    }
}

