import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    public static void main(String[] args) throws EncryptedDocumentException, IOException {
        // Create a list to store the file data
        List<String> fileNames = new ArrayList<>();
        List<Double> fileSize = new ArrayList<>();
        List<String> filePaths = new ArrayList<>();
        List<String> fileCreationTimes = new ArrayList<>();

        System.out.println(System.getProperty("os.name"));

        // Get the folder containing the files
        File folder = new File("../../gpmrks");
//        File folder = new File(args[1]);

        // Get a list of all the files in the folder
        File[] files = folder.listFiles();
        System.out.println();

        String pattern = "dd-MM-yyyy - HH:mm:ss";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);

        // Get the file names and store them in the list
        for (File file : files) {
            fileNames.add(file.getName());
            fileSize.add((double) file.length());
            filePaths.add(Arrays.toString(file.getPath().split(file.getName())).toString());

            BasicFileAttributes attributes = Files.readAttributes(file.toPath(), BasicFileAttributes.class);

            Date creationDate = new Date(attributes.creationTime().toMillis());

            fileCreationTimes.add(simpleDateFormat.format(creationDate));

        }

        // Create the Excel workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(); // create a sheet

        Row titles = sheet.createRow(0);
        Cell name = titles.createCell(0);
        name.setCellValue("Nome");
        Cell size = titles.createCell(1);
        size.setCellValue("Tamanho");
        Cell path = titles.createCell(2);
        path.setCellValue("Caminho");
        Cell creationDate = titles.createCell(3);
        creationDate.setCellValue("Criado em");

        // Write the file names to the sheet
        for (int i = 1; i < fileNames.size(); i++) {
            // create a row
            Row row = sheet.createRow(i);
            // create a cell
            Cell fileNameCell = row.createCell(0);
            Cell fileSizeCell = row.createCell(1);
            Cell filePathCell = row.createCell(2);
            Cell fileCreationDate = row.createCell(3);
            // set the cell value
            fileNameCell.setCellValue(fileNames.get(i));

            double sizeFormated = (fileSize.get(i) / (1024 * 1024));
            fileSizeCell.setCellValue(String.format("%.3f KB", sizeFormated));

            filePathCell.setCellValue(filePaths.get(i));

            fileCreationDate.setCellValue(fileCreationTimes.get(i));
        }

        // Write the workbook to a file
        File outputFile = new File("../../gpmrks/data.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }
}
