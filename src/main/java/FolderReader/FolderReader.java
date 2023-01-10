package FolderReader;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.SimpleDateFormat;
import java.util.*;

public class FolderReader implements Runnable {
    public void FolderReader() throws IOException {

        // Create a list to store the file data
        List<String> fileNames = new ArrayList<>();
        List<Double> fileSize = new ArrayList<>();
        List<String> filePaths = new ArrayList<>();
        List<String> fileCreationTimes = new ArrayList<>();

        System.setProperty(System.getProperty("os.name"), "os.name");
        System.out.println(System.getProperty("os.name"));

        // Get the folder containing the files
        JFileChooser chooseDirectory = new JFileChooser();
        chooseDirectory.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooseDirectory.setCurrentDirectory(new File(System.getProperty("user.home")));
        int result = chooseDirectory.showOpenDialog(null);
        File directory = null;
        if (result == JFileChooser.APPROVE_OPTION) {
            directory = chooseDirectory.getSelectedFile();
            System.out.println("Selected file: " + directory.getAbsolutePath());
        }

        // Get a list of all the files in the folder
        Collection<File> files = FileUtils.listFiles(directory, null, true);
//        File[] files = directory.listFiles();
        System.out.println();

        String pattern = "dd-MM-yyyy - HH:mm:ss";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);

        // Get the file names and store them in the list
        for (File file : files) {
            fileNames.add(file.getName());
            fileSize.add((double) file.length());
            filePaths.add(Arrays.toString(file.getPath().split(file.getName())));

            BasicFileAttributes attributes = Files.readAttributes(file.toPath(), BasicFileAttributes.class);

            Date creationDate = new Date(attributes.creationTime().toMillis());

            fileCreationTimes.add(simpleDateFormat.format(creationDate));

        }

        // Create the Excel workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(); // create a sheet

        Row titles = sheet.createRow(0);
        Cell name = titles.createCell(0);
        name.setCellValue("Nome do Arquivo");
        Cell size = titles.createCell(1);
        size.setCellValue("Tamanho");
        Cell path = titles.createCell(2);
        path.setCellValue("Caminho");
        Cell creationDate = titles.createCell(3);
        creationDate.setCellValue("Criado em");
        Cell totalSize = titles.createCell(4);
        totalSize.setCellValue("Tamanho Total");

        JProgressBar progressBar = LoadingScreen.ProgressBar();

        double sum = 0;
        // Write the file names to the sheet
        Cell totalSizeCell = null;

        for (int i = 1; i < fileNames.size(); i++) {
            // create a row
            Row row = sheet.createRow(i);
            // create a cell
            Cell fileNameCell = row.createCell(0);
            Cell fileSizeCell = row.createCell(1);
            Cell filePathCell = row.createCell(2);
            Cell fileCreationDateCell = row.createCell(3);
            totalSizeCell = row.createCell(4);

            // set the cell value
            fileNameCell.setCellValue(fileNames.get(i));

            double sizeFormated = (fileSize.get(i) / (1024 * 1024));
            fileSizeCell.setCellValue(String.format("%.3f KB", sizeFormated));

            filePathCell.setCellValue(filePaths.get(i));

            fileCreationDateCell.setCellValue(fileCreationTimes.get(i));

            sum += sizeFormated;

            int bar = (int) ((i / (double) fileNames.size()) * 100);
            progressBar.setValue(bar);
            progressBar.setString(bar + "%");

        }

        totalSizeCell.setCellValue(String.format("%.3f KB", sum));

        JFileChooser chooseFileDestination = new JFileChooser();
        chooseFileDestination.setCurrentDirectory(new

                File(System.getProperty("user.home")));
        chooseFileDestination.setFileFilter(new

                FileNameExtensionFilter("Excel Files", "xlsx"));
        int res = chooseFileDestination.showSaveDialog(null);
        File selectedFile = null;
        if (res == JFileChooser.APPROVE_OPTION) {
            selectedFile = chooseFileDestination.getSelectedFile();
            if (!selectedFile.getName().endsWith(".xlsx")) {
                selectedFile = new File(selectedFile.getAbsolutePath() + ".xlsx");
            }
        }

        // Write the workbook to a file
        File outputFile = new File(selectedFile.toURI());
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();

        JOptionPane.showMessageDialog(null, "Arquivo gerado com sucesso!");
    }

    @Override
    public void run() {
        //Code for process
        try {
            FolderReader();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
