package FolderReader;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.AccessDeniedException;
import java.nio.file.Files;
import java.nio.file.attribute.BasicFileAttributes;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class FolderReader implements Runnable {

    public void FolderReader() throws IOException, UnsupportedLookAndFeelException, ClassNotFoundException, InstantiationException, IllegalAccessException {

        JFrame frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        UIManager.getSystemLookAndFeelClassName();
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        // Create a JProgressBar
        JProgressBar progressBar = new JProgressBar();
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        progressBar.setStringPainted(true);

        // Add the progress bar to the frame
        frame.add(progressBar, BorderLayout.CENTER);

        try {
            // Create a list to store the file data
            List<String> fileNames = new ArrayList<>();
            List<Double> fileSize = new ArrayList<>();
            List<String> filePaths = new ArrayList<>();
            List<String> fileCreationTimes = new ArrayList<>();

            System.setProperty(System.getProperty("os.name"), "os.name");
            System.out.println(System.getProperty("os.name"));

            // Get the folder containing the files
            List<File> files = null;

            try {
                File directory = getFile();
                // Get a list of all the files in the folder
                try {
                    files = (List<File>) FileUtils.listFiles(directory, null, true);
                } catch (UncheckedIOException e) {
                    JOptionPane.showMessageDialog(null, "Acesso negado a pasta. \n" + e.getMessage());
                    throw new AccessDeniedException(e.getMessage());
                }
            } catch (NullPointerException e) {
                throw new NullPointerException();
            }

            // Set the size of the frame and make it visible
            frame.pack();
            frame.setTitle("Lendo arquivos...");
            frame.setLocationRelativeTo(null);
            frame.setSize(400, 100);
            frame.setVisible(true);

            String pattern = "dd-MM-yyyy - HH:mm:ss";
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);

            // Get the file names and store them in the lists
            assert files != null;
            for (int i = 0; i < files.size(); i++) {
                fileNames.add(files.get(i).getName());
                fileSize.add((double) files.get(i).length());
                filePaths.add(Arrays.toString(files.get(i).getPath().split(files.get(i).getName())));

                BasicFileAttributes attributes = Files.readAttributes(files.get(i).toPath(), BasicFileAttributes.class);

                Date creationDate = new Date(attributes.creationTime().toMillis());

                fileCreationTimes.add(simpleDateFormat.format(creationDate));

                int bar = (int) ((i / (double) files.size()) * 100);
                progressBar.setValue(bar);
                progressBar.setString(bar + "%");
            }

            // Create the Excel workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(); // create a sheet

            Row titles = sheet.createRow(0);
            Cell name = titles.createCell(0);
            name.setCellValue("Nome do Arquivo");
            Cell size = titles.createCell(1);
            size.setCellValue("Tamanho");
            Cell totalSize = titles.createCell(2);
            totalSize.setCellValue("Tamanho Total");
            Cell creationDate = titles.createCell(3);
            creationDate.setCellValue("Criado em");
            Cell path = titles.createCell(4);
            path.setCellValue("Caminho");


            // Write the file names to the sheet
            frame.setTitle("Escrevendo documento xlsx...");
            Cell totalSizeCell = null;
            double sum = 0;

            for (int i = 1; i < fileNames.size(); i++) {
                // create a row
                Row row = sheet.createRow(i);
                // create a cell
                Cell fileNameCell = row.createCell(0);
                Cell fileSizeCell = row.createCell(1);
                totalSizeCell = row.createCell(2);
                Cell fileCreationDateCell = row.createCell(3);
                Cell filePathCell = row.createCell(4);

                // set the cell value
                fileNameCell.setCellValue(fileNames.get(i));

                formatSize(fileSizeCell, fileSize.get(i));

                String filePath = filePaths.get(i);
                String pathFormatted = filePath
                        .replace("[/", "")
                        .replace("/]", "")
                        .replace("[", "")
                        .replace("\\]", "");

                filePathCell.setCellValue(pathFormatted);

                fileCreationDateCell.setCellValue(fileCreationTimes.get(i));

                sum += fileSize.get(i);

                int bar = (int) ((i / (double) files.size()) * 100);
                progressBar.setValue(bar);
                progressBar.setString(bar + "%");
            }

            assert totalSizeCell != null;

            formatSize(totalSizeCell, sum);

            // Close the frame
            frame.dispose();

            // Choose destination
            File selectedFile = selectDestinationOfFile();

            // Write the workbook to a file
            assert selectedFile != null;
            File outputFile = new File(selectedFile.toURI());
            FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            JOptionPane.showMessageDialog(null, "Arquivo gerado com sucesso!");
        } catch (HeadlessException | IOException e) {
            JOptionPane.showMessageDialog(null, "Arquivo não pôde ser criado.");
        }
    }

    private static void formatSize(Cell sizeCell, double sizeInBytes) {
        double kb = sizeInBytes / 1024.0;
        double mb = ((sizeInBytes / 1024.0) / 1024.0);
        double gb = (((sizeInBytes / 1024.0) / 1024.0) / 1024.0);
        double tb = ((((sizeInBytes / 1024.0) / 1024.0) / 1024.0) / 1024.0);

        if (tb > 1) {
            sizeCell.setCellValue(String.format("%.2f TB", tb));
        } else if (gb > 1) {
            sizeCell.setCellValue(String.format("%.2f GB", gb));
        } else if (mb > 1) {
            sizeCell.setCellValue(String.format("%.2f MB", mb));
        } else {
            sizeCell.setCellValue(String.format("%.2f KB", kb));
        }
    }

    private static File getFile() {
        JFileChooser chooseDirectory = new JFileChooser();
        chooseDirectory.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooseDirectory.setCurrentDirectory(new File(System.getProperty("user.home")));
        int result = chooseDirectory.showOpenDialog(null);
        File directory = null;
        if (result == JFileChooser.APPROVE_OPTION) {
            directory = chooseDirectory.getSelectedFile();
            System.out.println("Selected folder: " + directory.getAbsolutePath());
        }
        return directory;
    }

    private static File selectDestinationOfFile() {
        JFileChooser chooseFileDestination = new JFileChooser();
        chooseFileDestination.setCurrentDirectory(new File(System.getProperty("user.home")));
        chooseFileDestination.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        int res = chooseFileDestination.showSaveDialog(null);
        File selectedFile = null;
        if (res == JFileChooser.APPROVE_OPTION) {
            selectedFile = chooseFileDestination.getSelectedFile();
            if (!selectedFile.getName().endsWith(".xlsx")) {
                selectedFile = new File(selectedFile.getAbsolutePath() + ".xlsx");
                System.out.println("Selected destination: " + selectedFile);
            }
        }
        return selectedFile;
    }

    @Override
    public void run() {
        //Code for process
        try {
            FolderReader();
        } catch (IOException | UnsupportedLookAndFeelException | ClassNotFoundException | InstantiationException |
                 IllegalAccessException e) {
            JOptionPane.showMessageDialog(null, e.getMessage());
            throw new RuntimeException(e);
        }
    }
}
