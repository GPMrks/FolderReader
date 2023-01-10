package FolderReader;

import org.apache.poi.EncryptedDocumentException;

import javax.swing.*;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws EncryptedDocumentException, IOException {
        FolderReader folderReader = new FolderReader();
        try {
            folderReader.run();
        } catch (NullPointerException e) {
            JOptionPane.showMessageDialog(null, "Cancelado");
        }
    }
}

