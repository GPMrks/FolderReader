package FolderReader;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;

public class LoadingScreen extends JDialog {

    public static JProgressBar ProgressBar() {
        // Create a new JFrame
        JFrame frame = new JFrame("Progress...");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Create a JProgressBar
        JProgressBar progressBar = new JProgressBar();
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        progressBar.setStringPainted(true);

        // Add the progress bar to the frame
        frame.add(progressBar, BorderLayout.CENTER);

        // Set the size of the frame and make it visible
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setSize(400, 50);
        frame.setVisible(true);

        // Close the frame
        frame.dispose();
        return progressBar;
    }
}
