import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import java.io.*;
import org.apache.poi.xwpf.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;




public class LineBarChartExample extends JFrame {

    private JTextField filePathField;
    private JButton browseButton;
    private JButton uploadButton;
    private JLabel filePathLabel;
    private JTabbedPane tabbedPane;
    private DefaultCategoryDataset dataset;
    private DefaultCategoryDataset dataset2;

    public LineBarChartExample(String applicationTitle, String chartTitle) {
        super(applicationTitle);

        dataset = new DefaultCategoryDataset();
        dataset2 = new DefaultCategoryDataset();

        setTitle("CSV Uploader and Chart Generator");
        setSize(700, 300);
        setLayout(new GridBagLayout());
        setDefaultCloseOperation(EXIT_ON_CLOSE);

        // Using a GridBagConstraints object to specify the layout of the components
        GridBagConstraints c = new GridBagConstraints();

        filePathLabel = new JLabel("File Path: ");
        filePathLabel.setFont(new Font("Arial", Font.BOLD, 14));
        filePathLabel.setForeground(Color.DARK_GRAY);
        c.gridx = 0;
        c.gridy = 0;
        c.insets = new Insets(10,10,10,10);
        add(filePathLabel, c);

        filePathField = new JTextField(40);
        filePathField.setBackground(Color.WHITE);
        filePathField.setForeground(Color.BLACK);
        filePathField.setFont(new Font("Arial", Font.PLAIN, 14));
        c.gridx = 1;
        c.gridy = 0;
        add(filePathField, c);

        browseButton = new JButton("Browse");
        browseButton.addActionListener(new BrowseListener());
        browseButton.setBackground(new Color(0, 153, 153));
        browseButton.setForeground(Color.black);
        browseButton.setFont(new Font("Arial", Font.BOLD, 14));
        c.gridx = 2;
        c.gridy = 0;
        add(browseButton, c);
        uploadButton = new JButton("Upload");
        uploadButton.addActionListener(new UploadListener());
        uploadButton.setBackground(new Color(0, 153, 153));
        uploadButton.setForeground(Color.BLACK);
        uploadButton.setFont(new Font("Arial", Font.BOLD, 14));
        c.gridx = 3;
        c.gridy = 0;
        add(uploadButton, c);

        tabbedPane = new JTabbedPane();
        c.gridx = 0;
        c.gridy = 1;
        c.gridwidth = 4;
        c.insets = new Insets(20,10,10,10);
        add(tabbedPane, c);


    }

    private class BrowseListener implements ActionListener {
        public void actionPerformed(ActionEvent e) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            int result = fileChooser.showOpenDialog(LineBarChartExample.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                filePathField.setText(selectedFile.getAbsolutePath());
            }
        }
    }

    private class UploadListener implements ActionListener {
        public void actionPerformed(ActionEvent e) {
            String filePath = filePathField.getText();
            File file = new File(filePath);
            if (file.exists()) {
                try {
                    FileReader filereader = new FileReader(filePath);
                    CSVParser csvParser = new CSVParser(filereader, CSVFormat.DEFAULT);
                    for (CSVRecord record : csvParser) {
                        String month = record.get(0);
                        String value = record.get(1);
                        if (!value.isEmpty()) {
                            double doubleValue = Double.parseDouble(value);
                            dataset.addValue(doubleValue, "Deaths", month);
                            dataset2.addValue(doubleValue, "Deaths", month);
                        }
                    }
                    csvParser.close();
                } catch (NumberFormatException e1) {
                    System.out.println("Invalid value found in the file, please check the file for empty or non-numeric values");
                    e1.printStackTrace();
                } catch (IOException e1) {
                    System.out.println("No file found");
                    e1.printStackTrace();
                }
                // Creating line chart
                String chartTitle;
                JFreeChart lineChart = ChartFactory.createLineChart(
                        chartTitle="Covid Death record",
                        "Month", "Number of Deaths",
                        dataset,
                        PlotOrientation.VERTICAL,
                        true, true, false);

                ChartPanel chartPanel = new ChartPanel(lineChart);

                // Creating bar chart
                JFreeChart barChart = ChartFactory.createBarChart(
                        "Covid death Rate ",
                        "Month", "Number of death",
                        dataset2,
                        PlotOrientation.VERTICAL,
                        true, true, false);

                ChartPanel chartPanel2 = new ChartPanel(barChart);

                tabbedPane.addTab("Line Chart", chartPanel);
                tabbedPane.addTab("Bar Chart", chartPanel2);

                setContentPane(tabbedPane);
                JOptionPane.showMessageDialog(LineBarChartExample.this, "File uploaded and charts generated successfully!");
            } else {
                JOptionPane.showMessageDialog(LineBarChartExample.this, "Invalid file path. Please select a valid file.");
            }

        }

    }

    private static class ConclusionToWordConverter extends JFrame {

        private JTextArea conclusionTextArea;
        private JButton convertButton;



        public ConclusionToWordConverter() {
            super("Write your Conclusion and Convert to word");
            setSize(400, 300);
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setLocationRelativeTo(null);

            // create conclusion text area
            conclusionTextArea = new JTextArea();
            conclusionTextArea.setLineWrap(true);
            JScrollPane scrollPane = new JScrollPane(conclusionTextArea);
            scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            getContentPane().add(scrollPane, BorderLayout.CENTER);

            // create convert button
            convertButton = new JButton("Save as MS word");
            convertButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    convertToWord();
                }
            });
            getContentPane().add(convertButton, BorderLayout.SOUTH);
        }

        private void convertToWord() {
            try {
                // create Word document
                XWPFDocument document = new XWPFDocument();
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();

                // set conclusion text
                String conclusionText = conclusionTextArea.getText();
                run.setText(conclusionText);

                // save Word document to file
                JFileChooser fileChooser = new JFileChooser();
                int userSelection = fileChooser.showSaveDialog(this);
                if (userSelection == JFileChooser.APPROVE_OPTION) {
                    File fileToSave = fileChooser.getSelectedFile();
                    FileOutputStream out = new FileOutputStream(fileToSave);
                    document.write(out);
                    out.close();
                    JOptionPane.showMessageDialog(this, "File saved successfully.");
                }
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Error: " + ex.getMessage());
            }
        }


    }

    public static void main(String[] args) {
        LineBarChartExample chart = new LineBarChartExample("Covid-19 death rate", "Number of Death vs Months");
        chart.pack();
        chart.setVisible(true);
        ConclusionToWordConverter app = new ConclusionToWordConverter();
        app.setVisible(true);

    }



}


