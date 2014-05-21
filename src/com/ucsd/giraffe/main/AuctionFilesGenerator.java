package com.ucsd.giraffe.main;

import java.io.File;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.UIManager;

import com.ucsd.giraffe.gui.MainWindow;
import com.ucsd.giraffe.main.extractor.ExcelDataExtractor;
import com.ucsd.giraffe.main.extractor.ExtractedRow;

public class AuctionFilesGenerator {
    private File excelFile = null;
    private JFrame window = null;

    public static void main(String[] args) {
        AuctionFilesGenerator generator = new AuctionFilesGenerator();
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }
        generator.window = new MainWindow(generator);
        generator.window.setLocationRelativeTo(null);
        generator.window.setVisible(true);
    }

    public void setExcelFile(File f) {
        excelFile = f;
    }

    public void generate() {
        if (excelFile == null) {
            JOptionPane.showMessageDialog(window, "Error! Was the Excel file moved?", "Error!",
                    JOptionPane.ERROR_MESSAGE);
            return;
        }
        ExcelDataExtractor extractor = new ExcelDataExtractor(excelFile);
        List<ExtractedRow> rows = extractor.extract();
        WordDocumentEditor editor = new WordDocumentEditor(extractor.getTitle());
        // TODO to improve usability, make this a UI control or make it read
        // from a file
        // Add donor replacement fields
        WordDocumentProperties donorThanks = new WordDocumentProperties("donor letter blank.doc");
        donorThanks.addField("<  field c                > ", "donor name");
        donorThanks.addField("< field d                     > ", "description for bid card");
        donorThanks.addReplacement("< field b >", extractor.getTitle());
        donorThanks.addField("_____ Value", "value");
        editor.registerDocument(donorThanks);
        // Add buyer replacement fields
        WordDocumentProperties buyerThanks = new WordDocumentProperties(
                "purchaser letter blank.doc");
        buyerThanks.addField("<  field m                > ", "buyer's name");
        buyerThanks.addField("< field d                     >", "description for bid card");
        buyerThanks.addReplacement("< field b   >  ", extractor.getTitle());
        editor.registerDocument(buyerThanks);
        // Add bid cards and signs (TODO)

        try {
            editor.edit(rows);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(window,
                    "Error! Are you sure the files are named correctly?", "Error!",
                    JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
            return;
        }
        JOptionPane.showMessageDialog(window, "Successfully generated Word files.", "Success!",
                JOptionPane.PLAIN_MESSAGE);
    }

}
