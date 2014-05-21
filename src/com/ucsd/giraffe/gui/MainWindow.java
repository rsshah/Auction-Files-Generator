package com.ucsd.giraffe.gui;

import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;

import com.ucsd.giraffe.main.AuctionFilesGenerator;

public class MainWindow extends JFrame {

    final AuctionFilesGenerator generator;
    JButton generateButton = new JButton("Generate Word Documents");
    JFileChooser fileExplorer = new JFileChooser();
    JButton browseButton = new JButton("Browse for Excel File");

    public MainWindow(AuctionFilesGenerator generator) {
        this.generator = generator;
        JPanel panel = new JPanel();
        // add browse button and its mouse handler
        panel.add(browseButton);
        browseButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                getExcelFile();
            }
        });
        // add generate button and its mouse handler
        panel.add(generateButton);
        generateButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                MainWindow.this.generator.generate();
            }
        });
        this.add(panel);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setTitle("Auction Files Generator");
        this.pack();
    }

    private void getExcelFile() {
        int ret = fileExplorer.showOpenDialog(MainWindow.this);
        if (ret == JFileChooser.APPROVE_OPTION) {
            File file = fileExplorer.getSelectedFile();
            System.out.println("Chosen: " + file.getName());
            generator.setExcelFile(file);
        }
    }
}
