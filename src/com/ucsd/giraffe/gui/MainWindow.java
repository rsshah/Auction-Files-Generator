package com.ucsd.giraffe.gui;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;

import com.ucsd.giraffe.main.AuctionFilesGenerator;

public class MainWindow extends JFrame {

    final AuctionFilesGenerator generator;
    
    JFileChooser srcExplorer = new JFileChooser();
    JFileChooser dstExplorer = new JFileChooser();
    
    JLabel selectFile = new JLabel("Select target file:");
    JTextField targetFile = new JTextField();
    JButton browseSrcButton = new JButton("Browse");
    
    JLabel selectDstLocation = new JLabel("Select target location:");
    JTextField targetLocation = new JTextField();
    JButton browseDstButton = new JButton("Browse");
    
    JButton writeButton = new JButton("OK");
    JButton closeButton = new JButton("Close");
    
    public MainWindow(AuctionFilesGenerator generator) {
        this.setPreferredSize(new Dimension(300, 220));
        this.generator = generator;
        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        JPanel srcPanel = new JPanel();
        srcPanel.setLayout(new BoxLayout(srcPanel, BoxLayout.Y_AXIS));
        JPanel selectSrcPanel = new JPanel();
        selectSrcPanel.setLayout(new GridLayout(2, 1));
        selectSrcPanel.add(selectFile);
        selectSrcPanel.add(targetFile);
        srcPanel.add(selectSrcPanel);   
        // add src browse button and its mouse handler
        browseSrcButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                getExcelFile();
            }
        });
        srcPanel.add(browseSrcButton);
        browseSrcButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        mainPanel.add(srcPanel);
        
        JPanel dstPanel = new JPanel();
        dstPanel.setLayout(new BoxLayout(dstPanel, BoxLayout.Y_AXIS));
        JPanel selectDstPanel = new JPanel();
        selectDstPanel.setLayout(new GridLayout(2, 1));
        selectDstPanel.add(selectDstLocation);
        selectDstPanel.add(targetLocation);
        dstPanel.add(selectDstPanel);
        // add dst browse button and its handler
        dstExplorer.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        browseDstButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                getTargetDir();
            }
        });
        dstPanel.add(browseDstButton);
        browseDstButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        mainPanel.add(dstPanel);
        
        // add generate button and its mouse handler
        writeButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                generateFiles();
            }
        });
        closeButton.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                System.exit(0);
            }
        });
        JPanel commandPanel = new JPanel();
        commandPanel.add(closeButton);
        commandPanel.add(writeButton);
        commandPanel.setBorder(new EmptyBorder(10, 0, 0, 0));
        mainPanel.add(commandPanel);
        this.add(mainPanel);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setResizable(false);
        this.setTitle("Welcome to Excel Reader!");
        this.pack();
        this.setVisible(true);
    }

    private void getExcelFile() {
        int ret = srcExplorer.showOpenDialog(MainWindow.this);
        if (ret == JFileChooser.APPROVE_OPTION) {
            File file = srcExplorer.getSelectedFile();
            targetFile.setText(file.getPath());
        }
    }
    
    private void getTargetDir() {
        int ret = dstExplorer.showOpenDialog(MainWindow.this);
        if (ret == JFileChooser.APPROVE_OPTION) {
            File file = dstExplorer.getSelectedFile();
            targetLocation.setText(file.getPath());
        }
    }
    
    private void generateFiles() {
        File f = new File(targetFile.getText());
        if (!f.exists()) {
            JOptionPane.showMessageDialog(null, "Invalid excel file.");
        }
        generator.setExcelFile(f);
        File f2 = new File(targetLocation.getText());
        if (!f2.exists()) {
            f2.mkdirs();
        }
        generator.dir = f2.getPath();
        MainWindow.this.generator.generate();
    }
}
