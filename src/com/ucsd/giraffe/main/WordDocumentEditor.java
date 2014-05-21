package com.ucsd.giraffe.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.ucsd.giraffe.main.extractor.ExtractedRow;

public class WordDocumentEditor {
    private List<WordDocumentProperties> documents = new ArrayList<>();
    private File dir;

    public WordDocumentEditor(String title) {
        dir = new File(title);
    }

    public void registerDocument(WordDocumentProperties document) {
        documents.add(document);
    }

    public void edit(List<ExtractedRow> rows) throws Exception {
        // create the folder for this event
        if (!dir.exists()) {
            dir.mkdir();
        }
        for (int i = 0; i < rows.size(); i++) {
            ExtractedRow row = rows.get(i);
            // go through each file and see if this row is eligible to generate
            // the file
            for (WordDocumentProperties prop : documents) {
                boolean eligible = true;
                for (Map.Entry<String, String> field : prop.getFields().entrySet()) {
                    if (!row.getData().containsKey(field.getValue())) {
                        System.out.println("Row " + i + " does not contain: " + field.getValue());
                        System.out.println("\t" + row.getData());
                        eligible = false;
                        break;
                    }
                }
                if (!eligible) { // check eligibility for next file
                    continue;
                }
                String dn = prop.getDocumentName();
                // write things to file
                if (dn.endsWith(".doc")) { // Word 2003 file
                    generateDocFile(row, i + 1, prop);
                } else if (dn.endsWith(".docx")) { // Word 2007+
                    generateDocxFile(row, i + 1, prop);
                } else {
                    throw new IllegalArgumentException(
                            "Unhandled extension for registered document: " + dn);
                }
            }
        }
    }

    public void generateDocFile(ExtractedRow row, int num, WordDocumentProperties prop)
            throws Exception {
        HWPFDocument doc = new HWPFDocument(new FileInputStream(prop.getDocumentName()));
        // WordExtractor a = new WordExtractor(doc);
        // String text = a.getText();
        for (Map.Entry<String, String> field : prop.getFields().entrySet()) {
            doc.getRange().replaceText(field.getKey(), row.getData().get(field.getValue()));
        }
        for (Map.Entry<String, String> replacement : prop.getReplacements().entrySet()) {
            doc.getRange().replaceText(replacement.getKey(), replacement.getValue());
        }

        // System.out.println("Text of file: " + prop.getDocumentName());
        // System.out.println(text);
        // System.out.println("\r\n\r\n\r\n");
        File f = new File(dir + "/" + "Item" + num + "-" + prop.getDocumentName());
        f.createNewFile();
        FileOutputStream fout = new FileOutputStream(f);
        doc.write(fout);
        fout.close();
    }

    public void generateDocxFile(ExtractedRow row, int num, WordDocumentProperties prop)
            throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(prop.getDocumentName()));
        XWPFWordExtractor a = new XWPFWordExtractor(doc);
        String text = a.getText();
        // System.out.println("Text of file: " + prop.getDocumentName());
        // System.out.println(text);
        // System.out.println("\r\n\r\n\r\n");
        a.close();
    }
}
