package com.ucsd.giraffe.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;
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

    private void generateDocFile(ExtractedRow row, int num, WordDocumentProperties prop)
            throws Exception {
        HWPFDocument doc = new HWPFDocument(new FileInputStream(prop.getDocumentName()));
        for (Map.Entry<String, String> field : prop.getFields().entrySet()) {
            doc.getRange().replaceText(field.getKey(), row.getData().get(field.getValue()));
        }
        for (Map.Entry<String, String> replacement : prop.getReplacements().entrySet()) {
            doc.getRange().replaceText(replacement.getKey(), replacement.getValue());
        }
        String docName = prop.getDocumentName().split(" blank")[0];
        String directory = AuctionFilesGenerator.dir;
        File f = new File(directory + "/" + dir + "/Item" + num + "-" + docName + ".doc");
        f.createNewFile();
        FileOutputStream fout = new FileOutputStream(f);
        doc.write(fout);
        fout.close();
    }

    private void generateDocxFile(ExtractedRow row, int num, WordDocumentProperties prop)
            throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(prop.getDocumentName()));
        for (Map.Entry<String, String> field : prop.getFields().entrySet()) {

        }
    }

    public void generateCardsAndSigns(List<ExtractedRow> rows, WordDocumentProperties prop)
            throws Exception {
        HWPFDocument doc = new HWPFDocument(new FileInputStream(prop.getDocumentName()));
        TableIterator itr = new TableIterator(doc.getRange());
        for (; itr.hasNext();) {
            Table t = itr.next();
            
            System.out.println("Table: " + t);
        }
        // iterate through all items and generate a bid card
        // there will be 2 cards per page
        for (ExtractedRow row : rows) {
            
        }
    }
}
