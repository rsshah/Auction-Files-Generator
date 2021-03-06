package com.ucsd.giraffe.main;

import java.util.HashMap;
import java.util.Map;

public class WordDocumentProperties {
    private String docName;
    private Map<String, String> fields = new HashMap<>();
    private Map<String, String> replacements = new HashMap<>();
    private boolean hasTables = false;

    public WordDocumentProperties(String docName) {
        this.docName = docName;
    }
    
    public WordDocumentProperties(String docName, boolean hasTables) {
        this(docName);
        this.hasTables = hasTables;
    }

    public void addField(String key, String value) {
        fields.put(key, value);
    }

    public Map<String, String> getFields() {
        return fields;
    }

    public void addReplacement(String key, String value) {
        replacements.put(key, value);
    }

    public Map<String, String> getReplacements() {
        return replacements;
    }

    public String getDocumentName() {
        return docName;
    }
}
