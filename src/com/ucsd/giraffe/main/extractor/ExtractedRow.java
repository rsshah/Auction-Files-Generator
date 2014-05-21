package com.ucsd.giraffe.main.extractor;

import java.util.HashMap;
import java.util.Map;

public class ExtractedRow {
    private Map<String, String> fields = new HashMap<>();

    public void addField(String key, String value) {
        fields.put(key, value);
    }

    public Map<String, String> getData() {
        return fields;
    }
}
