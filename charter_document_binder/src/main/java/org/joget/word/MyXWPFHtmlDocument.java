package org.joget.word;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;

public class MyXWPFHtmlDocument extends POIXMLDocumentPart {

    String html;
    String id;

    public MyXWPFHtmlDocument(PackagePart part, String id) {
        super(part);
        this.html = "<!DOCTYPE html>"
                + "<html>"
                + "<head>"
                + "<style>"
                + "body { font-family: 'Calibri', sans-serif; font-size: 11pt; }"
                + "</style>"
                + "<title>HTML import</title>"
                + "</head>"
                + "<body></body>"
                + "</html>";
        this.id = id;
    }

    String getId() {
        return id;
    }

    String getHtml() {
        return html;
    }

    void setHtml(String html) {
        this.html = html;
    }

    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        Writer writer = new OutputStreamWriter(out, "UTF-8");
        writer.write(html);
        writer.close();
        out.close();
    }
}
