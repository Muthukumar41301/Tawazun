package org.joget.charter;

import org.apache.poi.ooxml.POIXMLRelation;

public class XWPFHtmlRelation extends POIXMLRelation {

    public XWPFHtmlRelation() {
        super(
                "text/html",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
                "/word/htmlDoc#.html");
    }
}
