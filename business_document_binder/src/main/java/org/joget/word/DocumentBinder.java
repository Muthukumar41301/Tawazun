package org.joget.word;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.joget.apps.app.dao.FormDefinitionDao;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.model.FormDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppResourceUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.dao.FormDataDao;
import org.joget.apps.form.model.*;
import org.joget.apps.form.service.FileUtil;
import org.joget.apps.form.service.FormService;
import org.joget.apps.form.service.FormUtil;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.UuidGenerator;
import org.joget.plugin.base.PluginManager;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAltChunk;

import javax.sql.DataSource;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class DocumentBinder extends FormBinder implements FormStoreBinder, FormStoreElementBinder, FormStoreMultiRowElementBinder {

    private static final String MESSAGE_PATH = "messages/DocumentBinder";

    private FormService formService;

    PreparedStatement stmt;
    ResultSet rs;
    Connection con;

    @Override
    public String getName() {
        return "Business Document Binder";
    }

    @Override
    public String getVersion() {
        return "1.0.0";
    }

    @Override
    public String getDescription() {
        return  AppPluginUtil.getMessage("org.joget.word.DocumentBinder.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("org.joget.word.DocumentBinder.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getClassName() {
        return this.getClass().getName();
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/DocumentBinder.json", null, true, MESSAGE_PATH);
    }

    protected String getFormPropertyName(String tableName, String propertyName) {
        if (propertyName != null && !propertyName.isEmpty()) {
            FormDataDao formDataDao = (FormDataDao)AppUtil.getApplicationContext().getBean("formDataDao");
            Collection<String> columnNames = formDataDao.getFormDefinitionColumnNames(tableName);
            if (columnNames.contains(propertyName) && !"id".equals(propertyName)) {
                propertyName = "customProperties." + propertyName;
            }
        }

        return propertyName;
    }

    protected Form getForm(String formDefId) {
        Form tempForm = null;
        FormDefinitionDao formDefinitionDao = (FormDefinitionDao)AppUtil.getApplicationContext().getBean("formDefinitionDao");
        if (formDefId != null) {
            AppDefinition appDef = AppUtil.getCurrentAppDefinition();
            FormDefinition formDef = formDefinitionDao.loadById(formDefId, appDef);
            if (formDef != null) {
                String formJson = formDef.getJson();
                if (formJson != null) {
                    tempForm = (Form)this.getFormService().createElementFromJson(formJson, true);
                }
            }
        }
        return tempForm;
    }

    protected FormService getFormService() {
        if (this.formService == null) {
            this.formService = (FormService)AppUtil.getApplicationContext().getBean("formService");
        }

        return this.formService;
    }


    @Override
    public FormRowSet store(Element element, FormRowSet rows, FormData formData) {
        if (rows == null || rows.isEmpty()) {
            return rows;
        }
        LogUtil.info("Started", "BusinessCase change store");

        FormRow originalRow = rows.get(0);
        LogUtil.info("originalRow", String.valueOf(originalRow));
        normalStoring(element, rows, formData);
        generateWordDocument(rows);

        String uuid = UuidGenerator.getInstance().getUuid();
        storeToOtherFormDataTable(element, rows, formData, uuid);
        return rows;
    }

    public void normalStoring(Element element, FormRowSet rows, FormData formData) {
        PluginManager pluginManager = (PluginManager) AppUtil.getApplicationContext().getBean("pluginManager");
        FormStoreBinder binder = (FormStoreBinder) pluginManager.getPlugin("org.joget.apps.form.lib.WorkflowFormBinder");
        binder.store(element, rows, formData);
    }

    public void storeToOtherFormDataTable(Element element, FormRowSet rows, FormData formData, String id) {
        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
        String formId = "epms_projectVersion"; // the table of database is configured in the form with this formId
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();

        appService.storeFormData(appDef.getId(), appDef.getVersion().toString(), formId, rows, id);
    }

    public void generateWordDocument(FormRowSet rows) {
        String fileHashVar = this.getPropertyString("templatePath");
        String documentPath = this.getPropertyString("documentPath");
        String templateFilePath = AppUtil.processHashVariable(fileHashVar, null, null, null);
        Path filePath = Paths.get(templateFilePath, new String[0]);
        String fileName = filePath.getFileName().toString();
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        File file = AppResourceUtil.getFile(appDef.getAppId(), String.valueOf(appDef.getVersion()), fileName);
        String tableName = "epms_document";
        UuidGenerator uuid = UuidGenerator.getInstance();
        String primaryKey = uuid.getUuid();
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyMMdd");
        String formattedDate = currentDate.format(formatter);
        FormRow jsonObject=rows.get(0);
        LogUtil.info("jsonObject",jsonObject.toString());
        String id=jsonObject.getProperty("id");
        LogUtil.info("project_id",id);
        FileInputStream fileInputStream;
        XWPFDocument apachDoc;
        try {
            fileInputStream = new FileInputStream(file);
            apachDoc = new XWPFDocument(fileInputStream);
            parseWordDocument(apachDoc,jsonObject);
            String fileNameGen = formattedDate + "_" + jsonObject.getProperty("project_code") + "_Business_Case_V" + jsonObject.getProperty("businesscase_major_version") + "." + jsonObject.getProperty("businesscase_minor_version") + ".docx";
            fileNameGen = fileNameGen.replaceAll(" ", "_").replace("/","_");
            LogUtil.info("fileNameGen", fileNameGen);

            File tempOutputFile = new File(documentPath + fileNameGen);

            FileOutputStream out = new FileOutputStream(tempOutputFile);
            apachDoc.write(out);
            FileUtil.storeFile(tempOutputFile, tableName, primaryKey);

            storeDocumentHistory(jsonObject,fileNameGen,primaryKey,id);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void parseWordDocument(XWPFDocument apachDoc, FormRow jsonObject)  {
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyMMdd");
        DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        String formattedDate = currentDate.format(formatter);
        String formatDate = currentDate.format(dateFormat);
        Map<String, Object> map = jsonObject.getCustomProperties();
        Set<String> keys = map.keySet();
        try {
            for (IBodyElement bodyElement : apachDoc.getBodyElements()) {
                if (bodyElement instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                    for (XWPFRun r : paragraph.getRuns()) {
                        String text = r.getText(0);
                        if (text != null && !text.isEmpty()) {
                            LogUtil.info("text", text);
                            for (String key : keys) {
                                if (text != null && !text.isEmpty() && text.contains("project_name_doc")) {
                                    text = text.replace("project_name_doc", jsonObject.getProperty("project_name"));
                                    r.setText(text, 0);
                                }
                                if (text.equals("YYMMDD")) {
                                    text = text.replace(text, formattedDate);
                                    r.setText(text, 0);
                                }
                                if (text.equals("V00")) {
                                    text = text.replace(text, "V"+jsonObject.getProperty("businesscase_major_version") + "." + jsonObject.getProperty("businesscase_minor_version"));
                                    r.setText(text, 0);
                                }
                                if (text.contains("PMO")) {
                                    LogUtil.info("PMOName",getPropertyString("PMOName"));
                                    text = text.replace("PMO", getPropertyString("PMOName"));
                                    r.setText(text, 0);
                                }
                            }
                        }
                    }
                } else if (bodyElement instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) bodyElement;
                    for (XWPFTableRow tableRow : table.getRows()) {
                        for (XWPFTableCell cell : tableRow.getTableCells()) {
                            for (XWPFParagraph p : cell.getParagraphs()) {
                                for (XWPFRun r : p.getRuns()) {
                                    String text = r.getText(0);
                                    if (text != null && !text.isEmpty()) {
                                        LogUtil.info("text", text);
                                        for (String key : keys) {
                                            if (text.contains("charter_date")) {
                                                text = text.replace("charter_date",formatDate);
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("doc_description")) {
                                                text = text.replace("doc_description", jsonObject.getProperty("project_description"));
                                                r.setText(text, 0);
                                            }
                                            if ((text.contains("project_scope") || text.contains("potential_benefits")) && jsonObject.getProperty(key).contains("<")) {
                                                setText(apachDoc, jsonObject, text);
                                                text = text.replace(text, " ");
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Project_manager")) {
                                                text = text.replace("Project_manager", getFullName(jsonObject.getProperty("project_manager")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Project_sponsor")) {
                                                text = text.replace("Project_sponsor", getFullName(jsonObject.getProperty("project_sponser")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Vendor_name")) {
                                                text = text.replace("Vendor_name", jsonObject.getProperty("project_vendor")== null ? "" : jsonObject.getProperty("project_vendor"));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Co_Ordinator")) {
                                                text = text.replace("Co_Ordinator", getFullName(jsonObject.getProperty("project_co_ordinator"))== null ? "" : getFullName(jsonObject.getProperty("project_co_ordinator")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Point_of_conduct")){
                                                text = text.replace("Point_of_conduct", getFullName(jsonObject.getProperty("businesspoint_contact"))== null ? "" : getFullName(jsonObject.getProperty("businesspoint_contact")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("System_integrator")) {
                                                text = text.replace("System_integrator", getFullName(jsonObject.getProperty("technical_integrator"))== null ? "" : getFullName(jsonObject.getProperty("technical_integrator")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("information_governance")) {
                                                text = text.replace("information_governance", getFullName(jsonObject.getProperty("information_governance"))== null ? "" : getFullName(jsonObject.getProperty("information_governance")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Scope_reviewer")) {
                                                text = text.replace("Scope_reviewer", getFullName(jsonObject.getProperty("scope_reviewer"))== null ? "" : getFullName(jsonObject.getProperty("scope_reviewer")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Plan_reviewer")) {
                                                text = text.replace("Plan_reviewer", getFullName(jsonObject.getProperty("plan_reviewer"))== null ? "" : getFullName(jsonObject.getProperty("plan_reviewer")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Cr_reviewer")) {
                                                text = text.replace("Cr_reviewer", getFullName(jsonObject.getProperty("change_request_reviewer"))== null ? "" : getFullName(jsonObject.getProperty("change_request_reviewer")));
                                                r.setText(text, 0);
                                            }
                                            if (text.equals("Uat_reviewer")) {
                                                text = text.replace("Uat_reviewer", getFullName(jsonObject.getProperty("uat_reviewer")) == null ? "" : getFullName(jsonObject.getProperty("uat_reviewer")));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("startDate")) {
                                                String startDate=jsonObject.getProperty("estimated_start_date");
                                                LocalDate givenDate = LocalDate.parse(startDate);
                                                String projectStartDate = givenDate.format(dateFormat);
                                                text = text.replace("startDate", projectStartDate);
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("endDate")) {
                                                String endDate=jsonObject.getProperty("estimated_end_date");
                                                LocalDate givenDate = LocalDate.parse(endDate);
                                                String projectEndDate = givenDate.format(dateFormat);
                                                text = text.replace("endDate", projectEndDate);
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Project_budget")) {
                                                DecimalFormat df=new DecimalFormat("#,##0.00");
                                                double number=Double.parseDouble(jsonObject.getProperty("estimated_budget"));
                                                String formatedNumber=df.format(number);
                                                text = text.replace("Project_budget", formatedNumber);
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("project_date")) {
                                                text = text.replace("project_date",getPropertyString("projectManagerDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("unitChief")) {
                                                text = text.replace("unitChief", getPropertyString("unitChiefName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("chief_date")) {
                                                text = text.replace("chief_date", getPropertyString("unitChiefDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("ceoName")) {
                                                text = text.replace("ceoName", getPropertyString("strategyName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("ceo_date")) {
                                                text = text.replace("ceo_date", getPropertyString("strategyDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("sponsorName")) {
                                                text = text.replace("sponsorName", getPropertyString("sponsorName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("sponsor_date")) {
                                                text = text.replace("sponsor_date", getPropertyString("sponsorDate"));
                                                r.setText(text, 0);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            String keyDefId= this.getPropertyString("keyRiskDefId");
            LogUtil.info("keyDefId",keyDefId);
            //Key Risks
            setKeyRisks(jsonObject.getProperty("id"),apachDoc,keyDefId);
            //File History
            addFileHistory(jsonObject.getProperty("id"),apachDoc);

        }catch (Exception ex){
            LogUtil.error("Message",ex,"DocumentBinder Error");
        }
    }

    public void setText(XWPFDocument apachDoc, FormRow jsonObject,String key) throws Exception {
        MyXWPFHtmlDocument htmlSet = createHtmlDoc(apachDoc, key);
        htmlSet.setHtml(htmlSet.getHtml().replace("<body></body>",jsonObject.getProperty(key)));
        replaceIBodyElementWithAltChunk(apachDoc, key, htmlSet);
    }

    public MyXWPFHtmlDocument createHtmlDoc(XWPFDocument document, String id) throws Exception {
        OPCPackage oPCPackage = document.getPackage();
        PackagePartName partName = PackagingURIHelper.createPartName("/word/" + id + ".html");
        PackagePart part = oPCPackage.createPart(partName, "text/html");
        MyXWPFHtmlDocument myXWPFHtmlDocument = new MyXWPFHtmlDocument(part, id);
        document.addRelation(myXWPFHtmlDocument.getId(), new XWPFHtmlRelation(), myXWPFHtmlDocument);
        return myXWPFHtmlDocument;
    }

    public void replaceIBodyElementWithAltChunk(XWPFDocument document, String textToFind, MyXWPFHtmlDocument myXWPFHtmlDocument) {
        List<IBodyElement> iBodyElements = new ArrayList<>();
        for (IBodyElement bodyElement : document.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                String text = paragraph.getText();
                if (text != null && text.contains(textToFind)) {
                    insertAltChunk(paragraph, myXWPFHtmlDocument);
                    iBodyElements.add(paragraph);
                }
            } else if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) bodyElement;
                for (XWPFTableRow tableRow : table.getRows()) {
                    for (XWPFTableCell tableCell : tableRow.getTableCells()) {
                        for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                            String text = paragraph.getText();
                            if (text != null && text.contains(textToFind)) {
                                insertAltChunk(paragraph, myXWPFHtmlDocument);
                                iBodyElements.add(paragraph);
                            }
                        }
                    }
                }
            }
        }
    }

    private void insertAltChunk(XWPFParagraph paragraph, MyXWPFHtmlDocument myXWPFHtmlDocument) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        cursor.toEndToken();
        while (cursor.toNextToken() != XmlCursor.TokenType.START);
        String uri = CTAltChunk.type.getName().getNamespaceURI();
        cursor.beginElement("altChunk", uri);
        cursor.toParent();
        CTAltChunk cTAltChunk = (CTAltChunk) cursor.getObject();
        cTAltChunk.setId(myXWPFHtmlDocument.getId());
    }

    private void storeDocumentHistory(FormRow jsonObject, String fileNameGen, String primaryKey,String projectId) throws SQLException {
        String currentDateTimeValue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(Calendar.getInstance().getTime());
        String userName = getPropertyString("userName");
        String docType="Business Case";
        String names = getPropertyString("userFullName");
        DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
        con = ds.getConnection();
        try {
            if (!con.isClosed()) {
                String insertSql = "INSERT INTO app_fd_epms_document(id,dateCreated,dateModified,createdBy,createdByName,modifiedBy,modifiedByName,c_project_id,c_upload_documents,c_document_type,c_project_document_version) VALUES (?,?,?,?,?,?,?,?,?,?,?)";
                PreparedStatement insertStmt = con.prepareStatement(insertSql);

                insertStmt.setString(1, primaryKey);
                insertStmt.setString(2, currentDateTimeValue);
                insertStmt.setString(3, currentDateTimeValue);
                insertStmt.setString(4, userName);
                insertStmt.setString(5, names);
                insertStmt.setString(6, userName);
                insertStmt.setString(7, names);
                insertStmt.setString(8, projectId);
                insertStmt.setString(9, fileNameGen);
                insertStmt.setString(10, docType);
                insertStmt.setString(11, "V" + jsonObject.getProperty("businesscase_major_version") + "." + jsonObject.getProperty("businesscase_minor_version"));

                //Execute SQL statement
                insertStmt.executeUpdate();

                //upDateAtachment
                String attachUpdateSQL = "UPDATE app_fd_epms_project SET c_doc_id = ?, c_doc_name = ? WHERE id = ?";
                PreparedStatement upstmt = con.prepareStatement(attachUpdateSQL);
                upstmt.setString(1, primaryKey);
                upstmt.setString(2, fileNameGen);
                upstmt.setString(3, projectId);
                upstmt.executeUpdate();
            }
        } catch (Exception ex) {
            LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
        }
    }

    public void setKeyRisks(String projectId, XWPFDocument apachDoc, String keyDefId){
        int targetTableIndex1 = 5;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex1);
        int rowIndex1 = 1;
        int rowIndex2 = 1;
        LogUtil.info("projectId", projectId);
        FormRowSet rows=new FormRowSet();
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        FormDefinitionDao formDefinitionDao = (FormDefinitionDao) FormUtil.getApplicationContext().getBean("formDefinitionDao");
        FormDefinition keyformDef = formDefinitionDao.loadById(keyDefId, appDef);

        String tableName = keyformDef.getTableName();

        if (tableName != null && projectId != null) {
            FormDataDao formDataDao = (FormDataDao) AppUtil.getApplicationContext().getBean("formDataDao");
            String propertyName = this.getFormPropertyName(tableName, "project_id");
            String condition = "";
            List<Object> paramsList = new ArrayList<>();

            if (propertyName != null && !propertyName.isEmpty()) {
                condition += " WHERE " + propertyName + " = ?";
                paramsList.add(projectId);
            }

            Object[] paramsArray = paramsList.toArray();
            rows = formDataDao.find(keyDefId, tableName, condition, paramsArray, "dateCreated", false, (Integer) null, (Integer) null);
        }

        int riskSize=rows.size();
        LogUtil.info("Risk rows", String.valueOf(rows));
        for (int i = 0; i < riskSize; i++) {
            FormRow formRow = rows.get(i);
            LogUtil.info("formRow", String.valueOf(formRow));
            String riskDescription=rows.get(i).getProperty("risk_title") == null ? "" : rows.get(i).getProperty("risk_title");
            String riskLevel=rows.get(i).getProperty("risk_likelihood") == null ? "" : rows.get(i).getProperty("risk_likelihood");
            XWPFTableRow row = table.getRow(rowIndex1);
            if (row == null) {
                row = table.createRow();
            }
            XWPFTableCell cell = row.getCell(1);
            cell.removeParagraph(0);
            XWPFParagraph addparagraph = cell.addParagraph();
            XWPFRun run = addparagraph.createRun();
            run.setFontFamily("calibri");
            run.setFontSize(9);
            run.setText(riskDescription);
            rowIndex1++;
            XWPFTableRow row1 = table.getRow(rowIndex2);
            if (row1 == null) {
                row1 = table.createRow();
            }
            XWPFTableCell cell1 = row1.getCell(2);
            cell1.removeParagraph(0);
            XWPFParagraph addparagraph1 = cell1.addParagraph();
            XWPFRun run1 = addparagraph1.createRun();
            run1.setFontFamily("calibri");
            run1.setFontSize(9);
            run1.setText(riskLevel);
            rowIndex2++;
        }
    }

    public void addFileHistory(String projectId, XWPFDocument apachDoc) throws SQLException {
        int targetTableIndex3 = 9;
        XWPFTable targetTable1 = apachDoc.getTables().get(targetTableIndex3);
        int rowIndex9 = 0;
        XWPFTableRow row1 = targetTable1.getRow(rowIndex9);
        XWPFTableCell cell1 = row1.getCell(1);
        XWPFParagraph paragraph1 = cell1.addParagraph();
        XWPFRun run1 = paragraph1.createRun();
        DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
        con = ds.getConnection();
        try {
            if (con != null && !con.isClosed()) {
                stmt = con.prepareStatement("select * from app_fd_epms_document WHERE c_project_id=? and c_document_type != 'Business Case' ");
                stmt.setObject(1, projectId);
                rs = stmt.executeQuery();
                int rowNum = 1;
                while (rs.next()) {
                    String attachedDocument = rs.getString("c_upload_documents") == null ? "" : rs.getString("c_upload_documents");
                    if (attachedDocument !=null && !attachedDocument.isEmpty()) {
                        run1.setFontFamily("calibri");
                        run1.setFontSize(9);
                        run1.setText(rowNum + ". " + attachedDocument);
                        String trimmedText = run1.getText(0).trim();
                        run1.setText(trimmedText, 0);
                        run1.addBreak();
                        rowNum++;
                    }
                }
            }
        }catch (Exception ex) {
            LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
        }
    }

    public String getFullName(String userIdData){
        String firstName;
        String lastName;
        String userFullName = null;
        String fullName;
        PreparedStatement stmt;
        ResultSet rs;
        Connection con;
        Collection userName=new ArrayList();
        String[] userIds = userIdData.split(";");
        for (String userId : userIds) {
            LogUtil.info("userId",userId);
            try {
                DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
                con = ds.getConnection();
                if (!con.isClosed()) {
                    stmt = con.prepareStatement("select firstName,lastName from dir_user WHERE id=? ");
                    stmt.setObject(1, userId);
                    rs = stmt.executeQuery();
                    while (rs.next()) {
                        firstName=rs.getString("firstName") == null ? "" : rs.getString("firstName");
                        lastName=rs.getString("lastName") == null ? "" : rs.getString("lastName");
                        userFullName=firstName+" "+lastName;
                    }
                }
            }catch (Exception ex) {
                LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
            }
            LogUtil.info("userFullName",userFullName);
            userName.add(userFullName);
        }
        fullName=userName.toString().replace("[","").replace("]","").replace(",","");
        LogUtil.info("fullName",fullName);
        return fullName;
    }
}
