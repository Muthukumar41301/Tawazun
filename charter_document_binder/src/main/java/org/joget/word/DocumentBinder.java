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
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.UuidGenerator;
import org.joget.plugin.base.PluginManager;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAltChunk;

import javax.sql.DataSource;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class DocumentBinder extends FormBinder implements FormStoreBinder, FormStoreElementBinder, FormStoreMultiRowElementBinder {

    private static final String MESSAGE_PATH = "messages/DocumentBinder";

    private FormService formService;

    XWPFDocument apachDoc;
    PreparedStatement stmt;
    ResultSet rs;
    Connection con;

    @Override
    public String getName() {
        return "Charter Document Binder";
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
        LogUtil.info("Started", "Charter change store");

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
        try {
            fileInputStream = new FileInputStream(file);
            apachDoc = new XWPFDocument(fileInputStream);
            parseWordDocument(apachDoc,jsonObject);
            String fileNameGen = formattedDate + "_" + jsonObject.getProperty("project_code") + "_Project_Charter_V" + jsonObject.getProperty("charter_major_version") + "." + jsonObject.getProperty("charter_minor_version") + ".docx";
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

    private void storeDocumentHistory(FormRow jsonObject, String fileNameGen, String primaryKey, String id) throws SQLException {
        String currentDateTimeValue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(Calendar.getInstance().getTime());
        String userName = getPropertyString("userName");
        String docType="Project Charter";
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
                insertStmt.setString(8, id);
                insertStmt.setString(9, fileNameGen);
                insertStmt.setString(10, docType);
                insertStmt.setString(11, "V" + jsonObject.getProperty("charter_major_version") + "." + jsonObject.getProperty("charter_minor_version"));

                //Execute SQL statement
                insertStmt.executeUpdate();

                //upDateAtachment
                String attachUpdateSQL = "UPDATE app_fd_epms_project SET c_doc_id = ?, c_doc_name = ? WHERE id = ?";
                PreparedStatement upstmt = con.prepareStatement(attachUpdateSQL);
                upstmt.setString(1, primaryKey);
                upstmt.setString(2, fileNameGen);
                upstmt.setString(3, id);
                upstmt.executeUpdate();
            }
        } catch (Exception ex) {
            LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
        }
    }

    private void parseWordDocument(XWPFDocument apachDoc, FormRow jsonObject) {
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        String requesting_unit=jsonObject.getProperty("requesting_unit");
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
                                if (text.contains("project_date")) {
                                    text = text.replace("project_date", formatDate);
                                    r.setText(text, 0);
                                }
                                if (text.contains("project_version")) {
                                    text = text.replace("project_version", "V" + jsonObject.getProperty("charter_major_version") + "." + jsonObject.getProperty("charter_minor_version"));
                                    r.setText(text, 0);
                                }
                                if (text.equals("project_summary") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc,jsonObject,text,"summary");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.equals("acceptance_criteria") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc, jsonObject, text,"criteria1");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.equals("business_objectives") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc, jsonObject, text,"objectives1");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.equals("business_measures") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc, jsonObject, text,"measures");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }if (text.equals("security_requirements") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc, jsonObject, text,"security");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.equals("project_scope") && jsonObject.getProperty(key).contains("<")) {
                                    LogUtil.info("project_scope",key);
                                    setText(apachDoc, jsonObject, text,"scope");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.equals("business_benefits") && jsonObject.getProperty(key).contains("<")) {
                                    setText(apachDoc, jsonObject, text,"benefits");
                                    text = text.replace(text, " ");
                                    r.setText(text, 0);
                                }
                                if (text.contains("Project_description")) {
                                    text = text.replace("Project_description", jsonObject.getProperty("project_description"));
                                    r.setText(text, 0);
                                }
                                if (text.contains("projectName")) {
                                    text = text.replace("projectName",jsonObject.getProperty("project_name"));
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
                                            if (text.contains("P_name")) {
                                                text = text.replace("P_name",jsonObject.getProperty("project_name"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Reference_id")) {
                                                text = text.replace("Reference_id",jsonObject.getProperty("project_code"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Version_id")) {
                                                text = text.replace("Version_id","V"+jsonObject.getProperty("charter_major_version") + "." + jsonObject.getProperty("charter_minor_version"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Charter_status")) {
                                                text = text.replace("Charter_status","Draft");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("project_date")) {
                                                text = text.replace("project_date",formatDate);
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Pm_version")) {
                                                text = text.replace("Pm_version",jsonObject.getProperty("charter_major_version") + "." + jsonObject.getProperty("charter_minor_version"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Pm_date")) {
                                                text = text.replace("Pm_date",getPropertyString("projectManagerDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Pm_remarks")) {
                                                text = text.replace("Pm_remarks","Draft");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("UC_version")) {
                                                text = text.replace("UC_version","");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Uc_date")) {
                                                text = text.replace("Uc_date",getPropertyString("unitChiefDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Uc_name")) {
                                                text = text.replace("Uc_name",getPropertyString("unitChiefName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Uc_remarks")) {
                                                text = text.replace("Uc_remarks","");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("ST_version")) {
                                                text = text.replace("ST_version", "");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("St_date")) {
                                                text = text.replace("St_date",getPropertyString("strategyDate"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("St_name")) {
                                                text = text.replace("St_name", getPropertyString("strategyName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("St_remarks")) {
                                                text = text.replace("St_remarks", "");
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Request_unit")) {
                                                text = text.replace("Request_unit", getRequestUnitName(requesting_unit));
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
                                            if (text.equals("Implementation_partner")) {
                                                text = text.replace("Implementation_partner", getFullName(jsonObject.getProperty("information_governance"))== null ? "" : getFullName(jsonObject.getProperty("information_governance")));
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
                                            if (text.contains("Unit_Chief_Name")) {
                                                text = text.replace("Unit_Chief_Name",getPropertyString("unitChiefName"));
                                                r.setText(text, 0);
                                            }
                                            if (text.contains("Strategy_Name")) {
                                                text = text.replace("Strategy_Name",getPropertyString("strategyName"));
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

            LogUtil.info("list", "list load start");
            //Project Deliverables
            setProjectDeliverable(jsonObject);
            // Dependencies
            setDependencies(jsonObject.getProperty("id"));
            // Key Assumptions
            setKeyAssumptions(jsonObject);
            //Key Risks
            setKeyRisks(jsonObject.getProperty("id"));
            //Constraints
            setProjectConstraints(jsonObject);
            //Affected Products
            setAffectedProducts(jsonObject);
            //Unaffected Products
            setUnaffectedProducts(jsonObject);
            //Timeline and Reports
//            setTimelinesAndReports(jsonObject);

            LogUtil.info("list", "list load end");

        }catch (Exception ex){
            LogUtil.error("Message",ex,"DocumentBinder Error");
        }
    }

    public void setText(XWPFDocument apachDoc, FormRow jsonObject, String key, String id) throws Exception {
        MyXWPFHtmlDocument htmlSet = createHtmlDoc(apachDoc, id);
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

    public void setProjectDeliverable(FormRow originalRow){
        int rowIndex = 0;
        int targetTableIndex = 5;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int targetTableIndex1 = 6;
        XWPFTable table1 = apachDoc.getTables().get(targetTableIndex1);
        int targetTableIndex2 = 7;
        XWPFTable table2 = apachDoc.getTables().get(targetTableIndex2);
        try {
            JSONArray deliverable=new JSONArray(originalRow.getProperty("charter_highlevel_deliverables"));
            for (int i = 0; i < deliverable.length(); i++) {
                JSONObject response = deliverable.getJSONObject(i);
                String projectDeliverables = response.getString("charter_project_deliverables");
                String clientDeliverables = response.getString("charter_client_deliverables");
                String charterOutScope = response.getString("charter_out_scope");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(projectDeliverables);

                XWPFTableRow row1 = table1.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table1.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(0);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(clientDeliverables);

                XWPFTableRow row2 = table2.getRow(rowIndex);
                if (row2 == null) {
                    row2 = table2.createRow();
                }
                XWPFTableCell cell2 = row2.getCell(0);
                cell2.removeParagraph(0);
                XWPFParagraph addparagraph2 = cell2.addParagraph();
                XWPFRun run2 = addparagraph2.createRun();
                run2.setFontFamily("calibri");
                run2.setFontSize(11);
                run2.setText(charterOutScope);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }

    }

    public void setDependencies(String projectId){
        int targetTableIndex = 8;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        LogUtil.info("projectId", projectId);
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select * from app_fd_epms_dependency WHERE c_project_id=? ");
                stmt.setObject(1, projectId);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String ownerOfDependency = rs.getString("c_owner") == null ? "" : rs.getString("c_owner");
                    String deliverable = rs.getString("c_details") == null ? "" : rs.getString("c_details");
                    String requiredDate = rs.getString("c_required_date") == null ? "" : rs.getString("c_required_date");
                    String criticality = rs.getString("c_criticality") == null ? "" : rs.getString("c_criticality");
                    XWPFTableRow row = table.getRow(rowIndex);
                    if (row == null) {
                        row = table.createRow();
                    }
                    XWPFTableCell cell = row.getCell(0);
                    cell.removeParagraph(0);
                    XWPFParagraph addparagraph = cell.addParagraph();
                    XWPFRun run = addparagraph.createRun();
                    run.setFontFamily("calibri");
                    run.setFontSize(11);
                    run.setText(ownerOfDependency);
                    XWPFTableRow row1 = table.getRow(rowIndex);
                    if (row1 == null) {
                        row1 = table.createRow();
                    }
                    XWPFTableCell cell1 = row1.getCell(1);
                    cell1.removeParagraph(0);
                    XWPFParagraph addparagraph1 = cell1.addParagraph();
                    XWPFRun run1 = addparagraph1.createRun();
                    run1.setFontFamily("calibri");
                    run1.setFontSize(11);
                    run1.setText(deliverable);
                    XWPFTableRow row2 = table.getRow(rowIndex);
                    if (row2 == null) {
                        row2 = table.createRow();
                    }
                    XWPFTableCell cell2 = row2.getCell(2);
                    cell2.removeParagraph(0);
                    XWPFParagraph addparagraph2 = cell2.addParagraph();
                    XWPFRun run2 = addparagraph2.createRun();
                    run2.setFontFamily("calibri");
                    run2.setFontSize(11);
                    run2.setText(requiredDate);
                    XWPFTableRow row3 = table.getRow(rowIndex);
                    if (row3 == null) {
                        row3 = table.createRow();
                    }
                    XWPFTableCell cell3 = row3.getCell(3);
                    cell3.removeParagraph(0);
                    XWPFParagraph addparagraph3 = cell3.addParagraph();
                    XWPFRun run3 = addparagraph3.createRun();
                    run3.setFontFamily("calibri");
                    run3.setFontSize(11);
                    run3.setText(criticality);
                    rowIndex++;
                }
            }
        }catch (Exception ex){
            ex.printStackTrace();
        } finally {
            try {
                if (con != null)
                    con.close();
            } catch (SQLException e) {
            }
        }

    }

    public void setKeyAssumptions(FormRow originalRow){
        int targetTableIndex = 9;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        try {
            JSONArray assumptions=new JSONArray(originalRow.getProperty("assumptions"));
            for (int i = 0; i < assumptions.length(); i++) {
                JSONObject response = assumptions.getJSONObject(i);
                String charterRelates = response.getString("charter_relates");
                String assumptionDescription = response.getString("charter_assumption_description");
                String status = response.getString("charter_status");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(charterRelates);
                XWPFTableRow row1 = table.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(1);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(assumptionDescription);
                XWPFTableRow row2 = table.getRow(rowIndex);
                if (row2 == null) {
                    row2 = table.createRow();
                }
                XWPFTableCell cell2 = row2.getCell(2);
                cell2.removeParagraph(0);
                XWPFParagraph addparagraph2 = cell2.addParagraph();
                XWPFRun run2 = addparagraph2.createRun();
                run2.setFontFamily("calibri");
                run2.setFontSize(11);
                run2.setText(status);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }

    public void setKeyRisks(String projectId){
        int targetTableIndex = 10;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        LogUtil.info("projectId1", projectId);
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select * from app_fd_epms_project_risk WHERE c_project_id=? ");
                stmt.setObject(1, projectId);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String riskDescription=rs.getString("c_risk_title") == null ? "" : rs.getString("c_risk_title");
                    String riskLevel=rs.getString("c_risk_likelihood") == null ? "" : rs.getString("c_risk_likelihood");
                    String impactLevel=rs.getString("c_risk_impact") == null ? "" : rs.getString("c_risk_impact");
                    String mitigationPlan=rs.getString("c_mitigation_plan") == null ? "" : rs.getString("c_mitigation_plan");
                    String ContingencyPlan=rs.getString("c_contingency_plan") == null ? "" : rs.getString("c_contingency_plan");
                    XWPFTableRow row = table.getRow(rowIndex);
                    if (row == null) {
                        row = table.createRow();
                    }
                    XWPFTableCell cell = row.getCell(0);
                    cell.removeParagraph(0);
                    XWPFParagraph addparagraph = cell.addParagraph();
                    XWPFRun run = addparagraph.createRun();
                    run.setFontFamily("calibri");
                    run.setFontSize(11);
                    run.setText(riskDescription);
                    LogUtil.info("projectId1", projectId);
                    XWPFTableRow row1 = table.getRow(rowIndex);
                    if (row1 == null) {
                        row1 = table.createRow();
                    }
                    XWPFTableCell cell1 = row1.getCell(1);
                    cell1.removeParagraph(0);
                    XWPFParagraph addparagraph1 = cell1.addParagraph();
                    XWPFRun run1 = addparagraph1.createRun();
                    run1.setFontFamily("calibri");
                    run1.setFontSize(11);
                    run1.setText(riskLevel);
                    XWPFTableRow row2 = table.getRow(rowIndex);
                    if (row2 == null) {
                        row2 = table.createRow();
                    }
                    XWPFTableCell cell2 = row2.getCell(2);
                    cell2.removeParagraph(0);
                    XWPFParagraph addparagraph2 = cell2.addParagraph();
                    XWPFRun run2 = addparagraph2.createRun();
                    run2.setFontFamily("calibri");
                    run2.setFontSize(11);
                    LogUtil.info("projectId1", projectId);
                    run2.setText(impactLevel);
                    XWPFTableRow row3 = table.getRow(rowIndex);
                    if (row3 == null) {
                        row3 = table.createRow();
                    }
                    XWPFTableCell cell3 = row3.getCell(3);
                    cell3.removeParagraph(0);
                    XWPFParagraph addparagraph3 = cell3.addParagraph();
                    XWPFRun run3 = addparagraph3.createRun();
                    run3.setFontFamily("calibri");
                    run3.setFontSize(11);
                    LogUtil.info("projectId1", projectId);
                    run3.setText("Mitigation Plan:"+mitigationPlan);
                    run3.addBreak();
                    run3.setBold(true);
                    run3.setText("Contingency Plan:"+ContingencyPlan);
                    rowIndex++;
                }
            }
        }catch (SQLException ex){
            ex.printStackTrace();
        } finally {
            try {
                if (con != null)
                    con.close();
            } catch (SQLException e) {
            }
        }
    }

    public void setProjectConstraints(FormRow originalRow){
        int targetTableIndex = 11;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        try {
            JSONArray constraints=new JSONArray(originalRow.getProperty("charter_constraints"));
            for (int i = 0; i < constraints.length(); i++) {
                JSONObject response = constraints.getJSONObject(i);
                String constraintNumber = response.getString("charter_constraint_number");
                String constrainRelates = response.getString("charter_constraint_relates");
                String constrainDescription = response.getString("charter_constraint_description");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(constraintNumber);
                XWPFTableRow row1 = table.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(1);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(constrainRelates);
                XWPFTableRow row2 = table.getRow(rowIndex);
                if (row2 == null) {
                    row2 = table.createRow();
                }
                XWPFTableCell cell2 = row2.getCell(2);
                cell2.removeParagraph(0);
                XWPFParagraph addparagraph2 = cell2.addParagraph();
                XWPFRun run2 = addparagraph2.createRun();
                run2.setFontFamily("calibri");
                run2.setFontSize(11);
                run2.setText(constrainDescription);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }

    public void setAffectedProducts(FormRow originalRow){
        int targetTableIndex = 12;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        try {
            JSONArray affectedProducts=new JSONArray(originalRow.getProperty("project_charter_affected_products"));
            for (int i = 0; i < affectedProducts.length(); i++) {
                JSONObject response = affectedProducts.getJSONObject(i);
                String products = response.getString("charter_products");
                String productsAffected = response.getString("charter_products_affected");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(products);
                XWPFTableRow row1 = table.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(1);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(productsAffected);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }

    public void setUnaffectedProducts(FormRow originalRow){
        int targetTableIndex = 13;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        try {
            JSONArray unaffectedProducts=new JSONArray(originalRow.getProperty("project_charter_unaffected_products"));
            for (int i = 0; i < unaffectedProducts.length(); i++) {
                JSONObject response = unaffectedProducts.getJSONObject(i);
                String products = response.getString("charter_unaffected_products");
                String productsUnaffected = response.getString("charter_products_unaffected");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(products);
                XWPFTableRow row1 = table.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(1);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(productsUnaffected);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }

    public void setTimelinesAndReports(FormRow originalRow){
        int targetTableIndex = 14;
        XWPFTable table = apachDoc.getTables().get(targetTableIndex);
        int rowIndex = 1;
        try {
            JSONArray timelineReport=new JSONArray(originalRow.getProperty("project_charter_timeline_report"));
            for (int i = 0; i < timelineReport.length(); i++) {
                JSONObject response = timelineReport.getJSONObject(i);
                String events = response.getString("charter_events");
                String dateCommenced = response.getString("charter_date_commenced");
                String dateFinalized = response.getString("charter_date_finalized");
                String responsibility = response.getString("charter_responsibility");
                XWPFTableRow row = table.getRow(rowIndex);
                if (row == null) {
                    row = table.createRow();
                }
                XWPFTableCell cell = row.getCell(0);
                cell.removeParagraph(0);
                XWPFParagraph addparagraph = cell.addParagraph();
                XWPFRun run = addparagraph.createRun();
                run.setFontFamily("calibri");
                run.setFontSize(11);
                run.setText(events);
                XWPFTableRow row1 = table.getRow(rowIndex);
                if (row1 == null) {
                    row1 = table.createRow();
                }
                XWPFTableCell cell1 = row1.getCell(1);
                cell1.removeParagraph(0);
                XWPFParagraph addparagraph1 = cell1.addParagraph();
                XWPFRun run1 = addparagraph1.createRun();
                run1.setFontFamily("calibri");
                run1.setFontSize(11);
                run1.setText(dateCommenced);
                XWPFTableRow row2 = table.getRow(rowIndex);
                if (row2 == null) {
                    row2 = table.createRow();
                }
                XWPFTableCell cell2 = row2.getCell(2);
                cell2.removeParagraph(0);
                XWPFParagraph addparagraph2 = cell2.addParagraph();
                XWPFRun run2 = addparagraph2.createRun();
                run2.setFontFamily("calibri");
                run2.setFontSize(11);
                run2.setText(dateFinalized);
                XWPFTableRow row3 = table.getRow(rowIndex);
                if (row3 == null) {
                    row3 = table.createRow();
                }
                XWPFTableCell cell3 = row3.getCell(3);
                cell3.removeParagraph(0);
                XWPFParagraph addparagraph3 = cell3.addParagraph();
                XWPFRun run3 = addparagraph3.createRun();
                run3.setFontFamily("calibri");
                run3.setFontSize(11);
                run3.setText(responsibility);
                rowIndex++;
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }


    public String getRequestUnitName(String requesting_unit) {
        String unit_name = "";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            LogUtil.info("requesting_unit", requesting_unit);
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select * from app_fd_epms_unit_master WHERE id=? ");
                stmt.setObject(1, requesting_unit);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    unit_name=rs.getString("c_unit_name") == null ? "" : rs.getString("c_unit_name");
                    LogUtil.info("unit_name", unit_name);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            try {
                if (con != null)
                    con.close();
            } catch (SQLException e) {
            }
        }
        return unit_name;
    }


}
