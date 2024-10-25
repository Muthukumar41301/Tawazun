package org.example;

import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.commons.util.LogUtil;
import org.joget.plugin.base.ApplicationPlugin;
import org.joget.plugin.base.Plugin;
import org.joget.plugin.base.PluginManager;
import org.joget.plugin.property.model.PropertyEditable;
import org.joget.workflow.util.WorkflowUtil;

import javax.servlet.http.HttpServletRequest;
import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Map;

public class SendEmail {

    PreparedStatement stmt;
    ResultSet rs;
    Connection con = null;

    public void getEmails(String recordId) throws SQLException {
        String query="SELECT processId, activityName, assignees from (SELECT p.processId, sact.Name AS activityName," +
                "sass.ResourceId AS assignee FROM app_fd_dsia_cl_letters a JOIN wf_process_link p on p.parentProcessId = a.id " +
                "JOIN SHKActivities sact on p.processId = sact.ProcessId JOIN SHKActivityStates ssta ON ssta.oid = sact.State " +
                "INNER JOIN SHKAssignmentsTable sass ON sact.Id = sass.ActivityId WHERE " +
                "ssta.KeyValue = 'open.not_running.not_started' group by sact.Name, sass.ResourceId, p.processId) AS A " +
                "CROSS APPLY (SELECT assignee + ';' FROM (SELECT p.processId, sact.Name AS activityName, sass.ResourceId AS " +
                "assignee FROM app_fd_dsia_cl_letters a JOIN wf_process_link p on p.parentProcessId = a.id JOIN SHKActivities " +
                "sact on p.processId = sact.ProcessId JOIN SHKActivityStates ssta ON ssta.oid = sact.State INNER JOIN " +
                "SHKAssignmentsTable sass ON sact.Id = sass.ActivityId WHERE ssta.KeyValue = 'open.not_running.not_started' " +
                "group by sact.Name, sass.ResourceId, p.processId) AS B WHERE A.processId = B.processId AND " +
                "A.activityName = B.activityName FOR XML PATH('')) D (assignees) where " +
                "processId = (select processId from wf_process_link where originProcessId = ?) " +
                "GROUP BY processId, activityName, assignees";
        DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
        con = ds.getConnection();
        try {
            stmt = con.prepareStatement(query);
            stmt.setObject(1, recordId);
            rs = stmt.executeQuery();
            while (rs.next()) {
                String emailId = rs.getString("assignees") == null ? "" : rs.getString("assignees");
                sendEmail(emailId);
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
    }
    public void sendEmail(String emailId) {

        LogUtil.info("Trigger Mail Task", "Start");

        String taskOwnerFullName = null;
        String mail_subject = null;

        String creater_full_name = "#currentUser.fullname#";
        String bgValue = "#form.bank_guarantee.bg_value#";
        String selectDc = "#form.bank_guarantee.select_dc#";
        String bgValidityDate = "#form.bank_guarantee.bg_validity_date#";
        String saEffectiveDate = "#form.bank_guarantee.sa_effective_date#";
        String sender = "#variable.dc_process#";
        String sender1 = "#form.bank_guarantee.request_user#";
        String id = "#form.bank_guarantee.id#";
        String bgId = "#form.bank_guarantee.bg_id#";
        String mail_email_id = sender;
        String mail_content;
        mail_subject = "New Bank Guarantee Request against " + bgId;

        LogUtil.info("Spacing Check", "Spacing End");
        mail_content = "<table cellspacing=\"0\" cellpadding=\"0pt\" style=\"width:468pt;table-layout:fixed;border-collapse:collapse;\"><tr align=\"left\" valign=\"top\"><td valign=\"middle\" style=\"width:468pt;height:18pt;\"><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';font-size:11pt;\">Dear #user.{form.bank_guarantee.select_dc}.fullName#</span><span style=\"font-family:'Calibri';\"></span><span style=\"font-family:'Calibri';font-size:11pt;\">, <br/></span></p><br/><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';\"></span></p><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';\">\t A new <b>#form.bank_guarantee.bg_id#</b> has been initiated by <b> #user.{form.bank_guarantee.request_user}.fullName#</b>.<br/><br/>Please find the below Bank Guarantee Details for your reference.</span></p><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';font-weight:bold;\">BG Value: </span><span style=\"font-family:'Calibri';\">" + bgValue + "</p><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';font-weight:bold;\">DC: </span><span style=\"font-family:'Calibri';\">" + selectDc + "</span></p><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';font-weight:bold;\">BG Validity Date: </span><span style=\"font-family:'Calibri';\">" + bgValidityDate + "</span></p><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';font-weight:bold;\">SA Effective Date: </span><span style=\"font-family:'Calibri';\">" + saEffectiveDate + "</span></p><br/><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';\">Please click on the link below to verify the bank guarantee.</span></p><br/><p style=\"margin-top:0pt;margin-bottom:0pt;\"><a  href=\"#envVariable.baseURL##envVariable.appURL#/_/bank_guarantee_detail_process?processId=#assignment.processId#\" style=\"font-family:'Calibri';color:#0070C0;text-decoration:underline ;\">Click Here</a></p><br/><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';\">#i18n.epEmailSignature#</span></p><br/><p style=\"margin-top:0pt;margin-bottom:0pt;\"><span style=\"font-family:'Calibri';\">***This is a System Generated Notification Message. Do-not-reply to this message.</span></p></td></tr></table>";

        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
        AppDefinition appDef = appService.getAppDefinition("tawazunEconomicProgram", null);
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        PluginManager pluginManager = (PluginManager) AppUtil.getApplicationContext().getBean("pluginManager");
        Plugin plugin = pluginManager.getPlugin("org.joget.apps.app.lib.EmailTool");

        Map propertiesMap = AppPluginUtil.getDefaultProperties(plugin, "", appDef, null);

        propertiesMap.put("pluginManager", pluginManager);
        propertiesMap.put("appDef", appDef);
        propertiesMap.put("request", request);
        propertiesMap.put("toSpecific", mail_email_id);
        propertiesMap.put("subject", mail_subject);
        propertiesMap.put("message", mail_content);
        propertiesMap.put("isHtml", "true");
        // propertiesMap.put("content-type", "text/html; charset=utf-8");

        ApplicationPlugin emailTool = (ApplicationPlugin) plugin;
        //set properties and execute the tool
        ((PropertyEditable) emailTool).setProperties(propertiesMap);
        emailTool.execute(propertiesMap);

        LogUtil.info("Trigger Mail Task", "End");
    }
}
