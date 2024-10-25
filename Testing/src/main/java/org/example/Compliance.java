package org.example;

import org.joget.apps.app.service.AppUtil;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.UuidGenerator;

import javax.sql.DataSource;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Calendar;

public class Compliance {

    Connection con = null;
    PreparedStatement stmt = null;
    ResultSet rs;

    public void getComplianceDetails() {
        getClearanceLetterIssued();
        getConceptPaper();
        getBusinessPlan();
        getBusinessPlanRequestInfo();
        getCreditAward();
        getCreditAwardRequestInfo();
    }

    public void getCreditAwardRequestInfo() {
        String key="CAR";
        Integer days=getTotalDays(key);
        LocalDate currentDate = LocalDate.now();
        LocalDate newDate = currentDate.minusDays(days);
        LogUtil.info("newDate",newDate.toString());
        String reason="Credit Award Pending";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, dateCreated, GETDATE()) AS delay_in_days,* from app_fd_dsia_credit_award where c_ca_status = ? and c_isReqInfo = 1 and dateCreated < ?");
                stmt.setObject(1, "Awaiting RM Review");
                stmt.setObject(2, newDate);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_ts_id") == null ? "" : rs.getString("c_ts_id");
                    String dcName = rs.getString("c_dc_user_name") == null ? "" : rs.getString("c_dc_user_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"CA",reason);
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
    }

    public void getCreditAward() {
        String key="CAN";
        Integer days=getTotalDays(key);
        LocalDate currentDate = LocalDate.now();
        LocalDate newDate = currentDate.minusDays(days);
        LogUtil.info("newDate",newDate.toString());
        String reason="Credit Award Pending";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, dateCreated, GETDATE()) AS delay_in_days,* from app_fd_dsia_credit_award where c_ca_status = ? and (c_isReqInfo IS NULL OR c_isReqInfo = 0) and dateCreated < ?");
                stmt.setObject(1, "Awaiting RM Review");
                stmt.setObject(2, newDate);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_ts_id") == null ? "" : rs.getString("c_ts_id");
                    String dcName = rs.getString("c_dc_user_name") == null ? "" : rs.getString("c_dc_user_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"CA",reason);
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
    }

    public void getBusinessPlanRequestInfo() {
        String key="BPR";
        Integer days=getTotalDays(key);
        LocalDate currentDate = LocalDate.now();
        LocalDate newDate = currentDate.minusDays(days);
        LogUtil.info("newDate",newDate.toString());
        String reason="Business Plan Pending";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, dateCreated, GETDATE()) AS delay_in_days,* from app_fd_dsia_business_plan where c_cp_status = ? and c_isReqInfo = 1 and dateCreated < ?");
                stmt.setObject(1, "Awaiting PM Review");
                stmt.setObject(2, newDate);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_bp_id") == null ? "" : rs.getString("c_bp_id");
                    String dcName = rs.getString("c_dc_name") == null ? "" : rs.getString("c_dc_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"BP",reason);
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
    }

    public void getBusinessPlan() {
        String key="BPN";
        Integer days=getTotalDays(key);
        LocalDate currentDate = LocalDate.now();
        LocalDate newDate = currentDate.minusDays(days);
        LogUtil.info("newDate",newDate.toString());
        String reason="Business Plan Pending";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, dateCreated, GETDATE()) AS delay_in_days,* from app_fd_dsia_business_plan where c_cp_status = ? and (c_isReqInfo IS NULL OR c_isReqInfo = 0) and dateCreated < ?");
                stmt.setObject(1, "Awaiting PM Review");
                stmt.setObject(2, newDate);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_bp_id") == null ? "" : rs.getString("c_bp_id");
                    String dcName = rs.getString("c_dc_name") == null ? "" : rs.getString("c_dc_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"BP",reason);
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
    }

    public void getConceptPaper() {
        String reason="Concept Paper Pending";
        String key="CPN";
        Integer days=getTotalDays(key);
        LocalDate currentDate = LocalDate.now();
        LocalDate newDate = currentDate.minusDays(days);
        LogUtil.info("newDate",newDate.toString());
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, dateCreated, GETDATE()) AS delay_in_days,* from app_fd_dsia_concept_paper where c_cp_status = ? and dateCreated < ?");
                stmt.setObject(1, "Awaiting RM Review");
                stmt.setObject(2, newDate);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_cp_id") == null ? "" : rs.getString("c_cp_id");
                    String dcName = rs.getString("c_dc_name") == null ? "" : rs.getString("c_dc_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"CP",reason);
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
    }

    public Integer getTotalDays(String key) {
        Integer days = 0;
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select * from app_fd_dsia_sla_config where c_keys = ?");
                stmt.setObject(1, key);
                rs = stmt.executeQuery();
                while (rs.next()) {
                    days = rs.getInt("c_limit_no_of_days") == 0 ? 0 : rs.getInt("c_limit_no_of_days");
                    LogUtil.info("days", String.valueOf(days));
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
        return days;
    }

    public void getClearanceLetterIssued() {
        LocalDate currentDate = LocalDate.now();
        LogUtil.info("currentDate ",currentDate.toString());
        String reason="Clearance Letter Not Issued";
        try {
            DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
            con = ds.getConnection();
            if (!con.isClosed()) {
                stmt = con.prepareStatement("select DATEDIFF(DAY, c_sla_date, GETDATE()) AS delay_in_days, id, c_contract_reference_no, c_dc_name from app_fd_dsia_cl_letters where c_sla_date < ? and c_status != ?");
                LogUtil.info("Query ",stmt.toString());
                stmt.setObject(1, currentDate);
                stmt.setObject(2, "Clearance Letter Issued");
                rs = stmt.executeQuery();
                while (rs.next()) {
                    String delayDays = rs.getString("delay_in_days") == null ? "" : rs.getString("delay_in_days");
                    String id = rs.getString("id") == null ? "" : rs.getString("id");
                    String refNo = rs.getString("c_contract_reference_no") == null ? "" : rs.getString("c_contract_reference_no");
                    String dcName = rs.getString("c_dc_name") == null ? "" : rs.getString("c_dc_name");
                    storeComplianceDetails(id, delayDays, refNo, dcName,"CL",reason);
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
    }

    public void deleteComplianceDetails() throws SQLException {
        PreparedStatement stmt;
        DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
        con = ds.getConnection();
        try {
            String deleteQuery="DELETE FROM app_fd_dsia_compliance_data";
            stmt = con.prepareStatement(deleteQuery);
            stmt.executeUpdate();
        } catch (Exception ex) {
            LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
        } finally {
            try {
                if (con != null)
                    con.close();
            } catch (SQLException e) {
            }
        }
    }

    public void storeComplianceDetails(String id, String delayDays, String refNo, String dcName,String module,String reason) throws SQLException {
        String currentDateTimeValue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(Calendar.getInstance().getTime());
        String userName = "#currentUser.username#";
        String userFullName = "#currentUser.fullName#";
        String clUUID = UuidGenerator.getInstance().getUuid();
        PreparedStatement insertStmt;
        DataSource ds = (DataSource) AppUtil.getApplicationContext().getBean("setupDataSource");
        con = ds.getConnection();
        try {
            String insertSql = "INSERT INTO app_fd_dsia_compliance_data(id,dateCreated,dateModified,createdBy,createdByName,modifiedBy,modifiedByName,c_record_id,c_module,c_delay_in_days,c_reason_for_non_compliance,c_ref_no,c_dc_name) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)";
            insertStmt = con.prepareStatement(insertSql);
            insertStmt.setString(1, clUUID);
            insertStmt.setString(2, currentDateTimeValue);
            insertStmt.setString(3, currentDateTimeValue);
            insertStmt.setString(4, userName);
            insertStmt.setString(5, userFullName);
            insertStmt.setString(6, userName);
            insertStmt.setString(7, userFullName);
            insertStmt.setString(8, id);
            insertStmt.setString(9, module);
            insertStmt.setString(10, delayDays);
            insertStmt.setString(11, reason);
            insertStmt.setString(12, refNo);
            insertStmt.setString(13, dcName);

            insertStmt.executeUpdate();
        } catch (Exception ex) {
            LogUtil.error("Your App/Plugin Name", ex, "Error storing using jdbc");
        } finally {
            try {
                if (con != null) {
                    con.close();
                }
            } catch (SQLException ex) {
                LogUtil.error("inside sub cont insert - finally", ex, ex.getMessage());
            }
        }
    }

//    deleteComplianceDetails();
//    getComplianceDetails();


}
