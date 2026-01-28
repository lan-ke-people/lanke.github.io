package com.qax.situation.asset.application.service.impl.excel.builder;

import com.qax.situation.asset.application.dto.excel.export.*;

import java.util.ArrayList;
import java.util.List;

/**
 * @author L-wangxinzhuo
 * @version 1.0
 * @description:
 * @date 2026/1/27 17:38
 */
public class SystemExportDataBuilder {

    /**
     * 构建复杂的Excel数据
     */
    public static List<List<Object>> buildComplexData(List<UnitGroupDto> unitGroups) {
        List<List<Object>> allData = new ArrayList<>();

        // 1. 添加标题行
        List<Object> titleRow = new ArrayList<>();
        titleRow.add("重点保护对象清单");
        // 填充34个空单元格
        for (int i = 1; i < 35; i++) {
            titleRow.add("");
        }
        allData.add(titleRow);

        // 2. 遍历每个单位组
        for (UnitGroupDto group : unitGroups) {
            // 2.1 单位信息行
            List<Object> unitRow = buildUnitInfoRow(group.getUnitInfo());
            allData.add(unitRow);

            // 2.2 责任处室行
            List<Object> deptRow = buildDepartmentRow(group.getDepartment());
            allData.add(deptRow);

            // 2.3 工作人员行
            List<Object> staffRow = buildStaffRow(group.getStaff());
            allData.add(staffRow);


            // 2.4 表头行
            List<Object> headerRow = buildSystemHeaderRow();
            allData.add(headerRow);

            // 2.5 系统数据行
            int serial = 1;
            for (SystemInfoDto system : group.getSystemList()) {
                system.setSerialNumber(serial++);
                List<Object> systemRow = buildSystemDataRow(system);
                allData.add(systemRow);
            }

            // 2.6 添加空行分隔（如果不是最后一组）
            if (unitGroups.indexOf(group) < unitGroups.size() - 1) {
                allData.add(new ArrayList<>());
            }
        }

        return allData;
    }

    private static List<Object> buildUnitInfoRow(UnitInfoDto unitInfo) {
        List<Object> row = new ArrayList<>();
        row.add("单位名称");
        row.add("");
        row.add(unitInfo.getUnitName());
        for (int i = 3; i < 7; i++) row.add("");
        row.add("是否具有行业主管监管职责");
        row.add("");
        row.add(unitInfo.getHasSupervisionDuty());
        for (int i = 10; i < 13; i++) row.add("");
        row.add("第一责任人姓名");
        row.add("");
        row.add(unitInfo.getFirstResponsiblePerson());
        for (int i = 16; i < 17; i++) row.add("");
        row.add("职务");
        row.add(unitInfo.getFirstResponsiblePosition());
        for (int i = 19; i < 25; i++) row.add("");
        row.add("直接责任人姓名");
        for (int i = 26; i < 29; i++) row.add("");
        row.add(unitInfo.getDirectResponsiblePerson());
        row.add("");
        row.add("职务");
        row.add(unitInfo.getDirectResponsiblePosition());
        row.add("");
        row.add("");
        return row;
    }

    private static List<Object> buildDepartmentRow(DepartmentDto dept) {
        List<Object> row = new ArrayList<>();
        row.add("责任处室名称");
        row.add("");
        row.add(dept.getDepartmentName());
        for (int i = 3; i < 7; i++) row.add("");
        row.add("专职从事网络安全工作人员数量");
        row.add("");
        row.add(dept.getSecurityStaffCount());
        for (int i = 10; i < 13; i++) row.add("");
        row.add("处室负责人姓名");
        row.add("");
        row.add(dept.getDepartmentHeadName());
        for (int i = 16; i < 17; i++) row.add("");
        row.add("职务");
        row.add(dept.getDepartmentHeadPosition());
        for (int i = 19; i < 25; i++) row.add("");
        row.add("办公电话");
        for (int i = 26; i < 29; i++) row.add("");
        row.add(dept.getOfficePhone());
        row.add("");
        row.add("手机");
        row.add(dept.getMobilePhone());
        row.add("");
        row.add("");
        return row;
    }

    private static List<Object> buildStaffRow(StaffDto staff) {
        List<Object> row = new ArrayList<>();
        for (int i = 0; i < 13; i++) row.add("");
        row.add("工作人员");
        row.add("");
        row.add(staff.getStaffName());
        for (int i = 16; i < 17; i++) row.add("");
        row.add("职务");
        row.add(staff.getStaffPosition());
        for (int i = 19; i < 25; i++) row.add("");
        row.add("办公电话");
        for (int i = 26; i < 29; i++) row.add("");
        row.add(staff.getStaffOfficePhone());
        row.add("");
        row.add("手机");
        row.add(staff.getStaffMobilePhone());
        row.add("");
        row.add("");
        return row;
    }

    private static List<Object> buildSystemHeaderRow() {
        List<Object> row = new ArrayList<>();
        row.add("序号");
        row.add("网络应用系统名称");
        row.add("一级隶属单位");
        row.add("二级隶属单位");
        row.add("系统负责人姓名");
        row.add("系统负责人联系电话");
        row.add("系统负责人邮箱");
        row.add("上线时间");
        row.add("行业类型");
        row.add("系统类型");
        row.add("ICP备案号");
        row.add("等保级别");
        row.add("域名");
        row.add("系统使用URL");
        row.add("IP地址");
        row.add("端口");
        row.add("是否上云");
        row.add("云服务商名称");
        row.add("是否连接互联网");
        row.add("是否面向公众服务");
        row.add("服务对象");
        row.add("用户规模");
        row.add("运维单位名称");
        row.add("工作责任划分");
        row.add("运维单位联系人信息");
        row.add("各类数据规模条数");
        row.add("各类数据规模大小(GB)");
        row.add("各类数据存储位置");
        row.add("是否为关键信息基础设施");
        row.add("是否为党政机关门户网站");
        row.add("是否为大型网络平台");
        row.add("覆盖30%以上人口");
        row.add("覆盖10万人以上");
        row.add("存储超过100万人敏感信息");
        row.add("存储超过100万条基础数据");
        return row;
    }

    private static List<Object> buildSystemDataRow(SystemInfoDto system) {
        List<Object> row = new ArrayList<>();
        row.add(system.getSerialNumber());
        row.add(system.getSystemName());
        row.add(system.getFirstLevelUnit());
        row.add(system.getSecondLevelUnit());
        row.add(system.getSystemResponsiblePerson());
        row.add(system.getSystemResponsiblePhone());
        row.add(system.getSystemResponsibleEmail());
        row.add(system.getOnlineTime());
        row.add(system.getIndustryType());
        row.add(system.getSystemType());
        row.add(system.getIcpRecordNumber());
        row.add(system.getSecurityLevel());
        row.add(system.getDomain());
        row.add(system.getSystemUrl());
        row.add(system.getIpAddress());
        row.add(system.getPort());
        row.add(system.getIsOnCloud());
        row.add(system.getCloudProvider());
        row.add(system.getIsConnectedToInternet());
        row.add(system.getIsPublicService());
        row.add(system.getServiceTarget());
        row.add(system.getUserScale());
        row.add(system.getMaintenanceUnit());
        row.add(system.getResponsibilityDivision());
        row.add(system.getMaintenanceContact());
        row.add(system.getDataRecordCount());
        row.add(system.getDataSizeGB());
        row.add(system.getDataStorageLocation());
        row.add(system.getIsCriticalInfrastructure());
        row.add(system.getIsGovernmentWebsite());
        row.add(system.getIsLargePlatform());
        row.add(system.getCoverOver30Percent());
        row.add(system.getCoverOver100k());
        row.add(system.getStoreOver1mSensitiveInfo());
        row.add(system.getStoreOver1mBasicData());
        return row;
    }
}