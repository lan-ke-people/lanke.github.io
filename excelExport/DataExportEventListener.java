package com.qax.situation.asset.application.service.impl.excel.event;

import cn.hutool.core.date.DateUtil;
import com.qax.dayu.asset.sdk.model.PageResult;
import com.qax.dayu.asset.sdk.model.cond.OrganizationCond;
import com.qax.dayu.asset.sdk.model.dto.OrganizationRelDto;
import com.qax.dayu.asset.sdk.model.dto.SystemRelDto;
import com.qax.situation.asset.application.dto.excel.export.*;
import com.qax.situation.asset.application.dto.request.DataPreCheckQuery;
import com.qax.situation.asset.application.dto.response.FileUploadResDto;
import com.qax.situation.asset.application.service.impl.KeyAssetExportServiceImpl;
import com.qax.situation.asset.application.service.impl.excel.util.ComplexExcelExportUtil;
import com.qax.situation.asset.infra.external.HakkeroClient;
import com.qax.situation.asset.infra.persistence.db.entity.KeyAssetExport;
import lombok.extern.slf4j.Slf4j;
import org.springframework.context.event.EventListener;
import org.springframework.http.ResponseEntity;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author L-wangxinzhuo
 * @version 1.0
 * @description:
 * @date 2026/1/27 16:48
 */
@Component
@Slf4j
public class DataExportEventListener {

    @Resource
    private KeyAssetExportServiceImpl keyAssetExportService;

    @Resource
    private HakkeroClient hakkeroClient;

    @Async
    @EventListener
    public void handleDataExportEvent(DataExportEvent event) {
        KeyAssetExport keyAssetExport = event.getKeyAssetExport();
        DataPreCheckQuery dataPreCheckQuery = event.getDataPreCheckQuery();

        try {
            // 执行实际的数据导出处理
            String fileId = performDataExport(dataPreCheckQuery);

            // 更新导出记录的文件ID和状态
            keyAssetExport.setFileId(fileId);
            keyAssetExport.setState(2); // 状态：完成
            keyAssetExport.setUpdateTime(LocalDateTime.now());
            keyAssetExportService.updateById(keyAssetExport);

            log.info("数据导出完成，任务ID: {}, 文件ID: {}", keyAssetExport.getTaskId(), fileId);
        } catch (Exception e) {
            log.error("数据导出失败，任务ID: {}", keyAssetExport.getTaskId(), e);
            // 更新失败状态和错误日志
            keyAssetExport.setState(3); // 状态：失败
            keyAssetExport.setLog(e.getMessage());
            keyAssetExport.setUpdateTime(LocalDateTime.now());
            keyAssetExportService.updateById(keyAssetExport);
        }
    }

    private String performDataExport(DataPreCheckQuery dataPreCheckQuery) {
        // 实现具体的数据导出逻辑
        OrganizationCond organizationCond = keyAssetExportService.getOrganizationCond(dataPreCheckQuery);

        ResponseEntity<PageResult<OrganizationRelDto>> orgPageResultResponseEntity = keyAssetExportService.getOrgPageResultResponseEntity(organizationCond);
        ResponseEntity<PageResult<SystemRelDto>> sysPageResultResponseEntity = keyAssetExportService.getSysPageResultResponseEntity(organizationCond);

        String tempFilePath = null;
        try {
            // 1. 查询数据
            //List<UnitGroupDto> unitGroups = buildUnitGroupsFromDatabase(orgPageResultResponseEntity, sysPageResultResponseEntity);

            List<UnitGroupDto> unitGroups = createTestData();

            // 2. 导出Excel
            String fileName = "重点单位资产清单" + DateUtil.format(new Date(), "yyyyMMddHHmmss");

            // 3. 导出到临时文件
            tempFilePath = ComplexExcelExportUtil.exportComplexExcelToTempFile(fileName, unitGroups);

            // 4. 从临时文件创建MultipartFile
            MultipartFile multipartFile = ComplexExcelExportUtil.createMultipartFileFromTemp(tempFilePath);

           // FileUploadResDto fileUploadResDto = hakkeroClient.uploadFile(multipartFile, "asset");
            log.info("系统清单导出成功，共{}个单位", unitGroups.size());

            // 返回Hakkero服务返回的文件ID
//            return fileUploadResDto.getFileId();
            return "1";
        } catch (IOException e) {
            log.error("导出数据异常：", e);
            throw new RuntimeException("导出失败：" + e.getMessage());
        } finally {
            // 6. 清理临时文件
//            if (tempFilePath != null) {
//                deleteTempFile(tempFilePath);
//            }
        }
    }

    private List<UnitGroupDto> createTestData() {
        List<UnitGroupDto> groups = new ArrayList<>();

        extracted(groups, 1);
        extracted(groups, 3);
        extracted(groups, 2);

        return groups;
    }

    private static void extracted(List<UnitGroupDto> groups, int i) {
        // 创建第一组测试数据
        UnitGroupDto group1 = new UnitGroupDto();

        UnitInfoDto unitInfo1 = new UnitInfoDto();
        unitInfo1.setUnitName("测试单位1");
        unitInfo1.setHasSupervisionDuty("是");
        unitInfo1.setFirstResponsiblePerson("张三");
        unitInfo1.setFirstResponsiblePosition("局长");
        unitInfo1.setDirectResponsiblePerson("李四");
        unitInfo1.setDirectResponsiblePosition("处长");
        group1.setUnitInfo(unitInfo1);

        DepartmentDto dept1 = new DepartmentDto();
        dept1.setDepartmentName("信息中心");
        dept1.setSecurityStaffCount(5);
        dept1.setDepartmentHeadName("王五");
        dept1.setDepartmentHeadPosition("主任");
        dept1.setOfficePhone("010-12345678");
        dept1.setMobilePhone("13800138000");
        group1.setDepartment(dept1);

        // 添加工作人员
        StaffDto staff1 = new StaffDto();
        staff1.setStaffName("赵六");
        staff1.setStaffPosition("工程师");
        staff1.setStaffOfficePhone("010-87654321");
        staff1.setStaffMobilePhone("13900139000");
        group1.setStaff(staff1);

        // 添加系统信息
        List<SystemInfoDto> systemList1 = new ArrayList<>();
        for (int j = 1; j <= i; j++){
            SystemInfoDto system1 = new SystemInfoDto();
            system1.setSystemName("测试系统1");
            system1.setFirstLevelUnit("一级单位");
            system1.setSecondLevelUnit("二级单位");
            system1.setSystemResponsiblePerson("钱七");
            system1.setSystemResponsiblePhone("13800000001");
            system1.setSystemResponsibleEmail("test1@example.com");
            system1.setOnlineTime("2023-01-01");
            system1.setIndustryType("政府");
            system1.setSystemType("门户网站");
            system1.setIcpRecordNumber("京ICP备12345678号");
            system1.setSecurityLevel("三级");
            system1.setDomain("test1.gov.cn");
            system1.setSystemUrl("https://test1.gov.cn");
            system1.setIpAddress("192.168.1.1");
            system1.setPort("80");
            system1.setIsOnCloud("是");
            system1.setCloudProvider("阿里云");
            system1.setIsConnectedToInternet("是");
            system1.setIsPublicService("是");
            system1.setServiceTarget("公众");
            system1.setUserScale("10000人");
            system1.setMaintenanceUnit("运维公司");
            system1.setResponsibilityDivision("全权运维");
            system1.setMaintenanceContact("孙八, 13900000002");
            system1.setDataRecordCount("100万条");
            system1.setDataSizeGB("50GB");
            system1.setDataStorageLocation("云平台");
            system1.setIsCriticalInfrastructure("是");
            system1.setIsGovernmentWebsite("是");
            system1.setIsLargePlatform("否");
            system1.setCoverOver30Percent("否");
            system1.setCoverOver100k("是");
            system1.setStoreOver1mSensitiveInfo("否");
            system1.setStoreOver1mBasicData("是");
            systemList1.add(system1);
        }

        group1.setSystemList(systemList1);

        groups.add(group1);
    }

    /**
     * 清理临时文件
     */
    private void deleteTempFile(String filePath) {
        try {
            File file = new File(filePath);
            if (file.exists() && file.delete()) {
                log.debug("临时文件已删除：{}", filePath);
            }
        } catch (Exception e) {
            log.warn("删除临时文件失败：{}", filePath, e);
        }
    }

    /**
     * 构建单位组数据（从数据库查询并转换）
     */
    private List<UnitGroupDto> buildUnitGroupsFromDatabase(ResponseEntity<PageResult<OrganizationRelDto>> orgPageResultResponseEntity, ResponseEntity<PageResult<SystemRelDto>> sysPageResultResponseEntity) {
        List<UnitGroupDto> unitGroups = new ArrayList<>();
        if (orgPageResultResponseEntity.getBody() != null){
            List<OrganizationRelDto> orgData = orgPageResultResponseEntity.getBody().getItems();
            if (sysPageResultResponseEntity.getBody() != null){
                List<SystemRelDto> sysData = sysPageResultResponseEntity.getBody().getItems();
                Map<String, List<SystemRelDto>> groupedByOrgId = sysData.stream()
                        .filter(s -> s.getOrganization() != null
                                && s.getOrganization().getId() != null)
                        .collect(Collectors.groupingBy(
                                s -> s.getOrganization().getId()
                        ));

                for (OrganizationRelDto org : orgData){
                    UnitGroupDto group = new UnitGroupDto();

                    // 构建单位信息
                    UnitInfoDto unitInfo = new UnitInfoDto();
                    unitInfo.setUnitName(org.getName());
                    //unitInfo.setHasSupervisionDuty(org.getHasSupervision() ? "是" : "否");
                    // ... 设置其他字段
                    group.setUnitInfo(unitInfo);

                    // 构建责任处室信息
                    DepartmentDto dept = new DepartmentDto();
//                    dept.setDepartmentName(org.getDepartment().getName());
//                    dept.setSecurityStaffCount(org.getDepartment().getSecurityStaffCount());
                    // ... 设置其他字段
                    group.setDepartment(dept);

                    // 构建工作人员列表
//                    List<StaffDto> staffList = org.getStaffList().stream().map(staff -> {
//                        StaffDto dto = new StaffDto();
//                        dto.setStaffName(staff.getName());
//                        dto.setStaffPosition(staff.getPosition());
//                        // ... 设置其他字段
//                        return dto;
//                    }).collect(Collectors.toList());
//                    group.setStaffList(staffList);
                    group.setStaff(new StaffDto());

                    // 构建系统清单
                    List<SystemInfoDto> systemList = groupedByOrgId.get(org.getId())
                            .stream()
                            .map(this::convertToSystemInfoDto)
                            .collect(Collectors.toList());
                    group.setSystemList(systemList);

                    unitGroups.add(group);
                }
            }
        }

        return unitGroups;
    }

    private SystemInfoDto convertToSystemInfoDto(SystemRelDto entity) {
        SystemInfoDto dto = new SystemInfoDto();
        dto.setSystemName(entity.getName());
//        dto.setFirstLevelUnit(entity.getFirstLevelUnit());
//        dto.setSecondLevelUnit(entity.getSecondLevelUnit());
        // ... 设置其他字段
        return dto;
    }
}

