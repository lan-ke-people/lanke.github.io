package com.qax.situation.asset.application.service.impl.excel.strategy;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import com.qax.situation.asset.application.dto.excel.export.UnitGroupDto;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * @description: 复杂表头合并策略 - 精确格式匹配
 */
public class ComplexHeaderMergeStrategy extends AbstractMergeStrategy {

    // 需要合并的单元格范围
    private final Set<CellRangeAddress> mergeRegions;
    private boolean hasMerged = false; // 添加标志位，确保只合并一次
    private static final ThreadLocal<Set<String>> mergedSheets = new ThreadLocal<>(); // 线程安全的已合并sheet记录

    // 常量定义
    private static final int TOTAL_COLUMNS = 34; // AI列对应索引34

    public ComplexHeaderMergeStrategy(List<UnitGroupDto> unitGroups) {
        this.mergeRegions = new LinkedHashSet<>(calculateMergeRegions(unitGroups));
    }

    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
        // 确保每个sheet只合并一次
        String sheetKey = sheet.getSheetName();
        Set<String> alreadyMergedSheets = mergedSheets.get();
        if (alreadyMergedSheets == null) {
            alreadyMergedSheets = new HashSet<>();
            mergedSheets.set(alreadyMergedSheets);
        }

        if (alreadyMergedSheets.contains(sheetKey)) {
            return; // 该sheet已经合并过
        }

        if (mergeRegions != null && !mergeRegions.isEmpty() && !hasMerged) {
            try {
                // 获取所有已存在的合并区域
                List<CellRangeAddress> existingRegions = new ArrayList<>();
                int numMergedRegions = sheet.getNumMergedRegions();
                for (int i = 0; i < numMergedRegions; i++) {
                    existingRegions.add(sheet.getMergedRegion(i));
                }

                // 只添加不存在的合并区域
                for (CellRangeAddress region : mergeRegions) {
                    // 检查是否已存在相同的合并区域
                    boolean alreadyExists = false;
                    for (CellRangeAddress existing : existingRegions) {
                        if (regionsEqual(existing, region)) {
                            alreadyExists = true;
                            break;
                        }
                    }

                    if (!alreadyExists && isValidMergeRegion(region)) {
                        try {
                            sheet.addMergedRegion(region); // 使用安全版本
                        } catch (Exception e) {
                            // 忽略合并异常
                        }
                    }
                }

                alreadyMergedSheets.add(sheetKey);
                hasMerged = true;
            } catch (Exception e) {
                // 记录错误但继续执行
                e.printStackTrace();
            }
        }
    }

    /**
     * 检查两个合并区域是否相等
     */
    private boolean regionsEqual(CellRangeAddress region1, CellRangeAddress region2) {
        return region1.getFirstRow() == region2.getFirstRow() &&
                region1.getLastRow() == region2.getLastRow() &&
                region1.getFirstColumn() == region2.getFirstColumn() &&
                region1.getLastColumn() == region2.getLastColumn();
    }

    /**
     * 检查合并区域是否有效
     */
    private boolean isValidMergeRegion(CellRangeAddress region) {
        int firstRow = region.getFirstRow();
        int lastRow = region.getLastRow();
        int firstCol = region.getFirstColumn();
        int lastCol = region.getLastColumn();

        // 必须至少覆盖两个单元格
        if (firstRow == lastRow && firstCol == lastCol) {
            return false;
        }

        // 确保列索引在有效范围内
        if (firstCol < 0 || lastCol > TOTAL_COLUMNS || firstRow < 0) {
            return false;
        }

        return true;
    }

    /**
     * 计算合并区域 - 精确格式匹配（不考虑数据内容）
     */
    private List<CellRangeAddress> calculateMergeRegions(List<UnitGroupDto> unitGroups) {
        List<CellRangeAddress> regions = new ArrayList<>();
        int currentRow = 0;

        // 1. 合并标题行（第1行）
        regions.add(new CellRangeAddress(currentRow, currentRow, 0, TOTAL_COLUMNS));
        currentRow++;

        // 遍历每个单位组
        for (int groupIndex = 0; groupIndex < unitGroups.size(); groupIndex++) {
            UnitGroupDto group = unitGroups.get(groupIndex);

            // 2. 单位信息块 - 第一行（单位名称行）
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    0, 1)); // A-B列
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    2, 6)); // C-G列
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    7, 8));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    9, 12));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    13, 14));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    15, 16));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    17, 17));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    18, 24));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    25, 28));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    29, 30));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    31, 31));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    32, 34));
            currentRow++;

            // 3. 单位信息块 - 第二行（责任处室行）
            int nextCurrentRow = currentRow + 1;
            regions.add(new CellRangeAddress(currentRow, nextCurrentRow,
                    0, 1)); // A-B列
            regions.add(new CellRangeAddress(currentRow, nextCurrentRow,
                    2, 6)); // C-G列
            regions.add(new CellRangeAddress(currentRow, nextCurrentRow,
                    7, 8));
            regions.add(new CellRangeAddress(currentRow, nextCurrentRow,
                    9, 12));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    13, 14));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    15, 16));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    17, 17));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    18, 24));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    25, 28));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    29, 30));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    31, 31));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    32, 34));
            currentRow++;

            // 4. 单位信息块 - 第三行（工作人员行）
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    13, 14));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    15, 16));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    17, 17));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    18, 24));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    25, 28));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    29, 30));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    31, 31));
            regions.add(new CellRangeAddress(currentRow, currentRow,
                    32, 34));
            currentRow++;

            // 5. 系统表头行（不需要合并）
            // 注意：表头行不需要任何合并，保持原样
            currentRow++;

            // 6. 系统数据行
            int systemCount = group.getSystemList() != null ? group.getSystemList().size() : 0;
            currentRow += systemCount;

            // 7. 空行
            if (groupIndex < unitGroups.size() - 1) {
                currentRow++;
            }
        }

        return regions;
    }
}