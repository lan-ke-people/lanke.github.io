package com.qax.situation.asset.application.service.impl.excel.event;

import com.qax.situation.asset.application.dto.request.DataPreCheckQuery;
import com.qax.situation.asset.infra.persistence.db.entity.KeyAssetExport;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author L-wangxinzhuo
 * @version 1.0
 * @description:
 * @date 2026/1/27 16:43
 */
@Getter
@AllArgsConstructor
public class DataExportEvent {
    private final KeyAssetExport keyAssetExport;
    private final DataPreCheckQuery dataPreCheckQuery;
}
