'''
file: IYATT_BatchMergeByColumns.py
description: 按列批量合并选中区域
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
from pytableenginesdk import utils

pluginInfo = {
    'name': 'BatchMergeByColumns',
    'author': 'IYATT',
    'description': '按列批量合并选中区域',
    'version': '0.0.2',
}

xlCenter = -4108

def _processMerge(appComHandle, targets, logger):
    totalMergedCols = 0
    for sheetName, address, mainRng in targets:
        logger.info(f'正在处理工作表 [{sheetName}] 的区域 [{address}]...')
        for subRng in utils.forEachArea(mainRng):
            for i in range(1, subRng.Columns.Count + 1):
                col = subRng.Columns.Item(i)
                col.Merge()
                col.HorizontalAlignment = xlCenter
                col.VerticalAlignment = xlCenter
                totalMergedCols += 1

    logger.info(f'合并完成！共处理了 {len(targets)} 组区域，累积合并 {totalMergedCols} 列。')

def run(appComHandle, logQueue):
    utils.runWithSelection(appComHandle, logQueue, pluginInfo, pickerTitle="按列合并 - 区域选取器", actionFunc=_processMerge)