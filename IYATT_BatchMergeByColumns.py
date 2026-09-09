'''
file: IYATT_BatchMergeByColumns.py
description: 按列批量合并选中区域
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
from pytableenginesdk.PluginLogger import PluginLogger
from pytableenginesdk.SelectionPickerUI import SelectionPickerUI

pluginInfo = {
    'name': 'BatchMergeByColumns',
    'author': 'IYATT',
    'description': '按列批量合并选中区域',
    'version': '0.0.1',
}

# Excel 常量
xlCenter = -4108

def run(appComHandle, logQueue):
    logger = PluginLogger(logQueue, pluginInfo)
    logger.info('开始执行【按列批量合并】插件...')

    if appComHandle.Workbooks.Count == 0:
        logger.warning('当前没有打开的工作簿。')
        return

    picker = SelectionPickerUI(
        appComHandle,
        logger,
        title="按列合并 - 区域选取器"
    )

    confirmed, targets = picker.show()

    if not confirmed or not targets:
        logger.info('用户取消了选取操作，插件退出。')
        return

    # 执行合并核心业务逻辑
    appComHandle.ScreenUpdating = False
    appComHandle.DisplayAlerts = False

    totalMergedCols = 0

    try:
        for sheetName, address, mainRng in targets:
            logger.info(f'正在处理工作表 [{sheetName}] 的区域 [{address}]...')
            areasCount = getattr(mainRng.Areas, 'Count', 1)

            for aIdx in range(1, areasCount + 1):
                subRng = mainRng.Areas.Item(aIdx) if areasCount > 1 else mainRng
                colCount = subRng.Columns.Count

                for i in range(1, colCount + 1):
                    col = subRng.Columns.Item(i)
                    col.Merge()
                    col.HorizontalAlignment = xlCenter
                    col.VerticalAlignment = xlCenter
                    totalMergedCols += 1

        logger.info(f'合并完成！共处理了 {len(targets)} 组区域，累积合并 {totalMergedCols} 列。')

    except Exception as e:
        logger.error(f'合并处理出现异常: {e}')
        raise e

    finally:
        appComHandle.DisplayAlerts = True
        appComHandle.ScreenUpdating = True