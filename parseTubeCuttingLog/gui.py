import config
import cutRecord
import json
import dispatch
import partList
import util

import dearpygui.dearpygui as dpg


def unmergeAllCellSave():
    dispatch.unmergeAllCell(dispatch.wb.active)
    util.saveWorkbook(dispatch.wb, config.DISPATCH_FILE_PATH)

dpg.create_context()
reg = dpg.add_font_registry()
fontName = dpg.add_font(file=r"C:\Windows\Fonts\msyh.ttc", size=18, parent=reg)
dpg.add_font_range(0x0001, 0x9FFF, parent=fontName)
dpg.bind_font(fontName)

if config.GUI_GEOMETRY_PATH.exists():
    with open(config.GUI_GEOMETRY_PATH, "r", encoding="utf-8") as f:
        geo = json.load(f)
    dpg.create_viewport(
            title="TubePro Aid",
            decorated=False,
            x_pos=geo["x_pos"],
            y_pos=geo["y_pos"],
            width=geo["width"],
            height=geo["height"],
            always_on_top=True,
            resizable=False,
        )
else:
    dpg.create_viewport(
            title="TubePro Aid",
            decorated=False,
            width=200,
            height=230,
            always_on_top=False,
        )

dpg.setup_dearpygui()

with dpg.window(
        label="TubePro辅助",
        autosize=True,
        no_close=True,
        no_title_bar=False,
        no_move=True,
        no_collapse=True
    ):
    dpg.add_text(f"编程: {config.AUTHOR}")
    dpg.add_text(f"版本号: {config.VERSION}")
    dpg.add_text(f"最后更新: {config.LASTUPDATED}")
    dpg.add_button(label="开料截图",             callback=cutRecord.takeScreenshot)
    dpg.add_button(label="更新所有开料截图",     callback=cutRecord.updateScreenshotRecords)
    dpg.add_button(label="重新链接所有开料截图", callback=cutRecord.relinkScreenshots)
    dpg.add_button(label="激光文件名称检查",     callback=partList.invalidNamingParts)
    dpg.add_button(label="激光文件规格列表",     callback=partList.exportDimensions)
    dpg.add_button(label="派工单填写",           callback=dispatch.fillPartInfo)
    dpg.add_button(label="派工单优化",           callback=dispatch.beautifyCells)
    dpg.add_button(label="派工单表格取消合并",   callback=unmergeAllCellSave)
    dpg.add_button(label="退出", callback=dpg.destroy_context)

dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
