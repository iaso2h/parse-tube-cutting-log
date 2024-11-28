import console
import config
import cutRecord
import dispatch
import partList
import util

import os
import json
import dearpygui.dearpygui as dpg
import win32api


def unmergeAllCellSave():
    dispatch.unmergeAllCell(dispatch.wb.active)
    util.saveWorkbook(dispatch.wb, config.DISPATCH_FILE_PATH)

dpg.create_context()
reg = dpg.add_font_registry()
fontName = dpg.add_font(file=r"C:\Windows\Fonts\msyh.ttc", size=17, parent=reg)
dpg.add_font_range(0x0001, 0x9FFF, parent=fontName)
dpg.bind_font(fontName)

if win32api.GetSystemMetrics(0) < win32api.GetSystemMetrics(1) and config.GUI_GEOMETRY_PATH.exists():
    with open(config.GUI_GEOMETRY_PATH, "r", encoding="utf-8") as f:
        geo = json.load(f)
else:
    geo = {
            "x_pos": 800,
            "y_pos": 600,
            "width": 290,
            "height": 192,
    }

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

dpg.setup_dearpygui()

with dpg.window(
        label="TubePro辅助 v" + config.VERSION,
        autosize=False,
        no_resize=True,
        width=290,
        no_close=True,
        no_title_bar=False,
        no_move=True,
        no_collapse=True,
    ):
    with dpg.group(horizontal=True, horizontal_spacing=60):
        dpg.add_text(f"编程: {config.AUTHOR}")
        dpg.add_text(f"最后更新: {config.LASTUPDATED}")
    dpg.add_separator(label="开料")
    with dpg.group(horizontal=True):
        dpg.add_button(label="程序截图",     callback=cutRecord.takeScreenshot)
        dpg.add_button(label="更新本地截图", callback=cutRecord.updateScreenshotRecords)
        dpg.add_button(label="重新链接截图", callback=cutRecord.relinkScreenshots)
    dpg.add_separator(label="排样文件")
    with dpg.group(horizontal=True):
        dpg.add_button(label="名称检查",     callback=partList.invalidNamingParts)
        dpg.add_button(label="工件规格总览", callback=partList.exportDimensions)
    if os.getlogin() == "OT03":
        dpg.add_separator(label="派工单")
        with dpg.group(horizontal=True):
            dpg.add_button(label="填写工件", callback=dispatch.fillPartInfo)
            dpg.add_button(label="表格优化", callback=dispatch.beautifyCells)
            dpg.add_button(label="取消合并", callback=unmergeAllCellSave)
    dpg.add_input_text(
            multiline=True,
            default_value=console.logFlow,
            tab_input=True,
            tracked=False,
            width=260,
            height=155,
            readonly=True,
            tag="log",
            )
    dpg.add_button(label="退出", callback=dpg.destroy_context)


dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
