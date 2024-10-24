import config
import cutRecord
import dispatch
import util

import dearpygui.dearpygui as dpg


def unmergeAllCellSave():
    dispatch.unmergeAllCell(dispatch.wb["OT计件表"])
    util.saveWorkbook(dispatch.dispatchFilePath, dispatch.wb)

dpg.create_context()
reg = dpg.add_font_registry()
fontName = dpg.add_font(file=r"C:\Windows\Fonts\msyh.ttc", size=18, parent=reg)
dpg.add_font_range(0x0001, 0x9FFF, parent=fontName)
dpg.bind_font(fontName)

dpg.create_viewport(
        title="TubePro Aid",
        decorated=False,
        x_pos=780,
        y_pos=1200,
        width=254,
        height=250,
        always_on_top=True,
        resizable=False,
    )
dpg.setup_dearpygui()

with dpg.window(
        label="功能列表",
        autosize=True,
        no_close=True,
        no_title_bar=False,
        no_move=True,
        no_collapse=True
    ):
    dpg.add_text(f"此TubePro日志分析程序由{config.AUTHOR}编写")
    dpg.add_text(f"版本号: {config.VERSION}")
    dpg.add_text(f"最后更新: {config.LASTUPDATED}")
    dpg.add_button(label="开料截图",             callback=cutRecord.takeScreenshot)
    dpg.add_button(label="开料记录",             callback=cutRecord.updateScreenshotRecords)
    dpg.add_button(label="开料记录截图重新链接", callback=cutRecord.relinkScreenshots)
    dpg.add_button(label="派工单",               callback=dispatch.fillPartInfo)
    dpg.add_button(label="派工单优化",           callback=dispatch.beautifyCells)
    dpg.add_button(label="派工单表格取消合并",   callback=unmergeAllCellSave)

dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
