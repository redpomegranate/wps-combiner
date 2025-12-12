import os
import time
import win32com.client as win32
from pathlib import Path


def kill_processes():
    """强制清理后台残留进程"""
    print("正在清理后台进程...")
    os.system("taskkill /f /im et.exe >nul 2>&1")
    os.system("taskkill /f /im excel.exe >nul 2>&1")
    os.system("taskkill /f /im wps.exe >nul 2>&1")
    time.sleep(1)


def merge_wps_fix_save(source_folder, output_filename):
    # 1. 清理环境
    kill_processes()

    source_folder = str(Path(source_folder).resolve())
    output_path = str(Path(source_folder) / output_filename)

    # 如果输出文件已存在，先删除，防止SaveAs报错
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except:
            print(f"⚠️ 无法删除旧文件: {output_filename}，可能会导致保存失败。")

    print(f"尝试启动 WPS 引擎...")

    try:
        # 优先尝试 WPS 表格核心进程
        app = win32.Dispatch("Ket.Application")
    except Exception:
        try:
            app = win32.Dispatch("Et.Application")
        except Exception:
            print("❌ 无法启动 WPS，请检查安装。")
            return

    # 关键设置
    app.Visible = False
    app.DisplayAlerts = False

    try:
        app.EnableEvents = False
    except:
        pass

    target_wb = None
    try:
        # 创建新工作簿
        target_wb = app.Workbooks.Add()

        files = [f for f in os.listdir(source_folder)
                 if f.lower().endswith(('.xls', '.xlsx', '.xlsm'))
                 and not f.startswith('~$')
                 and f != output_filename]
        files.sort()

        print(f"检测到 {len(files)} 个文件，开始合并...")

        for file_name in files:
            file_path = os.path.join(source_folder, file_name)
            print(f"处理中: {file_name}")

            source_wb = None
            try:
                # 只读打开
                source_wb = app.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=True)

                # 倒序复制 Sheet
                for i in range(source_wb.Sheets.Count, 0, -1):
                    sheet = source_wb.Sheets(i)

                    # 构造新名字 (处理过长文件名)
                    clean_fname = os.path.splitext(file_name)[0]
                    new_sheet_name = f"{clean_fname}_{sheet.Name}"[:30]

                    # 复制到目标工作簿最前面
                    # 注意：这里直接指定 target_wb 可能引用丢失，改用 app.Workbooks(1) 这种绝对引用更稳
                    sheet.Copy(Before=target_wb.Sheets(1))

                    try:
                        target_wb.Sheets(1).Name = new_sheet_name
                    except:
                        pass

                source_wb.Close(SaveChanges=False)

            except Exception as e:
                print(f"⚠️ 跳过文件 {file_name}: {str(e)}")
                if source_wb:
                    try:
                        source_wb.Close(SaveChanges=False)
                    except:
                        pass

        # 清理默认 Sheet
        try:
            for s in target_wb.Sheets:
                if s.Name == "Sheet1" and target_wb.Sheets.Count > 1:
                    s.Delete()
        except:
            pass

        # ==========================================
        # 🛠️ 修复核心：保存环节
        # ==========================================
        print("正在保存文件...")

        # 1. 激活目标工作簿，确保它处于焦点
        target_wb.Activate()

        # 2. 使用 SaveAs 保存为 .xls (FileFormat=56)
        # 56 = xlExcel8 (97-2003 format), 是 WPS 支持最好的带宏格式
        # 避免使用 .xlsm (FileFormat=52) 因为在 WPS COM 接口中经常出现"成员找不到"

        try:
            # 尝试方案 A: 明确指定格式 56 (xls)
            target_wb.SaveAs(output_path, FileFormat=56)
        except Exception as e:
            print(f"⚠️ 方案A失败 ({e})，尝试方案B...")
            try:
                # 尝试方案 B: 不指定格式，让 WPS 根据后缀名猜
                # 这种方式最"土"，但在 WPS 抽风时往往有效
                target_wb.SaveAs(Filename=output_path)
            except Exception as e2:
                print(f"⚠️ 方案B失败 ({e2})，尝试方案C (另存副本)...")
                # 尝试方案 C: SaveCopyAs (通常不会报错，但无法更改当前打开的文件名)
                target_wb.SaveCopyAs(output_path)

        print(f"✅ 合并成功！文件已保存为: {output_path}")

    except Exception as e:
        print(f"❌ 全局错误: {e}")

    finally:
        # 清理
        try:
            if target_wb: target_wb.Close(SaveChanges=False)
        except:
            pass
        app.Quit()
        kill_processes()


if __name__ == "__main__":
    FOLDER = r"D:\Work\建能院\_技术资源池V1.0\电源电气"
    # 🔴 注意：后缀名改为 .xls 以获得最佳兼容性
    OUTPUT_NAME = "最终合并版_WPS专用.xls"

    merge_wps_fix_save(FOLDER, OUTPUT_NAME)
