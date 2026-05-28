import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import shutil
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from threading import Thread
import subprocess
import json
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
import random
import multiprocessing

# EXE化対応: マルチプロセシングのfreeze_supportを追加
if __name__ == "__main__":
    multiprocessing.freeze_support()

# --- 定数 ---
THUMBNAIL_SCALE = 0.17
THUMBNAILS_PER_ROW = 5
THUMBNAIL_WIDTH = 120
THUMBNAIL_HEIGHT = 200
MAX_WORKER_THREADS = 4
UI_UPDATE_DELAY = 100

# --- サポートファイル形式定義 ---
SUPPORTED_EXTENSIONS = {
    '.pdf': 'PDF',
    '.doc': 'Word文書',
    '.docx': 'Word文書',
    '.docm': 'Word文書',
    '.dotx': 'Word文書',
    '.dotm': 'Word文書',
    '.xls': 'Excel文書',
    '.xlsx': 'Excel文書',
    '.xlsm': 'Excel文書',
    '.xlsb': 'Excel文書',
    '.ppt': 'PowerPoint文書',
    '.pptx': 'PowerPoint文書',
    '.pptm': 'PowerPoint文書',
    '.potx': 'PowerPoint文書',
    '.potm': 'PowerPoint文書',
    '.ppsx': 'PowerPoint文書',
    '.ppsm': 'PowerPoint文書',
    '.jtd': '一太郎文書',
    '.jtt': '一太郎文書'
}


def get_supported_extensions_text():
    """サポートされている拡張子のテキストを取得"""
    return "対応ファイル形式: PDF, Word, Excel, PowerPoint"


def check_file_support(file_path):
    """ファイルがサポートされているかチェック"""
    file_ext = os.path.splitext(file_path)[1].lower()
    return file_ext in SUPPORTED_EXTENSIONS


def get_file_type_name(file_path):
    """ファイルタイプ名を取得"""
    file_ext = os.path.splitext(file_path)[1].lower()
    return SUPPORTED_EXTENSIONS.get(file_ext, f"{file_ext}ファイル")


# --- パレットカラー定義 ---
PALETTE_COLORS = [
    "#FFE6E6",  # 薄いピンク
    "#E6F3FF",  # 薄い青
    "#E6FFE6",  # 薄い緑
    "#FFFACD",  # レモンシフォン
    "#F0E6FF",  # 薄い紫
    "#FFE6CC",  # 薄いオレンジ
    "#E6FFFF",  # 薄いシアン
    "#FFE6F0",  # 薄いマゼンタ
    "#F0FFE6",  # 薄いライム
    "#FFE6E6",  # 薄いコーラル
    "#E6E6FF",  # 薄いラベンダー
    "#F5F5DC"  # ベージュ
]


def get_random_palette_color():
    """ランダムなパレットカラーを取得"""
    return random.choice(PALETTE_COLORS)


# --- グローバル変数 ---
main_label = None
file_name_entry = None
input_frame = None
thumbnail_frame = None
pdf_data_list = []
pdf_file_names = []
pdf_rotations = []
pdf_colors = []
thumbnail_containers = []
dragged_item_idx = None
dragged_image = None
drag_offset_x = 0
drag_offset_y = 0
app = None
executor = ThreadPoolExecutor(max_workers=MAX_WORKER_THREADS)
dnd_frame = None
content_frame = None
button_frame = None
thumbnails_wrapper = None
processing_label = None
progress_bar = None
progress_frame = None
page_number_var = None
start_page_var = None
copyright_label = None


# --- UI更新ヘルパー関数 ---
def update_main_message(text):
    if main_label:
        main_label.config(text=text)


def bring_copyright_to_front():
    """コピーライトを最前面に表示"""
    if copyright_label:
        copyright_label.lift()
        copyright_label.tkraise()


def show_dnd_area():
    dnd_frame.pack(fill="both", expand=True, padx=20, pady=20)
    dnd_frame.lift()
    content_frame.pack_forget()
    button_frame.pack_forget()
    if progress_frame:
        progress_frame.pack_forget()
    app.after(10, bring_copyright_to_front)


def show_content_area():
    dnd_frame.pack_forget()
    content_frame.pack(fill="both", expand=True, padx=20, pady=10)
    button_frame.pack(side="bottom", fill="x", pady=10)
    if processing_label:
        processing_label.pack_forget()
    if progress_frame:
        progress_frame.pack_forget()
    app.after(10, bring_copyright_to_front)


def show_ichitaro_error_dialog(filename):
    """一太郎変換エラー時の詳細ダイアログ"""
    message = f"""一太郎ファイルの変換をスキップしました。

ファイル: {filename}

【現在の状況】
・TaroView（一太郎ビューアー）はインストール済み
・自動PDF変換用のCOMインターフェースが利用不可
・他のファイル（Word、Excel、PowerPoint、PDF）は正常変換可能

【一太郎ファイルを変換したい場合】
1. 手動でTaroViewを起動してPDF変換
2. 一太郎本体をインストール
3. Word 2016以降で.jtdファイルを開いてPDF保存

【このまま続行】
一太郎ファイル以外を結合して続行できます。"""

    result = messagebox.askquestion("一太郎ファイル変換",
                                    message + "\n\n一太郎ファイルをスキップして続行しますか？",
                                    icon='question')
    return result == 'yes'


def convert_office_to_pdf_direct(file_path, output_dir, file_type):
    """Office文書を直接PDFに変換（subprocess を使わない方法）"""
    try:
        import win32com.client as win32
        import pythoncom

        # COM初期化
        pythoncom.CoInitialize()

        # ファイルパスの正規化
        normalized_file_path = os.path.abspath(file_path)
        normalized_output_dir = os.path.abspath(output_dir)

        # 出力ファイル名の作成
        base_name = os.path.basename(normalized_file_path)
        output_file_name = os.path.splitext(base_name)[0] + '.pdf'
        output_path = os.path.join(normalized_output_dir, output_file_name)

        print(f"直接変換開始: {normalized_file_path} -> {output_path}")

        app_com = None
        doc = None

        try:
            if file_type == 'word':
                print("Word文書を処理中...")
                app_com = win32.DispatchEx("Word.Application")
                app_com.Visible = False
                app_com.DisplayAlerts = False

                doc = app_com.Documents.Open(normalized_file_path, ReadOnly=True,
                                         AddToRecentFiles=False)
                doc.ExportAsFixedFormat(output_path, 17, False, 0, 0, 0, 0, 7,
                                        True,
                                        True, 2, True, True, False)

            elif file_type == 'excel':
                print("Excel文書を処理中...")
                app_com = win32.DispatchEx("Excel.Application")
                app_com.Visible = False
                app_com.DisplayAlerts = False

                doc = app_com.Workbooks.Open(normalized_file_path,
                                         UpdateLinks=False, ReadOnly=True)
                doc.ExportAsFixedFormat(0, output_path, IgnorePrintAreas=False)

            elif file_type == 'powerpoint':
                print("PowerPoint文書を処理中...")
                app_com = win32.DispatchEx("PowerPoint.Application")
                app_com.Visible = 1

                doc = app_com.Presentations.Open(normalized_file_path,
                                             ReadOnly=True, WithWindow=False)
                import time
                time.sleep(1)
                doc.SaveAs(output_path, FileFormat=32)

            print(f"直接変換成功: {output_path}")
            return output_path

        finally:
            try:
                if doc:
                    if file_type == 'word':
                        doc.Close(SaveChanges=False)
                    elif file_type == 'excel':
                        doc.Close(SaveChanges=False)
                    elif file_type == 'powerpoint':
                        doc.Close()
            except:
                pass

            try:
                if app_com:
                    if file_type == 'powerpoint':
                        import time
                        time.sleep(0.5)
                    app_com.Quit()
            except:
                pass

            try:
                pythoncom.CoUninitialize()
            except:
                pass

    except Exception as e:
        print(f"直接変換エラー: {e}")
        return None


def convert_via_subprocess(file_path, temp_dir, file_type):
    """subprocess経由での変換（フォールバック用）"""
    try:
        # EXE化対応: converter.pyのパスを正しく取得
        if getattr(sys, 'frozen', False):
            converter_script_path = os.path.join(sys._MEIPASS, 'converter.py')
        else:
            converter_script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 'converter.py')

        if not os.path.exists(converter_script_path):
            print(f"converter.pyが見つかりません: {converter_script_path}")
            return None

        args_json = json.dumps({
            "file_path": file_path,
            "output_dir": temp_dir,
            "file_type": file_type
        }, ensure_ascii=False)

        print(f"subprocess変換開始...")
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUNBUFFERED'] = '1'

        # EXE化対応: Python実行可能ファイルのパスを正しく取得
        if getattr(sys, 'frozen', False):
            import shutil as sh
            python_executable = sh.which('python')
            if not python_executable:
                python_executable = sh.which('python3')

            if not python_executable:
                print(
                    "システムのPythonが見つかりません。Office文書変換をスキップします。")
                return None

            print(f"使用するPython: {python_executable}")
        else:
            python_executable = sys.executable

        startupinfo = None
        creationflags = 0
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            [python_executable, converter_script_path, args_json],
            capture_output=True, text=True, encoding='utf-8',
            check=True, timeout=300, env=env,
            startupinfo=startupinfo, creationflags=creationflags)

        output_path = result.stdout.strip()
        print(f"subprocess変換完了: {output_path}")

        if not output_path or not os.path.exists(output_path):
            print(f"subprocess変換から無効な出力パス: {output_path}")
            print(f"stderr: {result.stderr}")
            return None

        return output_path

    except subprocess.CalledProcessError as e:
        print(f"subprocess変換エラー: {e.returncode}")
        print(f"stdout: {e.stdout}")
        print(f"stderr: {e.stderr}")
        return None
    except Exception as e:
        print(f"subprocess変換例外: {e}")
        return None


# --- PDF処理関数 ---
def convert_to_pdf_in_memory(file_path):
    temp_dir = tempfile.mkdtemp()
    original_filename = os.path.basename(file_path)
    try:
        print(f"変換開始: {file_path}")
        file_ext = os.path.splitext(file_path)[1].lower()
        output_path = None

        if file_ext in ['.doc', '.docx', '.docm', '.dotx', '.dotm', '.xls',
                        '.xlsx', '.xlsm', '.xlsb', '.ppt', '.pptx', '.pptm',
                        '.potx', '.potm', '.ppsx', '.ppsm']:
            print(f"Office文書として処理: {file_path}")

            file_type = 'word'
            if file_ext in ['.xls', '.xlsx', '.xlsm', '.xlsb']:
                file_type = 'excel'
            elif file_ext in ['.ppt', '.pptx', '.pptm', '.potx', '.potm',
                              '.ppsx', '.ppsm']:
                file_type = 'powerpoint'

            if getattr(sys, 'frozen', False):
                print("EXE環境: 直接変換を試行中...")
                output_path = convert_office_to_pdf_direct(file_path, temp_dir,
                                                           file_type)

                if not output_path or not os.path.exists(output_path):
                    print("直接変換失敗、subprocess方式にフォールバック...")
                    output_path = convert_via_subprocess(file_path, temp_dir,
                                                         file_type)
            else:
                print("開発環境: subprocess方式を使用...")
                output_path = convert_via_subprocess(file_path, temp_dir,
                                                     file_type)

        elif file_ext in ['.jtd', '.jtt']:
            # 一太郎ファイルはメインスレッドでダイアログを出してスキップ
            print(f"一太郎ファイルは現在サポートされていません: {file_path}")
            app.after(0, lambda fn=original_filename: show_ichitaro_error_dialog(fn))
            return None, None

        elif file_ext == '.pdf':
            print(f"PDFファイルとして処理: {file_path}")
            output_path = file_path

        if output_path and os.path.exists(output_path):
            print(f"PDFファイル読み込み開始: {output_path}")
            with open(output_path, 'rb') as f:
                pdf_data = f.read()
            print(f"PDFファイル読み込み完了: {len(pdf_data)} bytes")
            return pdf_data, original_filename
        else:
            print(f"出力ファイルが見つからない: {output_path}")
            return None, None

    except subprocess.TimeoutExpired as e:
        print(f"変換タイムアウト: {file_path} - {e}")
        return None, None
    except Exception as e:
        print(f"変換エラー: {file_path} - {type(e).__name__}: {e}")
        import traceback
        print("詳細なエラー情報:")
        traceback.print_exc()
        return None, None
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            print(f"一時ディレクトリ削除エラー: {e}")


def process_files_async(items):
    """バックグラウンドスレッドで呼ばれる。UI操作は必ずapp.after経由で行う。"""
    new_pdf_data = []
    new_pdf_names = []
    unsupported_files = []

    print(f"処理開始: {len(items)}個のアイテム")

    total_files = 0

    for item_path in items:
        item_path = os.path.normpath(item_path.strip('"'))
        print(f"アイテム処理: {item_path}")

        if os.path.isfile(item_path):
            if check_file_support(item_path):
                print(f"サポート対象ファイルとして追加: {item_path}")
                total_files += 1
            else:
                print(f"サポート対象外ファイル: {item_path}")
                unsupported_files.append(item_path)

        elif os.path.isdir(item_path):
            print(f"ディレクトリとして処理: {item_path}")
            for root, _, files in os.walk(item_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    if check_file_support(file_path):
                        print(f"ディレクトリ内サポート対象ファイル: {file_path}")
                        total_files += 1
                    else:
                        print(f"ディレクトリ内サポート対象外ファイル: {file_path}")
                        unsupported_files.append(file_path)

    if unsupported_files:
        def show_unsupported_files_warning():
            if len(unsupported_files) == 1:
                file_name = os.path.basename(unsupported_files[0])
                file_type = get_file_type_name(unsupported_files[0])
                message = f"""サポート対象外のファイルがありました：

ファイル: {file_name}
種類: {file_type}

{get_supported_extensions_text()}

サポート対象のファイルのみ処理を続行します。"""
            else:
                file_list = []
                for i, file_path in enumerate(unsupported_files[:5]):
                    file_name = os.path.basename(file_path)
                    file_type = get_file_type_name(file_path)
                    file_list.append(f"• {file_name} ({file_type})")

                if len(unsupported_files) > 5:
                    file_list.append(f"... 他 {len(unsupported_files) - 5} 件")

                message = f"""サポート対象外のファイルが {len(unsupported_files)} 件ありました：

{chr(10).join(file_list)}

{get_supported_extensions_text()}

サポート対象のファイルのみ処理を続行します。"""

            messagebox.showwarning("サポート対象外ファイル", message)

        app.after(0, show_unsupported_files_warning)

    print(f"合計{total_files}個のサポート対象ファイル")

    if total_files == 0:
        if unsupported_files:
            app.after(UI_UPDATE_DELAY, lambda: handle_no_supported_files())
        else:
            app.after(UI_UPDATE_DELAY, lambda: handle_no_supported_files())
        return

    # progress_bar の設定はメインスレッドで行う
    app.after(0, lambda t=total_files: progress_bar.config(maximum=t, value=0))

    completed_count = 0

    def process_single_file(file_path):
        nonlocal completed_count
        try:
            print(f"ファイル変換開始: {file_path}")
            result, file_name = convert_to_pdf_in_memory(file_path)
            completed_count += 1
            print(f"ファイル変換完了: {file_name} ({completed_count}/{total_files})")

            # 進捗バーをメインスレッドで更新
            app.after(0, lambda c=completed_count: progress_bar.config(value=c))

            if result:
                new_pdf_data.append(result)
                new_pdf_names.append(file_name)
                print(f"PDF変換成功: {file_name}")
            else:
                print(f"PDF変換失敗: {file_path}")
        except Exception as e:
            completed_count += 1
            print(f"ファイル変換エラー: {file_path} - {type(e).__name__}: {e}")
            app.after(0, lambda c=completed_count: progress_bar.config(value=c))

    for item_path in items:
        item_path = os.path.normpath(item_path.strip('"'))

        if os.path.isfile(item_path):
            if check_file_support(item_path):
                process_single_file(item_path)

        elif os.path.isdir(item_path):
            for root, _, files in os.walk(item_path):
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    if check_file_support(file_path):
                        process_single_file(file_path)

    print(
        f"全処理完了: 成功{len(new_pdf_data)}個, 失敗{total_files - len(new_pdf_data)}個, "
        f"サポート対象外{len(unsupported_files)}個")

    app.after(UI_UPDATE_DELAY,
              lambda: handle_processing_complete(new_pdf_data, new_pdf_names))


# --- ファイル名表示関数 ---
def format_filename_for_display(filename):
    """ファイル名を表示用にフォーマット"""
    name_without_ext = os.path.splitext(filename)[0]
    if len(name_without_ext) > 14:
        return name_without_ext[:13] + "..."
    return name_without_ext


# --- アイコン作成関数 ---
def create_rotation_icon(size=16):
    """回転矢印アイコンを作成"""
    img = Image.new('RGBA', (size, size), (255, 255, 255, 255))
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)

    cx, cy = size // 2, size // 2
    radius = size // 3

    draw.arc([(cx - radius, cy - radius), (cx + radius, cy + radius)],
             start=45, end=315, fill='#0066CC', width=3)

    arrow_size = 4
    arrow_x = cx + radius * 0.7
    arrow_y = cy - radius * 0.7

    arrow_points = [
        (arrow_x, arrow_y),
        (arrow_x - arrow_size, arrow_y + arrow_size),
        (arrow_x + arrow_size, arrow_y + arrow_size)
    ]
    draw.polygon(arrow_points, fill='#0066CC')

    draw.line(
        [(arrow_x, arrow_y + arrow_size), (arrow_x, arrow_y + arrow_size + 2)],
        fill='#0066CC', width=2)

    return ImageTk.PhotoImage(img)


def create_delete_icon(size=16):
    """削除用×アイコンを作成"""
    img = Image.new('RGBA', (size, size), (255, 230, 230, 255))
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)

    line_width = 3
    margin = 3

    draw.line([(margin, margin), (size - margin, size - margin)],
              fill='#CC0000', width=line_width)

    draw.line([(size - margin, margin), (margin, size - margin)],
              fill='#CC0000', width=line_width)

    return ImageTk.PhotoImage(img)


# --- サムネイル操作関数 ---
def delete_thumbnail(container):
    """指定コンテナのサムネイルを削除"""
    global pdf_data_list, pdf_file_names, pdf_rotations, pdf_colors, thumbnail_containers

    try:
        index = thumbnail_containers.index(container)

        if 0 <= index < len(pdf_data_list):
            pdf_data_list.pop(index)
            pdf_file_names.pop(index)
            pdf_rotations.pop(index)
            pdf_colors.pop(index)

            container.destroy()
            thumbnail_containers.pop(index)

            refresh_thumbnail_layout()

            if not pdf_data_list:
                update_main_message(
                    "Word, Excel, PowerPoint, PDFファイルを\nここにドラッグ＆ドロップしてください")
                show_dnd_area()
    except ValueError:
        print("削除対象のコンテナが見つかりません")


def rotate_thumbnail(container):
    """指定コンテナのサムネイルを90度回転"""
    global pdf_rotations

    try:
        index = thumbnail_containers.index(container)

        if 0 <= index < len(pdf_rotations):
            pdf_rotations[index] = (pdf_rotations[index] + 90) % 360
            recreate_thumbnail(index)
    except ValueError:
        print("回転対象のコンテナが見つかりません")


def recreate_thumbnail(index):
    """指定インデックスのサムネイルを再生成"""
    if 0 <= index < len(pdf_data_list):
        old_container = thumbnail_containers[index]
        old_container.destroy()

        new_container = create_thumbnail_widget(pdf_data_list[index],
                                                pdf_file_names[index], index)
        thumbnail_containers[index] = new_container

        row = index // THUMBNAILS_PER_ROW
        col = index % THUMBNAILS_PER_ROW
        new_container.grid(row=row, column=col, padx=5, pady=5, sticky="nw")


# --- サムネイル作成関数 ---
def create_thumbnail_widget(pdf_data, file_name, index):
    """個別のサムネイルウィジェットを作成"""
    try:
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        page = doc[0]

        rotation = pdf_rotations[index] if index < len(pdf_rotations) else 0

        matrix = fitz.Matrix(THUMBNAIL_SCALE, THUMBNAIL_SCALE)

        if rotation != 0:
            rotation_matrix = fitz.Matrix(1, 1).prerotate(rotation)
            matrix = matrix * rotation_matrix

        # RGB・アルファなしで明示指定することでモード不一致を防ぐ
        pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((100, 140), Image.Resampling.LANCZOS)
        photo_img = ImageTk.PhotoImage(img)

        if index < len(pdf_colors):
            bg_color = pdf_colors[index]
        else:
            bg_color = get_random_palette_color()
            if len(pdf_colors) <= index:
                pdf_colors.extend([get_random_palette_color() for _ in
                                   range(index - len(pdf_colors) + 1)])
            pdf_colors[index] = bg_color

        thumb_container = tk.Frame(thumbnail_frame, borderwidth=2,
                                   relief="groove",
                                   bg=bg_color, width=THUMBNAIL_WIDTH,
                                   height=THUMBNAIL_HEIGHT)
        thumb_container.pack_propagate(False)
        thumb_container.grid_propagate(False)

        image_frame = tk.Frame(thumb_container, bg=bg_color, height=150)
        image_frame.pack(side="top", fill="x", padx=2, pady=2)
        image_frame.pack_propagate(False)

        label = tk.Label(image_frame, image=photo_img, bg=bg_color)
        label.image = photo_img
        label.pack(expand=True)

        button_frame_widget = tk.Frame(image_frame, bg=bg_color)
        button_frame_widget.place(relx=1.0, rely=0.0, anchor='ne', x=-2, y=2)

        rotation_icon = create_rotation_icon()
        rotate_btn = tk.Button(button_frame_widget, image=rotation_icon,
                               command=lambda: rotate_thumbnail(
                                   thumb_container),
                               width=18, height=18, relief='raised',
                               bg='#E6F3FF', bd=1, activebackground='#CCE7FF')
        rotate_btn.image = rotation_icon
        rotate_btn.pack(side='left', padx=(0, 2))

        delete_icon = create_delete_icon()
        delete_btn = tk.Button(button_frame_widget, image=delete_icon,
                               command=lambda: delete_thumbnail(
                                   thumb_container),
                               width=18, height=18, relief='raised',
                               bg='#FFE6E6', bd=1, activebackground='#FFCCCC')
        delete_btn.image = delete_icon
        delete_btn.pack(side='left')

        filename_frame = tk.Frame(thumb_container, bg="#ffffff", height=40)
        filename_frame.pack(side="bottom", fill="x", padx=2, pady=(0, 2))
        filename_frame.pack_propagate(False)

        display_name = format_filename_for_display(file_name)
        filename_label = tk.Label(filename_frame, text=display_name,
                                  justify="center",
                                  bg="white",
                                  fg="black",
                                  font=("Helvetica", 9, "bold"),
                                  wraplength=110,
                                  anchor="center")
        filename_label.pack(expand=True, fill="both")

        def bind_drag_events(widget):
            widget.unbind("<Button-1>")
            widget.unbind("<B1-Motion>")
            widget.unbind("<ButtonRelease-1>")
            widget.bind("<Button-1>",
                        lambda event: start_drag(event, thumb_container))
            widget.bind("<B1-Motion>",
                        lambda event: continue_drag(event, thumb_container))
            widget.bind("<ButtonRelease-1>", lambda event: end_drag(event))

        bind_drag_events(thumb_container)
        bind_drag_events(image_frame)
        bind_drag_events(label)
        bind_drag_events(filename_frame)
        bind_drag_events(filename_label)

        doc.close()
        return thumb_container

    except Exception as e:
        print(f"サムネイル作成に失敗しました: {file_name} - エラー: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_thumbnails_and_display(pdf_datas, pdf_names):
    """サムネイル作成とグリッド表示"""
    global thumbnail_frame, pdf_data_list, pdf_file_names, pdf_rotations, pdf_colors, thumbnail_containers

    pdf_data_list.extend(pdf_datas)
    pdf_file_names.extend(pdf_names)
    pdf_rotations.extend([0] * len(pdf_datas))
    pdf_colors.extend(
        [get_random_palette_color() for _ in range(len(pdf_datas))])

    for widget in thumbnail_frame.winfo_children():
        widget.destroy()
    thumbnail_containers.clear()

    if not pdf_data_list:
        return

    for i, pdf_data in enumerate(pdf_data_list):
        thumb_widget = create_thumbnail_widget(pdf_data, pdf_file_names[i], i)
        if thumb_widget:
            row = i // THUMBNAILS_PER_ROW
            col = i % THUMBNAILS_PER_ROW
            thumb_widget.grid(row=row, column=col, padx=5, pady=5, sticky="nw")
            thumbnail_containers.append(thumb_widget)


def refresh_thumbnail_layout():
    """サムネイルのレイアウトを再配置"""
    for i, container in enumerate(thumbnail_containers):
        row = i // THUMBNAILS_PER_ROW
        col = i % THUMBNAILS_PER_ROW
        container.grid(row=row, column=col, padx=5, pady=5, sticky="nw")
        update_container_bindings(container, i)


def update_container_bindings(container, correct_index):
    """コンテナとその子要素のドラッグイベントを正しいインデックスで再バインド"""

    def bind_drag_events_with_index(widget, target_container):
        widget.unbind("<Button-1>")
        widget.unbind("<B1-Motion>")
        widget.unbind("<ButtonRelease-1>")
        widget.bind("<Button-1>",
                    lambda event: start_drag(event, target_container))
        widget.bind("<B1-Motion>",
                    lambda event: continue_drag(event, target_container))
        widget.bind("<ButtonRelease-1>", lambda event: end_drag(event))

    bind_drag_events_with_index(container, container)

    for child in container.winfo_children():
        if isinstance(child, tk.Button):
            continue
        bind_drag_events_with_index(child, container)

        for grandchild in child.winfo_children():
            if isinstance(grandchild, tk.Button):
                if hasattr(grandchild, 'image') and grandchild.image:
                    if grandchild['bg'] == '#E6F3FF':
                        grandchild.config(
                            command=lambda c=container: rotate_thumbnail(c))
                    elif grandchild['bg'] == '#FFE6E6':
                        grandchild.config(
                            command=lambda c=container: delete_thumbnail(c))
                continue
            bind_drag_events_with_index(grandchild, container)


# --- ドラッグ＆ドロップロジック ---
def continue_drag(event, container):
    pass


def _insert_item(from_idx, to_idx):
    """アイテムを指定位置に挿入"""
    global thumbnail_containers, pdf_data_list, pdf_file_names, pdf_rotations, pdf_colors

    if from_idx == to_idx:
        return

    print(f"移動: {from_idx} -> {to_idx}")

    moving_container = thumbnail_containers[from_idx]
    moving_pdf_data = pdf_data_list[from_idx]
    moving_file_name = pdf_file_names[from_idx]
    moving_rotation = pdf_rotations[from_idx]
    moving_color = pdf_colors[from_idx]

    thumbnail_containers.pop(from_idx)
    pdf_data_list.pop(from_idx)
    pdf_file_names.pop(from_idx)
    pdf_rotations.pop(from_idx)
    pdf_colors.pop(from_idx)

    thumbnail_containers.insert(to_idx, moving_container)
    pdf_data_list.insert(to_idx, moving_pdf_data)
    pdf_file_names.insert(to_idx, moving_file_name)
    pdf_rotations.insert(to_idx, moving_rotation)
    pdf_colors.insert(to_idx, moving_color)

    refresh_thumbnail_layout()


def start_drag(event, container):
    global dragged_item_idx, dragged_image, drag_offset_x, drag_offset_y

    if dragged_item_idx is not None:
        return

    try:
        idx = thumbnail_containers.index(container)
    except ValueError:
        print("ドラッグ対象のコンテナが見つかりません")
        return

    if idx >= len(thumbnail_containers):
        return

    dragged_item_idx = idx
    drag_offset_x = event.x
    drag_offset_y = event.y

    drag_container_orig = thumbnail_containers[dragged_item_idx]
    drag_container_orig.configure(relief="sunken", bg="#CCCCCC")

    app.bind("<B1-Motion>", drag_motion)
    app.bind("<ButtonRelease-1>", end_drag)

    create_visual_drag_image(event)


def create_visual_drag_image(event):
    global dragged_image, dragged_item_idx

    if dragged_item_idx is None or dragged_image is not None:
        return

    try:
        dragged_image = tk.Toplevel(app)
        dragged_image.overrideredirect(True)
        dragged_image.attributes('-alpha', 0.7)
        dragged_image.configure(bg="white")

        drag_container_orig = thumbnail_containers[dragged_item_idx]
        children = drag_container_orig.winfo_children()

        if len(children) >= 2:
            image_frame = children[0]
            image_children = image_frame.winfo_children()
            if image_children:
                thumb_label = image_children[0]

                img_label = tk.Label(dragged_image, image=thumb_label.image,
                                     relief="raised", bd=2, bg="white")
                img_label.image = thumb_label.image
                img_label.pack(pady=2)

            filename_frame = children[1]
            filename_children = filename_frame.winfo_children()
            if filename_children:
                filename_label = filename_children[0]

                file_name_label = tk.Label(dragged_image,
                                           text=filename_label['text'],
                                           wraplength=100, justify="center",
                                           bg="white",
                                           font=("Helvetica", 9, "bold"))
                file_name_label.pack(pady=2)

        dragged_image.update_idletasks()

        x = event.x_root - drag_offset_x
        y = event.y_root - drag_offset_y
        dragged_image.geometry(f"+{x}+{y}")

    except Exception as e:
        print(f"ドラッグ画像作成エラー: {e}")


def drag_motion(event):
    global dragged_item_idx, dragged_image, drag_offset_x, drag_offset_y

    if dragged_item_idx is None or dragged_image is None:
        return

    try:
        x = event.x_root - drag_offset_x
        y = event.y_root - drag_offset_y
        dragged_image.geometry(f"+{x}+{y}")
    except tk.TclError:
        pass


def find_drop_target(event):
    """ドロップターゲットを見つける"""
    try:
        for i, container in enumerate(thumbnail_containers):
            if i == dragged_item_idx:
                continue

            try:
                container_x = container.winfo_rootx()
                container_y = container.winfo_rooty()
                container_width = container.winfo_width()
                container_height = container.winfo_height()

                if (
                        container_x <= event.x_root <= container_x + container_width and
                        container_y <= event.y_root <= container_y + container_height):
                    return i
            except tk.TclError:
                continue

        return -1
    except Exception as e:
        print(f"ドロップターゲット検索エラー: {e}")
        return -1


def end_drag(event):
    global dragged_item_idx, dragged_image, drag_offset_x, drag_offset_y

    if dragged_item_idx is None:
        return

    if dragged_item_idx < len(thumbnail_containers) and dragged_item_idx < len(
            pdf_colors):
        try:
            original_color = pdf_colors[dragged_item_idx]
            thumbnail_containers[dragged_item_idx].configure(relief="groove",
                                                             bg=original_color)
        except (IndexError, tk.TclError):
            pass

    drop_idx = find_drop_target(event)

    if drop_idx != -1 and drop_idx != dragged_item_idx:
        _insert_item(dragged_item_idx, drop_idx)

    if dragged_image:
        try:
            dragged_image.destroy()
        except tk.TclError:
            pass
        dragged_image = None

    try:
        app.unbind("<B1-Motion>")
        app.unbind("<ButtonRelease-1>")
    except tk.TclError:
        pass

    dragged_item_idx = None
    drag_offset_x = 0
    drag_offset_y = 0


# --- ページ番号追加関数 ---
def add_page_numbers(pdf_merger, start_from_page):
    """PDFにページ番号を追加（回転に対応）"""
    try:
        total_pages = pdf_merger.page_count

        for page_num in range(total_pages):
            page = pdf_merger[page_num]

            if page_num >= start_from_page - 1:
                display_page_num = page_num - (start_from_page - 1) + 1
                page_number_text = f"- {display_page_num} -"

                rotation = page.rotation

                original_rotation = page.rotation
                page.set_rotation(0)
                original_rect = page.rect
                page.set_rotation(original_rotation)

                current_rect = page.rect

                print(f"ページ {page_num + 1}: 回転角度={rotation}度")
                print(
                    f"  元のサイズ: {original_rect.width:.1f} x {original_rect.height:.1f}")
                print(
                    f"  現在のサイズ: {current_rect.width:.1f} x {current_rect.height:.1f}")

                text_width = len(page_number_text) * 6
                margin = 30

                if rotation == 0:
                    x = (original_rect.width - text_width) / 2
                    y = original_rect.height - margin

                elif rotation == 90:
                    x = margin
                    y = (original_rect.height + text_width) / 2

                elif rotation == 180:
                    x = (original_rect.width + text_width) / 2
                    y = margin

                elif rotation == 270:
                    x = original_rect.width - margin
                    y = (original_rect.height - text_width) / 2

                else:
                    x = (original_rect.width - text_width) / 2
                    y = original_rect.height - margin

                print(f"  ページ番号配置座標: ({x:.1f}, {y:.1f})")

                page.insert_text(
                    (x, y),
                    page_number_text,
                    fontsize=10,
                    color=(0, 0, 0),
                    overlay=True
                )

    except Exception as e:
        print(f"ページ番号追加エラー: {e}")
        import traceback
        traceback.print_exc()
        raise e


def combine_pdfs():
    if not pdf_data_list:
        messagebox.showwarning("警告", "結合するPDFがありません。")
        return

    output_name = file_name_entry.get().strip()
    if not output_name:
        messagebox.showwarning("警告", "結合後のファイル名を入力してください。")
        return

    output_name_with_ext = output_name + ".pdf"

    output_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=output_name_with_ext,
        filetypes=[("PDF files", "*.pdf")]
    )

    if not output_path:
        return

    try:
        pdf_merger = fitz.open()
        current_page_count = 0

        for i, pdf_data in enumerate(pdf_data_list):
            temp_doc = fitz.open(stream=pdf_data, filetype="pdf")

            rotation = pdf_rotations[i] if i < len(pdf_rotations) else 0

            for page_num in range(temp_doc.page_count):
                page = temp_doc[page_num]
                global_page_num = current_page_count + page_num

                if page_number_var.get() and global_page_num >= start_page_var.get() - 1:
                    display_page_num = global_page_num - (start_page_var.get() - 1) + 1
                    page_number_text = f"- {display_page_num} -"

                    rect = page.rect
                    text_width = len(page_number_text) * 6
                    margin = 30

                    if rotation == 0:
                        x = (rect.width - text_width) / 2
                        y = rect.height - margin

                    elif rotation == 90:
                        if rect.width > rect.height:
                            x = rect.width - margin
                            y = (rect.height - text_width) / 2 + 30
                        else:
                            x = rect.width / 2
                            y = rect.height - margin

                    elif rotation == 180:
                        x = (rect.width + text_width) / 2
                        y = margin + 10

                    elif rotation == 270:
                        if rect.width > rect.height:
                            x = margin * 2 - 30
                            y = rect.height / 2 - 15
                        else:
                            x = rect.width / 2
                            y = rect.height - margin

                    else:
                        x = (rect.width - text_width) / 2
                        y = rect.height - margin

                    page.insert_text(
                        (x, y),
                        page_number_text,
                        fontsize=10,
                        color=(0, 0, 0),
                        overlay=True,
                        rotate=rotation
                    )

                if rotation != 0:
                    page.set_rotation(rotation)

            pdf_merger.insert_pdf(temp_doc)
            current_page_count += temp_doc.page_count
            temp_doc.close()

        pdf_merger.save(output_path)
        pdf_merger.close()

        messagebox.showinfo("成功", f"PDFが結合され、'{output_path}'として保存されました。")

        pdf_data_list.clear()
        pdf_file_names.clear()
        pdf_rotations.clear()
        pdf_colors.clear()

        update_main_message("Word, Excel, PowerPoint, PDFファイルを\nここにドラッグ＆ドロップしてください")
        show_dnd_area()

    except Exception as e:
        messagebox.showerror("エラー", f"PDFの結合に失敗しました: {e}")
        import traceback
        traceback.print_exc()


def handle_no_supported_files():
    """サポート対象ファイルがない場合の処理"""
    if processing_label:
        processing_label.pack_forget()
    if progress_frame:
        progress_frame.pack_forget()

    update_main_message(
        "Word, Excel, PowerPoint, PDFファイルを\nここにドラッグ＆ドロップしてください")

    if not pdf_data_list:
        show_dnd_area()


def handle_processing_complete(pdf_datas, pdf_names):
    global main_label, dnd_frame

    show_content_area()

    update_main_message(
        "別のファイルがあれば、引き続きドラッグ＆ドロップしてください。\n準備ができたら、サムネイルの順番を入れ替えて結合ボタンを押してください。")

    create_thumbnails_and_display(pdf_datas, pdf_names)

    if pdf_file_names:
        first_file_name = os.path.splitext(pdf_file_names[0])[0]
        file_name_entry.delete(0, tk.END)
        file_name_entry.insert(0, f"{first_file_name}_combined")

    processing_label.config(
        text="PDFの順番をドラッグ＆ドロップで変更してください。\nファイルを追加したい場合は、エリア内にデータをドロップしてください。",
        font=("Helvetica", 10))
    processing_label.pack(side="left", padx=10)

    button_frame.pack(side="bottom", fill="x", pady=10)


class DragDropApp(TkinterDnD.Tk):
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

        def update_init_progress(value, message):
            if self.progress_callback:
                self.progress_callback(value, message)

        super().__init__()
        update_init_progress(30, "パステルカラーの絵の具を用意しています...")

        self.title("PDFパレット Ver1.0.2")
        self.geometry("750x600")
        self.resizable(False, False)

        update_init_progress(40, "必要な絵の具を設定しています...")
        self.set_application_icon()

        global main_label, file_name_entry, input_frame, thumbnail_frame, app, dnd_frame, content_frame, button_frame, thumbnails_wrapper, processing_label, progress_bar, progress_frame, page_number_var, start_page_var, copyright_label
        app = self

        update_init_progress(50, "筆洗バケツをきれいにしています...")
        page_number_var = tk.BooleanVar(value=True)
        start_page_var = tk.IntVar(value=1)

        update_init_progress(65, "キャンバスを作成しています...")
        dnd_frame = tk.Frame(self, relief="solid", borderwidth=2, bg="#ffffff")
        self.label = tk.Label(dnd_frame,
                              text="Word, Excel, PowerPoint, PDFファイルを\nここにドラッグ＆ドロップしてください",
                              font=("Helvetica", 12), fg="gray", bg="#ffffff",
                              wraplength=600)
        self.label.pack(fill="both", expand=True)
        main_label = self.label

        update_init_progress(70, "筆洗バケツにたっぷり注水しています...")
        initial_close_button = tk.Button(dnd_frame, text="閉じる",
                                         command=self.destroy,
                                         font=("Helvetica", 10), width=10,
                                         height=1)
        initial_close_button.place(relx=1.0, rely=1.0, anchor='se', x=-10,
                                   y=-10)

        update_init_progress(75, "筆を用意しています...")
        progress_frame = tk.Frame(self)
        progress_label_widget = tk.Label(progress_frame, text="進捗:",
                                         font=("Helvetica", 10))
        progress_label_widget.pack(side="left", padx=(10, 5))
        progress_bar = ttk.Progressbar(progress_frame, length=400,
                                       mode='determinate')
        progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))

        update_init_progress(80, "穂先を整えています...")
        content_frame = tk.Frame(self)

        update_init_progress(85, "作品名の記入欄を作成しています...")
        input_frame = tk.Frame(content_frame)
        input_frame.pack(side="top", pady=10)

        file_name_label = tk.Label(input_frame, text="結合後のファイル名:",
                                   font=("Helvetica", 10))
        file_name_label.pack(side="left", padx=(0, 5))
        file_name_entry = tk.Entry(input_frame, width=50,
                                   font=("Helvetica", 10))
        file_name_entry.pack(side="left")

        ext_label = tk.Label(input_frame, text=".pdf", font=("Helvetica", 10))
        ext_label.pack(side="left")

        update_init_progress(88, "ページ番号の記入欄を作成しています...")
        page_number_frame = tk.Frame(content_frame)
        page_number_frame.pack(side="top", pady=(5, 10))

        page_number_checkbox = tk.Checkbutton(
            page_number_frame,
            text="ページ番号を追加する",
            variable=page_number_var,
            font=("Helvetica", 10)
        )
        page_number_checkbox.pack(side="left")

        start_page_label = tk.Label(page_number_frame, text="開始ページ:",
                                    font=("Helvetica", 10))
        start_page_label.pack(side="left", padx=(20, 5))

        start_page_spinbox = tk.Spinbox(
            page_number_frame,
            from_=1,
            to=999,
            width=5,
            textvariable=start_page_var,
            font=("Helvetica", 10)
        )
        start_page_spinbox.pack(side="left")

        start_page_help_label = tk.Label(
            page_number_frame,
            text="(何ページ目から1ページとして番号付けするか)",
            font=("Helvetica", 9),
            fg="gray"
        )
        start_page_help_label.pack(side="left", padx=(5, 0))

        update_init_progress(90, "絵の具をパレットに準備中...")
        canvas = tk.Canvas(content_frame, borderwidth=0, highlightthickness=0,
                           bg="#ffffff")
        scrollbar_v = tk.Scrollbar(content_frame, orient="vertical",
                                   command=canvas.yview)
        scrollbar_h = tk.Scrollbar(content_frame, orient="horizontal",
                                   command=canvas.xview)
        canvas.configure(yscrollcommand=scrollbar_v.set,
                         xscrollcommand=scrollbar_h.set)

        scrollbar_v.pack(side="right", fill="y")
        scrollbar_h.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True, pady=(5, 0))

        thumbnail_frame = tk.Frame(canvas, bg="#ffffff")
        canvas.create_window((0, 0), window=thumbnail_frame, anchor="nw")

        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        thumbnail_frame.bind("<Configure>", configure_scroll_region)

        def on_mouse_wheel(event):
            if sys.platform == "win32" or sys.platform == "linux":
                delta = -1 * (event.delta // 120)
            else:
                delta = -1 * event.delta
            canvas.yview_scroll(delta, "units")

        canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        update_init_progress(93, "絵の具を混ぜています...")
        button_frame = tk.Frame(self)

        processing_label = tk.Label(button_frame, text="", fg="blue",
                                    font=("Helvetica", 10))

        right_button_frame = tk.Frame(button_frame)
        right_button_frame.pack(side="right")

        combine_button = tk.Button(right_button_frame, text="結合",
                                   command=combine_pdfs,
                                   font=("Helvetica", 12, "bold"), width=10,
                                   height=2)
        combine_button.pack(side="left", padx=10)

        close_button = tk.Button(right_button_frame, text="閉じる",
                                 command=self.destroy,
                                 font=("Helvetica", 12, "bold"), width=10,
                                 height=2)
        close_button.pack(side="left", padx=10)

        update_init_progress(95, "混ぜた絵の具を配置しています...")
        show_dnd_area()

        update_init_progress(97, "配置した絵の具を有効にしています...")
        dnd_frame.drop_target_register(DND_FILES)
        dnd_frame.dnd_bind("<<Drop>>", self.handle_drop)

        content_frame.drop_target_register(DND_FILES)
        content_frame.dnd_bind("<<Drop>>", self.handle_drop)

        update_init_progress(99, "最終調整中...")
        self.after(100, bring_copyright_to_front)

    def set_application_icon(self):
        """アプリケーションアイコンを設定"""
        try:
            if getattr(sys, 'frozen', False):
                application_path = sys._MEIPASS
                icon_files = [
                    os.path.join(application_path, 'pdf_palette_icon.ico'),
                    os.path.join(application_path, 'icon.ico'),
                    os.path.join(application_path, 'app.ico')
                ]
            else:
                icon_files = ['pdf_palette_icon.ico', 'icon.ico', 'app.ico']

            for icon_file in icon_files:
                if os.path.exists(icon_file):
                    try:
                        self.iconbitmap(default=icon_file)
                        self.iconbitmap(icon_file)
                        print(f"アイコン設定成功: {icon_file}")

                        try:
                            from PIL import Image
                            icon_pil = Image.open(icon_file)
                            icon_pil = icon_pil.resize((32, 32),
                                                       Image.Resampling.LANCZOS)
                            icon_photo = ImageTk.PhotoImage(icon_pil)
                            self.iconphoto(True, icon_photo)
                            self.icon_photo_ref = icon_photo
                        except Exception as photo_error:
                            print(f"PhotoImage変換エラー（無視可能）: {photo_error}")

                        return
                    except Exception as ico_error:
                        print(f"ICOファイル設定エラー: {ico_error}")
                        continue

            if getattr(sys, 'frozen', False):
                png_files = [
                    os.path.join(sys._MEIPASS, 'pdf_palette_icon.png'),
                    os.path.join(sys._MEIPASS, 'icon.png'),
                    os.path.join(sys._MEIPASS, 'app.png')
                ]
            else:
                png_files = ['pdf_palette_icon.png', 'icon.png', 'app.png']

            for png_file in png_files:
                if os.path.exists(png_file):
                    try:
                        icon_img = tk.PhotoImage(file=png_file)
                        self.iconphoto(True, icon_img)
                        self.png_icon_ref = icon_img
                        print(f"PNGアイコン設定成功: {png_file}")
                        return
                    except Exception as png_error:
                        print(f"PNGファイル設定エラー: {png_error}")
                        continue

            self.create_and_set_custom_icon()

        except Exception as e:
            print(f"カスタムアイコン作成エラー: {e}")

    def create_and_set_custom_icon(self):
        """カスタムアイコンを動的に作成して設定"""
        try:
            from PIL import Image, ImageDraw

            sizes = [16, 32]
            icon_photos = []

            for icon_size in sizes:
                icon = Image.new('RGBA', (icon_size, icon_size),
                                 (255, 255, 255, 0))
                draw = ImageDraw.Draw(icon)

                colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57',
                          '#FF9FF3']

                if icon_size == 16:
                    square_size = 4
                    positions = [(2, 2), (8, 2), (2, 8), (8, 8)]
                    text_pos = (1, 10)
                else:
                    square_size = 8
                    positions = [
                        (4, 4), (14, 4), (24, 4),
                        (4, 14), (14, 14), (24, 14)
                    ]
                    text_pos = (6, 24)

                for i, (x, y) in enumerate(positions):
                    color = colors[i % len(colors)]
                    rgb = tuple(int(color[j:j + 2], 16) for j in (1, 3, 5))
                    draw.rectangle([x, y, x + square_size, y + square_size],
                                   fill=rgb, outline=(100, 100, 100))

                if icon_size >= 24:
                    draw.text(text_pos, "PDF", fill=(50, 50, 50))

                icon_photo = ImageTk.PhotoImage(icon)
                icon_photos.append(icon_photo)

            if icon_photos:
                self.iconphoto(True, *icon_photos)
                self.custom_icon_refs = icon_photos
                print("カスタムアイコン作成・設定成功")

        except Exception as e:
            print(f"カスタムアイコン作成エラー: {e}")

    def handle_drop(self, event):
        file_paths = self.splitlist(event.data)
        update_main_message("PDFに変換中です...")

        # UI操作はメインスレッド（ここ）で実行する
        processing_label.config(text="PDFに変換中です...")
        processing_label.pack(side="left", padx=10)
        progress_frame.pack(side="bottom", fill="x", pady=5)
        progress_bar.config(value=0)

        def process_drop():
            process_files_async(file_paths)

        thread = Thread(target=process_drop, daemon=True)
        thread.start()


def show_splash_screen():
    """スプラッシュスクリーンを表示（画像使用版・進捗バー付き）"""
    root = tk.Tk()
    root.withdraw()

    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.attributes('-topmost', True)
    splash.lift()
    splash.focus_force()

    MAX_SPLASH_WIDTH = 600
    MAX_SPLASH_HEIGHT = 400

    progress_var = None
    progress_label = None
    progress_bar = None

    try:
        if getattr(sys, 'frozen', False):
            image_path = os.path.join(sys._MEIPASS, 'splash.png')
        else:
            image_path = 'splash.png'

        if os.path.exists(image_path):
            img = Image.open(image_path)

            original_width, original_height = img.size

            if original_width > MAX_SPLASH_WIDTH or original_height > MAX_SPLASH_HEIGHT:
                ratio = min(MAX_SPLASH_WIDTH / original_width,
                            MAX_SPLASH_HEIGHT / original_height)
                new_width = int(original_width * ratio)
                new_height = int(original_height * ratio)
                img = img.resize((new_width, new_height),
                                 Image.Resampling.LANCZOS)
            else:
                new_width, new_height = original_width, original_height

            photo = ImageTk.PhotoImage(img)

            screen_width = splash.winfo_screenwidth()
            screen_height = splash.winfo_screenheight()
            x = (screen_width - new_width) // 2
            y = (screen_height - new_height) // 2

            splash.geometry(f"{new_width}x{new_height}+{x}+{y}")

            if sys.platform == "win32":
                splash.attributes('-transparentcolor', 'white')
                splash.configure(bg='white')
            else:
                splash.configure(bg='white')

            image_label = tk.Label(splash, image=photo, bg='white', bd=0)
            image_label.image = photo
            image_label.place(x=0, y=0)

            progress_frame_height = 50
            progress_frame_width = new_width - 150
            progress_y = new_height - progress_frame_height - 20
            progress_x = 40

            bg_color = '#FDFDFD'

            bg_panel = tk.Frame(splash, bg=bg_color,
                                height=progress_frame_height,
                                width=progress_frame_width,
                                relief='flat',
                                bd=1,
                                highlightbackground='#B0D4FF',
                                highlightthickness=1)
            bg_panel.place(x=progress_x, y=progress_y)

            progress_frame = tk.Frame(splash, bg=bg_color,
                                      height=progress_frame_height,
                                      width=progress_frame_width)
            progress_frame.place(x=progress_x, y=progress_y)

            progress_label = tk.Label(
                progress_frame,
                text="初期化中...",
                font=("Meiryo", 7),
                bg=bg_color,
                fg='#000000'
            )
            progress_label.place(x=15, y=20, anchor='w')

            style = ttk.Style()
            style.theme_use('default')
            style.configure("Thin.Horizontal.TProgressbar", thickness=5)

            progress_var = tk.DoubleVar(value=0)
            progress_bar = ttk.Progressbar(
                progress_frame,
                variable=progress_var,
                maximum=100,
                length=progress_frame_width - 20,
                mode='determinate',
                style="Thin.Horizontal.TProgressbar"
            )
            progress_bar.place(x=15, y=30)

        else:
            splash_width = 500
            splash_height = 350
            screen_width = splash.winfo_screenwidth()
            screen_height = splash.winfo_screenheight()
            x = (screen_width - splash_width) // 2
            y = (screen_height - splash_height) // 2
            splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
            splash.configure(bg="#4A90E2")

            title_label = tk.Label(
                splash,
                text="PDFパレット",
                font=("Helvetica", 24, "bold"),
                bg="#4A90E2",
                fg="white"
            )
            title_label.pack(pady=(60, 10))

            version_label = tk.Label(
                splash,
                text="ver 1.0.1",
                font=("Helvetica", 12),
                bg="#4A90E2",
                fg="white"
            )
            version_label.pack(pady=(0, 20))

            loading_label = tk.Label(
                splash,
                text="起動中...",
                font=("Helvetica", 14),
                bg="#4A90E2",
                fg="white"
            )
            loading_label.pack(pady=10)

            progress_frame_widget = tk.Frame(splash, bg='#4A90E2')
            progress_frame_widget.pack(fill='x', padx=40, pady=10)

            progress_label = tk.Label(
                progress_frame_widget,
                text="初期化中...",
                font=("Helvetica", 10),
                bg='#4A90E2',
                fg="white"
            )
            progress_label.pack(pady=(0, 5))

            progress_var = tk.DoubleVar(value=0)
            progress_bar = ttk.Progressbar(
                progress_frame_widget,
                variable=progress_var,
                maximum=100,
                length=400,
                mode='determinate'
            )
            progress_bar.pack()

            copyright_label = tk.Label(
                splash,
                text="© Junichi Ono",
                font=("Helvetica", 9),
                bg="#4A90E2",
                fg="white"
            )
            copyright_label.pack(side="bottom", pady=10)

    except Exception as e:
        print(f"スプラッシュ画像読み込みエラー: {e}")

        splash_width = 500
        splash_height = 350
        screen_width = splash.winfo_screenwidth()
        screen_height = splash.winfo_screenheight()
        x = (screen_width - splash_width) // 2
        y = (screen_height - splash_height) // 2
        splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
        splash.configure(bg="#4A90E2")

        label = tk.Label(splash, text="PDFパレット\n起動中...",
                         font=("Helvetica", 18), bg="#4A90E2", fg="white")
        label.pack(expand=True)

    splash.update_idletasks()
    splash.update()

    return root, splash, progress_var, progress_label


if __name__ == "__main__":
    multiprocessing.freeze_support()

    splash_root, splash, progress_var, progress_label = show_splash_screen()

    app = None


    def update_progress(value, message):
        if progress_var:
            progress_var.set(value)
        if progress_label:
            progress_label.config(text=message)
        splash_root.update()


    def launch_main_app():
        global app

        try:
            update_progress(10, "初期化を開始...")

            update_progress(20, "パレットを準備しています...")
            app = DragDropApp(update_progress)

            update_progress(100, "さあ、始めましょう！")
            import time
            time.sleep(0.3)

            splash.destroy()
            splash_root.quit()
            splash_root.destroy()
        except Exception as e:
            print(f"起動エラー: {e}")
            import traceback
            traceback.print_exc()

        app.mainloop()


    splash_root.after(100, launch_main_app)

    splash_root.mainloop()
