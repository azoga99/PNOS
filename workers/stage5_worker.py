import os
import time
import pythoncom
import win32com.client
import tempfile
import re
from PIL import Image
from PySide6.QtCore import QThread, Signal

class Stage5Worker(QThread):
    """Этап 6 (был 5): Вставка картинок в Отчет.docx на основе цвета."""
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(bool)
    info = Signal(str, str)
    action_update = Signal(str)
    report_ready = Signal(list)

    def __init__(self, local_root: str, parent=None):
        super().__init__(parent)
        self.local_root = local_root
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def analyze_color(self, img_path):
        """Определяет доминантный цвет на схеме (красный/синий)."""
        try:
            with Image.open(img_path) as img:
                img = img.convert("RGB")
                img.thumbnail((250, 250)) # Чуть крупнее для тонких линий
                pixels = list(img.getdata())
                
                red_pixels = 0
                blue_pixels = 0
                
                for r, g, b in pixels:
                    # Игнорируем почти белые и серые пиксели
                    if r > 230 and g > 230 and b > 230: continue
                    # Игнорируем слишком темные
                    if r < 40 and g < 40 and b < 40: continue
                    
                    # Проверяем на цветовое преобладание
                    # Красный: R значительно больше G и B (увеличили чувствительность до 15%)
                    if r > g * 1.15 and r > b * 1.15:
                        red_pixels += 1
                    # Синий: B значительно больше R и G
                    elif b > r * 1.15 and b > g * 1.15:
                        blue_pixels += 1
                
                # Если цветных пикселей совсем мало - значит схема ч/б или нейтральная
                if (red_pixels + blue_pixels) < 10:
                    return "other"
                
                if red_pixels > blue_pixels * 1.15:
                    return "red"
                if blue_pixels > red_pixels * 1.15:
                    return "blue"
                    
                return "other"
        except Exception as e:
            print(f"Error analyzing {img_path}: {e}")
            return "other"

    def run(self):
        self.log.emit("═" * 40)
        self.log.emit("ЭТАП 6: Вставка картинок (на основе цвета)")
        self.log.emit("═" * 40)

        if not os.path.isdir(self.local_root):
            self.log.emit("❌ Корневая папка не найдена.")
            self.finished_ok.emit(False)
            return

        point_folders = [os.path.join(self.local_root, d) for d in os.listdir(self.local_root) 
                        if os.path.isdir(os.path.join(self.local_root, d)) and d.startswith("п.")]

        if not point_folders:
            self.log.emit("⚠ Папки пунктов не найдены.")
            self.finished_ok.emit(False)
            return

        pythoncom.CoInitialize()
        wd_app = None
        try:
            wd_app = win32com.client.DispatchEx("Word.Application")
            wd_app.Visible = False
            wd_app.DisplayAlerts = 0

            processed_count = 0
            report_data = []

            for i, p_folder in enumerate(point_folders):
                if self._is_cancelled: break
                
                p_name = os.path.basename(p_folder)
                pt_num = p_name[2:] if p_name.startswith("п.") else p_name
                self.info.emit(f"Обработка фото для {p_name}...", "wait")
                self.log.emit(f"\n── {p_name} ──")
                
                point_res = {
                    "№ Пункта": pt_num,
                    "Фото": "0",
                    "Результат": "Ожидание"
                }

                # 1. Поиск фото
                pervichka = os.path.join(p_folder, "Первичка")
                photos = []
                if os.path.exists(pervichka):
                    self.log.emit(f"   📂 Поиск фото в: {pervichka}")
                    for root, dirs, files in os.walk(pervichka):
                        for f in files:
                            if f.lower().endswith((".jpg", ".jpeg", ".png")):
                                photos.append(os.path.join(root, f))
                else:
                    self.log.emit(f"   ⚠ Папка 'Первичка' не найдена в {p_name}")
                
                if not photos:
                    self.log.emit("   ⚠ Фото в папке 'Первичка' не найдены.")
                    point_res["Результат"] = "⚠️ Нет фото"
                    report_data.append(point_res)
                    continue

                # Сортируем фото по суффиксу _1, _2 и т.д.
                def sort_key(fpath):
                    m = re.search(r'_(\d+)\.', fpath)
                    return int(m.group(1)) if m else 0
                photos.sort(key=sort_key)

                # 2. Анализ цветов
                red_list = []
                blue_list = []
                for p in photos:
                    color = self.analyze_color(p)
                    if color == "red": red_list.append(p)
                    elif color == "blue": blue_list.append(p)
                
                self.log.emit(f"   📊 Найдено: Красных={len(red_list)}, Синих={len(blue_list)}")
                point_res["Фото"] = f"К:{len(red_list)}, С:{len(blue_list)}"

                # 3. Поиск файла отчета
                report_file = None
                for f in os.listdir(p_folder):
                    if f.startswith("Отчет_") and f.endswith(".docx"):
                        report_file = os.path.join(p_folder, f)
                        break
                
                if not report_file:
                    self.log.emit("   ❌ Файл 'Отчёт_*.docx' не найден. Сначала выполните Этап 5.")
                    point_res["Результат"] = "❌ Нет отчета"
                    report_data.append(point_res)
                    continue

                # 4. Вставка в Word
                doc = None
                try:
                    doc = wd_app.Documents.Open(report_file)
                    
                    def process_marker(photo_list, marker):
                        # Находим маркер
                        find_range = doc.Content
                        find_range.Find.Execute(FindText=marker)
                        if not find_range.Find.Found:
                            self.log.emit(f"   ⚠ Маркер {marker} не найден.")
                            return 0

                        if not photo_list:
                            # Если фото нет, просто удаляем маркер
                            find_range.Text = ""
                            return 0

                        self.log.emit(f"   📝 Обработка {marker} ({len(photo_list)} фото)...")
                        
                        # 1. Определяем диапазон страницы, где находится маркер
                        marker_range = find_range.Duplicate
                        marker_range.Select()
                        page_range = wd_app.Selection.Bookmarks("\\page").Range
                        
                        # Копируем страницу как эталон
                        page_range.Copy()
                        
                        # 2. Дублируем страницу, если фото больше одного
                        insert_pos = page_range.End
                        for _ in range(len(photo_list) - 1):
                            r_paste = doc.Range(insert_pos, insert_pos)
                            r_paste.Paste()
                            insert_pos = r_paste.End

                        # 3. Заменяем маркеры на фото
                        inserted = 0
                        for img_path in photo_list:
                            fresh_find = doc.Content
                            fresh_find.Find.Execute(FindText=marker)
                            if fresh_find.Find.Found:
                                try:
                                    anchor_range = fresh_find.Duplicate
                                    # Меняем на пробел, чтобы не удалять абзац и не смещать заголовки
                                    anchor_range.Text = " " 
                                    anchor_range.Collapse(1) 
                                    # Читаем размеры из PIL и считаем точные пропорции
                                    # Максимальные габариты внутри рамки
                                    max_w = 520.0
                                    max_h = 750.0
                                    
                                    orig_w, orig_h = 0, 0
                                    
                                    # Переворачиваем картинку на 90 градусов через PIL (в Temp)
                                    temp_path = os.path.join(tempfile.gettempdir(), f"rot_{os.path.basename(img_path)}")
                                    with Image.open(img_path) as im:
                                        # Поворот на 90 градусов
                                        im_rotated = im.transpose(Image.Transpose.ROTATE_90)
                                        orig_w, orig_h = im_rotated.size
                                        im_rotated.save(temp_path, format=im.format or "PNG")
                                    
                                    # Считаем точные размеры для вставки, жестко сохраняя пропорции
                                    ratio = min(max_w / orig_w, max_h / orig_h)
                                    target_w = orig_w * ratio
                                    target_h = orig_h * ratio
                                    
                                    # Вставляем как ПЛАВАЮЩИЙ объект СРАЗУ с рассчитанными размерами
                                    shape = doc.Shapes.AddPicture(temp_path, False, True, 0, 0, target_w, target_h, anchor_range)
                                    
                                    # Настройка обтекания (поверх текста)
                                    shape.WrapFormat.Type = 3 # wdWrapFront
                                    shape.LockAspectRatio = True # На всякий случай блокируем

                                    # Центрирование относительно ПОЛЕЙ СТРАНИЦЫ (Margin = 0)
                                    # Это учтет ваши 2см слева и 0.5см справа
                                    shape.RelativeHorizontalPosition = 0 # wdRelativeHorizontalPositionMargin
                                    shape.Left = -999995 # wdShapeCenter
                                    
                                    shape.RelativeVerticalPosition = 0 # wdRelativeVerticalPositionMargin
                                    shape.Top = -999995 # wdShapeCenter
                                    
                                    inserted += 1
                                    self.log.emit(f"   ✓ Вставлено и перевернуто: {os.path.basename(img_path)}")
                                except Exception as e:
                                    self.log.emit(f"   ❌ Ошибка вставки: {e}")
                            else:
                                self.log.emit("   ⚠ Не удалось найти область для очередного фото.")
                        
                        # Подчищаем хвосты (убираем лишние маркеры, если они остались пустые)
                        while True:
                            cleanup = doc.Content
                            if cleanup.Find.Execute(FindText=marker):
                                cleanup.Text = " "
                            else:
                                break
                                
                        return inserted

                    r_in = process_marker(red_list, "[ФОТО1]")
                    b_in = process_marker(blue_list, "[ФОТО2]")

                    doc.Save()
                    doc.Close()
                    doc = None

                    # 5. Очистка таблиц (Удаление пустых строк и Задвижек) через python-docx
                    self.log.emit("   🧹 Очистка таблиц ультразвуковой толщинометрии...")
                    import docx
                    from docx.oxml.table import CT_Tbl
                    from docx.oxml.text.paragraph import CT_P
                    from docx.table import Table
                    from docx.text.paragraph import Paragraph
                    
                    try:
                        clean_doc = docx.Document(report_file)
                        
                        def iter_block_items(parent):
                            parent_elm = parent.element.body if isinstance(parent, docx.document.Document) else parent._element
                            for child in parent_elm.iterchildren():
                                if isinstance(child, CT_P):
                                    yield Paragraph(child, parent)
                                elif isinstance(child, CT_Tbl):
                                    yield Table(child, parent)

                        target_table = None
                        waiting_for_table = False
                        
                        for block in iter_block_items(clean_doc):
                            if isinstance(block, Paragraph):
                                if "Таблица результатов ультразвуковой толщинометрии" in block.text:
                                    waiting_for_table = True
                            elif isinstance(block, Table) and waiting_for_table:
                                target_table = block
                                break
                        
                        if target_table:
                            rows_deleted = 0
                            # Идем с конца, чтобы не сбить индексы при удалении XML узлов
                            for row in reversed(target_table.rows[1:]): # пропускаем первую строку (заголовок)
                                is_empty = True
                                # Проверяем ячейки с 6 по 12 (индексы 5-11, счет с 0)
                                if len(row.cells) > 5:
                                    for col_idx in range(5, 12):
                                        if col_idx < len(row.cells):
                                            if row.cells[col_idx].text.strip() != "":
                                                is_empty = False
                                                break
                                else:
                                    is_empty = False # если столбцов меньше, не трогаем
                                
                                # Проверяем Задвижку (3 столбец, индекс 2)
                                is_zadvizhka = False
                                if len(row.cells) > 2 and "задвижка" in row.cells[2].text.lower():
                                    is_zadvizhka = True
                                
                                # Удаляем из дерева xml
                                if is_empty or is_zadvizhka:
                                    tbl = target_table._element
                                    tbl.remove(row._tr)
                                    rows_deleted += 1
                                else:
                                    # Задаем фиксированную высоту оставшимся строкам (0.45 см)
                                    row.height = docx.shared.Cm(0.45)
                                    row.height_rule = docx.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY
                            
                            if rows_deleted > 0:
                                self.log.emit(f"   ✓ Из таблицы удалено строк: {rows_deleted}")
                            else:
                                self.log.emit("   ✓ Таблица чистая, удалять нечего.")
                            
                            # Сохраняем изменения поверх файла отчета
                            clean_doc.save(report_file)
                        else:
                            self.log.emit("   ⚠ Таблица результатов ультразвуковой толщинометрии не найдена.")
                            
                    except Exception as clean_ex:
                        self.log.emit(f"   ❌ Ошибка при очистке таблиц: {clean_ex}")

                    total = r_in + b_in
                    self.log.emit(f"   ✅ Отчет {p_name} готов. ФОТО: {total}")
                    point_res["Результат"] = "✅ Успешно" if total > 0 else "⚠️ Пропущено"
                    processed_count += 1
                except Exception as ex:
                    self.log.emit(f"   ❌ Ошибка Word: {ex}")
                    point_res["Результат"] = "❌ Ошибка Word"
                    if doc: doc.Close(False)
                
                report_data.append(point_res)
                prog_val = 10 + int((i + 1) / len(point_folders) * 90)
                self.progress.emit(prog_val)

            self.report_ready.emit(report_data)
            self.progress.emit(100)
            self.finished_ok.emit(True)
        except Exception as e:
            self.log.emit(f"❌ Критическая ошибка: {e}")
            self.finished_ok.emit(False)
        finally:
            if wd_app: wd_app.Quit()
            pythoncom.CoUninitialize()
