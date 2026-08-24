# -*- coding: utf-8 -*-
"""
Конвертация старого формата .doc (Word 97-2003) в .docx через установленный
Microsoft Word (COM-автоматизация, требуется pywin32).

Зачем: apply_docx / find_replace работают только с .docx (это ZIP из XML), а
.doc — другой бинарный формат (OLE), его нельзя ни прочитать zipfile'ом, ни
безопасно править. Поэтому такие файлы («Приложение №…» и т.п.) сначала
пересохраняются Word'ом в .docx, а дальше обрабатываются как обычно.

pywin32 импортируется ЛЕНИВО (внутри функций), чтобы модуль импортировался и на
машинах без Word — там available() вернёт False, а обработка .docx продолжит
работать.
"""
import os

# wdFormatDocumentDefault = 16 — формат по умолчанию (.docx в Word 2007+).
WD_FORMAT_DOCX = 16


class ConversionUnavailable(RuntimeError):
    pass


def available():
    """True, если доступен Word через COM (установлен Word + есть pywin32)."""
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


class WordConverter:
    """Один экземпляр Word на пакет файлов (быстрее, чем открывать на каждый).
    Использовать в потоке, где создан (COM инициализируется в __init__)."""

    def __init__(self):
        try:
            import win32com.client as client
            import pythoncom
            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
            self._word = client.DispatchEx('Word.Application')
            self._word.Visible = False
            self._word.DisplayAlerts = False
        except Exception as e:
            raise ConversionUnavailable(str(e))

    def convert(self, doc_path):
        """Конвертирует один .doc в .docx рядом (тот же каталог и имя).
        Если .docx уже существует — не трогает, просто возвращает его путь.
        Возвращает путь к .docx."""
        doc_path = os.path.abspath(doc_path)
        target = os.path.splitext(doc_path)[0] + '.docx'
        if os.path.exists(target):
            return target
        doc = self._word.Documents.Open(doc_path, ReadOnly=True)
        try:
            try:
                doc.SaveAs2(target, FileFormat=WD_FORMAT_DOCX)
            except Exception:
                doc.SaveAs(target, FileFormat=WD_FORMAT_DOCX)  # Word < 2010
        finally:
            doc.Close(False)
        return target

    def close(self):
        try:
            self._word.Quit()
        except Exception:
            pass
        try:
            self._pythoncom.CoUninitialize()
        except Exception:
            pass
