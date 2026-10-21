"""Regressões de identidade usando arquivos XLSX serializados somente em memória."""

from io import BytesIO
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pdf_import
from main import PortalApp
from storage import ExcelStorageError, StudentStorage, STUDENT_COLUMNS, STUDENT_SHEET


class MemoryStudentStorage(StudentStorage):
    def __init__(self):
        super().__init__("not_written.xlsx")
        self.contents = None

    def _open(self):
        openpyxl, Alignment, Font, PatternFill = self._load_openpyxl()
        if self.contents is None:
            wb = openpyxl.Workbook()
            wb.active.title = STUDENT_SHEET
        else:
            wb = openpyxl.load_workbook(BytesIO(self.contents))
        self._ensure_sheet(wb, STUDENT_SHEET, STUDENT_COLUMNS, Font, PatternFill, Alignment)
        return wb

    def _safe_save(self, wb, path):
        output = BytesIO()
        wb.save(output)
        self.contents = output.getvalue()
        wb.close()


def student(student_id="A", cpf="11111111111", **fields):
    return {"id_aluno": student_id, "cpf": cpf, "nome_aluno": "Aluno de teste", **fields}


def extracted_records(records):
    reader = SimpleNamespace(is_encrypted=False, pages=[Mock(extract_text=Mock(return_value="Relatório"))])
    with patch("pypdf.PdfReader", return_value=reader), patch.object(pdf_import, "_extract_page", return_value=records):
        return pdf_import.extract_students_data("not_read.pdf")


class StudentImportTests(unittest.TestCase):
    def setUp(self):
        self.db = MemoryStudentStorage()
        self.db.save_students([
            student(ids_processos="P-ANTIGO", email="teste@example.invalid"),
            student("B", "22222222222"),
        ])

    def assert_rejected_without_changes(self, records):
        before = self.db.contents
        with self.assertRaises(ExcelStorageError):
            self.db.save_students(records)
        self.assertEqual(self.db.contents, before)

    def test_crossed_id_and_cpf_do_not_overwrite_either_student(self):
        self.assert_rejected_without_changes([student("B", "11111111111")])
        self.assertEqual(self.db.get_student("A")["cpf"], "11111111111")
        self.assertEqual(self.db.get_student("B")["cpf"], "22222222222")

    def test_existing_id_cannot_replace_its_cpf(self):
        self.assert_rejected_without_changes([student("A", "33333333333")])

    def test_new_id_for_existing_cpf_preserves_canonical_id_and_links(self):
        returned = self.db.save_students([student("EXTERNO", "111.111.111-11", nome_aluno="Nome atualizado")])
        self.assertEqual(returned, ["A"])
        self.assertEqual(len(self.db.list_students()), 2)
        saved = self.db.get_student("A")
        self.assertEqual(saved["nome_aluno"], "Nome atualizado")
        self.assertEqual(saved["ids_processos"], "P-ANTIGO")
        self.assertEqual(saved["email"], "teste@example.invalid")
        self.assertIsNone(self.db.get_student("EXTERNO"))

    def test_missing_id_reuses_cpf_match(self):
        self.assertEqual(self.db.save_students([student("", "111.111.111-11")]), ["A"])
        self.assertEqual(len(self.db.list_students()), 2)

    def test_id_case_and_cpf_punctuation_are_ignored(self):
        self.assertEqual(self.db.save_students([student(" a ", "111.111.111-11")]), ["A"])
        self.assertEqual(len(self.db.list_students()), 2)

    def test_missing_cpf_does_not_clear_existing_cpf(self):
        self.db.save_students([student("A", "", nome_aluno="Nome atualizado")])
        self.assertEqual(self.db.get_student("A")["cpf"], "11111111111")

    def test_empty_existing_cpf_can_be_filled(self):
        self.db.save_students([student("C", "")])
        self.db.save_students([student("C", "33333333333")])
        self.assertEqual(self.db.get_student("C")["cpf"], "33333333333")

    def test_one_source_id_with_different_cpfs_rejects_whole_batch(self):
        self.assert_rejected_without_changes([
            student("NOVO", "44444444444"),
            student("EXTERNO", "11111111111"),
            student(" externo ", "33333333333"),
        ])

    def test_late_conflict_does_not_save_earlier_valid_rows(self):
        self.assert_rejected_without_changes([
            student("NOVO", "44444444444"),
            student("B", "11111111111"),
        ])
        self.assertIsNone(self.db.get_student("NOVO"))

    def test_repeated_aliases_return_unique_canonical_ids(self):
        returned = self.db.save_students([
            student("EXTERNO-1", "11111111111"),
            student("EXTERNO-2", "111.111.111-11"),
            student("A", "11111111111"),
        ])
        self.assertEqual(returned, ["A"])
        self.assertEqual(len(self.db.list_students()), 2)

    def test_matching_historical_duplicates_are_blocked(self):
        for duplicate, incoming in [
            (student("C", "11111111111"), student()),
            (student("C", "11111111111"), student("A", "")),
            (student("A", "33333333333"), student()),
            (student("A", "33333333333"), student("EXTERNO", "11111111111")),
        ]:
            with self.subTest(duplicate=duplicate, incoming=incoming):
                db = MemoryStudentStorage()
                db.save_students([student()])
                wb = db._open()
                wb[STUDENT_SHEET].append([duplicate.get(field, "") for _, field in STUDENT_COLUMNS])
                db._safe_save(wb, db.path)
                before = db.contents
                with self.assertRaises(ExcelStorageError):
                    db.save_students([incoming])
                self.assertEqual(db.contents, before)

    def test_generated_ids_do_not_collide_with_later_source_ids(self):
        db = MemoryStudentStorage()
        returned = db.save_students([
            student("", "11111111111"),
            student("ALU-000001", "22222222222"),
            student("", ""),
        ])
        self.assertEqual(len(set(returned)), 3)
        self.assertEqual(len(db.list_students()), 3)
        self.assertEqual(db.get_student("ALU-000001")["cpf"], "22222222222")

    def test_cpf_leading_zeroes_are_preserved(self):
        self.db.save_students([student("C", "001.111.111-11")])
        self.assertEqual(self.db.save_students([student("EXTERNO", "00111111111")]), ["C"])
        self.assertEqual(self.db.get_student("C")["cpf"], "00111111111")

    def test_manual_student_save_returns_canonical_id(self):
        self.assertEqual(self.db.save_student(student("EXTERNO", "11111111111")), "A")

    def test_process_link_uses_existing_id(self):
        payload = student("EXTERNO", "11111111111")
        before = self.db.contents
        self.assertEqual(self.db.resolve_student_id(payload), "A")
        self.assertEqual(self.db.contents, before)
        self.assertEqual(self.db.save_process_student(payload, "P-NOVO"), "A")
        self.assertEqual(self.db.get_student("A")["ids_processos"], "P-ANTIGO; P-NOVO")
        self.assertIsNone(self.db.get_student("EXTERNO"))


class PDFIdentityTests(unittest.TestCase):
    def test_different_ids_with_same_cpf_reach_storage_validation(self):
        db = MemoryStudentStorage()
        db.save_students([student(), student("B", "22222222222")])
        before = db.contents
        records = extracted_records([student("EXTERNO", "11111111111"), student("B", "11111111111")])
        self.assertEqual(len(records), 2)
        with self.assertRaises(ExcelStorageError):
            db.save_students(records)
        self.assertEqual(db.contents, before)

    def test_repeated_pair_merges_complementary_fields(self):
        records = extracted_records([
            student("A", "111.111.111-11", rg="123456"),
            student("a", "11111111111", email="teste@example.invalid"),
        ])
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["rg"], "123456")
        self.assertEqual(records[0]["email"], "teste@example.invalid")

    def test_conflicting_id_on_different_pages_is_not_hidden(self):
        records = extracted_records([student(), student("A", "22222222222")])
        db = MemoryStudentStorage()
        with self.assertRaises(ExcelStorageError):
            db.save_students(records)
        self.assertIsNone(db.contents)

    def test_numeric_id_is_not_confused_with_another_students_cpf(self):
        records = extracted_records([student("11111111111", ""), student("B", "11111111111")])
        self.assertEqual(len(records), 2)


class ProcessIdentityTests(unittest.TestCase):
    def setUp(self):
        self.db = MemoryStudentStorage()
        self.db.save_students([student(), student("B", "22222222222")])

    def app(self, payload):
        app = SimpleNamespace(
            root=None,
            data_ready=True,
            current_process_id=None,
            student_storage=self.db,
            storage=Mock(),
            validate_duplicate_offer=Mock(get=Mock(return_value=True)),
            _collect_form=lambda: {"id_processo": "P-NOVO", **payload},
            _show_storage_error=Mock(),
            _set_form=Mock(),
            _refresh_process_selector=Mock(),
            _refresh_student_process_selector=Mock(),
            _update_request_preview=Mock(),
            _load_equivalences=Mock(),
        )
        app.storage.list_processes.return_value = []
        return app

    def test_conflict_is_rejected_before_process_is_written(self):
        app = self.app(student("B", "11111111111"))
        before = self.db.contents
        PortalApp.save_process(app)
        app.storage.save_process.assert_not_called()
        app._show_storage_error.assert_called_once()
        self.assertEqual(self.db.contents, before)

    def test_process_and_student_link_share_canonical_id(self):
        app = self.app(student("EXTERNO", "11111111111"))
        with patch("main.messagebox.showinfo"):
            PortalApp.save_process(app)
        app._show_storage_error.assert_not_called()
        saved = app.storage.save_process.call_args.args[0]
        self.assertEqual(saved["id_aluno"], "A")
        self.assertEqual(self.db.get_student("A")["ids_processos"], "P-NOVO")
        self.assertIsNone(self.db.get_student("EXTERNO"))

    def test_duplicate_offer_uses_resolved_student_id(self):
        app = self.app(student("EXTERNO", "11111111111", codigo_oferta="OFERTA"))
        app.storage.list_processes.return_value = [
            {"id_processo": "P-ANTIGO", "id_aluno": "A", "codigo_oferta": "OFERTA"}
        ]
        with patch("main.messagebox.askyesno", return_value=False) as question:
            PortalApp.save_process(app)
        question.assert_called_once()
        app.storage.save_process.assert_not_called()

    def test_new_student_with_only_cpf_receives_the_process_student_id(self):
        app = self.app(student("", "33333333333"))
        with patch("main.messagebox.showinfo"):
            PortalApp.save_process(app)
        app._show_storage_error.assert_not_called()
        saved = app.storage.save_process.call_args.args[0]
        self.assertTrue(saved["id_aluno"])
        self.assertEqual(self.db.get_student(saved["id_aluno"])["cpf"], "33333333333")


if __name__ == "__main__":
    unittest.main()
