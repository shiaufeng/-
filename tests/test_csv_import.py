import copy
import inspect
import io
from pathlib import Path
import unittest
import zipfile
from datetime import date
from unittest.mock import patch

from openpyxl import Workbook

import server


VALID_CSV = (
    '\ufeff姓名,手機,Email,公司名稱,職稱,地區,職階,桌號,備註\r\n'
    '王小明,0912345678,ming@example.com,範例公司,經理,台北,中高階,第12桌,"素食,不吃香菜"\r\n'
    ',,,,,,,,只有備註會被略過\r\n'
    '陳美玲,0987654321,mei@example.com,測試科技,專員,新竹,基層,A桌,\r\n'
).encode('utf-8')


def make_xlsx(headers=None, rows=None, sheet_name='報到名單', extra_sheet=None):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers or ['姓名', '手機', 'Email', '公司名稱', '職稱', '地區', '職階', '桌號', '備註'])
    for row in rows or [
        ['王小明', '0912345678', 'ming@example.com', '範例公司', '經理', '台北', '中高階', '第12桌', '素食'],
        ['陳美玲', '0987654321', 'mei@example.com', '測試科技', '專員', '新竹', '基層', 'A桌', ''],
    ]:
        worksheet.append(row)
    if extra_sheet:
        extra = workbook.create_sheet(extra_sheet)
        extra.append(headers or ['姓名', '手機', '公司名稱'])
        extra.append(['第二張', '0900000000', '第二公司'])
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def roster_row(row_id=1, status='pending'):
    return {
        'id': row_id,
        'admin_username': 'admin',
        'google_sheet_name': '活動A',
        'name': f'既有名單{row_id}',
        'phone': f'090000000{row_id}',
        'company': '既有公司',
        'status': status,
        'updated_at': '2026-08-12 10:00:00',
    }


class FakeDatabase:
    def __init__(self, roster=None, allowed_events='活動A'):
        self.roster = copy.deepcopy(roster or [])
        self.allowed_events = allowed_events
        self.admin_exists = True
        self.previews = {}
        self.connections = []
        self.fail_on_insert = None
        self.fail_on_snapshot = False

    def connect(self):
        connection = FakeConnection(self)
        self.connections.append(connection)
        return connection


class FakeCursor:
    def __init__(self, connection):
        self.connection = connection
        self.db = connection.db
        self._one = None
        self._all = []
        self.rowcount = 0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def execute(self, sql, args=()):
        compact_sql = ' '.join(sql.split())
        self.connection.statements.append((compact_sql, args))
        self._one = None
        self._all = []
        self.rowcount = 0

        if compact_sql.startswith('SELECT allowed_events FROM admins'):
            self._one = {'allowed_events': self.db.allowed_events} if self.db.admin_exists else None
            return

        if compact_sql.startswith('SELECT id FROM event_configs'):
            self._one = {'id': 1}
            return

        if compact_sql.startswith('SELECT allowed_events, current_event FROM admins'):
            self._one = {
                'allowed_events': self.db.allowed_events,
                'current_event': '活動A',
            } if self.db.admin_exists else None
            return

        if compact_sql.startswith('SELECT DISTINCT google_sheet_name FROM event_registrations'):
            self._all = [
                {'google_sheet_name': sheet}
                for sheet in sorted({row.get('google_sheet_name') for row in self.db.roster if row.get('google_sheet_name')})
            ]
            return

        if compact_sql.startswith('SELECT DISTINCT google_sheet_name FROM event_configs'):
            self._all = []
            return

        if compact_sql.startswith('UPDATE admins SET current_event'):
            _sheet, allowed_events, _admin = args
            self.db.allowed_events = allowed_events
            self.rowcount = 1
            return

        if compact_sql.startswith('SELECT * FROM event_registrations'):
            if self.db.fail_on_snapshot:
                raise RuntimeError('simulated snapshot failure')
            self._all = copy.deepcopy(self.db.roster)
            return

        if compact_sql.startswith('UPDATE csv_import_previews SET used_at=NOW() WHERE admin_username'):
            admin, sheet = args
            for record in self.db.previews.values():
                if record['admin_username'] == admin and record['google_sheet_name'] == sheet and not record['used_at']:
                    record['used_at'] = True
                    self.rowcount += 1
            return

        if compact_sql.startswith('DELETE FROM csv_import_previews WHERE expires_at'):
            return

        if compact_sql.startswith('INSERT INTO csv_import_previews'):
            (
                token_hash, admin, sheet, file_sha256, normalized_rows_sha256,
                source_format, worksheet_name, parser_version, valid_count,
                roster_count, revision, _ttl,
            ) = args
            self.db.previews[token_hash] = {
                'admin_username': admin,
                'google_sheet_name': sheet,
                'file_sha256': file_sha256,
                'normalized_rows_sha256': normalized_rows_sha256,
                'source_format': source_format,
                'worksheet_name': worksheet_name,
                'parser_version': parser_version,
                'valid_count': valid_count,
                'roster_count': roster_count,
                'roster_revision': revision,
                'used_at': None,
            }
            self.rowcount = 1
            return

        if compact_sql.startswith('SELECT admin_username, google_sheet_name, file_sha256'):
            record = self.db.previews.get(args[0])
            self._one = copy.deepcopy(record) if record and not record['used_at'] else None
            return

        if compact_sql.startswith('UPDATE csv_import_previews SET used_at=NOW() WHERE token_hash'):
            record = self.db.previews.get(args[0])
            if record and not record['used_at']:
                record['used_at'] = True
                self.rowcount = 1
            return

        if compact_sql.startswith('DELETE FROM event_registrations'):
            self.db.roster = []
            self.rowcount = 1
            return

        if compact_sql.startswith('INSERT INTO event_registrations'):
            self.connection.insert_count += 1
            if self.db.fail_on_insert == self.connection.insert_count:
                raise RuntimeError('simulated insert failure')
            self.db.roster.append({
                'id': self.connection.insert_count,
                'admin_username': args[0],
                'google_sheet_name': args[1],
                'name': args[4],
                'phone': args[5],
                'company': args[7],
                'status': 'pending',
            })
            self.rowcount = 1
            return

        raise AssertionError(f'Unexpected SQL in test: {compact_sql}')

    def fetchone(self):
        return self._one

    def fetchall(self):
        return self._all


class FakeConnection:
    def __init__(self, db):
        self.db = db
        self.statements = []
        self.commits = 0
        self.rollbacks = 0
        self.closes = 0
        self.insert_count = 0
        self.last_insert_id = 0
        self._snapshot = (copy.deepcopy(db.roster), copy.deepcopy(db.previews))

    def cursor(self):
        return FakeCursor(self)

    def commit(self):
        self.commits += 1
        self._snapshot = (copy.deepcopy(self.db.roster), copy.deepcopy(self.db.previews))

    def rollback(self):
        self.rollbacks += 1
        self.db.roster, self.db.previews = copy.deepcopy(self._snapshot)

    def close(self):
        self.closes += 1


class CsvImportTests(unittest.TestCase):
    def setUp(self):
        server.app.config.update(TESTING=True, SECRET_KEY='csv-import-test-secret')
        self.client = server.app.test_client()
        self.log_in()

    def log_in(self, username='admin'):
        with self.client.session_transaction() as user_session:
            user_session['admin_logged_in'] = True
            user_session['username'] = username
            user_session['allowed_sheets'] = ['活動A']
            user_session['current_admin_sheet'] = '活動A'

    @staticmethod
    def upload_data(raw=VALID_CSV, token=None, filename='people.csv'):
        data = {'file': (io.BytesIO(raw), filename)}
        if token is not None:
            data['preview_token'] = token
        return data

    def post_import(self, db, mode='preview', raw=VALID_CSV, token=None, admin='admin', sheet='活動A', filename='people.csv'):
        with (
            patch.object(server, 'get_db_connection', side_effect=db.connect),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'ensure_config'),
        ):
            return self.client.post(
                f'/api/sheets/import_csv?mode={mode}&admin={admin}&sheet={sheet}',
                data=self.upload_data(raw, token, filename),
                content_type='multipart/form-data',
            )

    def preview(self, db, raw=VALID_CSV, filename='people.csv'):
        response = self.post_import(db, raw=raw, filename=filename)
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        return response.get_json()

    def test_parser_normalizes_headers_and_counts_only_valid_rows(self):
        parsed = server.parse_registration_csv(VALID_CSV)
        self.assertEqual(parsed['total_count'], 3)
        self.assertEqual(parsed['valid_count'], 2)
        self.assertEqual(parsed['skipped_count'], 1)
        self.assertEqual(parsed['rows'][0]['seat'], '12')
        self.assertEqual(parsed['rows'][0]['special_notes'], '素食,不吃香菜')

    def test_parser_imports_industry_and_meal_from_csv_and_xlsx(self):
        csv_raw = (
            '姓名,手機,公司名稱,產業類別,葷素,桌次\r\n'
            '王小明,0912345678,範例公司,資訊科技,素,第12桌\r\n'
            '林美華,0922000111,另一公司,專業服務,,第8桌\r\n'
        ).encode('utf-8')
        csv_rows = server.parse_registration_csv(csv_raw)['rows']
        csv_row = csv_rows[0]
        self.assertEqual(csv_row['industry_category'], '資訊科技')
        self.assertEqual(csv_row['meal_preference'], '素食')
        self.assertEqual(csv_row['seat'], '12')
        self.assertEqual(csv_rows[1]['meal_preference'], '不用餐')

        xlsx_raw = make_xlsx(
            headers=['姓名', '手機', '公司名稱', '產業', '餐飲偏好', '桌號'],
            rows=[['陳美玲', '0987654321', '測試科技', '專業服務', '蛋奶素', 'A桌']],
        )
        xlsx_row = server.parse_registration_upload(xlsx_raw, 'people.xlsx')['rows'][0]
        self.assertEqual(xlsx_row['industry_category'], '專業服務')
        self.assertEqual(xlsx_row['meal_preference'], '蛋奶素')
        self.assertEqual(xlsx_row['seat'], 'A')

    def test_parser_accepts_spaced_headers_cp950_and_export_headers(self):
        spaced = ' 姓名 , 手機 ,公司名稱\r\n王小明,0912345678,範例公司\r\n'.encode('cp950')
        self.assertEqual(server.parse_registration_csv(spaced)['valid_count'], 1)

        exported = (
            '姓名,手機,公司/單位,Email,職稱,桌號/座位,報到狀態,報到時間,肖像權狀態,代理人姓名,代理人手機,備註\r\n'
            '陳美玲,0987654321,測試科技,mei@example.com,專員,A桌,pending,,,,,輪椅席\r\n'
        ).encode('utf-8')
        result = server.parse_registration_csv(exported)
        self.assertEqual(result['rows'][0]['company'], '測試科技')
        self.assertEqual(result['rows'][0]['seat'], 'A')

    def test_xlsx_parser_reads_unicode_sheet_and_preserves_text_phone(self):
        raw = make_xlsx(sheet_name='貴賓名單')
        parsed = server.parse_registration_upload(raw, 'people.xlsx')
        self.assertEqual(parsed['source_format'], 'xlsx')
        self.assertEqual(parsed['worksheet_name'], '貴賓名單')
        self.assertEqual(parsed['valid_count'], 2)
        self.assertEqual(parsed['rows'][0]['phone'], '0912345678')
        self.assertEqual(parsed['rows'][0]['seat'], '12')

    def test_xlsx_preview_and_commit_share_safe_import_flow(self):
        raw = make_xlsx(sheet_name='貴賓名單')
        db = FakeDatabase(roster=[roster_row(1)])
        preview = self.preview(db, raw, 'people.xlsx')
        self.assertEqual(preview['source_format'], 'xlsx')
        self.assertEqual(preview['worksheet_name'], '貴賓名單')
        self.assertEqual(preview['valid_count'], 2)
        token_record = db.previews[server.csv_preview_token_hash(preview['preview_token'])]
        self.assertEqual(token_record['source_format'], 'xlsx')
        self.assertEqual(token_record['worksheet_name'], '貴賓名單')
        self.assertEqual(token_record['parser_version'], server.REGISTRATION_PARSER_VERSION)
        self.assertEqual(token_record['normalized_rows_sha256'], server.registration_rows_hash(
            server.parse_registration_upload(raw, 'people.xlsx')['rows']
        ))

        response = self.post_import(
            db, mode='commit', raw=raw, token=preview['preview_token'], filename='people.xlsx'
        )
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        self.assertEqual(response.get_json()['source_format'], 'xlsx')
        self.assertEqual([row['name'] for row in db.roster], ['王小明', '陳美玲'])

    def test_xlsx_rejects_formulas_numeric_phone_and_multiple_roster_sheets(self):
        invalid_files = [
            (make_xlsx(rows=[['=1+1', '0912345678', '', '公司', '', '', '', '', '']]), '含公式'),
            (make_xlsx(rows=[['王小明', 912345678, '', '公司', '', '', '', '', '']]), '數字格式'),
            (make_xlsx(rows=[['王小明', date(2026, 8, 12), '', '公司', '', '', '', '', '']]), '不是文字格式'),
            (make_xlsx(extra_sheet='第二份名單'), '多張可匯入'),
        ]
        for raw, message in invalid_files:
            with self.subTest(message=message):
                with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                        data=self.upload_data(raw, filename='people.xlsx'),
                        content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, 400)
                self.assertIn(message, response.get_json()['message'])

    def test_xlsx_rejects_bad_zip_zip_bomb_and_unsupported_extensions(self):
        bomb = io.BytesIO()
        with zipfile.ZipFile(bomb, 'w', zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('[Content_Types].xml', b'x')
            archive.writestr('xl/workbook.xml', b'x' * 10000)
        macro = io.BytesIO(make_xlsx())
        with zipfile.ZipFile(macro, 'a', zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('xl/VBAProject.bin', b'not-a-real-macro')
        cases = [
            (b'not a zip', 'people.xlsx', '不是有效'),
            (bomb.getvalue(), 'people.xlsx', '壓縮比例異常'),
            (macro.getvalue(), 'people.xlsx', '含巨集'),
            (VALID_CSV, 'people.xls', '只支援 .csv 或 .xlsx'),
            (VALID_CSV, 'people.txt', '只支援 .csv 或 .xlsx'),
        ]
        for raw, filename, message in cases:
            with self.subTest(filename=filename, message=message):
                with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                        data=self.upload_data(raw, filename=filename),
                        content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, 400)
                self.assertIn(message, response.get_json()['message'])

    def test_xlsx_parser_limits_rows_columns_and_cell_length(self):
        too_wide_headers = ['姓名'] + [f'欄{i}' for i in range(server.MAX_XLSX_COLUMNS)]
        cases = [
            (make_xlsx(headers=too_wide_headers, rows=[['王小明']]), '欄位超過'),
            (make_xlsx(rows=[['王小明', '0912345678', '', '公司', '', '', '', '', 'x' * 11]]), '文字超過'),
        ]
        for raw, message in cases:
            with self.subTest(message=message):
                with (
                    patch.object(server, 'MAX_XLSX_CELL_CHARACTERS', 10),
                    patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')),
                ):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                        data=self.upload_data(raw, filename='people.xlsx'),
                        content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, 400)
                self.assertIn(message, response.get_json()['message'])

        row_limited = make_xlsx(rows=[
            ['甲', '0911111111', '', '公司', '', '', '', '', ''],
            ['乙', '0922222222', '', '公司', '', '', '', '', ''],
        ])
        with patch.object(server, 'MAX_CSV_ROWS', 1):
            with self.assertRaisesRegex(ValueError, '超過 1 筆上限'):
                server.parse_registration_upload(row_limited, 'people.xlsx')

    def test_malformed_or_uneven_csv_is_rejected_before_database(self):
        malformed_files = [
            '姓名,手機,公司名稱\r\n"王小明,0912345678,公司\r\n'.encode('utf-8'),
            '姓名,手機,公司名稱\r\n王小明,0912345678,公司,多餘欄位\r\n'.encode('utf-8'),
            '姓名,手機,公司名稱\r\n王小明,0912345678\r\n'.encode('utf-8'),
        ]
        for raw in malformed_files:
            with self.subTest(raw=raw):
                with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=commit&admin=admin&sheet=活動A',
                        data=self.upload_data(raw),
                        content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, 400)

    def test_csv_row_limit_is_rejected_before_database(self):
        raw = ('姓名,手機,公司名稱\r\n' + '王小明,0912345678,公司\r\n' * 3).encode('utf-8')
        with (
            patch.object(server, 'MAX_CSV_ROWS', 2),
            patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')),
        ):
            response = self.client.post(
                '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                data=self.upload_data(raw), content_type='multipart/form-data',
            )
        self.assertEqual(response.status_code, 400)
        self.assertIn('超過 2 筆上限', response.get_json()['message'])

    def test_render_requires_session_and_admin_secrets(self):
        source = inspect.getsource(server)
        self.assertIn("if IS_RENDER and not SECRET_KEY:", source)
        self.assertIn("if IS_RENDER and (not ADMIN_PASSWORD or ADMIN_PASSWORD == 'admin123'):", source)
        self.assertIn("UPDATE admins SET password=%s WHERE password=%s", source)

    def test_preview_returns_summary_and_only_stores_one_time_token(self):
        db = FakeDatabase(roster=[roster_row(i) for i in range(1, 8)])
        data = self.preview(db)
        self.assertEqual(data['total_count'], 3)
        self.assertEqual(data['valid_count'], 2)
        self.assertEqual(data['skipped_count'], 1)
        self.assertEqual(data['existing_count'], 7)
        self.assertTrue(data['preview_token'])
        self.assertEqual(len(db.previews), 1)
        all_sql = [sql for connection in db.connections for sql, _ in connection.statements]
        self.assertFalse(any(sql.startswith('DELETE FROM event_registrations') for sql in all_sql))
        self.assertFalse(any(sql.startswith('INSERT INTO event_registrations') for sql in all_sql))

    def test_preview_fails_closed_if_roster_snapshot_is_unavailable(self):
        db = FakeDatabase()
        db.fail_on_snapshot = True
        response = self.post_import(db)
        self.assertEqual(response.status_code, 500)
        self.assertEqual(db.previews, {})

    def test_import_requires_login_matching_admin_and_allowed_sheet(self):
        db = FakeDatabase()
        with self.client.session_transaction() as user_session:
            user_session.clear()
        with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
            anonymous = self.client.post(
                '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                data=self.upload_data(), content_type='multipart/form-data',
            )
        self.assertEqual(anonymous.status_code, 401)

        self.log_in()
        with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
            other_admin = self.client.post(
                '/api/sheets/import_csv?mode=preview&admin=other&sheet=活動A',
                data=self.upload_data(), content_type='multipart/form-data',
            )
        self.assertEqual(other_admin.status_code, 403)

        forbidden_sheet = self.post_import(db, sheet='活動B')
        self.assertEqual(forbidden_sheet.status_code, 403)
        self.assertFalse(any('DELETE FROM event_registrations' in sql for sql, _ in db.connections[-1].statements))

    def test_sheet_admin_routes_use_authenticated_identity(self):
        db = FakeDatabase(roster=[roster_row(1)])
        with (
            patch.object(server, 'get_db_connection', side_effect=db.connect),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'ensure_config'),
        ):
            listed = self.client.get('/api/sheets/list?admin=other')
            switched = self.client.post('/api/session/sheet', json={'admin': 'other', 'sheet': '活動B'})
        self.assertEqual(listed.status_code, 200)
        self.assertEqual(listed.get_json()['username'], 'admin')
        self.assertEqual(switched.status_code, 200)
        update_sql = [
            args for connection in db.connections for sql, args in connection.statements
            if sql.startswith('UPDATE admins SET current_event')
        ]
        self.assertEqual(update_sql[-1][2], 'admin')

        with self.client.session_transaction() as user_session:
            user_session.clear()
        with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
            anonymous_list = self.client.get('/api/sheets/list')
            anonymous_switch = self.client.post('/api/session/sheet', json={'sheet': '活動A'})
        self.assertEqual(anonymous_list.status_code, 401)
        self.assertEqual(anonymous_switch.status_code, 401)

    def test_all_roster_mutators_take_the_same_event_lock(self):
        for function in [
            server.api_registration_add,
            server.api_checkin,
            server.import_csv_api,
            server.delete_sheet_data_api,
        ]:
            with self.subTest(function=function.__name__):
                body = inspect.getsource(function)
                self.assertIn('lock_event_mutations(cur, admin, sheet)', body)

        checkin_body = inspect.getsource(server.api_checkin)
        self.assertNotIn('WHERE id=%s LIMIT 1', checkin_body)
        self.assertIn('FOR UPDATE', checkin_body)
        self.assertIn('cur.rowcount != 1', checkin_body)

    def test_invalid_headers_and_zero_valid_rows_never_connect_to_database(self):
        invalid_files = [
            b'foo,bar\r\n1,2\r\n',
            '姓名,手機,公司名稱\r\n,,\r\n'.encode('utf-8'),
        ]
        for raw in invalid_files:
            with self.subTest(raw=raw):
                with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=commit&admin=admin&sheet=活動A',
                        data=self.upload_data(raw), content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, 400)

    def test_commit_replaces_roster_and_token_cannot_be_replayed(self):
        db = FakeDatabase(roster=[roster_row(1), roster_row(2)])
        preview = self.preview(db)
        response = self.post_import(db, mode='commit', token=preview['preview_token'])
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        self.assertEqual(response.get_json()['inserted_count'], 2)
        self.assertEqual([row['name'] for row in db.roster], ['王小明', '陳美玲'])

        replay = self.post_import(db, mode='commit', token=preview['preview_token'])
        self.assertEqual(replay.status_code, 409)
        self.assertEqual([row['name'] for row in db.roster], ['王小明', '陳美玲'])

    def test_changed_file_or_target_is_rejected_without_deleting_roster(self):
        db = FakeDatabase(roster=[roster_row(1)])
        preview = self.preview(db)
        original = copy.deepcopy(db.roster)
        changed = VALID_CSV.replace('王小明'.encode(), '王大明'.encode())
        response = self.post_import(db, mode='commit', raw=changed, token=preview['preview_token'])
        self.assertEqual(response.status_code, 409)
        self.assertEqual(db.roster, original)

    def test_roster_change_after_preview_requires_a_new_preview(self):
        db = FakeDatabase(roster=[roster_row(1)])
        preview = self.preview(db)
        db.roster[0]['status'] = 'checked_in'
        response = self.post_import(db, mode='commit', token=preview['preview_token'])
        self.assertEqual(response.status_code, 409)
        self.assertIn('名單在預覽後已被更新', response.get_json()['message'])
        self.assertEqual(db.roster[0]['status'], 'checked_in')

    def test_database_error_rolls_back_roster_and_preview_token(self):
        original = [roster_row(1)]
        db = FakeDatabase(roster=original)
        preview = self.preview(db)
        db.fail_on_insert = 2
        response = self.post_import(db, mode='commit', token=preview['preview_token'])
        self.assertEqual(response.status_code, 500)
        self.assertEqual(db.roster, original)
        token_hash = server.csv_preview_token_hash(preview['preview_token'])
        self.assertFalse(db.previews[token_hash]['used_at'])
        self.assertEqual(db.connections[-1].rollbacks, 1)

    def test_template_is_utf8_bom_and_parser_compatible(self):
        response = self.client.get('/api/sheets/import_csv/template')
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.data.startswith(b'\xef\xbb\xbf'))
        self.assertIn('text/csv', response.content_type)
        self.assertIn('attachment', response.headers['Content-Disposition'])
        header = response.data + '王小明,0912345678,,範例公司,,,,資訊科技,素食,12,\r\n'.encode('utf-8')
        self.assertEqual(server.parse_registration_csv(header)['valid_count'], 1)

    def test_utf16_and_oversized_files_are_rejected_before_database(self):
        invalid_files = [
            ('姓名,手機\r\n王小明,0912345678\r\n'.encode('utf-16'), 400),
            (b'a' * (server.MAX_CSV_BYTES + 1), 413),
        ]
        for raw, expected_status in invalid_files:
            with self.subTest(expected_status=expected_status):
                with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
                    response = self.client.post(
                        '/api/sheets/import_csv?mode=preview&admin=admin&sheet=活動A',
                        data=self.upload_data(raw), content_type='multipart/form-data',
                    )
                self.assertEqual(response.status_code, expected_status)

    def test_frontend_shows_imported_meal_and_table_without_meal_selection(self):
        frontend = Path(server.app.root_path, '活動報到系統.html').read_text(encoding='utf-8')
        self.assertIn('async function loadSystemConfig()', frontend)
        self.assertIn('await loadSystemConfig();', frontend)
        render_start = frontend.index('async function renderMealOptions()')
        render_end = frontend.index('function toggleIdentity', render_start)
        confirmation_source = frontend[render_start:render_end]
        self.assertNotIn("selectMeal(this", confirmation_source)
        self.assertNotIn("submitCheckin('不需要')", confirmation_source)
        self.assertIn("selectedUser.meal_preference", confirmation_source)
        self.assertIn("selectedUser.seat", confirmation_source)
        self.assertIn('id="proxySeat"', confirmation_source)
        self.assertIn('由工作人員填寫', confirmation_source)
        self.assertIn("const isOriginal = window.isOriginal !== false;", frontend)
        self.assertIn("const proxyNameEl = document.getElementById('proxyName');", frontend)
        self.assertIn("const proxySeatEl = document.getElementById('proxySeat');", frontend)
        self.assertIn('✓ 工作人員確認', frontend)
        self.assertIn('onsubmit="event.preventDefault(); handleAction();"', frontend)
        self.assertIn('<button type="submit" class="ck-btn"', frontend)

    def test_proxy_public_user_keeps_original_and_exposes_attendance_name(self):
        user = server.public_user({
            'name': '原主管', 'proxy_name': '替代主管', 'proxy_phone': '0900111222',
            'status': '替代', 'is_original': 0, 'seat': '第3桌', 'meal_preference': '素食',
        })
        self.assertEqual(user['name'], '原主管')
        self.assertEqual(user['original_name'], '原主管')
        self.assertEqual(user['display_name'], '替代主管')
        self.assertEqual(user['attendance_name'], '替代主管')
        self.assertTrue(user['is_proxy'])

    def test_checked_in_companies_only_include_checked_status_and_merge_duplicates(self):
        rows = [
            {'name': '甲', 'company': '範例公司', 'status': 'checked_in', 'industry_category': '科技'},
            {'name': '乙', 'company': '範例公司', 'status': '替代', 'proxy_name': '乙代理', 'industry_category': '科技'},
            {'name': '丙', 'company': '尚未公司', 'status': 'pending', 'industry_category': '服務'},
            {'name': '丁', 'company': '', 'status': 'checked_in'},
        ]
        companies = server.aggregate_checked_in_companies(rows)
        self.assertEqual(len(companies), 1)
        self.assertEqual(companies[0]['name'], '範例公司')
        self.assertEqual(companies[0]['checked_in_count'], 2)
        self.assertEqual(companies[0]['attendees'], ['甲', '乙代理'])

    def test_frontend_keeps_checkin_method_but_uses_boarding_terms_elsewhere(self):
        source = Path('活動報到系統.html').read_text(encoding='utf-8')
        self.assertIn('請 選 擇 登 機 方 式', source)
        remaining = source.replace('請 選 擇 登 機 方 式', '')
        self.assertNotRegex(remaining, r'登\s*機')
        self.assertIn('onclick="returnFromProducts()"', source)
        self.assertIn("showProducts('success')", source)

    def test_projection_header_fits_and_donut_stays_at_one_hundred_percent(self):
        source = Path('dashboard.html').read_text(encoding='utf-8')
        self.assertIn('--header-height:128px', source)
        self.assertIn('min-height:var(--header-height)', source)
        self.assertIn('flex-wrap:wrap', source)
        self.assertIn('<div class="donut-num" id="donut-num">', source)
        self.assertIn('100%', source)
        self.assertNotIn('id="donut-lbl"', source)
        self.assertNotIn('已報到／總名單', source)
        self.assertNotIn("setText('donut-num'", source)
        self.assertIn('font-size:1.7rem', source)
        self.assertIn('letter-spacing:-.06em', source)
        self.assertIn('white-space:nowrap', source)
        self.assertIn('max-width:62%', source)
        self.assertNotIn('font-size:2.35rem', source)
        self.assertIn('s.industry_stats?.checked_in', source)
        self.assertIn('buildDonut(src)', source)
        self.assertIn('renderAgenda()', source)
        self.assertIn('renderEnterprises()', source)

    def test_config_defaults_to_showing_meal_options(self):
        config = server.serialize_config(
            {'admin_username': 'admin', 'google_sheet_name': '活動A'},
            admin='admin', sheet='活動A',
        )
        self.assertIs(config['show_meal_options'], True)

        hidden = server.serialize_config(
            {'admin_username': 'admin', 'google_sheet_name': '活動A', 'show_meal_options': False},
            admin='admin', sheet='活動A',
        )
        self.assertIs(hidden['show_meal_options'], False)


if __name__ == '__main__':
    unittest.main()
