import os
import unittest
from datetime import datetime
from unittest.mock import patch

import server


class TimeCaptureCursor:
    def __init__(self, connection):
        self.connection = connection
        self.lastrowid = 0
        self.rowcount = 0
        self._one = None
        self._all = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def execute(self, sql, args=()):
        compact = ' '.join(sql.split())
        self.rowcount = 0
        self._one = None
        self._all = []

        if compact.startswith('INSERT INTO event_registrations'):
            self.connection.insert_args = args
            self.lastrowid = 41
            self.rowcount = 1
            self.connection.row = {
                'id': self.lastrowid,
                'admin_username': args[0],
                'google_sheet_name': args[1],
                'name': args[4],
                'phone': args[5],
                'company': args[7],
                'status': args[14],
                'checked_in_at': args[16],
                'checkin_time': args[17],
                'portrait_consent_time': args[20],
            }
            return

        if compact.startswith('UPDATE event_registrations'):
            self.connection.update_args = args
            self.rowcount = 1
            self.connection.row.update({
                'status': args[0],
                'is_original': args[1],
                'proxy_name': args[2],
                'proxy_phone': args[3],
                'checked_in_at': args[4],
                'checkin_time': args[5],
                'portrait_consent_time': args[8],
                'seat': args[9],
                'seating_chart': args[10],
            })
            return

        if compact.startswith('SELECT * FROM event_registrations'):
            self._one = dict(self.connection.row) if self.connection.row else None
            self._all = [dict(self.connection.row)] if self.connection.row else []
            return

        raise AssertionError(f'Unexpected SQL in timezone test: {compact}')

    def fetchone(self):
        return self._one

    def fetchall(self):
        return self._all


class TimeCaptureConnection:
    def __init__(self, row=None):
        self.row = dict(row) if row else None
        self.insert_args = None
        self.update_args = None
        self.cursor_instance = TimeCaptureCursor(self)
        self.commits = 0
        self.rollbacks = 0

    def cursor(self):
        return self.cursor_instance

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rollbacks += 1

    def close(self):
        pass


class TimezoneTests(unittest.TestCase):
    def setUp(self):
        server.app.config.update(TESTING=True, SECRET_KEY='timezone-test-secret')
        self.client = server.app.test_client()

    def test_existing_naive_utc_value_displays_in_taipei(self):
        row = server.public_user({
            'name': '測試人員',
            'checked_in_at': datetime(2026, 8, 19, 3, 18, 48),
            'portrait_consent_time': datetime(2026, 8, 19, 3, 18, 48),
        })

        self.assertEqual(row['checked_in_at'], '2026-08-19 11:18:48')
        self.assertEqual(row['checkin_time'], '2026-08-19 11:18:48')
        self.assertEqual(row['checkedInAt'], '11:18:48')
        self.assertEqual(row['checked_in_at_iso'], '2026-08-19T11:18:48+08:00')
        self.assertEqual(row['portrait_consent_time'], '2026-08-19 11:18:48')
        self.assertEqual(row['portrait_consent_time_iso'], '2026-08-19T11:18:48+08:00')

    def test_utc_conversion_handles_taipei_midnight(self):
        converted = server.utc_db_datetime_to_taipei(datetime(2026, 8, 19, 16, 30, 0))

        self.assertEqual(converted.isoformat(timespec='seconds'), '2026-08-20T00:30:00+08:00')

    def test_taipei_aware_value_is_not_converted_twice(self):
        original = datetime(2026, 8, 20, 0, 30, 0, tzinfo=server.TAIPEI_TZ)
        converted = server.utc_db_datetime_to_taipei(original)

        self.assertEqual(converted.isoformat(timespec='seconds'), '2026-08-20T00:30:00+08:00')

    def test_non_datetime_legacy_fixture_is_left_unchanged(self):
        row = server.public_user({'checked_in_at': '2026-08-19 03:18:48'})

        self.assertEqual(row['checked_in_at'], '2026-08-19 03:18:48')

    def test_health_explicitly_reports_taipei_timezone(self):
        fixed = datetime(2026, 8, 19, 11, 18, 48, tzinfo=server.TAIPEI_TZ)
        with patch.object(server, 'taipei_now', return_value=fixed):
            response = self.client.get('/api/health')

        self.assertEqual(response.status_code, 200)
        body = response.get_json()
        self.assertEqual(body['time'], '2026-08-19 11:18:48')
        self.assertEqual(body['time_iso'], '2026-08-19T11:18:48+08:00')
        self.assertEqual(body['timezone'], 'Asia/Taipei')

    def test_mysql_session_timestamps_are_forced_to_utc(self):
        with patch.dict(os.environ, {
            'DATABASE_URL': 'mysql://app:secret@mysql.internal:3306/registration',
        }, clear=True):
            params = server.db_params()

        self.assertEqual(params['init_command'], "SET time_zone = '+00:00'")

    def test_on_site_registration_stores_utc_and_returns_taipei(self):
        fixed_utc = datetime(2026, 8, 19, 3, 20, 30)
        connection = TimeCaptureConnection()
        with (
            patch.object(server, 'utc_now_naive', return_value=fixed_utc),
            patch.object(server, 'get_db_connection', return_value=connection),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'ensure_config'),
            patch.object(server, 'lock_event_mutations'),
        ):
            response = self.client.post('/api/registrations/add', json={
                'admin': 'admin',
                'sheet': '活動A',
                'name': '現場報名',
                'phone': '0912345678',
            })

        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        self.assertEqual(connection.insert_args[16], fixed_utc)
        self.assertEqual(connection.insert_args[17], fixed_utc)
        self.assertEqual(connection.insert_args[20], fixed_utc)
        self.assertEqual(response.get_json()['data']['checkedInAt'], '11:20:30')

    def test_roster_checkin_stores_utc_and_returns_taipei(self):
        fixed_utc = datetime(2026, 8, 19, 16, 30, 0)
        connection = TimeCaptureConnection({
            'id': 7,
            'admin_username': 'admin',
            'google_sheet_name': '活動A',
            'name': '名單人員',
            'phone': '0987654321',
            'company': '測試公司',
            'status': 'pending',
            'seat': '8',
        })
        with (
            patch.object(server, 'utc_now_naive', return_value=fixed_utc),
            patch.object(server, 'get_db_connection', return_value=connection),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'ensure_config'),
            patch.object(server, 'lock_event_mutations'),
        ):
            response = self.client.post(
                '/api/checkin/7?admin=admin&sheet=活動A',
                json={'is_original': True},
            )

        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        self.assertEqual(connection.update_args[4], fixed_utc)
        self.assertEqual(connection.update_args[5], fixed_utc)
        self.assertEqual(connection.update_args[8], fixed_utc)
        self.assertEqual(response.get_json()['data']['checked_in_at'], '2026-08-20 00:30:00')

    def test_csv_export_uses_taipei_time_in_content_and_filename(self):
        connection = TimeCaptureConnection({
            'id': 9,
            'admin_username': 'admin',
            'google_sheet_name': '活動A',
            'name': '匯出測試',
            'status': 'checked_in',
            'checked_in_at': datetime(2026, 8, 19, 3, 18, 48),
        })
        fixed_taipei = datetime(2026, 8, 19, 11, 30, 0, tzinfo=server.TAIPEI_TZ)
        with self.client.session_transaction() as user_session:
            user_session['admin_logged_in'] = True
            user_session['username'] = 'admin'
            user_session['allowed_sheets'] = ['活動A']
            user_session['current_admin_sheet'] = '活動A'
        with (
            patch.object(server, 'get_db_connection', return_value=connection),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'taipei_now', return_value=fixed_taipei),
        ):
            response = self.client.get('/api/sheets/export_csv?admin=admin&sheet=活動A')

        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        csv_text = response.data.decode('utf-8-sig')
        self.assertIn('2026-08-19 11:18:48', csv_text)
        self.assertIn('20260819_113000.csv', response.headers['Content-Disposition'])


if __name__ == '__main__':
    unittest.main()
