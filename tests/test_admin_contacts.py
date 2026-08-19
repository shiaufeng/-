import copy
import unittest
from pathlib import Path
from unittest.mock import patch

import server


ROSTER = [
    {
        'id': 1,
        'name': '本人主管',
        'phone': '0911222333',
        'email': 'owner@example.com',
        'company': '本人公司',
        'company_name': '本人公司',
        'job_title': '主管',
        'industry_category': '資訊科技',
        'meal_preference': '葷食',
        'seat': '1',
        'status': 'checked_in',
        'is_original': 1,
        'proxy_name': '',
        'proxy_phone': '',
        'checked_in_at': None,
        'checkin_time': None,
        'portrait_consent_status': '同意',
        'special_notes': '',
    },
    {
        'id': 2,
        'name': '原主管',
        'phone': '0922000000',
        'email': 'original@example.com',
        'company': '代理公司',
        'company_name': '代理公司',
        'job_title': '總經理',
        'industry_category': '專業服務',
        'meal_preference': '素食',
        'seat': '2',
        'status': '替代',
        'is_original': 0,
        'proxy_name': '替代主管',
        'proxy_phone': '0988777666',
        'checked_in_at': None,
        'checkin_time': None,
        'portrait_consent_status': '未填',
        'special_notes': '',
    },
]


class DashboardCursor:
    def __init__(self, rows):
        self.rows = rows
        self._one = None
        self._all = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def execute(self, sql, args=()):
        compact = ' '.join(sql.split())
        self._one = None
        self._all = []
        if 'COUNT(*) AS total' in compact:
            checked = sum(1 for row in self.rows if server.status_checked(row.get('status')))
            self._one = {'total': len(self.rows), 'checked': checked}
            return
        if compact.startswith('SELECT id, name, phone, email'):
            self._all = copy.deepcopy(self.rows)
            return
        if compact.startswith('SELECT * FROM event_registrations'):
            self._all = copy.deepcopy(self.rows)
            return
        raise AssertionError(f'Unexpected SQL in admin contact test: {compact}')

    def fetchone(self):
        return self._one

    def fetchall(self):
        return self._all


class DashboardConnection:
    def __init__(self, rows=None):
        self.rows = copy.deepcopy(rows or ROSTER)

    def cursor(self):
        return DashboardCursor(self.rows)

    def close(self):
        pass


class AdminContactTests(unittest.TestCase):
    def setUp(self):
        server.app.config.update(TESTING=True, SECRET_KEY='admin-contact-test-secret')
        self.client = server.app.test_client()

    def log_in(self, username='admin', sheets=None):
        with self.client.session_transaction() as user_session:
            user_session['admin_logged_in'] = True
            user_session['username'] = username
            user_session['allowed_sheets'] = sheets or ['活動A']
            user_session['current_admin_sheet'] = (sheets or ['活動A'])[0]

    def get_stats(self, path):
        connection = DashboardConnection()
        with (
            patch.object(server, 'get_db_connection', return_value=connection),
            patch.object(server, 'ensure_core_tables'),
            patch.object(server, 'load_industry_mappings', return_value=[]),
        ):
            return self.client.get(path)

    def test_public_dashboard_only_returns_aggregate_data(self):
        response = self.get_stats('/api/dashboard_stats?admin=admin&sheet=活動A')

        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        stats = response.get_json()['stats']
        self.assertEqual(stats['logs'], [])
        self.assertEqual(stats['table_details'], [])
        self.assertNotIn('attendees', stats['checked_in_companies'][0])
        serialized = response.get_data(as_text=True)
        for private_value in ['0911222333', 'owner@example.com', '本人主管', '0988777666']:
            self.assertNotIn(private_value, serialized)

    def test_admin_dashboard_requires_matching_session_and_sheet(self):
        response = self.get_stats('/api/admin/dashboard_stats?admin=admin&sheet=活動A')
        self.assertEqual(response.status_code, 401)

        self.log_in(username='other', sheets=['活動A'])
        response = self.get_stats('/api/admin/dashboard_stats?admin=admin&sheet=活動A')
        self.assertEqual(response.status_code, 403)

        self.log_in(username='admin', sheets=['活動B'])
        response = self.get_stats('/api/admin/dashboard_stats?admin=admin&sheet=活動A')
        self.assertEqual(response.status_code, 403)

    def test_roster_export_requires_matching_admin_session(self):
        with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
            response = self.client.get('/api/sheets/export_csv?admin=admin&sheet=活動A')
        self.assertEqual(response.status_code, 401)

        self.log_in(username='other', sheets=['活動A'])
        with patch.object(server, 'get_db_connection', side_effect=AssertionError('DB must not be called')):
            response = self.client.get('/api/sheets/export_csv?admin=admin&sheet=活動A')
        self.assertEqual(response.status_code, 403)

    def test_admin_dashboard_returns_actual_attendee_contacts(self):
        self.log_in()
        response = self.get_stats('/api/admin/dashboard_stats?admin=admin&sheet=活動A')

        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        stats = response.get_json()['stats']
        by_name = {row['attendance_name']: row for row in stats['logs']}
        self.assertEqual(by_name['本人主管']['attendance_phone'], '0911222333')
        self.assertEqual(by_name['本人主管']['attendance_email'], 'owner@example.com')
        self.assertEqual(by_name['替代主管']['attendance_phone'], '0988777666')
        self.assertEqual(by_name['替代主管']['attendance_email'], '')
        self.assertEqual(by_name['替代主管']['original_email'], 'original@example.com')
        table_members = [
            member
            for group in stats['table_details']
            for member in group['members']
        ]
        proxy = next(row for row in table_members if row['attendance_name'] == '替代主管')
        self.assertEqual(proxy['attendance_phone'], '0988777666')

    def test_admin_ui_renders_contact_fields_and_uses_private_endpoint(self):
        source = Path('admin.html').read_text(encoding='utf-8')

        self.assertIn('<th>實際出席姓名</th><th>聯絡資料</th>', source)
        self.assertIn('function contactHtml(x)', source)
        self.assertIn("x?.attendance_phone", source)
        self.assertIn("x?.attendance_email", source)
        self.assertIn('/admin/dashboard_stats', source)
        self.assertIn('id="addEmail"', source)
        self.assertIn('<th>E-mail</th>', source)
        self.assertIn('colspan="9"', source)


if __name__ == '__main__':
    unittest.main()
