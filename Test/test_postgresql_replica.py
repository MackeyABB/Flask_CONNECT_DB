import unittest
from html.parser import HTMLParser
from pathlib import Path
from unittest.mock import MagicMock, patch

import db_mgt


class DatabaseOptionsParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_database_select = False
        self.values = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if tag == 'select':
            self.in_database_select = attributes.get('name') == 'DB_Select'
        elif tag == 'option' and self.in_database_select:
            self.values.append(attributes.get('value'))

    def handle_endtag(self, tag):
        if tag == 'select':
            self.in_database_select = False


class PostgreSQLReplicaTests(unittest.TestCase):
    def make_database(self):
        database = db_mgt.Database()
        database.cursor = MagicMock()
        database.cursor.description = [('PartNumber',), ('SAP_Number',)]
        database.cursor.fetchall.return_value = [('CAP_1', '2TFU123')]
        return database

    def test_connection_uses_distinct_dsns_with_identical_options(self):
        with patch.object(db_mgt.pypyodbc, 'connect') as connect:
            for dbindex, server in ((4, 'Master'), (5, 'Replica')):
                with self.subTest(dbindex=dbindex):
                    database = db_mgt.Database()
                    self.assertTrue(database.openDB(dbindex, db_mgt.DBList, MagicMock()))
                    self.assertIn(
                        f'DSN=Connect ePDMS ODBC-{server};', connect.call_args.args[0]
                    )
                    self.assertEqual(
                        connect.call_args.kwargs, {'timeout': 20, 'readonly': True}
                    )
                    self.assertIs(database.cursor, connect.return_value.cursor.return_value)

    def test_view_listing_is_identical(self):
        for dbindex in (4, 5):
            with self.subTest(dbindex=dbindex):
                database = self.make_database()
                self.assertEqual(database.listTable(dbindex), [('CAP_1', '2TFU123')])
                database.cursor.execute.assert_called_once_with(
                    "SELECT viewname FROM pg_views WHERE schemaname = 'public';"
                )

    def test_single_and_all_table_sql_and_results_are_identical(self):
        for table_name in ('01-Capacitors', '---All----'):
            for filters in ([''] * 8, ['cap', '2tf', '1u', 'mpn', 'desc', 'tech', 'editor', 'manufacturer']):
                with self.subTest(table_name=table_name, filters=filters):
                    queries = []
                    for dbindex in (4, 5):
                        database = self.make_database()
                        result = database.fetch(table_name, dbindex, *filters)
                        self.assertEqual(
                            result, ([('CAP_1', '2TFU123')], ['PartNumber', 'SAP_Number'])
                        )
                        queries.append(database.cursor.execute.call_args.args[0])
                    self.assertEqual(queries[0], queries[1])
                    self.assertIn('"01-Capacitors"', queries[1])

    def test_connection_failure_returns_false_for_both_servers(self):
        with patch.object(db_mgt.pypyodbc, 'connect', side_effect=RuntimeError('test failure')):
            for dbindex in (4, 5):
                with self.subTest(dbindex=dbindex):
                    app = MagicMock()
                    self.assertFalse(db_mgt.Database().openDB(dbindex, db_mgt.DBList, app))
                    app.logger.error.assert_called_once()

    def test_avl_selector_offers_both_servers(self):
        template = Path(__file__).resolve().parents[1] / 'templates' / 'AVLHandle.html'
        parser = DatabaseOptionsParser()
        parser.feed(template.read_text(encoding='utf-8'))
        self.assertIn('4', parser.values)
        self.assertIn('5', parser.values)


if __name__ == '__main__':
    unittest.main()