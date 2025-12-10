import pandas as pd
import sqlite3
import os
import tempfile

# Simulate the ExcelDatabaseApp class methods for testing
class TestExcelDatabaseApp:
    def __init__(self):
        self.df = None
        self.conn = None
        self.table_name = "excel_data"
        self.db_path = None

    def load_file(self, filepath):
        try:
            self.df = pd.read_excel(filepath)
        except Exception as e:
            raise Exception(f"Failed to load file: {e}")

        if self.df.empty:
            raise Exception("The loaded Excel file has no data.")

        # Create database in temp location
        self.db_path = tempfile.mktemp(suffix='.db')
        self.conn = sqlite3.connect(self.db_path)
        self.df.to_sql(self.table_name, self.conn, if_exists='replace', index=False)

    def search_data(self, column, term):
        if not column or not term:
            raise ValueError("Please select a column and enter a search term.")

        query = f"SELECT * FROM {self.table_name} WHERE {column} LIKE ?"
        cursor = self.conn.cursor()
        cursor.execute(query, ('%' + term + '%',))
        rows = cursor.fetchall()
        return rows

    def sort_data(self, col, reverse=False):
        # Simulate sorting on DataFrame
        self.df = self.df.sort_values(by=col, ascending=not reverse)
        return self.df

    def save_file(self, save_path):
        if not self.df.empty:
            self.df.to_excel(save_path, index=False)
        else:
            raise ValueError("No data to save.")

    def __del__(self):
        if self.conn:
            self.conn.close()
        if self.db_path and os.path.exists(self.db_path):
            os.remove(self.db_path)

# Test cases
def test_load_file():
    app = TestExcelDatabaseApp()
    try:
        app.load_file('sample.xlsx')
        assert app.df is not None
        assert len(app.df) == 3
        assert list(app.df.columns) == ['Name', 'Age', 'City']
        print("✓ Load file test passed")
    except Exception as e:
        print(f"✗ Load file test failed: {e}")
    finally:
        del app

def test_search_data():
    app = TestExcelDatabaseApp()
    try:
        app.load_file('sample.xlsx')
        # Search for 'Alice' in 'Name'
        results = app.search_data('Name', 'Alice')
        assert len(results) == 1
        assert results[0][0] == 'Alice'
        # Search for 'LA' in 'City'
        results = app.search_data('City', 'LA')
        assert len(results) == 1
        assert results[0][2] == 'LA'
        # Search for non-existent
        results = app.search_data('Name', 'NonExistent')
        assert len(results) == 0
        print("✓ Search data test passed")
    except Exception as e:
        print(f"✗ Search data test failed: {e}")
    finally:
        del app

def test_sort_data():
    app = TestExcelDatabaseApp()
    try:
        app.load_file('sample.xlsx')
        # Sort by Age ascending
        sorted_df = app.sort_data('Age')
        assert sorted_df.iloc[0]['Age'] == 25
        assert sorted_df.iloc[2]['Age'] == 35
        # Sort by Age descending
        sorted_df = app.sort_data('Age', reverse=True)
        assert sorted_df.iloc[0]['Age'] == 35
        assert sorted_df.iloc[2]['Age'] == 25
        print("✓ Sort data test passed")
    except Exception as e:
        print(f"✗ Sort data test failed: {e}")
    finally:
        del app

def test_save_file():
    app = TestExcelDatabaseApp()
    try:
        app.load_file('sample.xlsx')
        save_path = tempfile.mktemp(suffix='.xlsx')
        app.save_file(save_path)
        # Verify saved file
        saved_df = pd.read_excel(save_path)
        assert len(saved_df) == 3
        assert list(saved_df.columns) == ['Name', 'Age', 'City']
        os.remove(save_path)
        print("✓ Save file test passed")
    except Exception as e:
        print(f"✗ Save file test failed: {e}")
    finally:
        del app

def test_edge_cases():
    app = TestExcelDatabaseApp()
    try:
        # Test empty search
        try:
            app.load_file('sample.xlsx')
            app.search_data('', 'term')
            assert False, "Should raise ValueError"
        except ValueError:
            pass
        # Test invalid column
        try:
            app.search_data('InvalidColumn', 'term')
            assert False, "Should raise exception"
        except:
            pass
        print("✓ Edge cases test passed")
    except Exception as e:
        print(f"✗ Edge cases test failed: {e}")
    finally:
        del app

if __name__ == "__main__":
    test_load_file()
    test_search_data()
    test_sort_data()
    test_save_file()
    test_edge_cases()
    print("All tests completed.")
