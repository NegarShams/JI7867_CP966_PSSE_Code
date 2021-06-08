import unittest
import os
import sys
import time

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

import optimisation.file_handling as TestModule
import optimisation.constants as constants

SKIP_SLOW_TESTS = False


# ----- UNIT TESTS -----
class TestImportContingencies(unittest.TestCase):
	"""
		Functions to check that PSSE import and initialisation is possible
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Imports the contingency full workbook here as multiple tests carried out on it
		"""
		file_path = os.path.join(TESTS_DIR, 'Contingencies_full.xlsx')
		cls.contingency_data = TestModule.ImportContingencies(pth=file_path)

	def test_import_cont_empty(self):
		"""
			Test that importing an empty contingency file raises an error
		:return:
		"""
		file_path = os.path.join(TESTS_DIR, 'Contingencies_empty.xlsx')

		with self.assertRaises(ValueError):
			TestModule.ImportContingencies(pth=file_path)

	def test_import_partial(self):
		"""
			Test that importing an partially completed imports without issue
		:return:
		"""
		file_path = os.path.join(TESTS_DIR, 'Contingencies_partial.xlsx')

		contingency_data = TestModule.ImportContingencies(pth=file_path)
		self.assertTrue(len(contingency_data.circuits) == 3)
		self.assertTrue(len(contingency_data.tx2) == 2)
		self.assertTrue(contingency_data.tx3.empty)
		self.assertTrue(len(contingency_data.contingency_names) == 4)
		print(contingency_data.circuits)

	def test_import_full(self):
		"""
			Test that importing a complete contingnecy file operates correctly
		:return:
		"""
		self.assertEqual(len(self.contingency_data.circuits), 5)
		self.assertEqual(len(self.contingency_data.tx2), 3)
		self.assertEqual(len(self.contingency_data.tx3), 2)
		self.assertEqual(len(self.contingency_data.busbars), 1)
		self.assertEqual(len(self.contingency_data.contingency_names), 10)

	def test_contingency_grouping_contH(self):
		"""
			Test extraction of a single contingency works as expected
		:return:
		"""
		cont_name = 'Test H'
		circuit_data, tx2_data, tx3_data, busbars, fixed_shunts, switched_shunts = self.contingency_data.group_contingencies_by_name(cont_name)

		self.assertTrue(len(circuit_data) == 1)
		self.assertTrue(len(tx2_data) == 1)
		self.assertTrue(len(tx3_data) == 0)
		self.assertTrue(len(busbars) == 0)

	def test_contingency_grouping_contD(self):
		"""
			Test extraction of a single contingency works as expected
		:return:
		"""
		cont_name = 'Test D'
		circuit_data, tx2_data, tx3_data, busbars, fixed_shunts, switched_shunts = self.contingency_data.group_contingencies_by_name(cont_name)

		self.assertEqual(len(circuit_data), 0)
		self.assertEqual(len(tx2_data), 2)
		self.assertEqual(len(tx3_data), 0)
		self.assertEqual(len(busbars), 0)

	def test_imported_datatype(self):
		"""
			Tests importing of contingencies and confirms that data types converted correctly
		:return:
		"""
		self.assertTrue(self.contingency_data.circuits[constants.Excel.id].dtype == object)
		self.assertTrue(self.contingency_data.tx2[constants.Excel.id].dtype == object)
		self.assertTrue(self.contingency_data.tx3[constants.Excel.id].dtype == object)
		self.assertTrue(self.contingency_data.fixed_shunts[constants.Excel.id].dtype == object)
		self.assertTrue(self.contingency_data.switched_shunts[constants.Excel.id].dtype == object)

		print('{}'.format(self.contingency_data.circuits[constants.Excel.id][1]))

if __name__ == '__main__':
	unittest.main()
