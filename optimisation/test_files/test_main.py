"""
	Complete testing for Main.py
"""

import unittest
import os
import sys
import time

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')
TESTS_CONTINGENCY_SINGLE = os.path.join(TESTS_DIR, 'Contingencies_single.xlsx')
TESTS_CONTINGENCY_FULL = os.path.join(TESTS_DIR, 'Contingencies_full.xlsx')
TESTS_SAV = os.path.join(TESTS_DIR, 'test_sav_complete.sav')
TESTS_EXPORT_SINGLE = os.path.join(TESTS_DIR, 'test_export_single.xlsx')
TESTS_EXPORT_FULL = os.path.join(TESTS_DIR, 'test_export_full.xlsx')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

DELETE_LOG_FILES = True

import Main as TestModule
import optimisation

optimisation.constants.DEBUG_MODE = True

# ----- UNIT TESTS -----
class TestCompleteStudy(unittest.TestCase):
	"""
		Functions to check that PSSE import and initialisation is possible
	"""
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Imports the contingency full workbook here as multiple tests carried out on it
		"""

		cls.logger = optimisation.Logger(pth_logs=TEST_LOGS, uid='TestCompleteIntegration', debug=True)

	def test_single_study(self):
		"""
			Tests that runs if a single contingency is provided as an input
		:return:
		"""
		if os.path.exists(TESTS_EXPORT_SINGLE):
			mod_time = os.path.getmtime(TESTS_EXPORT_SINGLE)
		else:
			mod_time = 0
		results_path = TestModule.main(cont_workbook=TESTS_CONTINGENCY_SINGLE,
									   psse_sav_case=TESTS_SAV,
									   target_workbook=TESTS_EXPORT_SINGLE)

		self.assertTrue(results_path == TESTS_EXPORT_SINGLE)
		self.assertTrue(os.path.exists(results_path))
		self.assertTrue(os.path.getmtime(results_path) > mod_time)

	def test_complete_study(self):
		"""
			Tests that operational for all contingencies
		:return:
		"""
		if os.path.exists(TESTS_EXPORT_FULL):
			mod_time = os.path.getmtime(TESTS_EXPORT_FULL)
		else:
			mod_time = 0
		results_path = TestModule.main(cont_workbook=TESTS_CONTINGENCY_FULL,
									   psse_sav_case=TESTS_SAV,
									   target_workbook=TESTS_EXPORT_FULL)

		self.assertTrue(results_path == TESTS_EXPORT_FULL)
		self.assertTrue(os.path.exists(results_path))
		self.assertTrue(os.path.getmtime(results_path) > mod_time)

	@classmethod
	def tearDownClass(cls):
		# Delete log files created by logger
		if DELETE_LOG_FILES:
			paths = [cls.logger.pth_debug_log,
					 cls.logger.pth_progress_log,
					 cls.logger.pth_error_log]
			del cls.logger
			for pth in paths:
				if os.path.exists(pth):
					os.remove(pth)


if __name__ == '__main__':
	unittest.main()
