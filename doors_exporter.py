import sys
import os
import argparse
import subprocess
import json
import copy

class DoorsExporter:
	def __init__(self, config):
		self.config = copy.deepcopy(config)
		self.python_exe_path = sys.executable.replace('\\', '/')
		current_dir_path = os.path.dirname(__file__).replace('\\', '/')
		self.save_documents_script_path = current_dir_path + '/save_documents.py'
	##

	def run(self):
		doors_cmd = self.build_doors_cmd()
		subprocess.call(doors_cmd, shell=False)
	##

	def build_dxl_string_array(self, name, str_list):
		# declare and open array
		dxl = 'const string {}[] = {{'.format(name)

		for item in str_list:
			dxl = dxl + '\\"{}\\",'.format(item)

		# remove last comma
		dxl = dxl[:-1]

		# close array
		dxl = dxl + '}; '

		return dxl
	##

	def build_doors_cmd(self):
		# copy parameters from config or use default values
		doors_exe_path = self.config['doors_path'] if 'doors_path' in self.config else 'C:/Program Files/IBM/Rational/DOORS/9.7/bin/doors.exe'
		doors_user_options = self.config['doors_options'] if 'doors_options' in self.config else ''
		export_format = self.config['format'] if 'format' in self.config else 'xlsx'
		time_limit = self.config['time_limit'] if 'time_limit' in self.config else '0'

		# prepare doors options for dxl script execution
		doors_export_options = '-f "%TEMP%" -D "'
		doors_export_options = doors_export_options + 'pragma runLim,{}; '.format(time_limit)
		doors_export_options = doors_export_options + 'const string PYTHON_EXE = \\"{}\\"; '.format(self.python_exe_path)
		doors_export_options = doors_export_options + 'const string SAVE_DOCUMENTS_PY = \\"{}\\"; '.format(self.save_documents_script_path)
		doors_export_options = doors_export_options + 'const string EXPORT_FORMAT = \\"{}\\"; '.format(export_format)


		# add dxl string arrays with export modules configuration
		doors_export_options = doors_export_options + self.build_dxl_string_array('EXPORT_MODULES', [el['path'] for el in self.config['modules']])
		doors_export_options = doors_export_options + self.build_dxl_string_array('EXPORT_MODULES_VIEWS', [(el['view'] if 'view' in el else '') for el in self.config['modules']])

		# include dxl export script
		doors_export_options = doors_export_options + '#include <doors_exporter.dxl>"'

		return '"{}" {} {}'.format(doors_exe_path, doors_user_options, doors_export_options)
	##

def parse_args():
	parser = argparse.ArgumentParser(description='Export IBM Rational DOORS modules')
	parser.add_argument( 'config',
						help='Configuration file in JSON format.')

	args = parser.parse_args()
	return args
##

def main():
	args = parse_args()

	# read config from file
	with open(args.config) as json_data:
		config = json.load(json_data)

		exporter = DoorsExporter(config)
		exporter.run()
	
	return 0
##

if __name__ == '__main__':
	sys.exit(main())