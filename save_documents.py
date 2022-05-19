import os
import sys

supported_formats = ['xlsx', 'csv']

def main():
	# prepare output filename and format
	output_filename = sys.argv[1] if len(sys.argv) >= 2 else 'doors_export'
	output_format = sys.argv[2] if len(sys.argv) >= 3 else 'xlsx'

	if output_format not in supported_formats:
		raise Exception("Unsupported output format '{}'".format(output_format))

	# prepare output directory
	output_dir = os.path.dirname(__file__)+'\\out\\'
	if not os.path.exists(output_dir):
		os.makedirs(output_dir)

	if output_format == 'xlsx' or output_format == 'csv':
		# use xlwings for communication with excel application
		import xlwings as xw
		from xlwings.constants import FileFormat

		# save all excel workbooks
		wbt_cnt = ''
		for wb in xw.books:
			cr_path = '{}{}{}.{}'.format(output_dir, output_filename, wbt_cnt, output_format)
			
			if output_format == 'xlsx':
				wb.save(cr_path)
			elif output_format == 'csv':
				wb.api.SaveAs(cr_path, FileFormat:=FileFormat.xlCSV)
			else:
				raise Exception("Unsupported format {}".format(output_format))
			
			wbt_cnt = 1 if wbt_cnt == '' else wbt_cnt + 1

		# close all excel applications
		for app in xw.apps:
			app.kill()
		
	return 0
##

if __name__ == '__main__':
	sys.exit(main())