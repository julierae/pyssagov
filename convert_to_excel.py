"""
Usage:
    convert_to_excel.py [options]

Options:
  -h --help               Show this screen
  -v --version            Show version
  -f, --file <file-name>  Path to SSA.gov site XML file
"""
import os
import sys
import xmltodict

from docopt import docopt
from excel_helper import Excel


class EarningsData(object):

    def __init__(self, filename):
        self.data_sets = []
        self.data = {}
        self.filename = filename
        with open(self.filename) as fd:
            self.data = xmltodict.parse(fd.read())
        file, ext = os.path.splitext(os.path.basename(self.filename))
        self.user_name = self.data.get('osss:OnlineSocialSecurityStatementData').get('osss:UserInformation').get(
            'osss:Name').replace('.', '').replace(',', '').replace('-', '')
        self.build_data_sets()
        self.excel_file = Excel(response='{}.xlsx'.format(file),
                                workbook_name='{}'.format(self.user_name), data_sets=self.data_sets,
                                timezone_support=False, local_tz=None)
        print('Writing file: {}/{}.xlsx'.format(os.getcwd(), file))

    def build_data_sets(self):
        earnings_record = self.data.get('osss:OnlineSocialSecurityStatementData').get('osss:EarningsRecord').get(
            'osss:Earnings')
        column_order = ['Year', 'Fica Earnings', 'Medicare Earnings', ]
        data = []
        for record in earnings_record:
            row = {'Year': record.get('@endYear'),
                   'Fica Earnings': record.get('osss:FicaEarnings'),
                   'Medicare Earnings': record.get('osss:MedicareEarnings')}
            data.append(row)

        self.data_sets.append(
            {'sheet_name': 'Earnings History'.format(self.user_name)[:30], 'column_order': column_order,
             'data': data})


if __name__ == "__main__":
    arguments = docopt(__doc__, version='pyssagov 0.0.1') or {}
    filename = arguments.get('--file')
    if not filename:
        sys.exit('Please supply a file with source data from SSA.gov XML export file')
    earnings_data = EarningsData(filename=filename)
    earnings_data.excel_file.make()
