import pytz
import xlsxwriter
from datetime import datetime

WORKSHEET_FOOTER = 'Confidential'


class Excel(object):

    def __init__(self, response, workbook_name, data_sets, local_tz, timezone_support=True, **kwargs):
        """
        :param response: can be an HttpResponse or a file name
        :param workbook_name: name for the workbook file
        :param data_sets: data for all pages, a set per page.  A set is a list of dictionaries
            Example:
                data_sets = [
                    {
                        'sheet_name': 'Sheet1',
                        'data': [{}],
                        'column_order': ['Col1', 'Col2', 'Col3'],
                        'column_formats': {'Col1': {'border': 1}},
                        'column_label_overrides': {'Col1': 'ID', 'Col2': 'Item Category', 'Col3': 'Price'},
                    },
                ]
        :param timezone_support: Are datetimes in the source data timezone aware?  If so, answer True
        :param local_tz: the timezone you want items displayed
        """

        self.right_now = None
        self.workbook = None
        self.worksheet = None
        self.file_name = ''
        self.extension = 'xlsx'
        self.output = None
        self.total_sheets = 0
        self.next_sheet_index = 0
        self.sheet_name = ''
        self.column_formats = {}
        self.column_order = []
        self.column_label_overrides = {}
        self.column_header_format = None
        self.footer_format = None
        self.date_time_format = None
        self.body_format = None
        self.currency_format = None
        self.data_set = {}
        self.data = []
        self.max_cols = 0
        self.col_index = 0
        self.row_index = 0

        if timezone_support:
            self.local_tz = pytz.timezone(local_tz)
            self.right_now = datetime.now(tz=self.local_tz)
        else:
            self.right_now = datetime.now(tz=None)
        self.response = response
        self.workbook_name = workbook_name
        self.data_sets = data_sets
        self.total_sheets = len(data_sets)

    def make(self):
        self.create_workbook()
        while self.next_sheet_index < self.total_sheets:
            self.next_sheet()
        self.workbook.close()

    def create_workbook(self):
        self.workbook = xlsxwriter.Workbook(self.response, {'in_memory': True})
        self.workbook.set_properties({
            'title': self.workbook_name,
            'company': '',
            'comments': 'Created with Python and XlsxWriter'})
        self.create_formats()

    def next_sheet(self):
        self.data_set = self.data_sets[self.next_sheet_index]
        self.data = self.data_set.get('data')
        self.sheet_name = self.data_set.get('sheet_name', 'Sheet{}'.format(self.next_sheet_index+1))[:30]
        self.worksheet = self.workbook.add_worksheet(self.sheet_name)
        self.column_order = self.data_set['column_order']
        self.column_formats = self.data_set.get('column_formats') or {}
        self.column_label_overrides = self.data_set.get('column_label_overrides') or {}
        self.max_cols = len(self.column_order) - 1
        self.col_index = 0
        self.row_index = 0
        self.write_sheet()
        self.add_scatter_chart()
        self.next_sheet_index += 1

    def create_formats(self):

        self.column_header_format = self.workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': 'black',
            'border': 1,
        })

        self.footer_format = self.workbook.add_format({
            'align': 'right',
            'num_format': 'dd-mmm-yyyy hh:mm',
            'italic': True
        })

        self.date_time_format = self.workbook.add_format({
            'num_format': 'yyyy-mm-dd hh:mm AM/PM',
            'border': 1,
            'font_size': 12,
        })

        self.body_format = self.workbook.add_format({
            'font_size': 12,
            'border': 1,
            'num_format': '@'
        })

        self.currency_format = self.workbook.add_format({
            'num_format': '#,###',
            'font_size': 12,
            'border': 1,
            'align': 'right',
        })

    def write_sheet(self):
        default_col_width = 20
        try:
            self.worksheet.hide_gridlines(2)
            self.worksheet.set_margins(top=0.5, bottom=0.75, left=0.5, right=0.5)
            self.worksheet.set_footer(WORKSHEET_FOOTER, {'align_with_margins': True})
            self.worksheet.fit_to_pages(width=1, height=0)
            self.worksheet.repeat_rows(1, 1)
            self.worksheet.set_column(0, self.max_cols, default_col_width)

            if self.column_label_overrides:
                column_labels = []
                for label in self.column_order:
                    label = self.column_label_overrides.get(label, label)
                    column_labels.append(label)
            else:
                column_labels = [col.replace('_', ' ').title() for col in self.column_order]

            self.worksheet.write_row(row=self.row_index, col=self.col_index, data=column_labels,
                                     cell_format=self.column_header_format)
            self.row_index += 1

            for row in self.data:
                for col in self.column_order:
                    data = row.get(col)
                    cell_format = self.column_formats.get(col) or self.body_format
                    if 'date' in col.lower():
                        if isinstance(data, datetime):
                            data = data.astimezone(self.local_tz)
                            data = data.replace(tzinfo=None)
                        self.worksheet.write_datetime(self.row_index, self.col_index, data, self.date_time_format)
                    elif 'earnings' in col.lower():
                        self.worksheet.write_number(self.row_index, self.col_index, int(data), self.currency_format)
                    else:
                        self.worksheet.write(self.row_index, self.col_index, data, cell_format)
                        if len(data) > default_col_width:
                            self.worksheet.set_column(self.col_index, self.col_index, len(data) + 5)
                    self.col_index += 1
                self.row_index += 1
                self.col_index = 0

            self.worksheet.merge_range(first_row=self.row_index, first_col=self.col_index, last_row=self.row_index,
                                       last_col=self.max_cols,
                                       data='Prepared {}'.format(self.right_now.strftime('%d %b %Y %I:%M %p %Z')),
                                       cell_format=self.footer_format)
            self.row_index += 1

            print('{} rows output'.format(self.row_index))

        except Exception as e:
            print(e)

    def add_scatter_chart(self):
        """
        Create a scatter chart sub-type with smooth lines and markers.

        Data Specification in Excel Table:
        Col A must be the categories or X-axis values
        Col B is Series 1
        Col C is Series 2
        """
        chart_definition = {'type': 'scatter',
                            'subtype': 'smooth_with_markers'}
        chart_sheet = self.workbook.add_chartsheet()
        chart = self.workbook.add_chart(chart_definition)

        last_row = self.row_index - 1

        # Configure the first series
        chart.add_series({
            'name': "='{}'!$B$1".format(self.sheet_name),
            'categories': "='{}'!$A$2:$A${}".format(self.sheet_name, last_row),
            'values': "='{}'!$B$2:$B${}".format(self.sheet_name, last_row),
            'line': {'color': 'black'},
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill': {'color': 'black'},
                       },
            'data_labels': {'legend_key': True}
        })

        # Configure second series
        chart.add_series({
            'name': "='{}'!$C$1".format(self.sheet_name),
            'categories': "='{}'!$A$2:$A${}".format(self.sheet_name, last_row),
            'values': "='{}'!$C$2:$C${}".format(self.sheet_name, last_row),
            'line': {'color': 'green'},
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'green'},
                       'fill': {'color': 'green'},
                       },
            'data_labels': {'legend_key': True}
        })

        # Add a chart title and some axis labels.
        chart.set_title({'name': '{} Salary Trend'.format(self.workbook_name)})
        chart.set_x_axis({'name': 'Years'})
        chart.set_y_axis({'name': 'Dollars'})

        # Set an Excel chart style.
        chart.set_style(14)
        chart_sheet.set_chart(chart)
