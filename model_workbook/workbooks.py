from __future__ import absolute_import, unicode_literals

from functools import reduce

from django.core.exceptions import ImproperlyConfigured
from six import BytesIO
from xlsxwriter import Workbook

from .utils import apply_outer_border_to_range


class ModelWorkbook(object):
    """
    A helper for expressing Django models in Excel Workbooks.

    Initializes a workbook object for output by default.
    """
    fields = None
    model = None
    queryset = None
    landscape = False
    fit_width = False

    # Default methods to be called on ALL sheets. Values are passed as args.
    workbook_defaults = {'hide_gridlines': [0]}

    # Worksheet settings
    worksheets = []
    worksheet_obj_dict = {}

    # Table settings
    default_header_fmt = {'text_wrap': True, 'bold': True, 'bg_color': '#BDD7EE', 'valign': 'vcenter'}
    default_data_fmt = {'valign': 'vcenter'}
    table_striped = True
    table_stripe_fmt = {'bg_color': '#D9D9D9'}
    # fill table_fields with {
    #   'header': 'Verbose Name',
    #   'field_lkup': 'field.lookup',
    #   'header_fmts': {header_formats},
    #   'data_fmts': {data_formats}),
    # }
    table_fields = []
    table_offset = (0, 0)

    def __init__(self, queryset=None):
        if queryset is not None:
            self.queryset = queryset
        else:
            self.queryset = self.get_queryset()

        self.first_row_index, self.first_col_index = self.table_offset
        self.last_row_index = self.first_row_index + self.queryset.count()
        self.last_col_index = self.first_col_index + len(self.table_fields) - 1
        self.table_data_range = (
            self.first_row_index + 1,
            self.first_col_index,
            self.last_row_index,
            self.last_col_index - 1,
        )

        self.file_obj = BytesIO()
        self.workbook = Workbook(self.file_obj, {'in_memory': True})

        for sheet in self.worksheets:
            if isinstance(sheet, (tuple, list)):
                sheet_obj = self.workbook.add_worksheet(*sheet)
                self.worksheet_obj_dict[sheet[0]] = sheet_obj
            elif isinstance(sheet, str):
                sheet_obj = self.workbook.add_worksheet(sheet)
                self.worksheet_obj_dict[sheet] = sheet_obj

        for method_name, args in self.workbook_defaults.items():
            for sheet_obj in self.worksheet_obj_dict.values():
                setup_func = getattr(sheet_obj, method_name)
                if callable(setup_func):
                    if not args:
                        setup_func()
                        continue
                    if not isinstance(args, (tuple, list)):
                        args = list(args)
                    setup_func(*args)

        for sheet_obj in self.worksheet_obj_dict.values():
            if self.landscape:
                sheet_obj.set_landscape()
            if self.fit_width:
                sheet_obj.fit_to_pages(1, 0)

    def export(self):
        self.workbook.close()
        self.file_obj.seek(0)
        return self.file_obj

    def write_headers(self, ws):
        # Write headers
        for i, field in enumerate(self.table_fields):
            format_dict_with_default = field['header_fmts'].copy()
            format_dict_with_default.update(self.default_header_fmt)
            fmt = self.workbook.add_format(format_dict_with_default)
            ws.write(self.first_row_index, self.first_col_index + i, field['header'], fmt)

    def write_table_data(self, ws):
        queryset = self.get_queryset()
        for row_offset, obj in enumerate(queryset.iterator()):
            for col_offset, field in enumerate(self.table_fields):
                format_dict_with_default = field['data_fmts'].copy()
                format_dict_with_default.update(self.default_data_fmt)
                if self.table_striped and row_offset % 2 == 1:
                    format_dict_with_default.update(self.table_stripe_fmt)
                fmt = self.workbook.add_format(format_dict_with_default)
                val = reduce(getattr, field['field_lkup'].split('.'), obj)
                if callable(val):
                    val = val()
                ws.write(
                        self.first_row_index + 1 + row_offset,
                        self.first_col_index + col_offset,
                        val,
                        fmt,
                )

    def write_border_to_table(self, ws):
        options = {
            'first_row_index': self.first_row_index,
            'first_col_index': self.first_col_index,
            'last_row_index': self.last_row_index,
            'last_col_index': self.last_col_index,
        }
        apply_outer_border_to_range(self.workbook, ws, options=options)

    def get_queryset(self):
        """
        Copy from almost any Django model view (e.g. DetailView).

        Return the `QuerySet` that will be used to look up the object.
        """
        if self.queryset is None:
            if self.model:
                return self.model._default_manager.all()
            else:
                raise ImproperlyConfigured(
                    "%(cls)s is missing a QuerySet. Define "
                    "%(cls)s.model, %(cls)s.queryset, or override "
                    "%(cls)s.get_queryset()." % {
                        'cls': self.__class__.__name__
                    }
                )
        return self.queryset.all()
