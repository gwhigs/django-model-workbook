# django-model-workbook
A wrapper for XlsxWriter to help express Django models as formatted tables in Excel

## Requirements
- XlsxWriter
- Django

## Usage
``` Python
from workbooks import ModelWorkbook

class MyWorkbook(ModelWorkbook):
    # Define a queryset. Can also be passed as kwarg (queryset=...) during instantiation
    queryset = MyModel.objects.all()
    
    # Or just set `model = ...` (same as in Django CBV)
    # model = MyModel
    
    # Define queryset fields and verbose headers for the table,
    #  along with any formatting you'd like for headers/data as dicts.
    # For details on formats see http://xlsxwriter.readthedocs.io/format.html#format-methods-and-format-properties
    table_fields = [
      {
        'header': "My First Field's Header",
        'field_lkup': 'my_numeric_model_field',  # Also accepts dotted lookups and functions
        'header_fmts': {'align': 'right'},
        'data_fmts': {'num_format': '#,##0'},
      },
      # ...
    
    # Name sheets to be created
    worksheets = ['Sheet1']
    
    def __init__(self, **kwargs):
      super(MyWorkbook, self).__init__(**kwargs)
      # Write our table to one of the sheets listed in self.worksheets
      self.write_table_to_sheet('Sheet1')
      # Any other non-model related sheet changes can go here as well,
      # you can access XlsxWriter WorkSheet objects with:
      # xlsxwriter_sheet = self.get_sheet_by_name('Sheet1')
      

# Now we can use our ModelWorkbook to serve Excel files in views (or just save them to disk)
# All file handling is done in memory by default
workbook = MyWorkbook()
output = workbook.export()
response = http.HttpResponse(
  output.read(),
  content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
response['Content-Disposition'] = 'attachment; filename=ModelWorkbook.xlsx'

```
