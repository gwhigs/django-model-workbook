# django-model-workbook
A wrapper for XlsxWriter to help express Django models as tables in Excel (with formatting)

## Requirements
- XlsxWriter
- Django

## Usage
``` Python
from workbooks import ModelWorkbook

class MyWorkbook(ModelWorkbook):
    # Define a queryset (can also be passed in __init__ kwargs)
    queryset = MyModel.objects.all()
    
    # Define queryset fields for table
    table_fields = [
      {
        'header': "My First Field's Header",
        'field_lkup': 'mynumericmodelfield',
        'header_fmts': {'align': 'right'},
        'data_fmts': {'num_format': '#,##0'},
      },
      # ...
    
    # Name sheets to be created
    worksheets = ['Sheet1']
    
    def __init__():
      super(MyWorkbook, self).__init__()
      # Write headers
      self.write_headers(plan_ws)
      # Write table data
      self.write_table_data(plan_ws)
      # Add a border
      self.write_border_to_table(plan_ws)
      # Any other non-model related sheet tweaks go here as well...
      

# Now we can use our ModelWorkbook to serve Excel files in views (or just save them to disk)
# All file handling is done in memory by default
workbook = MyWorkbook()
output = wb.export()
response = http.HttpResponse(
  output.read(),
  content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
response['Content-Disposition'] = 'attachment; filename=ModelWorkbook.xlsx'

```
