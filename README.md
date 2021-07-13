# XlwingsTools

Xlwings is a great tool when needing to interact with Excel, however some functionalities are missing.
This script adds the following functionalities:

### Resizing table
By default xlwings doesn't allow to resize a table or a label to fit the desire dataframe.
This new method resize the table to the length of the dataframe and then paste the dataframe to the resized table
```
import xlwings as xw
sheet = xw.Book().sheets[0]
df = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
sheet.range('Table1').xsize[df]
```

### Save range as picture
```
import xlwings as xw
sheet = xw.Book().sheets[0]
sheet.range('A1:B5').save['path.png']
```
