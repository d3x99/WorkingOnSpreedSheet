from spreadsheet_operations.discount_and_chart import ProcessWorkbook as ps


ps('transaction2.xlsx', 'Sheet1').add_discounted_column(0.9, 3, 2)
ps('transaction2.xlsx', 'Sheet1').add_bar_chart(2, 4, 'b8')

