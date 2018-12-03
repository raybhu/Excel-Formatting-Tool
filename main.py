import xlwings as xw
import sys

# app = xw.App(visible=False)
# wb_original = xw.Book('sample.xlsx')
# wb_exported = xw.Book()
app = xw.App()
wb_original = app.books.open('sample.xlsx')
wb_exported = app.books.add()
sht1_ori = wb_original.sheets['Key word']
sht1_ex = wb_exported.sheets[0]

# row 1
sht1_ex.range('A1').value = 'Campaign Name'
sht1_ex.range('B1').value = 'Campaign Daily Budget'
sht1_ex.range('C1').value = 'Campaign Start Date'
sht1_ex.range('D1').value = 'Campaign End Date'
sht1_ex.range('E1').value = 'Campaign Targeting Type'
sht1_ex.range('F1').value = 'Ad Group Name'
sht1_ex.range('G1').value = 'Max Bid'
sht1_ex.range('H1').value = 'SKU'
sht1_ex.range('I1').value = 'Keyword'
sht1_ex.range('J1').value = 'Match Type'
sht1_ex.range('K1').value = 'Campaign Status'
sht1_ex.range('L1').value = 'Ad Group Status'
sht1_ex.range('M1').value = 'Status'
sht1_ex.range('N1').value = 'Bid+'
# row 2
sht1_ex.range('B2').value = '等待填写'
sht1_ex.range('C2').value = '等待填写'
sht1_ex.range('E2').value = 'Manual'
sht1_ex.range('K2').value = 'Enabled'




# 算出元数据一共多少行
table_range = sht1_ori.range('A1').expand('table')
rawDataRowNumber = table_range.rows.count
print('row number of original data:' + str(rawDataRowNumber))
#对每一行数据进行处理

var_A = 0
var_F = 0
var_G = 0
var_H = 0
var_I = 0
var_J = 0
var_L = 0
var_M = 0

for x in range(2, rawDataRowNumber+1):
    #取到每一行的range值
    selected_Range = sht1_ori.range('B' + str(x)).expand('right')

    # F
    fColumnAmount = (selected_Range.columns.count * 2) + 2
    f_ColumnNumber = 3 + var_F
    for fNum in range(0, fColumnAmount):
        sht1_ex.range('F' + str(f_ColumnNumber)).raw_value = sht1_ori.range('A' + str(x)).raw_value
        f_ColumnNumber = f_ColumnNumber + 1
    var_F = var_F + fColumnAmount

    #G
    g_ColumnNumber = 3 + var_G
    sht1_ex.range('G' + str(g_ColumnNumber)).raw_value = '0.1'
    var_G = var_G + fColumnAmount

    # H
    h_columnNumber = 2 + fColumnAmount + var_H
    sht1_ex.range('H' + str(h_columnNumber)).raw_value = sht1_ori.range('A' + str(x)).raw_value
    var_H = var_H + fColumnAmount
    #I
    iColumnAmount = selected_Range.columns.count * 2
    i_ColumnNumber = 4 + var_I
    for cell_range in selected_Range:
        sht1_ex.range('I' + str(i_ColumnNumber)).raw_value = cell_range.raw_value
        sht1_ex.range('I' + str(i_ColumnNumber + selected_Range.columns.count)).raw_value = str(cell_range.raw_value) + ' battery'
        i_ColumnNumber = i_ColumnNumber + 1
    var_I = var_I + iColumnAmount + 2

    #J
    j_ColumnNumber = 4 + var_J
    for y in range(0, iColumnAmount):
        sht1_ex.range('J' + str(j_ColumnNumber)).raw_value = 'Broad'
        j_ColumnNumber = j_ColumnNumber + 1
    var_J = var_J + iColumnAmount + 2

    #L
    l_columnNumber = 3 + var_L
    sht1_ex.range('L' + str(l_columnNumber)).raw_value = 'Enabled'
    var_L = var_L + fColumnAmount

    #M
    m_columnNumber = 4 + var_M
    for y in range(0, iColumnAmount+1):
        sht1_ex.range('M' + str(m_columnNumber)).raw_value = 'Enabled'
        m_columnNumber = m_columnNumber + 1
    var_M = var_M + iColumnAmount + 2

    #A
    a_columnNumber = 2 + var_A
    for y in range(0, fColumnAmount + 1):
        sht1_ex.range('A' + str(a_columnNumber)).raw_value = 'Speaker'
        a_columnNumber = a_columnNumber + 1
    var_A = var_A + fColumnAmount

sht1_ex.autofit()
if sys.platform == 'darwin':
    wb_exported.save(r'exported.xlsx')
else:
    wb_exported.save(r'.\exported.xlsx')
    wb_exported.close()
wb_original.close()
app.quit()