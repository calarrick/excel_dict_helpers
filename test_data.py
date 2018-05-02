import logging
import excel_dict
import random
# import unittest
import timeit

# todo, adopt this into proper unit testing format and framework


workbook_name = 'CQ_2yrData.xlsx'
sheet_name = 'Page Views'
starting_row = 4
col_n = {
    'A': 'Resource',
    'C': 'Activity Date',
    'D': 'Page Title',
    'E': 'URL',
    'F': 'Researcher',
    'G': 'Department',
    'H': 'Office',
    'I': 'Job Title',
    'J': 'Session Start',
    'K': 'Session End',
    'M': 'ProfitCentre Name',
    'N': 'Employee Type',
    'S': 'SessionId'
}

# sheets = {}

logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)
print('hello')

# page_views = excel_dict.read_sheet(workbook_name, sheet_name, start_row = 4)
# logging.debug(page_views)
# sheets['Page Views'] = page_views
#
# processed_views = excel_dict.read_sheet(workbook_name, 'ProcessedViews', start_row = 1)
# sheets['ProcessedViews'] = processed_views
#
# page_views = excel_dict.read_sheet(workbook_name, sheet_name, start_row=starting_row, col_names=col_n)
# logging.debug(page_views)
#
# page_views = excel_dict.read_sheet(workbook_name, sheet_name, start_row = 4, has_headers=False)
# logging.debug(page_views)

page_views_6 = excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True, has_headers=False)
excel_dict.rewrite_single_sheet(page_views_6, 'test_out_6')

page_views_a = excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True)
logging.debug(page_views_a)
sheets_a = {'a': page_views_a}
#
#
print('not write-only', timeit.repeat("excel_dict.rewrite_sheets(sheets_a, 'test_out_def', wo=False)",
                    globals=globals(), repeat=3, number=100))
print('write only', timeit.repeat("excel_dict.rewrite_sheets(sheets_a, 'test_out_abc', wo=True)",
                    globals=globals(), repeat=3, number=100))
#
#
# case_abc = excel_dict.read_sheet('test_out_abc.xlsx', 'a')
# case_def = excel_dict.read_sheet('test_out_def.xlsx', 'a')
# print('write method results equal ', str(case_abc == case_def))

col_o = {
    'A':'Resource',
    'B':'Simple Date',
    'C':'Activity Date',
    'D':'Page Title',
    'E':'URL',
    'F':'Researcher',
    'G':'Department',
    'H':'Office',
    'I':'Job Title',
    'J':'Session Start',
    'K':'Session End',
    'L':'Country',
    'M':'ProfitCentre Name',
    'N':'Employee Type',
    'O':'Location',
    'P':'Start',
    'Q':'End',
    'R':'ResourceId',
    'S':'SessionId'
}
print('prep p_views for equality test')
page_views_3 = excel_dict.read_sheet(workbook_name, sheet_name, start_row=starting_row, col_names=col_o)
logging.debug(page_views_3)
page_views_4 = excel_dict.read_sheet(workbook_name, sheet_name, start_row=starting_row)
# page_views_6 = excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True, has_headers=False)
# logging.debug(page_views_6)

# for p in page_views:
#     p['Test1'] = random.random()
#     p['Test2'] = 'spam'
#
# page_views[0]['Test1'] = 100

# excel_dict.rewrite_page_views(page_views, 'test_out.xlsx')
# excel_dict.rewrite_page_views(page_views, 'test2_out')
# excel_dict.rewrite_page_views(page_views, 'test3_out', '.bad')
#
#
# page_views = excel_dict.read_sheet('Book1.xlsx','owssvr')
# logging.debug(page_views)
# excel_dict.rewrite_page_views(page_views, 'Book1out')
#
page_views = excel_dict.read_sheet('Book1.xlsx','owssvr', data_only=False)
# logging.debug(page_views)
excel_dict.rewrite_single_sheet(page_views, 'Book1out2')
#
# excel_dict.rewrite_sheets(sheets,'test_out_4')

page_views_5 = excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True)
logging.debug(page_views_5)

page_views_1 = excel_dict.read_sheet(workbook_name, sheet_name, start_row = 4)
logging.debug(page_views_1)


sheets2 = {
    1: page_views_1,
    2: page_views_5
}
excel_dict.rewrite_single_sheet(page_views_5, 'test_out_5')
excel_dict.rewrite_sheets(sheets2, 'test_out_1')
sheets3 = {
    '1': page_views_3,
    '2': page_views_4
}
sheets4 = {
    '1': page_views_5
}

excel_dict.rewrite_sheets(sheets3, 'test_out_123456')
logging.info('versions match ' + str(page_views_1 == page_views_5))
logging.info('versions match ' + str(page_views_3 == page_views_4))


# print('read_sheet, read-only, headers',
#       timeit.repeat("excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True)",
#                     repeat=3, number=100, globals=globals()))
#
# print('read_sheet, read-only, no headers',
#       timeit.repeat("excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=True, has_headers=False)",
#                     repeat=3, number=100, globals=globals()))
#
# print('read sheet, not read-only',
#       timeit.repeat('excel_dict.read_sheet(workbook_name, sheet_name, start_row=4, ro=False)',
#                     repeat=3, number=100, globals=globals()))

# print('rewrite multi sheet (w one item dict)',
#       timeit.repeat("excel_dict.rewrite_multi_sheet(sheets4, 'test_out_1')",
#                     number=12, repeat=3, globals=globals()))
#
# print('rewrite_page_views (1 sheet, same list as above)',
#       timeit.repeat("excel_dict.rewrite_page_views(page_views_5, 'test_out_5')",
#                     repeat=3, number=12, globals=globals()))



# print('.read_sheet with column names param (but listed all cols): ')
# print(timeit.repeat('excel_dict.read_sheet(workbook_name, sheet_name, start_row=starting_row, col_names=col_o,'
#                     'ro=True)', globals=globals(), number=100, repeat=3))
# print('.read_sheet without column names param (all cols auto):')
# print(timeit.repeat('excel_dict.read_sheet(workbook_name, sheet_name, start_row=starting_row, ro=True)',
#                     globals=globals(), number=100, repeat=3))
#
# for view in page_views:
#     if view['URL'].startswith('https://'):
#         url = view['URL'][8:]
#         logging.debug(url)
#     elif view['URL'].startswith('http://'):
#         url = view['URL'][7:]
#         logging.debug(url)
#     else:
#         url = view['URL']
#
#     url_parts = url.strip('/').split('/')
#     view['Content Set'] = eval_url_parts(url_parts)
#     (view['utm_medium'], view['utm_source']) = eval_query_string(url_parts)
# rewrite_page_views(page_views)
