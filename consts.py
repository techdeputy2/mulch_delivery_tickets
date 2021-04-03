#style templates
ROW_STYLE_TALL = { 'tableStartLocation': { 'index': 2 }, 
                    'rowIndices': [0, 1], 'fields': '*', 
                    'tableRowStyle': { 
                        'minRowHeight': { 'magnitude': 55, 'unit': 'PT' } 
                    } 
                 }
ROW_STYLE_SHORT = { 'tableStartLocation': { 'index': 2 }, 
                    'rowIndices': [2], 'fields': '*', 
                    'tableRowStyle': { 
                        'minRowHeight': { 'magnitude': 25, 'unit': 'PT' } 
                    } 
                  }

COLUMN_PROPERTIES_WIDE = { 'tableStartLocation': { 'index': 2 }, 
                            'columnIndices': [0], 'fields': '*', 
                            'tableColumnProperties': { 
                                'widthType': 'FIXED_WIDTH', 
                                'width': { 'magnitude': 345, 'unit': 'PT' } 
                            } 
                          }

COLUMN_PROPERTIES_THIN = { 'tableStartLocation': { 'index': 2 }, 
                            'columnIndices': [0], 'fields': '*', 
                            'tableColumnProperties': { 
                                'widthType': 'FIXED_WIDTH', 
                                'width': { 'magnitude': 345, 'unit': 'PT' } 
                            } 
                          }  

CELL_STYLE_TOPLEFT = { 'tableCellStyle': { 'borderTop': { 'color': { 'color': { 'rgbColor': {} } }, 
                                           'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' } }, 
                                           'borderLeft': { 'color': { 'color': { 'rgbColor': {} } }, 
                                           'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                        'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                       'tableStartLocation': { 'index': 2 }, 'rowIndex': 0, 'columnIndex': 0 }, 
                                                       'rowSpan': 1, 'columnSpan': 1 } }

CELL_STYLE_TOPRIGHT = { 'tableCellStyle': { 'borderTop': { 'color': { 'color': { 'rgbColor': {} } }, 
                                            'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' } }, 
                                            'borderRight': { 'color': { 'color': { 'rgbColor': {} } }, 
                                            'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                        'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                       'tableStartLocation': { 'index': 2 }, 'rowIndex': 0, 'columnIndex': 1 }, 
                                                       'rowSpan': 1, 'columnSpan': 1 } }

CELL_STYLE_MIDLEFT = { 'tableCellStyle': { 'borderLeft': { 'color': { 'color': { 'rgbColor': {} } }, 
                                           'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                       'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                      'tableStartLocation': { 'index': 2 }, 
                                                      'rowIndex': 1, 'columnIndex': 0 }, 'rowSpan': 1, 'columnSpan': 1 } }

CELL_STYLE_MIDRIGHT = { 'tableCellStyle': { 'borderRight': { 'color': { 'color': { 'rgbColor': {} } }, 
                                            'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                        'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                       'tableStartLocation': { 'index': 2 }, 
                                                       'rowIndex': 1, 'columnIndex': 1 }, 'rowSpan': 1, 'columnSpan': 1 } }

CELL_STYLE_BOTTOMLEFT = { 'tableCellStyle': { 'borderBottom': { 'color': { 'color': { 'rgbColor': {} } }, 
                                              'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' } }, 
                                              'borderLeft': { 'color': { 'color': { 'rgbColor': {} } }, 
                                              'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                          'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                         'tableStartLocation': { 'index': 2 }, 
                                                         'rowIndex': 2, 'columnIndex': 0 }, 'rowSpan': 1, 'columnSpan': 1 } }

CELL_STYLE_BOTTOMRIGHT = { 'tableCellStyle': { 'borderBottom': { 'color': { 'color': { 'rgbColor': {} } }, 
                                               'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' } }, 
                                               'borderRight': { 'color': { 'color': { 'rgbColor': {} } }, 
                                               'dashStyle': 'SOLID', 'width': { 'magnitude': 0, 'unit': 'PT' }  } }, 
                            'fields': '*', 'tableRange': { 'tableCellLocation': { 
                                                           'tableStartLocation': { 'index': 2 }, 
                                                           'rowIndex': 2, 'columnIndex': 1 }, 'rowSpan': 1, 'columnSpan': 1 } }

TEXT_STYLE_BOTTOMRIGHT = { 'range': { 'startIndex': 17, 'endIndex': 18 }, 
                           'textStyle': { 'bold': 'true', 'weightedFontFamily': { 
                                                          'fontFamily': 'Times New Roman' }, 
                                                          'fontSize': { 'magnitude': 16, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_BOTTOMLEFT = { 'range': { 'startIndex': 15, 'endIndex': 17 }, 
                          'textStyle': { 'weightedFontFamily': { 
                                         'fontFamily': 'Times New Roman' }, 
                                         'fontSize': { 'magnitude': 14, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_BOTTOMLEFT_SMALL = { 'range': { 'startIndex': 15, 'endIndex': 17 }, 
                          'textStyle': { 'weightedFontFamily': { 
                                         'fontFamily': 'Times New Roman' }, 
                                         'fontSize': { 'magnitude': 10, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_MIDRIGHT = { 'range': { 'startIndex': 12, 'endIndex': 13 }, 
                        'textStyle': { 'weightedFontFamily': { 
                                       'fontFamily': 'Times New Roman' }, 
                                       'fontSize': { 'magnitude': 20, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_MIDLEFT = { 'range': { 'startIndex': 10, 'endIndex': 11 }, 
                       'textStyle': { 'weightedFontFamily': { 
                                      'fontFamily': 'Times New Roman' }, 
                                      'fontSize': { 'magnitude': 18, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_TOPRIGHT = { 'range': { 'startIndex': 7, 'endIndex': 8 }, 
                        'textStyle': { 'bold': 'true', 'weightedFontFamily': { 
                                                       'fontFamily': 'Times New Roman' }, 
                                                       'fontSize': { 'magnitude': 20, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_TOPRIGHT_SMALL = { 'range': { 'startIndex': 7, 'endIndex': 8 }, 
                        'textStyle': { 'bold': 'true', 'weightedFontFamily': { 
                                                       'fontFamily': 'Times New Roman' }, 
                                                       'fontSize': { 'magnitude': 16, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_TOPLEFT = { 'range': { 'startIndex': 5, 'endIndex': 6 }, 
                      'textStyle': { 'weightedFontFamily': { 
                                     'fontFamily': 'Times New Roman' }, 
                                     'fontSize': { 'magnitude': 18, 'unit': 'PT' } }, 'fields': '*' }

TEXT_STYLE_TOPLEFT_SMALL = { 'range': { 'startIndex': 5, 'endIndex': 6 }, 
                      'textStyle': { 'weightedFontFamily': { 
                                     'fontFamily': 'Times New Roman' }, 
                                     'fontSize': { 'magnitude': 18, 'unit': 'PT' } }, 'fields': '*' }

PARAGRAPH_STYLE_MID_LEFT = { 'range': { 'startIndex': 10, 'endIndex': 11 }, 
                             'paragraphStyle': { 'namedStyleType': 'NORMAL_TEXT', 'alignment': 'END' }, 'fields': '*' }

#ticket data fields
STEEP_DRIVEWAY = 'Steep'
DELIVERY_INSTR = 'Instructions'
DELIVERY_ZONE = 'Zone'
NAME = 'Name'
PHONE = 'Phone'
MULCH_TYPE = 'Mulch Type'
ORDER_NUMBER = 'Invoice'
ADDRESS = 'Address'
TOTAL_BAGS = 'Total Bags'
ZIP = 'Zip'
GATE_CODE = 'Gate Code'
