import openpyxl.styles as xl_styles

font = xl_styles.Font(name="Times New Roman",
                      size="14",
                      bold=False,
                      italic = False,
                      vertAlign = None,
                      underline = 'none',
                      strike = False,
                      color = '00000000')


alignment = xl_styles.Alignment(horizontal="center",vertical="center",wrapText=True)
