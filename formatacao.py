from openpyxl.styles import Border, Font, PatternFill, Side, Alignment

azul = PatternFill(start_color="1F4E78", fill_type="solid")
branco = Font(color="FFFFFF", bold=True)

linha = Side(style='thin', color='000000')
borda = Border(left=linha, right=linha, top=linha, bottom=linha)

alinhar = Alignment(horizontal='center')