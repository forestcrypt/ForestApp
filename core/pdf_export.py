import os
import sqlite3
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT


def export_table_to_pdf(screen, filename=None):
    if not filename:
        filename = f'reports/Перечетная_ведомость_{screen.current_section}.pdf'
    os.makedirs('reports', exist_ok=True)

    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle('Title2', parent=styles['Title'],
                                  fontSize=16, spaceAfter=12, alignment=TA_CENTER,
                                  textColor=colors.HexColor('#2E7D32'))
    heading_style = ParagraphStyle('Heading2', parent=styles['Heading2'],
                                    fontSize=12, spaceAfter=8, spaceBefore=12,
                                    textColor=colors.HexColor('#1565C0'))
    normal_style = ParagraphStyle('Normal2', parent=styles['Normal'],
                                   fontSize=9, spaceAfter=4)

    elements.append(Paragraph(f'Перечётная ведомость — Участок {screen.current_section}', title_style))
    elements.append(Spacer(1, 6*mm))

    all_data = []
    for page in sorted(screen.page_data.keys()):
        all_data.extend(screen.page_data[page])

    if all_data:
        col_names = screen.column_names
        data_table = [col_names]
        for row in all_data:
            data_table.append([str(c) if c else '' for c in row])

        col_widths = [min(120/len(col_names)*cm, 3*cm)] * len(col_names)
        t = Table(data_table, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E7D32')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 6*mm))

    totals = screen.calculate_totals()
    if totals and totals.get('total_trees', 0) > 0:
        elements.append(PageBreak())
        elements.append(Paragraph('Итоги по перечётной ведомости', heading_style))

        summary_data = [
            ['Показатель', 'Значение'],
            ['Всего деревьев', str(totals['total_trees'])],
            ['Средний диаметр', f'{totals["avg_diameter"]:.1f} см'],
            ['Средняя высота', f'{totals["avg_height"]:.1f} м'],
        ]
        summary_table = Table(summary_data, colWidths=[6*cm, 6*cm])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 6*mm))

        species_summary = totals.get('species_summary', {})
        if species_summary:
            elements.append(Paragraph('Распределение по породам', heading_style))
            species_data = [['Порода', 'Кол-во', 'D ср, см', 'H ср, м']]
            for species, data in sorted(species_summary.items()):
                d = data.get('diameters', [])
                h = data.get('heights', [])
                avg_d = f'{sum(d)/len(d):.1f}' if d else '-'
                avg_h = f'{sum(h)/len(h):.1f}' if h else '-'
                species_data.append([species, str(data['count']), avg_d, avg_h])
            species_table = Table(species_data, colWidths=[4*cm, 3*cm, 3*cm, 3*cm])
            species_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E7D32')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            elements.append(species_table)

    try:
        doc.build(elements)
        return filename
    except Exception as e:
        raise e
